import win32com.client
import pandas as pd
import datetime as dt
import caldav
import passwords as ps
from pytz import UTC
import os
import keyring


def get_calendar(begin,end):
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    calendar = outlook.getDefaultFolder(9).Items
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')

    restriction = "[Start] >= '" + begin.strftime('%d/%m/%Y') + "' AND [END] <= '" + end.strftime('%d/%m/%Y') + "'"
    calendar = calendar.Restrict(restriction)
    return calendar

remote_url = # dav
nc_calendar_path = # remote.php
paivat = 40
alku = dt.datetime.now()
loppu = dt.datetime.now() + dt.timedelta(days = paivat)
cal = get_calendar(alku, loppu)

cal_subject = [app.subject for app in cal]
cal_start = [app.startUTC for app in cal]
cal_end = [app.endUTC for app in cal]
cal_prio = [int(app.BusyStatus) for app in cal]
cal_uid = [app.GlobalAppointmentID[41:].replace("0100000000000000001", "") + app.startUTC.strftime("%Y%m%d%H%M")  for app in cal]
df = pd.DataFrame(data = {'subject': cal_subject,
                       'start': cal_start,
                       'end': cal_end,
                       'prio': cal_prio,
                       'uid': cal_uid})
df["start"] = df["start"].dt.tz_convert(None)
df["end"] = df["end"].dt.tz_convert(None)

client = caldav.DAVClient(url=remote_url, username=ps.user, password=ps.psword, ssl_verify_cert=False)
calendar = caldav.Calendar(client=client,url=nc_calendar_path)

events_fetched = calendar.date_search(
    start=alku, end=loppu, expand=True)

fetc_subject = []
fetc_start = []
fetc_end = []
fetc_prio = []
fetc_uid = []
for info in events_fetched:
    info = info.data.split("\r")
    subject = info[9].split(":")[1].replace("\\","")
    start = dt.datetime.strptime(info[7].split(":")[1], "%Y%m%dT%H%M%SZ")
    end = dt.datetime.strptime(info[8].split(":")[1], "%Y%m%dT%H%M%SZ")
    prio = int(info[10].split(":")[1])
    uid = info[5].split(":")[1]
    fetc_subject.append(subject)
    fetc_start.append(start)
    fetc_end.append(end)
    fetc_prio.append(prio)
    fetc_uid.append(uid)
df_fetc = pd.DataFrame(data = {'subject': fetc_subject,
                       'start': fetc_start,
                       'end': fetc_end,
                       'prio': fetc_prio,
                       'uid':fetc_uid})


new_events = df.merge(df_fetc, on=["subject","start","end","prio","uid"], how='left', indicator="Ind")
new_events = new_events[new_events['Ind'] == 'left_only'].drop("Ind", 1)

df_remove = df_fetc.merge(df, on=["subject","start","end","prio","uid"], how='left', indicator="Ind")
df_remove = df_remove[df_remove['Ind'] == 'left_only'].drop("Ind", 1)
print("\n\n")

if len(df_remove) > 0:
    print("Poistetaan seuraavat tapahtumat")
    print(df_remove[["subject", "start"]])
    for index, row in df_remove.iterrows():
        events_fetched[index].delete()

if len(new_events) == 0:
    print("Ei päivitettävää")
    
else:
    print("Seuraavat tapahtumat lisätään")
    print(new_events[["subject", "start"]])
    for index, row in new_events.iterrows():
        alku = row["start"].strftime("%Y%m%dT%H%M%SZ")
        loppu = row["end"].strftime("%Y%m%dT%H%M%SZ")
        teksti = """BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:""" + row["uid"] + """
DTSTAMP:""" + alku + """
DTSTART:""" + alku + """
DTEND:""" + loppu + """
SUMMARY:""" + row["subject"] + """
PRIORITY:""" + str(row["prio"]) + """
END:VEVENT
END:VCALENDAR
"""
        caldav.Event(client, data = teksti, parent = calendar).save()
