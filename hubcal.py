'''
Script to automatically send email to brown/black belt
class instructors when there's a need for an instructor
in the near future.

For setup see:
   https://developers.google.com/calendar/quickstart/python

This script assumes google secret and credentials files
are in the same folder as the script, as well
as an email.json that has the sender email
address and password.

@author Kendall Bailey
@created 2018-06-30
@license MIT
'''

import sys, pdb
from os import path

INSTDIR = path.dirname(__file__)
#sys.path = [ path.join( INSTDIR,'lib' ) ] + sys.path

from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import six, json
import pandas as pd
from datetime import timedelta, datetime, date
import textwrap

credfile = path.join(INSTDIR,'credentials.json')
secret = path.join(INSTDIR,'client_secret.json')
emailcredfile = path.join(INSTDIR,'email.json')
emailcred = {}
if path.exists( emailcredfile ) :
    with open(emailcredfile,'r') as f :
       emailcred = json.load(f)

#----------------------------------------------------------------------
def main() :
    cal = HubCal()
    ev = cal.events()
    today = datetime.today().date()
    dow = today.weekday() # 0 is Monday, 6 is Sunday
    toSun = 6-dow
    toFri = (11-dow)%7
    thisFri = today + timedelta(days=toFri)
    thisSun = today + timedelta(days=toSun)

    sundays = [ (today + timedelta(days=(i*7 + toSun))) for i in range(6) ]
    fridays = [ (today + timedelta(days=(i*7 + toFri))) for i in range(6) ]
    #toTeach = sorted(sundays + fridays)
    toTeach = sorted(fridays)
    taken = {}

    for event in ev:
        start = event['start'].get('dateTime', event['start'].get('date'))
        d = pd.to_datetime(start).date()
        taken[d] = event['summary']

    msg = textwrap.dedent('''\
          This is an automated notification of the next few weeks of
          the brown and black belt class instructor sign-up schedule.
          Please consider signing up for an open date on the calender.
          If you are not sure how, reply to this mail and it will be
          done for you. You must be at least third dan to teach a
          brown and black belt class.  Check the sign-ups at
          http://friday.trma.us, which includes further instructions.
          ''').replace('\n',' ')
    msg = ('Friday instructors,\n' + msg 
           + '\n\n' 
           + '\n'.join( teachInfo( d, taken ) for d in toTeach ))
    msg += '\n\nPil sung.'
    six.print_(msg)
    
    
    to = []
    for i,v in enumerate(sys.argv) :
        if v == '--to' :
           to.extend( [ x.strip() for x in sys.argv[i+1].split(",") ] )

    send = ('--send' in sys.argv) and to
    if not send : return
    if dow == 0 : # Monday
        send = thisFri not in taken
    elif dow == 2 : # Wednesday
        send = (thisFri not in taken) #or (thisSun not in taken)
    #elif dow == 5 : # Saturday
    #    send = (thisSun not in taken)
    else : send = False
    if send or ('--force' in sys.argv) : 
        email( msg, sendTo = to )

#----------------------------------------------------------------------
def teachInfo( d, taken ) :
   return '%s%s : %s' % (
      d.weekday()==4 and 'Friday, ' or 'Sunday, ',
      d.isoformat()[5:],
      taken.get(d, '** needs an instructor' ))

#----------------------------------------------------------------------
def email(body, sendFrom=None, sendTo=[], subject=None, creds = emailcred) :
   import smtplib

   six.print_("Sending email to %s" % (sendTo, ))
   if isinstance(creds,six.string_types) :
      with open(creds,'r') as f :
         creds = json.load(f)
   gmail = smtplib.SMTP_SSL('smtp.gmail.com', 465)
   gmail.ehlo()
   msg = "From: %s\nTo: %s\nSubject: %s\n\n%s"

   Subject = subject or 'Brown/Black belt Hub class instructor sign-up'
   From = sendFrom or creds['uid']
   To = sendTo
   gmail.login(creds['uid'],creds['pwd'])
   gmail.sendmail( From, To, msg % (From, ', '.join(To),Subject, body) )
   gmail.close()

#----------------------------------------------------------------------
class HubCal :
    SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
    def __init__(self, cred = credfile, sec = secret, id = emailcred.get('cal') ):
        store = file.Storage('credentials.json')
        creds = store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets(sec, self.SCOPES)
            creds = tools.run_flow(flow, store)
        self.service = build('calendar', 'v3', http=creds.authorize(Http()))
        self.calid = id
    def events(self):
        # Call the Calendar API
        now = (datetime.utcnow() - timedelta(hours=2)).isoformat() + 'Z'
        events_result = self.service.events().list(calendarId=self.calid, timeMin=now,
                                      maxResults=50, singleEvents=True,
                                      orderBy='startTime').execute()
        events = events_result.get('items', [])
        return events

#----------------------------------------------------------------------
if __name__ == "__main__" :
    main()

