#!/usr/bin/python3

#Author: Nathaniel Campan
import newpdf
from bs4 import BeautifulSoup
import requests
import sys
import re
import datetime
import pytz
from icalendar import Calendar, Event
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import os
import argparse
#Générez les api google agenda avant toute chose: https://developers.google.com/calendar/quickstart/python
#Editez les champs ci-dessous:

LOGIN="mchouteau.ira2024"
PASS="586351c3"
CALENDAR_ID="8tbmcpqjct0jkl900ubh7lh6l4@group.calendar.google.com"
os.system('rm /home/mchouteau/Feuille-temps/*')
global cal
cal = Calendar()
def main(argv):
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    parser = argparse.ArgumentParser(description="Synchronise Alcuin sur Google Agenda")
    parser.add_argument('days', type=int, help="Nombre de jours a synchroniser")
    parser.add_argument('-o', '--output', help="Fichier de log")
    parser.add_argument('date', type=int, nargs='?', const=0, default=0, help="Nombre de jours depuis lequels on synchronise")
    
    args = parser.parse_args()

    if args.output:
        sys.stdout = open(args.output, 'w+')


    print("[*] Connexion à Alcuin")
    data, session = getInputs("https://esaip.alcuin.com/OpDotNet/Noyau/Login.aspx")
    session = loginAlcuin("https://esaip.alcuin.com/OpDotNet/Noyau/Login.aspx", data, session)  #Connexion successful
    print("[*] Extraction des données et création du calendrier")
    for delta in range(args.days):
        datefrom = datetime.datetime.today() + datetime.timedelta(args.date)
        date = datefrom + datetime.timedelta(days=delta)
        print("[*] Synchronisation du {}".format(date.strftime("%d/%m/%Y")))
        cal = retrieveCal("", session, date)
        for i in cal:
            calData = extractCalData(i)
            if calData:
                newpdf.main(date.strftime('%d/%m/%Y'),  calData[0], calData[2], calData[5], date.strftime('%d_%m_%y'), date)
#                print (calendrier)
#                cal.add_component(calendrier)
#                debut = datetime.datetime.strftime(debut, '%Y%m%d%H%M%S')
#                print (debut)
                build_event(date, calData[0], calData[1], calData[2], calData[3], calData[4])
    print("[*] Agenda synchronisé!")

def getInputs(url): #Get the inputs to send back to the login page (Tokens etc)
    s = requests.Session()
    r = s.get(url)

    soup = BeautifulSoup(r.text, "html.parser")
    inputs = soup.find_all("input")

    data={}
    for i in inputs:    #Extract inputs and store them in data
        try:
            data[i["name"]]=i["value"]
        except:
            pass
    return(data, s)

def loginAlcuin(url,data,s):    #login
    data["UcAuthentification1$UcLogin1$txtLogin"] = LOGIN
    data["UcAuthentification1$UcLogin1$txtPassword"] = PASS
    try:
        s.post(url, data=data)
        print("[*] Connexion à Alcuin réussie")
    except:
        print("[-] Impossible de se connecter à Alcuin")
        sys.exit()
    return(s)

def retrieveCal(url, s, d):
    r = s.get('https://esaip.alcuin.com/OpDotNet/Context/context.jsx')  #Get user ID to show the right calendar
    usrId = re.search('\w+[0-9]', r.text).group(0)  #Regex to extract user ID
    data = {'IdApplication': '190', 'TypeAcces': 'Utilisateur', 'url': '/EPlug/Agenda/Agenda.asp', 'session_IdCommunaute': '561', 'session_IdUser': usrId, 'session_IdGroupe': '786', 'session_IdLangue': '1'}
    s.post("https://esaip.alcuin.com/commun/aspxtoasp.asp", data=data)  #Retrieve the calendar and create the necessary token
    r = s.post("https://esaip.alcuin.com/EPlug/Agenda/Agenda.asp", data={"NumDat": d.strftime('%Y%m%d'), "DebHor": "08", "FinHor": "18", "ValGra": "60", "NomCal":"PRJ13230", "TypVis":"Vis-Jrs.xsl"}) #Extract a specific day
    soup = BeautifulSoup(r.text, "html.parser")
    cal = soup.find_all("td", {"class": "GEDcellsouscategorie", "valign": None})
    return(cal)

def extractCalData(cal):
    course = cal.get_text() #Retrieve a specific course
    if os.name == 'nt':
        course = course.encode('cp1252').decode('latin9')    #Decode if os is windows
    time = []
    try:
        [time.append(re.split('H', i)) for i in re.findall('\d\dH\d\d', course)]   #[['08', '15'], ['09', '45']]
        courseSplit = re.split('- ', re.split('\d\dH\d\d', course)[0])
        courseSplit.append(re.split('[A-Z]\d|Amphi', re.split('\d\dH\d\d', course)[2])[0])
        courseName = courseSplit[1]
        courseSplit2 = re.split('([A-Z][a-zà-ÿ]{2,}(?=[A-Z]{2,}))', courseSplit[2])
        length = len(courseSplit2)-1
        if length == 0:
            courseSplit3 = courseSplit[2]
            courseSplit[2] = courseSplit[0] + "\x0a" + courseSplit[2]
        else:
            i = 0
            courseSplit3 = ""
            while i < length:
                courseSplit3 = courseSplit3 + courseSplit2[i]
                i += 1
                courseSplit3 = courseSplit3 + courseSplit2[i] + "\x0a"
                i += 1
            courseSplit3 = courseSplit3 + courseSplit2[length]
            courseSplit[2] = courseSplit[0] + "\x0a" + courseSplit3
        if re.search('Examen', courseSplit[0]):
            colorId = '11'
            description = courseSplit[2]
        elif re.search('Cours', courseSplit[0]):
            colorId = '5'
            description = courseSplit[2]
        else:
#            courseName = '{} - {}'.format(courseSplit[0], courseName)
            courseName = courseSplit[0]
            colorId = '4'
            description = courseSplit[2]

        salle1 = re.sub('[A-Z]{2,}\d','',  re.split('\d\dH\d\d', course)[-1])    #Get second part
                
        if re.search('Amphi A', course):
            salle = 'Amphi A'
        elif re.search('Amphi E', course):
            salle = 'Amphi E'
        elif re.search('[ABCEF]\d+', salle1):
            salle = re.search('[ABCEF]\d{2,3}', salle1).group(0)  #Retrieve classroom
        else:   #No classroom
            salle = 'Non renseigné'
        return(time, salle, courseName, colorId, description, courseSplit3)
    except:
        pass

def build_event(d, time, salle, courseName, colorId, description):
    filename = '/home/mchouteau/Feuille-temps/calendrier.ics'
    event = Event()
    event.add('summary', courseName)
    event.add('description', description)
    debut = d.replace(hour=int(time[0][0]),minute=int(time[0][1]), second=0)
    fin = d.replace(hour=int(time[1][0]), minute=int(time[1][1]), second=0)
    event.add('dtstart', debut)
    event.add('dtend', fin)
    event.add('location', salle)
    cal.add_component(event)
    f = open(filename, 'wb')
    f.write(cal.to_ical())
    f.close()

def usage():
    print("Usage: main.py [-h] [-o output] days")
    
if __name__ == "__main__":
    main(sys.argv[1:])
