"""pinger"""

import os
import sys
import re
import smtplib
import time
import getpass

try:
    import openpyxl
except ImportError as IE:
    print ("Nie wykryto modułu openpyxl. Zainstaluj openpyxl.")
    sys.exit(1)
    
def openfile():
    """Tworzy listę plików .xlsx w katalogu, zwraca ich listę pozwalając na wybór
    gdy istnieje tylko jeden plik, natychmiast go parsuje. Zamyka program w razie
    nie odnalezienia żadnego .xlsx."""
    
    filelist = [file for file in os.listdir('.') if os.path.isfile(file) \
                    and re.search(r".xlsx", file)]
                    
    if len(filelist) > 1:
    
        print ("Odnaleziono więcej niż jeden plik pingu w folder, wybierz z poniższej listy wpisując daną liczbę:")
        
        for file in filelist:
            print ("{}. - {}".format(filelist.index(file) + 1, file))
            
        while True:
        
            decision = input("wpisz numer pliku > ")
            
            try:
                return filelist[int(decision) - 1]
            except ValueError:
                print("musisz wprowadzić liczbę!")
            except IndexError:
                print("Liczba wykracza poza zasięg listy!")
                
    elif len(filelist) == 0:
        print ("Brak pliku excel")
        exit(1)
    else:
        print ("Odnaleziono plik: {}".format(filelist[0]))
        return filelist[0]

def checkmail(cell):
    """Funkcja używa wyrazenia regularnego do sprawdzenia, czy string/komórka
    jest adresem mailowym.
    ad.1 - funkcja domyśnie obrabia komórwki, ale wykorzystujemy ją również
    dla weryfukacji poprawności adresu mailowego, dlatego akceptuje też string."""

    if type(cell) != str:   #patrz ad.1
        cell = cell.value
        
    mailre = re.compile(r'([^\s]+)@([\w]+)\.([\w]+)')
    if cell and cell.startswith("mailto:"):
        print ("Uwaga: w wierszu {} odnaleziono bezpośredni odnośnik '{}',\n"\
        "sprawdź wiersz (może wyglądać na pusty, ale tak nie jest.)".format(cell.row, cell))
        time.sleep(1)
    elif cell and mailre.match(cell.strip()):
        return True
    else:
        return False
        
        
def checkurl(cell):
    """Funkcja używa wyrazenia regularnego do sprawdzenia, czy string
    jest linkiem do redmine."""
    
    checkurlre = re.compile(r'https://redmine.beyond.pl/issues/[0-9]{6}')
    
    if cell and checkurlre.match(cell.strip()):
        return True
    else:
        return False

        
def rowverification(row, mail):
    """Funkcja werfikuje dany wiersz. Sprawdza, czy istnieje
    link do redmine (sprawdzany poprzez checkurl() oraz opis
    do ticketu, którym może byc dowolny niepusty string.
    Funkcja sprawdza też, czy nie występują duplikaty lub braki,
    oraz informuje o tym użytkownika."""
    
    desc = ""
    link = ""
    

    if not checkemptyrow(row):
        return False
    
    for cell in row:
        if checkmail(cell):# Musimy ponownie sprwadzić, czy jest mail we fukncji, by nie został zinterpretowany jako opis.
            pass
        elif checkmail(cell) and cell.value != mail:
            return False
        elif checkurl(cell.value) and link:
            raise ValueError("W wierszu {} istnieją zduplikowane linki".format(cell.row))
        elif checkurl(cell.value) and not link:
            link = cell.value
        elif cell.value and desc:
            raise ValueError("W wierszu {} istnieją zduplikowane opisy".format(cell.row))
        elif cell.value:
            desc = cell.value.strip()

    if desc and link:
        return ' - '.join([link, desc])
    elif not desc:
        raise ValueError("W wierszu {} brakuje opisu.".format(row[0].row))
    elif not link:
        raise ValueError("W wierszu {} brakuje linku.".format(row[0].row))
         
         
def readxlsfile(xlsfile):
    """funkcja zwraca nam obiekt arkusza excelowskiego 
    do dalszej obróbki"""

    workbook = openpyxl.load_workbook(xlsfile)
    sheet = workbook.get_sheet_names()[0] # analizujemy tylko pierwszy skoroszyt
    activesheet = workbook.get_sheet_by_name(sheet)
    
    return activesheet
    

def checkemptyrow(row):	
    """Funkcja sprwadza, czy wierss jako całość jest pusty"""
    rowtocheck = [cell.value for cell in row if cell.value != None]
    if rowtocheck:
        return True
    else:
        return False
    
def tablemap(worksheet):
    """ Konstruuje listę wszystkich maili wraz z nr wiersza
    w którym występuje."""
    
    mailmap = []
    mailadresses = []
    messagedict = {}
    
    for column in tuple(worksheet.columns):
        for cell in column:
            if checkmail(cell) and cell.value not in mailadresses:
                mailmap.append(cell.row)
                mailadresses.append(cell.value)
    
    for mapping, mail in zip(mailmap, mailadresses):
            messagedict.update(sheetwalker(mapping, mail))
            
    return messagedict

def sheetwalker(mapping, mail):
    """wykorzystuje dane z tablemap do stworzenia kompletnych wiadomości 
    w postaci dictionary w którym klucz jest adresem mailowym, a wartość
    już sformatowanym stringiem stanowiącym pełną treść wiadomości."""
    
    message = {}
    response = ["Cześć,\nProsimy o aktualizacje zgłoszeń/złoszenia zgodnie z poniższą listą:\n"]
    counter = mapping
    
    while counter <= worksheet.max_row:
        row = rowverification(worksheet[counter], mail)
        counter += 1
        if not row or counter == worksheet.max_row:
            message[mail] = '\n'.join(response)
            return message
        else:
            response.append(row)

  
def getsmtpdata():
    """Funkcja odpytująca użytkownika o dane do smtp."""
    
    while True:
    
        username = input("Wprowadź swój adres email > ")
        
        if checkmail(username):
            password = getpass.getpass()
            return tuple([username, password])
        else:
            print ("Wprowadź poprawny adres email!")

        
def messagesaccept(messagedict):
    """Funkcja listuje kompletną treść wszystich pingów, odpytujać użtykownika, 
    czy chce rozpocząć wysyłke."""
    
    print("Arkusz poprawnie odczytany, przejrzyj czy wszystko się zgadza:\n")
    
    for key in messagedict.keys():
        print ("Adres email: {} \nTreść  pingu:".format(key))
        print (messagedict[key])
        print ("\n")
        
    while True:    
        accept = input("Czy chcesz rozesłać ten ping? (T/N) >")
        if accept == "T":
            return True
        elif accept == "N":
            print ("Exiting...")
            exit(0)
        print ("Wpisz T lub N!")


        
def smtpconnect(login, password):
    """Funkcja logująca się do serwera SMTP"""
    
    print("Łącze do serera SMTP: xxx, port 465")
    mailserver = smtplib.SMTP_SSL('xxx', 465, timeout=10)
    
    print("Połączono z serwerem SMTP")
    mailserver.ehlo()
    
    print("Wymieniono wiadomości powitalne")
    mailserver.login(login, password)
    
    print("Pomyślne logowanie via SSL.")
    return mailserver
    
def sendping(messagedict, fromaddr, ccaddres="xxx"):
    """Funkcja wysyłająca wiadomości po udanym logowaniu SMTP."""
    
    for key in messagedict:
        msg = "\r\n".join([
            "Content-Type: text/plain; charset=utf-8"
            "From: {}".format(fromaddr),
            "To: {}".format(key),
            "Cc: {}".format(ccaddres),
            "Subject: Ping Weekendowy",
            "",
            "{}".format(messagedict[key]),
            ""
            ])
        smtserver.sendmail(fromaddr, [key, ccaddres], msg.encode("utf-8"))
        print ("Wysłano wiadomość do {}".format(key))
        
    print ("Rozesłano wszystkie pingi.")
    smtserver.quit()
    print ("Zamknięto połączenie SMTP")
    print ("Exiting...")


if __name__ == "__main__":
    pingfile = openfile()
    worksheet = readxlsfile(pingfile)
    pingsend = tablemap(worksheet)
    messagesaccept(pingsend)
    smtplogindata = getsmtpdata()
    #smtserver = smtpconnect(smtplogindata[0], smtplogindata[1])
    #sendping(pingsend, smtplogindata[0])

