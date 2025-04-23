import pyautogui
import time
import cv2
import numpy as np
import sys      
import ctypes   

# Referenzbilder laden
new_mail = cv2.imread('pictures/new_mail.png', cv2.IMREAD_COLOR)
new_mail_read = cv2.imread('pictures/new_mail_read.png', cv2.IMREAD_COLOR)
mail_open = cv2.imread('pictures/mail_open.png', cv2.IMREAD_COLOR)
verlaengert = cv2.imread('pictures/verlaengert.png', cv2.IMREAD_COLOR)
mailprogramm_start = cv2.imread('pictures/mailprogramm_start.png', cv2.IMREAD_COLOR)
mail_loeschen = cv2.imread('pictures/mail__loeschen.png', cv2.IMREAD_COLOR)

# Toleranzen festlegen
toleranz_new_mail = 0.9
toleranz_new_mail_read = 0.9
toleranz_mail_open = 0.9
toleranz_verlaengert = 0.8
toleranz_mailprogramm_start = 0.8
toleranz_mail_loeschen = 0.9

# Anzahl verlägerter Anzeigen speichern
anzahl_verlaengert = 0

# Zeitpunkt des letzten Klicks speichern
letzter_klick = time.time()

# Array mit allen Referenzbildern, Toleranz und ID und ob klick erlaubt ist
referenzbilder = [
    {'id': 'new_mail', 'bild': new_mail, 'toleranz': toleranz_new_mail, 'allowed': True},
    {'id': 'new_mail_read', 'bild': new_mail_read, 'toleranz': toleranz_new_mail_read, 'allowed': True},
    {'id': 'mail_open', 'bild': mail_open, 'toleranz': toleranz_mail_open, 'allowed': False},
    {'id': 'verlaengert', 'bild': verlaengert, 'toleranz': toleranz_verlaengert, 'allowed': False},
    {'id': 'mail_loeschen', 'bild': mail_loeschen, 'toleranz': toleranz_mail_loeschen, 'allowed': False}, #Reihenfolge, damit mail gelöscht wird, auch wenn start-leiste schon draußen
    {'id': 'mailprogramm_start', 'bild': mailprogramm_start, 'toleranz': toleranz_mailprogramm_start, 'allowed': False}

    
]

# Funktion zum Vergleichen von Bildern
# brief: Funktion vergleicht aktuellen Bildschirminhalt mit allen Referenzbildern
# input: Array mit allen Referenzbildern
# output: ID und Ort der oberen linken Ecke eines passenden Referenzbildes  (vom ersten passenden)
def find_picture(referenzbilder):
    # Screenshot erstellen
    screenshot = pyautogui.screenshot()
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    
    # Schleife über alle Referenzbilder
    for ref in referenzbilder:
        if ref['allowed']:
            # Bildvergleich
            result = cv2.matchTemplate(screenshot, ref['bild'], cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            # Wenn Übereinstimmung gefunden, ID zurückgeben
            if max_val > ref['toleranz']:
                return (ref['id'], max_loc)
    return None

# Funktion zum Auswählen des passenden Klickpunktes
# brief: Funktion klickt, abhängig von der übergebenen Bild-ID, an entsprechender Postion
# input: ID und Ort der oberen linken Ecke eines passenden Referenzbildes
# return: None
def click(id, location):
    global anzahl_verlaengert
    if id == 'new_mail' and referenzbilder[0]['allowed']:
        referenzbilder[0]['allowed'] = False
        referenzbilder[1]['allowed'] = False
        referenzbilder[2]['allowed'] = True
        click_on_screen(location[0] + 50, location[1] + 20)
    elif id == 'new_mail_read' and referenzbilder[1]['allowed']:
        referenzbilder[1]['allowed'] = False
        referenzbilder[0]['allowed'] = False
        referenzbilder[2]['allowed'] = True
        click_on_screen(location[0] + 50, location[1] + 20)
    elif id == 'mail_open' and referenzbilder[2]['allowed']:
        referenzbilder[2]['allowed'] = False 
        referenzbilder[3]['allowed'] = True
        print('Mail geöffnet')
        click_on_screen(location[0] + 200, location[1] + 170)
    elif id == 'verlaengert' and referenzbilder[3]['allowed']:
        referenzbilder[3]['allowed'] = False
        referenzbilder[4]['allowed'] = True
        referenzbilder[5]['allowed'] = True
        anzahl_verlaengert += 1
        print('Anzeigen verlängert')
        click_on_screen(1900, 30)
    elif id == 'mailprogramm_start' and referenzbilder[5]['allowed']:
        referenzbilder[5]['allowed'] = False
        print('Mailprogramm gestartet')
        click_on_screen(location[0] + 90, location[1] + 30)
    elif id == 'mail_loeschen' and referenzbilder[4]['allowed']:
        referenzbilder[4]['allowed'] = False
        referenzbilder[0]['allowed'] = True
        referenzbilder[1]['allowed'] = True
        click_on_screen(location[0] + 220, location[1] + 120)
        time.sleep(0.5)
        click_on_screen(10, 780)

# Funktion zum Klicken
# brief: Funktion klickt auf den Bildschirm
# input: Position des Klicks, Dauer der Bewegung
# return: None
def click_on_screen(x, y, duration = 0.5):
    pyautogui.moveTo(x, y, duration=duration)
    pyautogui.click()
    time.sleep(0.5)


# Funktion zum Ende des Programms
# brief: Funktion zeigt Meldung, dass keine weitere Übereinstimmung gefunden wurde, fragt ob programm beendet werden soll
# input: None
# return: None
def end_program():
    # Flags für OK und Abbrechen
    MB_OKCANCEL = 0x00000001
    IDOK = 1
    IDCANCEL = 2

    # Messagebox anzeigen
    result = ctypes.windll.user32.MessageBoxW(0, f" Es wurden {anzahl_verlaengert} Anzeigen verlängert. Keine Übereinstimmung gefunden. Programm beenden?", "Programm beenden", MB_OKCANCEL)
    #time.sleep(1)
    if result == IDOK:
        print('Programm beendet')
        sys.exit()
    else:
        print('Programm läuft weiter')
        return

# Hauptfunktion
# brief: Funktion startet den Ablauf
# input: None
# return: None
def main():
    print('Autoclicker gestartet')
    letzter_klick = time.time()
    global anzahl_verlaengert
    anzahl_verlaengert = 0
    #anzahl_verlaengert += 1
    time.sleep(2)
    while True:
        # Bildvergleich
        result = find_picture(referenzbilder)
        
        # Wenn Übereinstimmung gefunden, klicken
        if not result == None:
            click(result[0], result[1])
            # Zeitpunkt des letzten Klicks speichern
            letzter_klick = time.time()
        elif result == None and (time.time() - letzter_klick) > 10: 
            # Wenn kein Bild gefunden, zeige statusmeldung und frage ob programm beendet werden soll
            end_program()
            letzter_klick = time.time()
        time.sleep(1)

if __name__ == "__main__":
    main()
