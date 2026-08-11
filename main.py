#Modules
#GUI generation
from tkinter import *
from tkinter.ttk import *

#email sending
#import win32com.client

#reading the HTML
from bs4 import BeautifulSoup

#Get the current date and time
from datetime import datetime

#Getting info from the web
import requests
import urllib3
import threading

#Excel file generation
from openpyxl import Workbook
from openpyxl import load_workbook
import os
import sys

#Disable warnings that come with unverified requests
#(This line of code is for peace of mind, the 'verify=false' line within requests generates a warning
#since this line is used every time it pulls info from a page in this program, it generates A LOT of warnings
#and this just makes them not show up)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# Scrape 'National Weather Service' to get current local temperatures and convert to Fahrenheit
def get_outside_temp():


    user_agent = "NJTransitWeatherApp (ggenaoperez@njtransit.com)"
    latitude = 40.74392
    longitude = -74.1029
    point_url = f"https://api.weather.gov/points/{latitude},{longitude}"
    headers = {"User-Agent": user_agent}

    points_response = requests.get(point_url, headers=headers)
    points_data = points_response.json()

    stations_url=points_data["properties"]["observationStations"]
    stations_response = requests.get(stations_url, headers=headers)
    stations_data = stations_response.json()
    station_id = stations_data["features"][0]["properties"]["stationIdentifier"]

    obs_url=f"https://api.weather.gov/stations/{station_id}/observations/latest"
    obs_response = requests.get(obs_url, headers=headers)
    obs_data = obs_response.json()

    temp_c = obs_data["properties"]["temperature"]["value"]
    temp_f = (temp_c * 9 / 5) + 32
    return round(temp_f, 1)

# Current Date and Time Local
def get_date_and_time():

    now = datetime.now()
    formatted_date = now.strftime("%m/%d/%Y, %I:%M %p")
    return formatted_date

#Scraping data from 'https://njt.vehicledb.com/' for report generation
#alp 45dp scrape
def get_alp45dp_temps():
    today = datetime.now().strftime("%Y-%m-%d")
    all_data = []

    for loco in range(4500, 4535):
        url = f"https://njt.vehicledb.com/converterReport.php?loco={loco}&date={today}"
        response = requests.get(url, auth=("njt", "njtdb"), verify = False)
        soup = BeautifulSoup(response.text, "html.parser")
        table = soup.find("table", id = "table")
        columns = table.find_all("td")

        con1 = float('-inf')
        con2 = float('-inf')

        if len(columns) >= 5:
            con1 = columns[3].text
            con2 = columns[4].text

        con1 = float(con1)
        con2 = float(con2)

        all_data.append([loco, con1, con2])

    print('alp 45dp data collected')

    return all_data

#alp 46 scrape
def get_alp46_temps():
    today = datetime.now().strftime("%Y-%m-%d")
    all_data = []

    for loco in range(4600, 4629):
        url = f"https://njt.vehicledb.com/converterReport_ALP46.php?loco={loco}&date={today}"
        response = requests.get(url, auth=("njt", "njtdb"), verify=False)
        soup = BeautifulSoup(response.text, "html.parser")
        table = soup.find("table", id="table")
        columns = table.find_all("td")

        con1 = float('-inf')
        con2 = float('-inf')

        if len(columns) >= 5:
            con1 = columns[2].text
            con2 = columns[3].text

        con1 = float(con1)
        con2 = float(con2)

        all_data.append([loco, con1, con2])

    print('alp 46 data collected')

    return all_data

#alp 46a scrape
def get_alp46a_temps():
    today = datetime.now().strftime("%Y-%m-%d")
    all_data = []

    for loco in range(4629, 4665):
        url = f"https://njt.vehicledb.com/converterReport_ALP46A.php?loco={loco}&date={today}"
        response = requests.get(url, auth=("njt", "njtdb"), verify=False)
        soup = BeautifulSoup(response.text, "html.parser")
        table = soup.find("table", id="table")
        columns = table.find_all("td")

        con1 = float('-inf')
        con2 = float('-inf')

        if len(columns) >= 5:
            con1 = columns[2].text
            con2 = columns[3].text

        con1 = float(con1)
        con2 = float(con2)

        all_data.append([loco, con1, con2])

    print('alp 46a data collected')

    return all_data

#alp 45a scrape
def get_alp45a_temps():

    today = datetime.now().strftime("%Y-%m-%d")

    all_data = []

    for loco in range(4535, 4561):
        url = f"https://njt.vehicledb.com/converterReport_alp45a.php?loco={loco}&date={today}"
        response = requests.get(url, auth=("njt", "njtdb"), verify=False)
        soup = BeautifulSoup(response.text, "html.parser")
        table = soup.find("table", id="table")
        columns = table.find_all("td")

        con1 = float('-inf')
        con2 = float('-inf')

        if len(columns) >= 5:
            con1 = columns[3].text
            con2 = columns[4].text

        con1 = float(con1)
        con2 = float(con2)

        all_data.append([loco, con1, con2])

    print('alp 45a data collected')

    return all_data

#generate the Excel file
def generate_excel_file():
    # generate ALP Temps Report Excel sheet
    # I tried to copy the format of the original as closely as possible, this is just a base sheet with formulas that is populated
    today = datetime.now().strftime("%m.%d.%y %H%M %p")

    pb['value'] = 0

    wb = load_workbook(resource_path("Automated ALP Temps TEMPLATE.xlsx"))

    ws = wb["Sheet1"]
    all_data = get_alp45dp_temps()
    all_data.sort(reverse = True, key=lambda x: x[2])
    for row in all_data:
        ws.append(row)

    pb['value'] += 20

    ws = wb["Sheet2"]
    all_data = get_alp46_temps()
    all_data.sort(reverse=True, key=lambda x: x[2])
    for row in all_data:
        ws.append(row)

    pb['value'] += 20

    ws = wb["Sheet3"]
    all_data = get_alp46a_temps()
    all_data.sort(reverse=True, key=lambda x: x[2])
    for row in all_data:
        ws.append(row)

    pb['value'] += 20

    ws = wb["Sheet4"]
    all_data = get_alp45a_temps()
    all_data.sort(reverse=True, key=lambda x: x[2])
    for row in all_data:
        ws.append(row)

    pb['value'] += 20

    print('data appended \ngenerating excel file...')

    year = datetime.now().strftime("%Y")

    wb.save(f"F:\\42 ALPs Converter Temp\\NJTDB Temps\\{year}\\ALP TEMPS {today}.xlsx")

    print('excel file generated')

    pb['value'] += 20
    status.config(text="File Generated to F: Drive")

    os.startfile(f"F:\\42 ALPs Converter Temp\\NJTDB Temps\\{year}\\ALP TEMPS {today}.xlsx")

#--------------------------------------------------------------------------------------------------------------------------------
#Button Commands
def button1commands():

    status.config(text="Generating..")
    threading.Thread(target=generate_excel_file).start()

def button2commands():

    """
    todaye = datetime.now().strftime("%M/%D/%Y %H:%M %p")
    temp = get_outside_temp()
    year = datetime.now().strftime("%Y")
    today = datetime.now().strftime("%Y.%m.%d %H%M %p")

    ol = win32com.client.DispatchEx("Outlook.Application")
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = f"ALP Temps {today}"
    newmail.To = ("Rail Mech Tech Services <RailMechTechServices@njtransit.com>; "
                  "Rail Mech QA/QC <RailMechQA_QC@njtransit.com>; "
                  "Rail Weekend Duty Officer <RailWeekendDutyOfficer@njtransit.com>; "
                  "Rail Mech MMC Locomotive Shop Foremen <Rail_Mech_MMC_Locomotive_Shop_Foremen@njtransit.com>; "
                  "Rail Mechanical Desk <RailMechanicalDesk@njtransit.com>; "
                  "Rail Mech Dover Yard Group <Rail_Mech_Dover_Yard_Group@njtransit.com>; "
                  "Rail Mech Gladstone Yard Group <Rail_Mech_Gladstone_Yard_Group@njtransit.com>; "
                  "Rail Mech Great Notch Yard Group <Rail_Mech_Great_Notch_Yard_Group@njtransit.com>; "
                  "Rail Mech Hoboken Yard Group <Rail_Mech_Hoboken_Yard_Group@njtransit.com>; "
                  "Rail Mech County Yard Group <Rail_Mech_County_Yard_Group@njtransit.com>; "
                  "Rail Mech Long Branch Yard Group <Rail_Mech_Long_Branch_Yard_Group@njtransit.com>; "
                  "Rail Mech Morrisville Yard Group <Rail_Mech_Morrisville_Yard_Group@njtransit.com>; "
                  "Rail Mech Port Morris Yard Group <Rail_Mech_Port_Morris_Yard_Group@njtransit.com>")

    newmail.CC = ("DeGennaro, David P.   (CROPDPD) <DDegennaro@njtransit.com>; "
                  "Rogust, Daniel G.     (CROPDGR) <DRogust@njtransit.com>; "
                  "Breen, Robert L.      (CROPRLB) <RBreen@njtransit.com>; "
                  "Kunchandy, George M.  (CROPGMK) <GKunchandy@njtransit.com>; "
                  "Panza, Adam J.        (CROPAJP) <APanza@njtransit.com>; "
                  "Ortland, Milena M.    (CROPMMO) <MOrtland@njtransit.com>; "
                  "Patel, Yogesh R.      (CROPYRP) <YPatel@njtransit.com>")

    newmail.Body = (f"All, "
                    f"\nAttached is the ALP Converter temperature report for {todaye}."
                    f"The current outside temperature is {temp}."
                    f"All converter readings are within normal operating limits."
                    f"\nRegards,")

    attach = f"F:\\42 ALPs Converter Temp\\NJTDB Temps\\{year}\\ALP Temps {today}.xlsx"

    newmail.Attachments.Add(attach)
    newmail.Display()
    """

#--------------------------------------------------------------------------------------------------------------------------------
# create the main window
window = Tk()
window.title("ALP Temperatures Email Report Generator - NJT Tech Services")
window.geometry("190x150")
window.config(bg="Blue")


#widget creation
button1 = Button(window, text="Generate ALP Temps Excel File", command=button1commands)
button2 = Button(window, text="Send ALP Temps Email Report", command=button2commands)
progress_text = Label(window, text = "Press either button to start")
pb = Progressbar(window, mode = "determinate", length = 100, maximum = 100)
status = Label(window, text="Click to use")

#widget placing

status.grid(row = 0, column = 0, padx = 10, pady = 5)
button1.grid(row = 1, column = 0, padx = 10, pady = 10)
button2.grid(row = 2, column = 0 , padx = 10)
pb.grid(row = 3, column = 0, padx = 10, pady = 10)


#open window!
window.mainloop()
