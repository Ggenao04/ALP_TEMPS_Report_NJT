#Modules
#GUI generation
from tkinter import *
from tkinter.ttk import *

#email sending
import urllib.parse

#reading the HTML
from bs4 import BeautifulSoup

#Get the current date and time
from datetime import datetime

#Getting info from the web
import requests
import urllib3
import threading

#Excel file generation
from openpyxl import load_workbook
import os
import sys

#Disable warnings that come with unverified requests
"""
(This line of code is for peace of mind, the 'verify=false' line within requests generates a warning
since this line is used every time it pulls info from a page in this program, it generates A LOT of warnings
and this just makes them not show up)
"""
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Getting the absolute file path for the Excel file pre-packaging
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
    concern_locos = []

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

        if con1 >= 125 or con2 >= 125:
            concern_locos.append(loco)

        all_data.append([loco, con1, con2])

    print('alp 45dp data collected')

    return all_data, concern_locos

#alp 46 scrape
def get_alp46_temps():
    today = datetime.now().strftime("%Y-%m-%d")
    all_data = []
    concern_locos = []

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

        if con1 >= 125 or con2 >= 125:
            concern_locos.append(loco)

        all_data.append([loco, con1, con2])

    print('alp 46 data collected')

    return all_data, concern_locos

#alp 46a scrape
def get_alp46a_temps():
    today = datetime.now().strftime("%Y-%m-%d")
    all_data = []
    concern_locos = []

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

        if con1 >= 125 or con2 >= 125:
            concern_locos.append(loco)

        all_data.append([loco, con1, con2])

    print('alp 46a data collected')

    return all_data, concern_locos

#alp 45a scrape
def get_alp45a_temps():

    today = datetime.now().strftime("%Y-%m-%d")

    all_data = []
    concern_locos = []

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

        if con1 >= 125 or con2 >= 125:
            concern_locos.append(loco)

        all_data.append([loco, con1, con2])

    print('alp 45a data collected')

    return all_data, concern_locos

#generate the Excel file
def generate_excel_file():
    # generate ALP Temps Report Excel sheet
    # I tried to copy the format of the original as closely as possible, this is just a base sheet with formulas that is populated
    today = datetime.now().strftime("%m.%d.%y %H%M %p")
    year = datetime.now().strftime("%Y")
    all_concern_locos = []

    pb['value'] = 0 #resets the progress bar to 0

    # Loading the Excel template
    wb = load_workbook(resource_path("Automated ALP Temps TEMPLATE.xlsx"))

    ws = wb["Sheet1"] #Accessing 'Sheet1' from the template
    all_data, concern_locos = get_alp45dp_temps() # Calling scraping function (line 81), storing returns
    all_data.sort(reverse = True, key=lambda x: x[2]) #Sorting the list of locos (high to low) by temp of conv 2
    #appending sorted data ([loco, conv1, conv2]) to 'Sheet1'
    for row in all_data:
        ws.append(row)
    #storing any locos with conv temps above 125 for later use in email report
    if concern_locos:
        all_concern_locos.append(concern_locos)

    pb['value'] += 20 #once the editing of the sheet is done, the progress bar is increased

    ws = wb["Sheet2"]
    all_data, concern_locos = get_alp46_temps()
    all_data.sort(reverse=True, key=lambda x: x[2])
    for row in all_data:
        ws.append(row)
    if concern_locos:
        all_concern_locos.append(concern_locos)

    pb['value'] += 20

    ws = wb["Sheet3"]
    all_data, concern_locos = get_alp46a_temps()
    all_data.sort(reverse=True, key=lambda x: x[2])
    for row in all_data:
        ws.append(row)
    if concern_locos:
        all_concern_locos.append(concern_locos)

    pb['value'] += 20

    ws = wb["Sheet4"]
    all_data, concern_locos = get_alp45a_temps()
    all_data.sort(reverse=True, key=lambda x: x[2])
    for row in all_data:
        ws.append(row)
    if concern_locos:
        all_concern_locos.append(concern_locos)

    pb['value'] += 20

    status.config(text="File Generated to F: Drive")
    wb.save(f"F:\\42 ALPs Converter Temp\\NJTDB Temps\\{year}\\ALP TEMPS {today}.xlsx")
    os.startfile(f"F:\\42 ALPs Converter Temp\\NJTDB Temps\\{year}\\ALP TEMPS {today}.xlsx")
    pb['value'] += 20

    return all_concern_locos

#formatting the list of concerned locos
def format_concerns(all_concern_locos):
    flat = []
    for group in all_concern_locos:
        flat.extend(group) if isinstance(group, list) else flat.append(group)
    return ", ".join(str(loco) for loco in flat) if flat else "(none)"

#building message to be sent depending on if theres are concerning locos or not
def build_message(all_concern_locos, today, temp):
    if all_concern_locos:
        body = (f"All,\n\n"
                f"Attached is the ALP converter temperature report for {today}. "
                f"The current outside temperature is {temp}\u00b0F. The unit(s) listed "
                f"below have converter temperatures of 125\u00b0F or higher and require "
                f"immediate attention. All other converter readings are within normal "
                f"operating limits.\n\n"
                f"    {format_concerns(all_concern_locos)}\n\n"
                f"Regards,")
    else:
        body = (f"All,\n\n"
                f"Attached is the ALP converter temperature report for {today}. "
                f"The current outside temperature is {temp}\u00b0F. All converter "
                f"readings are within normal operating limits.\n\n"
                f"Regards,")
    return f"ALP Temps {today}", body

#Send the email report
def send_email():
    all_concern_locos = generate_excel_file()
    today = datetime.now().strftime("%m.%d.%y %H%M %p")
    year = datetime.now().strftime("%Y")
    temp = get_outside_temp()
    filepath = f"F:\\42 ALPs Converter Temp\\NJTDB Temps\\{year}\\ALP TEMPS {today}.xlsx"

    to_line = ("RailMechTechServices@njtransit.com; "
               "RailMechQA_QC@njtransit.com; "
               "RailWeekendDutyOfficer@njtransit.com; "
               "Rail_Mech_MMC_Locomotive_Shop_Foremen@njtransit.com; "
               "RailMechanicalDesk@njtransit.com; "
               "Rail_Mech_Dover_Yard_Group@njtransit.com; "
               "Rail_Mech_Gladstone_Yard_Group@njtransit.com; "
               "Rail_Mech_Great_Notch_Yard_Group@njtransit.com; "
               "Rail_Mech_Hoboken_Yard_Group@njtransit.com; "
               "Rail_Mech_County_Yard_Group@njtransit.com; "
               "Rail_Mech_Long_Branch_Yard_Group@njtransit.com; "
               "Rail_Mech_Morrisville_Yard_Group@njtransit.com; "
               "Rail_Mech_Port_Morris_Yard_Group@njtransit.com; "
               "Rail_Mech_Raritan_Yard_Group@njtransit.com; "
               "Rail_Mech_Atlantic_City_Yard_Group@njtransit.com; "
               "Rail_Mech_Suffern_Yard_Group@njtransit.com; "
               "Rail_Mech_Spring_Valley_Yard_Group@njtransit.com; "
               "Rail_Mech_New_York-SSYD_Yard_Group@njtransit.com; "
               "Rail_Mech_Port_Jervis_Yard_Group@njtransit.com; "
               "Rail_Mech_Bay_Head_Yard_Group@njtransit.com")

    cc_line = ("DDegennaro@njtransit.com; DRogust@njtransit.com; "
               "RBreen@njtransit.com; GKunchandy@njtransit.com; "
               "APanza@njtransit.com; MOrtland@njtransit.com; "
               "YPatel@njtransit.com")

    subject, body = build_message(all_concern_locos, today, temp)
    status.config(text="Review the email")

    win = Toplevel(window)
    win.title("Review before sending")
    win.geometry("680x520")

    display = Text(win, wrap="word")
    display.insert("1.0",
                   f"To:      {to_line}\n\n"
                   f"Cc:      {cc_line}\n\n"
                   f"Subject: {subject}\n"
                   f"Attach:  {os.path.basename(filepath)}\n"
                   f"{'-' * 80}\n{body}")
    display.config(state=DISABLED)
    display.pack(fill=BOTH, expand=True, padx=8, pady=8)

    def copy(value, label):
        window.clipboard_clear()
        window.clipboard_append(value)
        window.update()
        status.config(text=f"{label} copied")

    row = Frame(win)
    row.pack(pady=6)
    Button(row, text="Copy To", command=lambda: copy(to_line, "To")).pack(side=LEFT, padx=4)
    Button(row, text="Copy Cc", command=lambda: copy(cc_line, "Cc")).pack(side=LEFT, padx=4)
    Button(row, text="Copy body", command=lambda: copy(body, "Body")).pack(side=LEFT, padx=4)
    Button(row, text="Copy all", command=lambda: copy(f"{to_line}\n\n{cc_line}\n\n{subject}\n\n{body}","Everything")).pack(side=LEFT, padx=4)
    Button(row, text="Cancel", command=win.destroy).pack(side=LEFT, padx=4)

#--------------------------------------------------------------------------------------------------------------------------------
#Button Commands
def button1commands():

    status.config(text="Generating..")
    thread  = threading.Thread(target=generate_excel_file).start()


def button2commands():

    status.config(text="Generating Email...")
    thread = threading.Thread(target=send_email).start()


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


