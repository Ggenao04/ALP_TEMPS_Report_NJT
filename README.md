# NJ TRANSIT ALP TEMPERATURES REPORT & EMAIL GENERATOR
### Overview

This app is intended to eventually replace the automated alp temperature report email generator that operates using VBA within Microsoft Excel. Although currently in crude operating, at its current state the app completes all the same functions as the 'Automated ALP Temps v4.2' Excel sheet. 

## Possible issues

  ### Unable to get current outside temperature:
  
  This means the website is not receiving a proper header when the request for information is being sent. likely the email written in the source code has been taken out of service.
  To fix, simply edit the 'get_outside_temp' function in the main.py file and replace the email in the header with either your, or an updated email that is accessible. 

    def get_outside_temp():
        user_agent = "NJTransitWeatherApp (YOUR EMAIL HERE)"




## Known Issues

  ### Chart Design lost upon generation of report:
This is a known issue with 'openpyxl', Python, and Microsoft Excel. When openinng up the template for data appending, openpyxl does NOT store chart design data so upon saving a copy for the newly created report, the chart design data appears as blank, no data affected it is just a graphical error.

A possible fix would be to instead generate a new chart each time the report is created through python. Below is some documentation in case the former should be attempted.

https://www.geeksforgeeks.org/python/creating-charts-using-openpyxl/
  
  ### No signs of activity upon generating email report:
The app appears to freeze when generating the email report. Originally I had the email report generate within a separate thread similar to the actual temperature generation however when generating the email report in a thread, a timing mismatch occurs between when the report is generated and the email is created because of the amount of time it takes to parse the data from each locomotive report website. The app still works as intended, the progress bar just does not update so the app appears unresponsive.

## Packaging the app for updates or bugfixes:

1. Download all the files into a folder you can access
2. open cmd, cd into the folder you just created which should include all the files.
3. Ensure all proper libraries are installed, check through an IDE (PyCharm Preferred)
4. Run the following command in cmd: *python -m PyInstaller --onefile --add-data "Automated ALP Temps TEMPLATE.xlsx;." --name ALPTempsV2.0 --windowed main.py*


  ## **NOTE:**

  This app is expandable! If an issue ever occurs such that python stops supporting being able to open outloook (possibly due to any changes with the new outlook), a simple fix would be to have python itself send out the email using the known smtp library. My advise would be to have python generate a new window that displays the email in a similar formatting to a regular email application if vieweing before sending is preffered, or to have python just directly send out the email upon report generation. I'll include some documentation that I believe would be neccesary to implement this update.

  https://docs.python.org/3/library/smtplib.html

  https://realpython.com/python-send-email/

  https://docs.python.org/3/library/tkinter.html

  *Final Thoughts,*

  *Feel free to contact me for any bugs that come up in the current code and I could take a look. This was a fun project to work on*
  
  
