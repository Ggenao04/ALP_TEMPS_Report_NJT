Possible issues:

  Unable to get current outside temperature:
    This means the website is not receiving a proper header when the request for information is being sent. likely the email written in the source code has been taken out of service.
    To fix, simply edit the 'get_outside_temp' function in the main.py file and replace the email in the header with either your, or an updated email that is accessible. 

    def get_outside_temp():
        user_agent = "NJTransitWeatherApp (YOUR EMAIL HERE)"


python -m PyInstaller --onefile --add-data "Automated ALP Temps TEMPLATE.xlsx;." --name ALPTempsV2.0 --windowed main.py
