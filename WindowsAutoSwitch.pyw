import sys
import os
import winreg
import time
import ctypes
import threading
from datetime import datetime
from datetime import timezone
from suntime import Sun
import geocoder
import pystray
import winotify
from winotify import Notification
from PIL import Image, ImageDraw

# ---------------- Timezone Capitals (in UTC) ----------------
TIMEZONE_CAPITALS = {
    -12: (0.193635, -176.476894), # Baker Island
    -11: (-13.848202, -171.760454), # Apia, Samoa
    -10: (21.301821, -157.858052), # Honolulu, Hawaii
    -9.5: (8.54424, 140.06135), # Taioha'e, French Polynesia
    -9: (58.301986, -134.418464), # Juneau, Alaska
    -8: (38.571181, -121.489869), # Sacramento, California
    -7: (39.741834, -104.997956), # Denver, Colorado
    -6: (30.270916, -97.741936), # Austin, Texas
    -5: (37.544046, -77.441989), # Richmond, Virginia
    -4: (44.645680, -63.614636), # Halifax, Nova Scotia
    -3.5: (47.547054, -52.740050), # St. John's, Newfoundland
    -3: (-15.803536, -47.891805), # Brasilia, Brazil
    -2: (64.174579, -51.731977), # Nuuk, Greenland
    -1: (14.914903, -23.510122), # Praia, Cape Verde
    0: (51.507188, -0.127841), # London, England
    +1: (50.844518, 4.352950), # Brussels, Belgium
    +2: (52.31102, 13.24239), # Berlin, Germany
    +3: (37.59030, 23.43398), # Athens, Greece
    +3.5: (35.43227, 51.19568), # Tehran, Iran
    +4: (24.27138, 54.22426), # Abu Dhabi, United Arab Emirates
    +4.5: (34.33186, 69.12249), # Kabul, Afghanistan
    +5: (33.41572, 73.02130), # Islamabad, Pakistan
    +5.5: (28.36487, 77.12310), # Delhi, India
    +5.75: (27.42431, 85.19101), # Kathmandu, Nepal
    +6: (23.48144, 90.24522), # Dhaka, Bangladesh
    +6.5: (19.45451, 96.04436), # Naypyidaw, Myanmar
    +7: (13.45201, 100.30180), # Bangkok, Thailand 
    +8: (39.51186, 116.20374), # Beijing, China
    +8.75: (31.40347, 128.53170), # Ecula, Western Australia
    +9: (35.39378, 139.45462), # Tokyo, Japan
    +9.5: (12.27535, 130.50413), # Darwin, Australia
    +10: (27.27485, 153.01100), # Brisbane, Australia
    +10.5: (34.55401, 138.35110), # Adelaide, Australia 
    +11: (22.16292, 166.26224), # Noumea, New Caladonia 
    +12: (18.07376, 178.26440), # Suva, Fiji
    +12.75: (43.57072, 176.33328), # Waitangi, Chatham Islands
    +13: (41.17136, 174.46432), # Wellington, New Zealand
    +14: (1.58525, 157.28417), # London, Kirimati
}

# ----------- Settings ---------
CHECK_INTERVAL_TIMER = 60 # Check every minute (60 seconds)
CHANGE_APP_THEME = True # Change the current app theme
CHANGE_SYSTEM_THEME = True # Change the current system theme
MANUAL_OVERRIDE = None # Override the program to allow a theme to stay permanent.
wakeEvent = threading.Event() # This ensures that changes happen instantly
HAS_SEEN_DAYNOTIFICATION = False # Ensures that the user only sees this once per session
HAS_SEEN_NIGHTNOTIFICATION = False # Ensures the user only sees this once per session

def getCurrentLocation():
    g = geocoder.ip('me') # Uses the current Internet Protocol address to determine the location
    if g.ok:
        return g.latlng # Return the coordinates of the IP address
    return None # No IP found - Location cannot be determined

def changeTheme(dark: bool):
    value = 0 if dark else 1 # If the mode is set to dark, then it stays as dark, otherwise it is light.
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize", 0, winreg.KEY_SET_VALUE) # use the Windows Registry to change the theme.

        if CHANGE_APP_THEME:
            winreg.SetValueEx(key, "AppsUseLightTheme", 0, winreg.REG_DWORD, value) # Set apps to use the light theme
        
        if CHANGE_SYSTEM_THEME:
            winreg.SetValueEx(key, "SystemUsesLightTheme", 0, winreg.REG_DWORD, value) # Set the system light theme to light.

        winreg.CloseKey(key) # Close the registry key
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x001A, 0, "ImmersiveColorSet", 0x0002, 5000, None) # Chnage the theme immediately - Do not delay
    except Exception:
        pass

def getSunTime(lat, lng):
    sun = Sun(lat, lng) # Get the current lat and lng coordinates
    today = datetime.now().date() # Get today's date
    sunrise = sun.get_local_sunrise_time(today).replace(tzinfo=None) # Find the local sunrise time according to the IP address
    sunset = sun.get_local_sunset_time(today).replace(tzinfo=None) # Find the local sunset time according to the IP address
    return sunrise, sunset # Return the current time

# ---------------- Start-up -------------------------
def getResourcePath(file):
    base = getattr(sys, '_MEIPASS', os.path.dirname(__file__)) # Get the requested file from the filepath
    return os.path.join(base, file) # Return the file

# --------------- Fallback Location -------------------
def getFallbackLocation():
    secondsOffset = -time.localtime().tm_gmtoff # Enables Winter and Summer switches
    UTCOffset = secondsOffset / 3600 # Compatability with .5 timezones
    return TIMEZONE_CAPITALS.get(UTCOffset, (51.5074, -0.1278)) # Uses London as the fallback capital - if none are found

# ------------------- System Tray --------------------------------
def createIcon(): # Create an icon for the system tray
    return Image.open(getResourcePath("WASIcon.png")) # Opens the WindowsAutoSwitch logo

def darkForce(icon, item): # Responsible for forcing the system / app theme to dark
    global MANUAL_OVERRIDE
    MANUAL_OVERRIDE = True # Allows users to force one theme on
    changeTheme(dark=True) # Set the theme to dark mode
    toast = Notification(app_id="Windows Auto Switch",
    title="The theme has been changed!",
    msg="The current theme has been changed from light mode, to dark mode",
    icon=getResourcePath(r"WASIcon.png"))
    toast.show()

def lightForce(icon, item): # Responsible for forcing the system / app theme to light
    global MANUAL_OVERRIDE
    MANUAL_OVERRIDE = False # Allows users to force one theme on
    changeTheme(dark=False) # Set the theme to light mode
    toast = Notification(app_id="Windows Auto Switch",
    title="The theme has been changed!",
    msg="The current theme has been changed from dark mode, to light mode",
    icon=getResourcePath(r"WASIcon.png"))
    toast.show()

def ChangeSystemTheme(icon, item): # Changes the system's theme
    global CHANGE_SYSTEM_THEME
    CHANGE_SYSTEM_THEME = not CHANGE_SYSTEM_THEME
    icon.update_menu()

def ChangeAppTheme(icon, item): # Changes the app's theme
    global CHANGE_APP_THEME
    CHANGE_APP_THEME = not CHANGE_APP_THEME
    icon.update_menu()

def resumeAutomaticSwitch(icon, item): # Allows the automatic switching of the themes again
    global MANUAL_OVERRIDE, currentMode
    MANUAL_OVERRIDE = None
    currentMode = None
    toast = Notification(app_id="Windows Auto Switch",
    title="Auto switch has been enabled!",
    msg="Automatic Switch has been enabled, the theme will now match the current time of day.",
    icon=getResourcePath(r"WASIcon.png"))
    toast.show()
    wakeEvent.set() # Change the theme instantly

def ExitApp(icon, item): # Quits the application
    icon.stop()

def buildAppMenu(): # For the actual system tray menu itself
    return pystray.Menu(
        pystray.MenuItem("WindowsAutoSwitch", None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Resume automatic switch", resumeAutomaticSwitch, checked=lambda item: MANUAL_OVERRIDE is None), # Checkbox for switching back to automatic
        pystray.MenuItem("Change App Theme", ChangeAppTheme, checked=lambda item: CHANGE_APP_THEME), # Change the app theme?
        pystray.MenuItem("Change System Theme", ChangeSystemTheme, checked=lambda item: CHANGE_SYSTEM_THEME), # Change the system theme?
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Force Light Mode", lightForce), # App never changes from light mode
        pystray.MenuItem("Force Dark Mode", darkForce), # App never changes from dark mode
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Exit application", ExitApp) # Closes the system app down.
    )

# ------------ Location Setup -----------------------
location = getCurrentLocation()
if location is None:
    location = list(getFallbackLocation()) # Fall back to the nearest capital to the corresponding timezone

REFRESH_LOCATION_HOURS = 2 # Refresh location every two hours
lastLocationCheckTime = datetime.now()
currentMode = None

# ------------------ Theme looper --------------------
def loopTheme():
    global location, lastLocationCheckTime, currentMode, HAS_SEEN_DAYNOTIFICATION, HAS_SEEN_NIGHTNOTIFICATION
    while True:
        now = datetime.now() # Get the exact time at this very moment

        # Update location periodically
        hourSinceLastCheckTime = (now - lastLocationCheckTime).seconds / 3600
        if hourSinceLastCheckTime >= REFRESH_LOCATION_HOURS: # If time checked exceeds ix
            locationNow = getCurrentLocation() # Get the user's current location
            if locationNow:
                location = locationNow
            lastLocationCheckTime = now

        lat, lng = location # Get the current latitude and longitude of the user's location
        sunrise, sunset = getSunTime(lat, lng)
        darkCheck = now < sunrise or now >= sunset # Check that the current time is less than the local sunrise, or is more than or equal to the local sunset

        if MANUAL_OVERRIDE is not None:
            pass # Skip
        else:
            if darkCheck != currentMode:
                changeTheme(dark=darkCheck) # Change the theme based on the time of day
                currentMode = darkCheck
                toast = None # Do not show a notification
                if not darkCheck: # It's now daytime, so it should fire the daytime response
                    if now.hour >= 12 and not HAS_SEEN_DAYNOTIFICATION: # Is it 12pm or beyond?
                        HAS_SEEN_DAYNOTIFICATION = True
                        if HAS_SEEN_NIGHTNOTIFICATION: # Has the user already seen the night-time notification?
                            HAS_SEEN_NIGHTNOTIFICATION = False
                        toast = Notification(app_id="Windows Auto Switch",
                                         title="Good Afternoon!",
                                         msg="The theme is now set to the light theme, as it is now daytime. Have a good afternoon! :)",
                                         icon=getResourcePath(r"WASIcon.png"))
                    elif now.hour >= 4 and not HAS_SEEN_DAYNOTIFICATION: # Is it 4am?
                        HAS_SEEN_DAYNOTIFICATION = True
                        if HAS_SEEN_NIGHTNOTIFICATION: # Has the user already seen the day-time notification?
                            HAS_SEEN_NIGHTNOTIFICATION = False
                        toast = Notification(app_id="Windows Auto Switch",
                                         title="Good Morning!",
                                         msg="The theme is now set to the light theme, as it is now day-time! Have a good morning! :)",
                                         icon=getResourcePath(r"WASIcon.png"))
                if darkCheck: # It's night-time! Time to enable dark mode
                    if now.hour >= 19 and not HAS_SEEN_NIGHTNOTIFICATION:
                            HAS_SEEN_NIGHTNOTIFICATION = True
                            if HAS_SEEN_DAYNOTIFICATION:
                                HAS_SEEN_DAYNOTIFICATION = False
                            toast = Notification(app_id="Windows Auto Switch",
                                         title="Good Night!",
                                         msg="The theme is now set to the dark theme, as it is now night-time. Have a good night! :)",
                                         icon=getResourcePath(r"WASIcon.png"))
                    elif now.hour >= 16 and not HAS_SEEN_NIGHTNOTIFICATION:
                        HAS_SEEN_NIGHTNOTIFICATION = True
                        if HAS_SEEN_DAYNOTIFICATION:
                            HAS_SEEN_DAYNOTIFICATION = False
                        toast = Notification(app_id="Windows Auto Switch",
                                         title="Good Evening!",
                                         msg="The theme is now set to the dark theme, as it is now night-time. Have a good evening! :D",
                                         icon=getResourcePath(r"WASIcon.png"))
                    else:  # Beyond 8pm, so just fall back to night-time.
                        HAS_SEEN_NIGHTNOTIFICATION = True
                        if HAS_SEEN_DAYNOTIFICATION:
                            HAS_SEEN_DAYNOTIFICATION = False
                        toast = Notification(app_id="Windows Auto Switch",
                                         title="Good Night!",
                                         msg="The theme is now set to the dark theme, as it is now night-time. Have a good night! :D",
                                         icon=getResourcePath(r"WASIcon.png"))
                if toast is not None:
                    toast.show()



        wakeEvent.wait(timeout=CHECK_INTERVAL_TIMER) # Change the theme quickly
        wakeEvent.clear() # Clear the current event

# -------------------- On system start ----------------
thread = threading.Thread(target=loopTheme, daemon=True) # Start the application.
thread.start() # Start the thread to open the application

icon = pystray.Icon("WindowsAutoSwitch", createIcon(), "WindowsAutoSwitch", buildAppMenu()) # Construct the app itself
icon.run() # Reveal the app icon