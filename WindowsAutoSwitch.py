import winreg
import time
from datetime import datetime
from suntime import Sun
import geocoder

# ----------- Settings ---------
CHECK_INTERVAL_TIMER = 60 # Check every minute (60 seconds)

def getCurrentLocation():
    g = geocoder.ip('me') # Uses the current Internet Protocol address to determine the location
    if g.ok:
        return g.latlng # Return the coordinates of the IP address
    return None # No IP found - Location cannot be determined

def changeTheme(dark: bool):
    value = 0 if dark else 1 # If the mode is set to dark, then it stays as dark, otherwise it is light.
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize", 0, winreg.KEY_SET_VALUE) # use the Windows Registry to change the theme.
        winreg.SetValueEx(key, "AppsUseLightTheme", 0, winreg.REG_DWORD, value) # Set apps to use the light theme
        winreg.SetValueEx(key, "SystemUsesLightTheme", 0, winreg.REG_DWORD, value) # Set the system light theme to light.
        winreg.CloseKey(key) # Close the registry key
    except Exception:
        pass

def getSunTime(lat, lng):
    sun = Sun(lat, lng) # Get the current lat and lng coordinates
    today = datetime.now().date() # Get today's date
    sunrise = sun.get_local_sunrise_time(today).replace(tzinfo=None) # Find the local sunrise time according to the IP address
    sunset = sun.get_local_sunset_time(today).replace(tzinfo=None) # Find the local sunset time according to the IP address
    return sunrise, sunset

# ------------ Location Setup -----------------------
location = getCurrentLocation()
if location is None:
    location = [51.5074, -0.1278] # Fall back to London's coordinates

REFRESH_LOCATION_HOURS = 6 # Refresh location every six hours
lastLocationCheckTime = datetime.now()
currentMode = None

while True:
    now = datetime.now() # Get the exact time at this very moment

    # Update location periodically
    hourSinceLastCheckTime = (now - lastLocationCheckTime).seconds / 3600
    if hourSinceLastCheckTime >= REFRESH_LOCATION_HOURS: # If time checked exceeds ix
        locationNow = getCurrentLocation() # Get the user's current location
        if locationNow:
            location = locationNow
        lastLocationCheckTime = now

    lat, lng = location
    sunrise, sunset = getSunTime(lat, lng)
    darkCheck = now < sunrise or now >= sunset

    if darkCheck != currentMode:
        changeTheme(dark=darkCheck)
        currentMode = darkCheck

    time.sleep(CHECK_INTERVAL_TIMER)
