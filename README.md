# Microsoft Graph Calendar API Integrations for Linda

> This is a library (graph.py) that I built  in a fortnight to provide our chatbot with a way to play with Outlook Calendar by hijacking the Microsoft Graph API. You too can import graph.py into your projects and do stuff!

## Dependencies
This junk relies on [Python 3.10.11](https://www.python.org/downloads/release/python-31011/) (that's the latest Python 3.10 release) to run properly. Any other Python version borks it. Don't ask me why, apparently Microsoft still relies on old libraries to do stuff. Install it using whichever way you are comfortable, I personally just uninstalled 3.12 and installed this one because I am too lazy to manage multiple Python versions.
 
## Setup
1. Git clone this repo: `git clone`
2. Inside the repo: `pip install -r requirements.txt`. There are lots of packages so this step may take a while.
3. Go to this website: https://developer.microsoft.com/en-us/graph/quick-start
4. Run through the wizard and download the zip file. Make sure you are logged in with a Microsoft 365 Home or organizational account.
5. Open `config.cfg` from the zip file and copy `clientId`
6. Paste it into the suspiciously similar `config.cfg` that lives in your cloned repo
7. You are good to go! `python main.py` should spawn the interface for testing all the calendar API functions!

## Project structure
- `main.py` runs the test interface and demonstrates a sample application
- `graph.py` provides all the library functionality that works with the Microsoft Graph API
- `config.cfg` is credentials and configuration for the API connection
- `requirements.txt` contains all python packages this project needs to run (there are a lot)