# FlightMemoryExporter

Export your full flight history from [flightmemory.com](https://www.flightmemory.com/) to an Excel spreadsheet.

Logs into your FlightMemory account, scrapes every page of your logged flights, and writes them out to a `.xlsx` file with columns for date, flight number, route, times, airline, aircraft, seat, and more.

![image](https://github.com/TobiasUr/FlightMemoryExporter/assets/68461330/c39bf5c8-19a4-40c1-80be-4def006241d3)

## Quick start (Windows, no Python required)

Download the latest build from the [Releases](https://github.com/TobiasUr/FlightMemoryExporter/releases/latest) page and run it. Enter your FlightMemory username and password, choose where to save the file, and wait — your flights will be written to `Flights.xlsx` in the location you chose.

> You'll need [Google Chrome](https://www.google.com/chrome/) installed, since the app drives a Chrome browser to log in and scrape your data.

## Setup (run from source)

```bash
git clone https://github.com/TobiasUr/FlightMemoryExporter.git
cd FlightMemoryExporter
pip install beautifulsoup4 openpyxl selenium chromedriver-autoinstaller
```

See the [wiki](https://github.com/TobiasUr/FlightMemoryExporter/wiki/Compiling-from-code) for full instructions on compiling from source.

## Usage

```bash
python FlightMemoryExporter.py
```

This opens a small window where you:
1. Enter your FlightMemory username and password
2. Click **OK** and choose a destination for the spreadsheet
3. Wait while the app signs in, pages through your entire flight history, and compiles it

The resulting spreadsheet includes: Date, Flight number, From, To, Dep time, Arr time, Duration, Airline, Aircraft, Registration, Seat number, Seat type, Flight class, Flight reason, Note, Dep_id, Arr_id, Airline_id, and Aircraft_id.

## Notes

- Chromedriver is installed and managed automatically via `chromedriver_autoinstaller` — no manual driver setup needed.
- Your username and password are only used to fill in FlightMemory's own login form in the automated browser session; they aren't saved or sent anywhere else.
- If the login fails, double-check your credentials — a popup will let you know if sign-in didn't succeed.

## Support

If you would like to support this and my other projects, you can [![Buy Me a Coffee](https://img.shields.io/badge/Buy%20Me%20a%20Coffee-orange?logo=buy-me-a-coffee)](https://www.buymeacoffee.com/tobiasurbanek)
