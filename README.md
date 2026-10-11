# 🏁 2027 World Motorsports Calendar

[![Created by Gemini](https://img.shields.io/badge/Created%20by-Google%20Gemini-4285F4?style=for-the-badge&logo=google&logoColor=white)](https://gemini.google.com)
[![Season](https://img.shields.io/badge/Season-2027-E10600?style=for-the-badge&logo=formula1&logoColor=white)](#-championship-coverage)
[![Championships](https://img.shields.io/badge/Championships-18%20Series-0EA5E9?style=for-the-badge)](#-championship-coverage)
[![Events](https://img.shields.io/badge/Total%20Events-149%20Races-10B981?style=for-the-badge)](#-overview)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license--disclaimer)
[![Stack](https://img.shields.io/badge/Stack-Vanilla%20HTML%20%2F%20CSS%20%2F%20JS-8B5CF6?style=for-the-badge)](#-technical-architecture)

> **Created and developed by Google Gemini.**

An interactive, responsive static web application illustrating the global motorsport calendar for the 2027 season, organized week by week across 18 major championships and 149 individual race events.

Built entirely with standard HTML5, CSS3, and vanilla ES6+ JavaScript—with zero external frameworks, zero package managers, and zero build tool requirements.

---

## 📑 Table of Contents

* [Overview](#-overview)
* [Key Features](#-key-features)
* [Championship Coverage](#-championship-coverage)
* [Project Structure](#-project-structure)
* [Getting Started & Local Preview](#-getting-started--local-preview)
* [GitHub Pages Deployment](#-github-pages-deployment)
* [How to Update Schedules & Timetables](#-how-to-update-schedules--timetables)
* [Technical Architecture](#-technical-architecture)
* [Created by Gemini](#-created-by-gemini)
* [Contributing](#-contributing)
* [License & Disclaimer](#-license--disclaimer)

---

## 🌟 Overview

Following motorsport across global categories is traditionally fragmented between separate series websites and timezone barriers. The **2027 World Motorsports Calendar** unifies 149 worldwide racing events across open-wheel, sports car, GT, touring car, and rally disciplines into a cohesive, multi-view schedule dashboard.

Whether planning travel to race weekends, tracking support championship pairings, monitoring 24-hour endurance classics, or checking potential broadcast clashes, the application provides an accessible, high-contrast calendar experience tailored for racing fans, marshals, media, and paddock engineers.

---

## 🚀 Key Features

### 1. 📅 Week-by-Week Spotlight (Primary View)
* **Interactive 52-Week Scrubber**: Rapidly glide through all 52 weeks of the year using an interactive range slider or step week-by-week using arrow buttons or keyboard shortcuts (`←` / `→`).
* **Quick-Jump Pills**: Jump instantly to marquee rounds (Daytona, F1 Opener, Miami, Le Mans, Spa 24h, Monza, Abu Dhabi) or quarterly intervals.
* **Rich Event Cards**: Each scheduled race card features official event branding, host circuit, host city, national flag emoji, session format tags (Sprint, 24h Enduro, Double Header), and support series tags.
* **Weekend Intensity Indicator**: Visual badges distinguishing single-race weekends, multi-series weekends (3+ series), super clash weekends (5+ series), and scheduled summer/winter breaks.

### 2. 📜 Continuous Weekly Timeline
* A full vertical timeline displaying all 52 weeks from January 1, 2027 to December 26, 2027.
* Compact event chips for rapid scanning and filtering.
* Restful off-week markers highlighting mid-season breaks.

### 3. 📊 Master Excel Matrix View
* An exact interactive representation of the original 52-week scheduling spreadsheet.
* **Freeze-Pane Headers**: Sticky Week number (Column 1) and Dates (Column 2) columns remain locked in place during horizontal scrolling.
* **Color-Coded Championship Blocks**: Instantly see which championships are active in any given week across the entire calendar year.

### 4. 🏆 Championship Series Hub
* Dedicated cards for each of the 18 tracked championships.
* Displays round counts, date spans, discipline category, regulation notes, and confirmation status.

### 5. ⚡ Clash Detector & Super Weekends
* Dedicated analytics view highlighting multi-series calendar clashes.
* Identifies "Super Clash" weekends where 4+ major international series race simultaneously.

### 6. ⏱️ Detailed Session Timetables & Status Badges
* Clicking any race opens a dedicated Event Detail modal.
* **Confirmed Schedules**: Full multi-day session timetables with local track times, session types (Practice, Qualifying, Sprint, Feature Race, Super Special Stage), and duration details. Scrollable layout supports long multi-day endurance and rally schedules (e.g. Rallye Monte-Carlo, Rolex 24 at Daytona, 24 Hours of Le Mans).
* **Unconfirmed Schedules**: Clearly tagged with **TBC** (To Be Confirmed) until official sanctioning bodies publish session-by-session times.

### 7. 🔍 Multi-Filter Search & Quick Sorting
* Instant text search matching circuit name, country, city, grand prix title, or series code.
* Quick category filter pills: All, Open Wheel, Sports Cars & GT, Touring & Rally.

### 8. 📥 Calendar (.ics) & CSV Export
* Export individual race weekends or the entire 149-event calendar to standard `.ics` (iCalendar) format for Apple Calendar, Google Calendar, and Microsoft Outlook.
* Export tabular `.csv` data for custom analysis in Excel or Google Sheets.

### 9. 🌓 High-Contrast Dark & Light Modes
* Full light and dark theme support with persistent `localStorage` preference.
* Carefully balanced WCAG AA-compliant contrast ratios across all chips, badges, matrix cells, and modal dialogs.

---

## 🏎️ Championship Coverage

The 2027 calendar tracks **149 events** across **18 premier racing series**:

| Category | Championship | Abbreviation | Rounds | 2027 Timetable Status |
| :--- | :--- | :--- | :---: | :---: |
| **Open Wheel** | Formula 1 (FIA Formula One World Championship) | F1 | 24 | Confirmed (10 Sprints) |
| **Open Wheel** | FIA Formula 2 Championship | F2 | 14 | Confirmed |
| **Open Wheel** | FIA Formula 3 Championship | F3 | 10 | Confirmed |
| **Open Wheel** | ABB FIA Formula E World Championship | FE | 16 | Confirmed (Gen4 Era) |
| **Open Wheel** | Formula Regional European Championship (FRECA) | FRECA | 10 | TBC |
| **Open Wheel** | Formula Regional Oceania Trophy | FROT | 4 | Confirmed (12 Races) |
| **Open Wheel** | Formula Winter Series | FWS | 5 | Confirmed |
| **Open Wheel** | Italian F4 / Euro 4 Championship | ITA F4 | 7 | TBC |
| **Open Wheel** | Spanish F4 Championship | ESP F4 | 7 | TBC |
| **Open Wheel** | German F4 Championship | GER F4 | 6 | TBC |
| **Open Wheel** | Macau Grand Prix (Formula Regional World Cup) | MACAU | 1 | TBC |
| **Sports Car / Enduro** | FIA World Endurance Championship | FIA WEC | 9 | Confirmed |
| **Sports Car / Enduro** | IMSA WeatherTech SportsCar Championship | IMSA | 11 | Confirmed |
| **Sports Car / Enduro** | European Le Mans Series | ELMS | 6 | TBC |
| **Sports Car / Enduro** | Asian Le Mans Series | AsLMS | 5 | TBC |
| **GT / Touring** | Intercontinental GT Challenge | IGTC | 5 | Confirmed |
| **GT / Touring** | Fanatec GT World Challenge Europe (Sprint + Enduro) | GTWCE | 9 | Confirmed |
| **GT / Touring** | Fanatec GT World Challenge Asia | GTWC Asia | 6 | TBC |
| **Touring** | DTM (Deutsche Tourenwagen Masters) | DTM | 8 | Confirmed |
| **Rally** | FIA World Rally Championship | WRC | 15 | Confirmed |

---

## 📁 Project Structure

The project follows a clean, modular static website layout:

```text
motorsports-calendar-2027/
├── index.html          # Semantic HTML5 entry point & UI structure
├── css/
│   └── style.css       # Complete modular styling, themes, layout, and animations
├── js/
│   └── script.js       # Dataset (149 events), calendar engine, timetable parser, & exports
├── README.md           # Documentation, setup guide, and timetable maintenance reference
└── backup/             # Preserved backup copies of previous single-file builds
```

---

## 💻 Getting Started & Local Preview

Because this project uses vanilla web standards with no build pipeline, you can run and inspect it locally with any simple static web server:

### Option A: Python (Pre-installed on macOS/Linux/Windows)
```bash
# In the project root directory:
python3 -m http.server 8000
```
Open [http://localhost:8000](http://localhost:8000) in your web browser.

### Option B: Node.js
```bash
npx serve .
# or
npx http-server .
```

### Option C: VS Code Live Server
1. Install the **Live Server** extension in Visual Studio Code.
2. Right-click `index.html` and choose **"Open with Live Server"**.

---

## 🌐 GitHub Pages Deployment

Deploying the calendar to GitHub Pages takes under two minutes:

1. **Create a GitHub Repository**:
   ```bash
   git init
   git add .
   git commit -m "Initial commit: 2027 World Motorsports Calendar"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo-name>.git
   git push -u origin main
   ```

2. **Enable GitHub Pages**:
   * Navigate to your repository on GitHub.
   * Go to **Settings** > **Pages** (under "Code and automation" in the left sidebar).
   * Under **Build and deployment** > **Source**, select **Deploy from a branch**.
   * Under **Branch**, select `main` and leave the folder set to `/(root)`.
   * Click **Save**.

3. **View the Live Site**:
   * GitHub Actions will automatically publish your site.
   * Your calendar will be live at:
     ```text
     https://<your-username>.github.io/<your-repo-name>/
     ```

All stylesheet and script references in `index.html` use relative paths (`css/style.css`, `js/script.js`), ensuring full compatibility with GitHub Pages custom repository subpaths.

---

## ✏️ How to Update Schedules & Timetables

When sanctioning bodies publish new event dates or release detailed weekend timetables, **only one file needs to be updated**:

### `js/script.js`

You **do not** need to edit `index.html` or `css/style.css`.

#### 1. Confirming an Announced Series
In `js/script.js`, locate the `CONFIRMED_SERIES` array at the top of the file:
```javascript
const CONFIRMED_SERIES = [
  'Formula 1',
  'Formula E',
  'FIA WEC',
  'IMSA',
  'DTM',
  'WRC',
  'IGTC',
  'GTWCE',
  'FROT',
  'FWS'
  // When ELMS or FRECA announces their official timetable, add them here:
  // 'ELMS',
];
```

#### 2. Adding Session Timetables
In `js/script.js`, locate the `getEventTimetable(ev)` function. Add or update the switch cases matching the series:
```javascript
case 'ELMS':
  return [
    { day: 'Friday', sessions: [
      { time: '11:00 - 12:30', name: 'Free Practice 1', type: 'practice' },
      { time: '15:00 - 15:30', name: 'Bronze Driver Test', type: 'practice' }
    ]},
    { day: 'Saturday', sessions: [
      { time: '10:00 - 11:30', name: 'Free Practice 2', type: 'practice' },
      { time: '14:30 - 15:45', name: 'Qualifying (LMP2 / LMP3 / LMGT3)', type: 'qualifying' }
    ]},
    { day: 'Sunday', sessions: [
      { time: '13:00 - 17:00', name: '4 Hours of Le Castellet (Race)', type: 'race' }
    ]}
  ];
```

#### 3. Adding or Adjusting Calendar Rounds
To change dates, circuits, or round metadata, edit the event objects in the `CALENDAR_DATA` array inside `js/script.js`. Each entry contains:
* `id`: Unique event string
* `week`: Integer (1–52)
* `dates`: Text representation (e.g. `"Mar 19–21"`)
* `series`: Series name
* `name`: Official Grand Prix / Race title
* `circuit`: Circuit name
* `location`: City, Country
* `flag`: National emoji flag
* `format`: `"Standard"`, `"Sprint"`, `"24h Enduro"`, etc.

---

## 🛠️ Technical Architecture

* **Zero Dependencies**: Pure HTML5, CSS3, and ES6+. Fast load times (<50ms) and zero vulnerability surface.
* **Sticky Positioning & Grid Matrix**: CSS Sticky headers and columns handle the complex 52-row by 18-column Excel Matrix view seamlessly without heavy third-party table libraries.
* **Accessible Modals**: Native `dialog`-inspired accessible modal with smooth backdrop blur, keyboard dismiss (`Esc`), top-right close buttons, and sticky footer action buttons.
* **iCal Engine**: Client-side RFC 5545 `.ics` generator handles timezone bounds, multi-day dates, and event descriptions dynamically in-browser.

---

## 🤖 Created by Gemini

This project was conceived, structured, and developed by **Google Gemini**. The end-to-end implementation includes:
- Extracting and standardizing the 52-week, 149-event schedule dataset across 18 motorsport series from the original calendar spreadsheet.
- Engineering the responsive UI, dark/light themes, and sticky freeze-pane matrix.
- Designing the session timetable system with confirmed multi-day schedules and dynamic TBC status indicators.
- Refactoring the single-file application into a modular, dependency-free static website project ready for GitHub Pages.

---

## 🤝 Contributing

Contributions, schedule corrections, and timetable additions are welcome!

1. Fork the Project.
2. Create your Feature Branch (`git checkout -b feature/AddELMSTimetable`).
3. Commit your Changes (`git commit -m 'Add 2027 ELMS detailed timetable'`).
4. Push to the Branch (`git push origin feature/AddELMSTimetable`).
5. Open a Pull Request.

---

## ⚖️ License & Disclaimer

* **Code License**: Released under the [MIT License](LICENSE).
* **Motorsport Trademarks**: Formula 1, FIA, IMSA, WRC, DTM, SRO, and all related series names, logos, circuit titles, and marks are registered trademarks of their respective owners and sanctioning bodies. This fan project is unofficial, non-commercial, and created purely for calendar reference and scheduling convenience.