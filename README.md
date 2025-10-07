# BCHS-TPL-Scheduling-Automation
During my time with **Bella Coola Heli Sports (BCHS)** and **Tweedsmuir Park Lodge (TPL)**, I identified opportunities to streamline internal operations by building **lightweight, spreadsheet-driven data pipelines** to eliminate repetitive manual work.

Upon my first innovative success with Daily Live Schedules, I was further tasked to develop additional data pipelines, from generating guest itineraries to coordinating helicopter capacity across multiple remote lodges. Each tool was designed to transform raw operational data into actionable, human-readable outputs for frontline teams.


## Project Overview
This project involved a series of **Extract–Transform–Load (ETL)** automations designed specifically for the hospitality and adventure tourism industry, with a focus on:

- **Extracting** structured data from daily guest activity schedules  
- **Transforming** it to meet downstream operational needs (kitchen prep, guide planning, helicopter logistics)  
- **Automating** the generation of usable output formats for on-the-ground teams (printable reports, email templates, etc.)


### Purpose 
- ✅ **Efficiency**: Eliminated dozens of hours of weekly manual data entry
- ✅ **Clarity**: Ensured staff received clear, consistent, and up-to-date information
- ✅ **Repeatability**: Enabled reusability across seasons, adaptable for new lodge operations

## 📦 Featured Automations

### TPL — Tweedsmuir Park Lodge

- **Automated Guest Activity Itinerary Generator**
  > Generates clean, email-ready itineraries from guest booking schedules,removing the need for manual formatting and lookup.
  - Saves **~24 hours** of manual task per season

- **Live Daily Guide & Guest Activity Schedule**
  > Consolidates guide assignments and guest bookings into a clear daily schedule for lodge managers and front desk staff.
- Saves **~46 hours** of manual task per season

- **Lunch Requirements Generator**
  > Uses formulas like **VLOOKUP**, **FILTER**, **MAP**, **LAMBDA** and **ARRAYFORMULA** to calculate total lunch package requirements by guest, guide, and activity—outputting a clean prep list for the kitchen team.
- Saves **~46 hours** of manual task per season
---

### BCHS — Bella Coola Heli Sports (Multi-Lodge Operations)

- **Condensed Heli-Skiing Package Tracker & Helicopter Scheduling View**
  > Aggregates guest counts, package types (e.g. 3-day Group vs. 5-day Private heli-skiing), and lodge transfers across multiple heli lodges. Outputs a condensed, scheduling-centric snapshot that helps management teams and senior operation allocate helicopters efficiently per lodge.


## 🛠️ Tools & Methods

- **Google Sheets**
  - **VLOOKUP**, **FILTER**, **ARRAYFORMULA**, **IMPORTRANGE** , **MAP**, **LAMBDA**, **MOD**, **INDEX**, **MATCH**, **REGEXMATCH**, 
  - SQL-like logic and named ranges
  - Conditional formatting for at-a-glance reporting
- **Data Design**
  - Structured guest & activity logs as base tables
  - Output dashboards segmented by audience (chefs, guides, managers)


## 👥 Audience & Use Cases

| Team             | Output Used For                                     |
|------------------|-----------------------------------------------------|
| **Front Desk**   | Guest itinerary distribution & check-ins            |
| **Guides**       | Daily activity briefings and scheduling             |
| **Kitchen Staff**| Meal prep based on activity & guest requirements    |
| **Ops Managers** | Helicopter usage planning across 4–5 remote lodges  |

---

## 🗂️ Repository Structure
```
TPL-BCHS-Guide-Schedule-Automation/
├── README.md
├── /images/
│   ├── TPL-daily-schedule.png
│   ├── BCHS-weekly-snapshot.png
│   └── ...
├── /docs/
│   └── pipeline_overview.md
└── /sheets/
    └── public_demo_link.txt
```



---

## 🌱 Next Steps

These tools were originally built in Google Sheets, but many could be **migrated to Python, Dash, or Looker Studio** for greater scalability, backend integration, or API-driven scheduling workflows.

