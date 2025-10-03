# TPL — Tweedsmuir Park Lodge

Three projects within the Tweedsmuir Park Lodge (TPL) system were produced. Each catering to a manual task requiring upwards of 30 minutes of daily duties for lodge or head office staff members. Leading up to and during a 16 week summer tourism season, these following projects are predicted to have saved 168 hours of manual labour tasks.


## File
With permisson, a copy of Tweedsmuir Park Lodge operations Google Sheet file has been produced to showcase the automation tasks presented below. 

### Access
![Link File Here](sdas)

#### Data
Names of guides and guests have been modified to provide privacy of those involved.

#### Tabs

- Activity Schedule
    - Prior schedule production from TPL, produced with Google Sheets logic from import CSV files
- References
    - Hosting data for **VLOOKUP**, **FILTER** and **INDEX / MATCH** for references within automation production.


## Automations

### Automated Guest Activity Itinerary Generator

Leading up to the tourism season, head office team members produce scheduled itinerary plans for each of the 292 reservations 1 month prior to arrival. Automation of this task reduced labour by outputing desired visual format in an transferable email template for each guest from the **Activity Schedule**. Using up the current activity itinerary information from guide scheduling.

#### Process

##### Filtering

- 01 | Produce an output array of all reservations present on chosen date.

``` excel-formula
= LET (
    reservations, 'Schedule - Reservations'!$D$8:$D,
    date, 'Schedule - Reservations'!$B$8:$B,
    chosen_date, $A3,

    UNIQUE(
        FILTER(
            {reservations},date = chosendate, date <> ""
            )
    )
)
```
- 02 | Produce Dropdown selection from the output array as Selected Cell Range to dicate chosen Reservation.

- 03 | Produce an array output of designated activities for chosen reservation.

``` excel-formula

=LET(
    date, 'Schedule - Reservations'!$B$8:$B,
    week_day, 'Schedule - Reservations'!$C$8:$C,
    morning_activity, 'Schedule - Reservations'!$O$8:$O,
    afternoon_activity, 'Schedule - Reservations'!$T$8:$T,
    reservations, 'Schedule - Reservations'!$D$8:$D,
    chosen_reservation, $A$27,

    Filter(
        {date, week_day, morning_activity, afternoon_activity}, reservations = chosen_reservation
    )
)
```

##### Index Matching

- 04 | Match internal operation activity type to categorised consumer language from References.

``` excel-formula

=LET(
  activities,    C28:C41,                                    
  keywords,      FILTER(Dropdowns!B$3:B,  LEN(Dropdowns!B$3:B)),                 
  categories,    FILTER(Dropdowns!F$3:F,  LEN(Dropdowns!B$3:B)),                 
  ARRAYFORMULA(
    IF(LEN(activities)=0,                                   
       "",
       IFNA(
         BYROW(activities,
           LAMBDA(act,
             INDEX( categories,
                    MATCH(TRUE, REGEXMATCH(LOWER(act), LOWER(keywords)), 0)
             )
           )
         ),
         "Uncategorized"                                   
    )
  )
)
)
```

#### Summary

The excel code provided above successfully automates a useable email template to communicate individual reservation activity interaries with up to date guided schedules.

