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
- Dropdowns
    - Hosting data for **VLOOKUP**, **FILTER** and **INDEX / MATCH** for references within automation production.


## Automations

### Guest Activity Itinerary Generator

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


### Daily Lunch Form

#### Process

##### Extracting & Transforming

- 01 | Extract the activity time, guide and **VLOOKUP** output value for lunch type for each reservation's daily activity.

Separate AM and PM activities into arrangements of 4 columns each. Each column extracting relavent AM or PM activity data. Code blocks below are only copies of the AM Activity Schedules. PM Activity Schedule code can be found in relevant formatting between columns **'Lunch Form'!$E:H**.

> AM Activity Lunches | Tab **'Lunch Form'!**

**'Lunch Form'!$A5:A52**
``` excel-formula
=LET(                                                                   /* One formula copied down column A */
AM_activity_time, $B5,                                                  /* Row's activity time */
IF(AM_activity_time <> "", "AM","")                                     /* Categorise Activity Time for future filtering */
)
```

**'Lunch Form'!$B5:B52**
``` excel-formula
=Let(
AM_activity_time, 'Schedule - Reservations'!$P$8:$P,                    /* Activity Time column for all AM Activities */
dates, 'Schedule - Reservations'!$B$8:$B,                               /* Date column for all reservations */
chosen_date, $B$3,                                                      /* Date chosen for daily lunch form input data */ 

FILTER(
   IF( {AM_activity_time}= "","",                                       /* If Activity Time is blank, return blank */
      {AM_activity_time}),                                              /* Extract all AM Activity Time entries /*
      dates = chosen_date,                                              /* Filter where Chosen_Date = Date */
      dates<>"")                                                        /* Filter where Date Not Blank */
)
```

**'Lunch Form'!$C5:C52**
``` excel-formula
=Let(
AM_guides, 'Schedule - Reservations'!$Q$8:$Q,                           /* Assigned Guide for each reservation on AM Activities */
dates, 'Schedule - Reservations'!$B$8:$B,                               /* Date column for all reservations */
chosen_date, $B$3,                                                      /* Date chosen for daily lunch form input data */ 

FILTER(
    IF( {AM_guides}= "", "",                                            /* If AM Guide is Blank, Return Blank */
    {AM_guides}),                                                       /* Extract All Guides assigned to AM Activites For Each Reservation */
    dates = chosen_date,                                                /* Filter Where Chosen_Date = Date */
    dates<>"")                                                          /* Filter Where Date Not Blank */
)                                                                       
```                                                                     

**'Lunch Form'!$D5:D52**
``` excel-formula
=Let(
AM_guides, 'Schedule - Reservations'!$Q$8:$Q,                           /* Assigned Guide for each reservation on AM Activities */
AM_activities, 'Schedule - Reservations'!$O$8:$O,                       /* AM Activity for each reservation across all dates */
activities_references, Dropdowns!$B$3:$E$69,                            /* References table for activity VLOOKUP values */
dates, 'Schedule - Reservations'!$B$8:$B,                               /* Date column for all reservations */
chosen_date, $B$3,                                                      /* Date chosen for daily lunch form input data */ 

FILTER(
   IFERROR(IF( 
        {AM_guides}= "", "",                                            /* If AM Guide is blank, Return blank */
        { VLOOKUP( AM_activities, activities_references, 4, FALSE)})    /* Lookup 4th column balue ("Lunch Form Naming") ror AM Activity */
        ,"None"),                                                       /* If Error, I.E no Activity Value, return "None" */
   dates = chosen_date,                                                 /* Filter where Chosen_Date = Date */
   dates <>""                                                           /* Filter where Date not blank */
  )
)
```

- 02 | Extract unique pairs of guide and lunch types and attach the earliest associated activity time. 

**'Lunch Form'!$A53:A76**
``` excel-formula
=LET(                                                                   /* One formula copied down column A */
AM_activity_time, $B5,                                                  /* Row's activity time */
IF(AM_activity_time <> "", "AM","")                                     /* Categorise Activity Time for future filtering */
)
```

**'Lunch Form'!$B53:B76**
``` excel-formula
=Let(
unique_AM_guides, $C$53:$C$76,                                          /* Guide Name entry from unique AM Guide-Lunch list for row i */
AM_guide_lunches, $D$53:$D$76,                                          /* Lunch Type entry from unique AM Guide-Lunch list for row i */
AM_activity_times, $B$5:$B$52,                                          /* Spilled AM Activity Times */
AM_guides, $C$5:$C$52,                                                  /* Spilled AM Guides */
AM_lunches, $D$5:$D$52,                                                 /* Spilled AM Lunch Types */


/* For each guide lunch pair (g,l) from the two query lists, compute the result.  */
/* MAP pairs g_i with l_i (no cross-join).*/

MAP(unique_AM_guides, AM_guide_lunches,                               /* Two lists To itterate through each row */
  LAMBDA(g, l,                                                        /* Array value name g = guide, l = lunch for each row */
    IFERROR(
           MIN( FILTER(AM_activity_times,                             /* Extract minimum AM_activity_time value among array rows
                      TRIM(AM_guides) = TRIM(g),                      /* TRIM guides against stray spaces. Filter where AM_guides==g */
                      AM_lunches = l,                                 /* Fiter where AM_lunches==l */
                      ISNUMBER(AM_activity_times)                     /* ISNUMBER ensures time is numeric. */
                     )
              ),
      "" )                                                            /* If No Matching rows, return blank */                                                      
      )
  )
)
```

**'Lunch Form'!$C53:D76**
``` excel-formula
=Let(
AM_activity_lunches, $C5:$D52,                                              /* AM Activity Lunch pairs */
AM_lunches, $D$5:$D$52,                                                     /* AM Lunches */

UNIQUE(                                                                     /* Extract unique unique AM Activity Lunch pairs */
  filter( AM_activity_lunches,
       AM_lunches <> "none"                                                 /* Filter where AM_Lunches does not = "none"
     )
  )
)
```

- 03 | Assign required lunch prepration time for each guided activity

**'Lunch Form'!$A77:A99**
``` excel-formula
=LET(                                                                   /* One formula copied down column A */
AM_activity_time, $B5,                                                  /* Row's activity time */
IF(AM_activity_time <> "", "AM","")                                     /* Categorise Activity Time for future filtering */
)
```

**'Lunch Form'!$B77:B99**
``` excel-formula
=Let(
lunch_type, $D77,                                                       /* Row associated Lunch Type */
AM_activity_time,$B53,                                                  /* Row associated AM Activity Time */
lunch_times, FILTER(Dropdowns!$J$3:$L, Dropdowns!$J$3:$J <>""),         /* Reference table for required lunch preparation time */

IF(lunch_type <> "",                                                    
    IF( lunch_type =Dropdowns!$J$4,                                     /* If Lunch Type = Dropdowns!$J$4 "Airport Snacks" */
            MOD( AM_activity_time -TIME(2,0,0),1),                      /* Output AM Activity Time subtracted by two hours */
            VLOOKUP(lunch_type,lunch_times,2,FALSE)                     /* Otheriwse, output second column value from VLOOKUP for AM lunch preparation time */   
        ),
    "")                                                                 /* If lunch type is blank, return blank */
)
```

**'Lunch Form'!$C77:D99**
=Let(
AM_lunches_form, $C$53:$D$76,                                           /* Previous unique filter guide lunch pairs */
ARRAYFORMULA(AM_lunches_form)                                           /* Output data for Column B reference */
)

#### Loading
- 04 | Present lunch form data for schedule requirements

**'Lunch Form'!$A102:D**
``` excel-formula
=Let(
AM_Lunches, $A$77:$D$99,                                                            /* Filtered AM Lunches */
PM_Lunches, $E$77:$H$99,                                                            /* Filtered PM Lunches */
AM_Lunch_Types, $D$77:$D$99,                                                        /* Lunch Types of filtered AM Lunches */
AM_Lunch_Times, $B$77:$B$99,                                                        /* Lunch Preparation Times of filtered AM Lunches */
PM_Lunch_Types, $H$77:$H$99,                                                        /* Lunch Types of filtered PM Lunches */
PM_Lunch_Times, $F$77:$F$99,                                                        /* Lunch Preparation Times of filtered PM Lunches */

SORT(
  UNIQUE(
    FILTER(
      {AM_Lunches; PM_Lunches},                                                     /* Extract unique AM and PM Lunches array values into singular output */
      {
        (LOWER(TRIM(SUBSTITUTE( AM_Lunch_Types ,CHAR(160)," "))) <> "none") *       /* Filter AM type <> none */
        (LOWER(TRIM(SUBSTITUTE( AM_Lunch_Times ,CHAR(160)," "))) <> "n/a");         /* Filter AM time <> n/a */
        (LOWER(TRIM(SUBSTITUTE( PM_Lunch_Types ,CHAR(160)," "))) <> "none") *       /* Filter PM type <> none */
        (LOWER(TRIM(SUBSTITUTE( PM_Lunch_Times ,CHAR(160)," "))) <> "n/a")          /* Filter PM time <> n/a */
      }
    )
  ),
  2, TRUE                                                                           /* Sort output by column 2 (lunch preparation time) in asecnding order */
 )
)

```

**'Lunch Form'!$E102:E**
``` excel-formula

=LET(                                                                               /* One formula copied down column E102:E */
activity_time_slot, $A102,                                                          /* Activity time slot AM or PM */
guide, $C102,                                                                       /* Guide Slot */
reservations, 'Schedule - Reservations'!$D$8:$D,                                    /* All guest reservations */
dates, 'Schedule - Reservations'!$B$8:$B,                                           /* Date column for all reservations */
chosen_date, $B$3,                                                                  /* Date chosen for daily lunch form input data */ 
AM_guides, 'Schedule - Reservations'!$Q$8:$Q,                                       /* Assigned Guide for each reservation on AM Activities */
PM_guides, 'Schedule - Reservations'!$V$8:$V,                                       /* Assigned Guide for each reservation on PM Activities */

IF(activity_time_slot = "AM",                                                       /* If lunch is for AM activity */
    IF(guide <>"",                                                                  /* If guide is not blank */
        TEXTJOIN(", ", TRUE,                                                        /* Textjoin all reservations by "," */
            UNIQUE(
            IFERROR(FILTER(
                ARRAYFORMULA(
                REGEXREPLACE( reservations, ".*(?: : | - )", "")                    /* Replace all ":" "-" values in reservations with blank for visual formatting */
                ), 
                ( dates = chosen_date ) *                                           /* Filter reservations for where dates = chosen_date */
                ( AM_guides = guide )                                               /* Filter reservations for where AM_guides = guide */
            ),"Error. No Matching Reservations"),
            )
        ), 
    ""),                                                                            /* If lunch is <> AM activity, therefore == PM activity */
    IF(guide <>"",                                                                  /* If guide is not blank */
        TEXTJOIN(", ", TRUE,                                                        /* Textjoin all reservations by "," */
            UNIQUE(
            FILTER(
                ARRAYFORMULA(
                REGEXREPLACE( reservations, ".*(?: : | - )", "")                    /* Replace all ":" "-" values in reservations with blank for visual formatting */
                ),
                ( dates = chosen_date) *                                            /* Filter reservations for where dates = chosen_date */
                ( PM_guides = guide )                                               /* Filter reservations for where PM_guides = guide */
            )
            )
        ), 
    "")
)
)
```

### Daily Activity Schedule

#### Process

##### Filtering
