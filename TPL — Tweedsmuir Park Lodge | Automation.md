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
Leading up to the tourism season, head office prepares itinerary plans for each of ~292 reservations one month prior to arrival. Automating this task outputs the required visual format as a transferable email template directly from the Activity Schedule, using the current guide scheduling. This reduced the work from ~30 minutes to a few minutes per guest.

#### Process

> Source Information | Sheet **'Schedule - Reservations'!**
> Pipeline Output | Sheet **'Guest Activity Itinerary'!**

##### 01 | Extract Source Data

- **'Guest Activity Itinerary'!B4:B30 - Reservations List**
    - Produce an output array of all reservations present on chosen date.
``` excel-formula
= LET (
    reservations, 'Schedule - Reservations'!$D$8:$D,                            /* All guest reservations */
    dates, 'Schedule - Reservations'!$B$8:$B,                                    /* Date column for all reservations */
    chosen_date, $A3,                                                           /* Date chosen for observe guest reservation interary */ 

    UNIQUE(
        FILTER(                                                                 /* Filter unique reservations into singular output */
            {reservations},dates = chosen_date, date <> ""
            )
    )
)
```

- **'Guest Activity Itinerary'!A27:D27 - Reservation Selector**
    - Produce Dropdown selection from the output array as Selected Cell Range to dicate chosen Reservation.
    - Data validation (list from a range): ='Guest Activity Itinerary'!$B$4:$B$26


- **'Guest Activity Itinerary'!A28:D47 - Itinerary Filtering**
    - Produce an array output of designated activities for chosen reservation from dates applicable

``` excel-formula

=LET(
    dates, 'Schedule - Reservations'!$B$8:$B,                                   /* Date column for all reservations */
    week_days, 'Schedule - Reservations'!$C$8:$C,                               /* Weekday column */
    AM_activities, 'Schedule - Reservations'!$O$8:$O,                           /* AM activities */ 
    PM_activities, 'Schedule - Reservations'!$T$8:$T,                           /* PM activities */
    reservations, 'Schedule - Reservations'!$D$8:$D,                            /* All guest reservations */
    chosen_reservation, $A$27,                                                  /* Reservation chosen present during chosen_date in above spill data */

    Filter(
        {dates, 
        week_days, 
        AM_activities, 
        PM_activities}, (reservations=chosen_res) * (dates<>"")                   /* Filter for selected column values from chosen reservation */
    )
)
```

##### 02 | Transform & Load
- Final data output into useable formatting for guest communication 

- **'Guest Activity Itinerary'!A50:A**
    - Array output of all dates reservation is present

``` excel-formula
=Let(
guest_dates, $A$28:$A$47,

ARRAYFORMULA(guest_dates)
)
```

- **'Guest Activity Itinerary'!B50:B**
    - Array output of all week days reservation is present

``` excel-formula
=Let(
weekdays, $B$28:$B$47,

ARRAYFORMULA(weekdays)
)
```

- **'Guest Activity Itinerary'!C50:D - Transform Activity Names**
    - Match internal operation activity type to categorised consumer language from References.

``` excel-formula

=LET(
  reservation_activities,    C28:C41,                                                       /* Reservation's activities */
  activities_raw,            FILTER(Dropdowns!B$3:B,  LEN(Dropdowns!B$3:B)),                /* Internal operation activity names (from References Table)*/ 
  activities_guest_names,    FILTER(Dropdowns!F$3:F,  LEN(Dropdowns!B$3:B)),                /* Approved guest named activities (from References Table)*/

  ARRAYFORMULA(
    IF(LEN(reservation_activities)=0,                                                       /* If activity length 0, return blank */      
       "",
       IFNA(
         BYROW(reservation_activities,                                                      /* Apply following function to each row of reservation_activities data */
           LAMBDA(act,                                                                      /* function */
             INDEX( activities_guest_names,                                             
                    MATCH(TRUE, REGEXMATCH(LOWER(act), LOWER(activities_raw)), 0)           /* Case-insensitive regex, contains match, first matching pattern in activities_guest_names returned from References Table */ 
             )
           )
         ),
         "Uncategorized"                                                                    /* If no match, return "Uncategorized" */
    )
  )
)
)
```

#### Summary

The excel code provided above successfully automates a useable email template to communicate individual reservation activity interaries with up to date guided schedules.





### Daily Lunch Form
Each day of operations at Tweedsmuir Park Lodge requires a variation of packed lunches and coolers for guest activities. Previously, receptionists read the  'Schedule - Reservations'ssheet and produced a typed/written lunch form—about 30 minutes per day for the kitchen to produce. 

The automation reduces the task to ~5 minutes per day. The procedure extracts guides, guest reservations and guests counts, transforms activity types in lunch types, and sorts by required preparation time for the chosen date. The only remaining manual step is adding dietary preferences, which aren’t available in the data pipeline.


#### Process

##### 01 | Extract Source Data

- Extract the activity time, guide and **VLOOKUP** output value for lunch type for each reservation's daily activity.

Separate AM and PM activities into arrangements of 4 columns each. Each column extracting relavent AM or PM activity data. Code blocks below are only copies of the AM Activity Schedules. PM Activity Schedule code can be found in relevant formatting between columns **'Lunch Form'!$E:H**.

> Source Information | Sheet **'Schedule - Reservations'!**
> Pipeline Output | Sheet **'Lunch Form'!**

- **'Lunch Form'!$A5:A99 - AM/PM Flag**
    - Produce Column A's Morning & Afternoon time period designation

``` excel-formula
=LET(                                                                   /* One formula copied down column A */
AM_activity_time, $B5,                                                  /* Row's activity time */
IF(AM_activity_time <> "", "AM","")                                     /* Categorise Activity Time for future filtering */
)
```

- **'Lunch Form'!B5:B52 - AM times (for chosen date)**
    - Extract all AM Activty Time entries
``` excel-formula
=Let(
AM_activity_time, 'Schedule - Reservations'!$P$8:$P,                    /* Activity Time (source table) all AM Activities */
dates, 'Schedule - Reservations'!$B$8:$B,                               /* Date column for all reservations (source table) */
chosen_date, $B$3,                                                      /* Date chosen for daily lunch form input data */ 

FILTER(
   IF( {AM_activity_time}= "","",                                       /* If Activity Time is blank, return blank */
      {AM_activity_time}),                                              /* Extract all AM Activity Time entries */
      dates = chosen_date,                                              /* Filter where Chosen_Date = Date */
      dates<>"")                                                        /* Filter where Date Not Blank */
)
```

- **'Lunch Form'!C5:C52 - AM guides (for chosen date)**
    - Extract all AM Guide entries
``` excel-formula
=Let(
AM_guides, 'Schedule - Reservations'!$Q$8:$Q,                           /* Assigned Guide for each reservation on AM Activities (source table) */
dates, 'Schedule - Reservations'!$B$8:$B,                               /* Date column for all reservations (source table) */
chosen_date, $B$3,                                                      /* Date chosen for daily lunch form input data */ 

FILTER(
    IF( {AM_guides}= "", "",                                            /* If AM Guide is Blank, Return Blank */
    {AM_guides}),                                                       /* Extract All Guides assigned to AM Activites For Each Reservation */
    dates = chosen_date,                                                /* Filter Where Chosen_Date = Date */
    dates<>"")                                                          /* Filter Where Date Not Blank */
)                                                                       
```                                                                     

- **'Lunch Form'!D5:D52 - Lunch type (LOOKUP)**
    - Extract and transform activity types into required lunch types
``` excel-formula
=Let(
AM_guides, 'Schedule - Reservations'!$Q$8:$Q,                           /* Assigned Guide for each reservation on AM Activities (source table)*/
AM_activities, 'Schedule - Reservations'!$O$8:$O,                       /* AM Activity for each reservation (soure table)*/
activities_references, Dropdowns!$B$3:$E$69,                            /* References VLOOKUP table */
dates, 'Schedule - Reservations'!$B$8:$B,                               /* Date column for all reservations (source table) */
chosen_date, $B$3,                                                      /* Date chosen for daily lunch form input data */ 

FILTER(
   IFERROR(IF( 
        {AM_guides}= "", "",                                            /* If AM Guide is blank, Return blank */
        { VLOOKUP( AM_activities, activities_references, 4, FALSE)})    /* Lookup 4th column value ("Lunch Form Naming") for AM Activity */
        ,"None"),                                                       /* If Error, I.E no Activity Value, return "None" */
   dates = chosen_date,                                                 /* Filter where Chosen_Date = Date */
   dates <>""                                                           /* Filter where Date not blank */
  )
)
```

##### 02 | Transform spill data 
- Extract unique pairs of guide and lunch types and attach the earliest associated activity time. 


- **'Lunch Form'!B53:B76 - Earliest time per (guide,lunch) pair**
    - Iterate through unique filtered spill data to source the minimum time value associated to each unique data pair.
    - References point to the spilled AM columns from step 01.

``` excel-formula
=Let(
unique_AM_guides, $C$53:$C$76,                                          /* Guide Name entry from unique AM Guide-Lunch list for row i */
AM_guide_lunches, $D$53:$D$76,                                          /* Lunch Type entry from unique AM Guide-Lunch list for row i */
AM_activity_times, $B$5:$B$52,                                          /* Spilled AM Activity Times */
AM_guides, $C$5:$C$52,                                                  /* Spilled AM Guides */
AM_lunches, $D$5:$D$52,                                                 /* Spilled AM Lunch Types */


/* For each guide lunch pair (g,l) from the two queried lists, compute the result.  */
/* MAP pairs g_i with l_i (no cross-join).*/

MAP(unique_AM_guides, AM_guide_lunches,                               /* Two lists to iterate through each row */
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

- **'Lunch Form'!C53:D76 — Unique guide–lunch pairs**
    - Filter spill data for unique guide-lunch pairs

``` excel-formula
=Let(
AM_activity_lunches, $C5:$D52,                                              /* AM Guide-Lunch pairs */
AM_lunches, $D$5:$D$52,                                                     /* AM Lunches */

UNIQUE(                                                                     /* Extract unique unique AM Guide-Lunch pairs */
  filter( AM_activity_lunches,
       AM_lunches <> "none"                                                 /* Filter where AM_Lunches do not equal "none" */
     )
  )
)
```

##### 03 | Continue Transformation
- Assign required lunch preparation time for each guided activity


- **'Lunch Form'!B77:B99 — Required prep time**
    - Assign required lunch-type preparation time

``` excel-formula
=Let(
lunch_type, $D77,                                                       /* Row associated Lunch Type */
AM_activity_time,$B77,                                                  /* Row associated AM Activity Time */
lunch_times, FILTER(Dropdowns!$J$3:$L, Dropdowns!$J$3:$J <>""),         /* Reference table for required lunch preparation time */

IF(lunch_type <> "",                                                    
    IF( lunch_type =Dropdowns!$J$4,                                     /* If Lunch Type = Dropdowns!$J$4 "Airport Snacks" */
            MOD( AM_activity_time -TIME(2,0,0),1),                      /* Output AM Activity Time subtracted by two hours */
            VLOOKUP(lunch_type,lunch_times,2,FALSE)                     /* Otherwise, output second column value from VLOOKUP for AM lunch preparation time */   
        ),
    "")                                                                 /* If lunch type is blank, return blank */
)
```

- **'Lunch Form'!$C77:D99**
    - Copy filtered unique guide-lunch pairs for associated functions
``` excel-formula
=Let(
AM_lunches_form, $C$53:$D$76,                                           /* Previous unique filter guide lunch pairs */
ARRAYFORMULA(AM_lunches_form)                                           /* Output data for Column B reference */
)
```

##### 04 | Load Form Output
- Present final output data in visual setting for interpretation 

- **'Lunch Form'!A102:D - Form Arrangements**
    - Load both AM and PM lunch form entries into singular array output
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
  2, TRUE                                                                           /* Sort output by column 2 (lunch preparation time) in ascending order */
 )
)

```

- **'Lunch Form'!E102:E - Guests Names**
    - Extract guest reservations associated with each AM/PM guided activity and lunch.
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

- **'Lunch Form'!F$102:$F - Guest Counts**
    - Sum the quantity of guests for each reservation associated with each row of guide-lunch data output.

``` excel-formula
=Let(
activity_time_slot, $A102,                                                          /* AM or PM activity slot */
guide, $C102,                                                                       /* Guide for relevant activity */
lunch_type, D$102,                                                                  /* Designated lunch type */
to_go_lunch_departure, Dropdowns!$J$8,                                              /* Specific lunch type for departing guests */
to_go_breakfast_departure, Dropdowns!$J$5,                                          /* Specific lunch type for early departing guests */
guest_qty, 'Schedule - Reservations'!$J$8:$J,                                       /* Guest quantity for each reservation */
dates, 'Schedule - Reservations'!$B$8:$B,                                           /* Date column for all reservations */
chosen_date, $B$3,                                                                  /* Date chosen for daily lunch form input data */ 
AM_guides, 'Schedule - Reservations'!$Q$8:$Q,                                       /* Assigned Guide for each reservation on AM Activities */
AM_activity_category, 'Schedule - Reservations'!$K$8:$K,                            /* Assigned AM activity category type */
departure_category, Dropdowns!$H$4,                                                 /* Specific category type for departing guests */
PM_guides, 'Schedule - Reservations'!$V$8:$V,                                       /* Assigned Guide for each reservation on PM Activities */
PM_activity_category, 'Schedule - Reservations'!$L$8:$L,                            /* Assigned PM activity category type */

IF(activity_time_slot = "AM",                                                       /* For AM Activity Slots */
  IF(guide <>"",
     SUM(                                                                           /* Sum guest quantities */
              IF(lunch_type = to_go_lunch_departure,                                /* Guide will == "None" */
             FILTER( {guest_qty},
                 (dates = chosen_date) *
                 (AM_guides = guide) *                                              
                 (AM_activity_category = departure_category)                        /* Therefore filter by "Departure" category */
                 ),
               IF(lunch_type = to_go_breakfast_departure ,
             FILTER( {guest_qty},
                 (dates = chosen_date) *
                 (AM_guides = guide) *                                              /* Guide will == "None" */
                (AM_activity_category = departure_category)                         /* Therefore filter by "Departure" category */
                 ),
            FILTER( {guest_qty},                                                    
                 (dates = chosen_date) *                                            /* Filter reservations for where dates = chosen_date */
                 (AM_guides = guide)                                                /* Filter reservations for where AM_guides = guide */
                )))
        )
 ,"")
 ,                                                                                  /* For PM Activity Slots */
  IF(guide <>"",
     SUM(                                                                           /* Sum guest quantities */
              IF(lunch_type = to_go_lunch_departure ,                               /* Guide will == "None" */
             FILTER( {guest_qty},
                 (dates = chosen_date) *
                 (PM_guides = guide) *                    
                 (PM_activity_category = departure_category)                        /* Therefore filter by "Departure" category */
                 ),
            FILTER( {guest_qty},
                 (dates = chosen_date) *                                            /* Filter reservations for where dates = chosen_date */
                (PM_guides = guide)                                                 /* Filter reservations for where AM_guides = guide */
                ))
        )
 ,""))
)
```

- **'Lunch Form'!G$102:G - Guide Counts**
    - Add quantity of guides of relevant activity to total lunch recipients 

``` excel-formula
=Let(
guide, $C102,                                                                       /* Activity guide */
total_guests, $F102,                                                                /* Total guests for activity */

IF(guide ="", "",                                                                   /* If guide is blank, return blank */
  IF(LOWER(TRIM(guide))="none", total_guests,                                       /* If guide is "none", add 0 additional recipients */
    total_guests + COUNTA(SPLIT(REGEXREPLACE(guide,"\s*,\s*",","),","))             /* For each guide separated by ",", count 1 additional recipient to lunch */
  )
)
)

```

#### Summary
The final workflow output is visual friendly printable/digital lunch form for kitchen employees to prepare during daily operations. 

This process saves ~25 minutes per day across the 16 week season, about ~46 hours of labour eliminated, to be used towards alternative operation tasks or enhancing customer service.

### Daily Activity Schedule

#### Process

##### Filtering
