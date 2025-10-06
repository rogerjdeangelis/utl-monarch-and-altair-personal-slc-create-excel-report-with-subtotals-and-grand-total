%let pgm=utl-monarch-and-altair-personal-slc-create-excel-report-with-subtotals-and-grand-total;

%stop_submission;

Monarch and altair personal slc create excel report with subtotals and grand total

Too long to post here, see github

github
https://github.com/rogerjdeangelis/utl-monarch-and-altair-personal-slc-create-excel-report-with-subtotals-and-grand-total

  Two Solutions (altair personal slc)

        1 proc report
        2 datastep

Other than using a sas datastep to create the inpout exec sheet, no sas dataset is used for
the remaining inputs and outputs. I treat excel like a database?

community.altair.com
https://community.altair.com/discussion/19000/how-to-create-a-subtotal-after-each-group-break-in-monarch-classic?tab=all&utm_source=community-search&utm_medium=organic-search&utm_term=monarch%20excel

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

INPUT                                OUTPUT

d:/cls/input.xlsx also in this repo  d:/cls/output.xlsx (datastep output)

--------+++----+                     -----------------------+
| A1| fx|DATE  |                     | A1| fx    | DATE     |
--------------------------------     -------------------------------------------------------+
[_] | A        | B |  C  |  D   |    [_] |          A         |    B    |   C     |    D    |
---------------------------------    -------------------------------------------------------|
 1  |DATE      |AGE|LUNCH|DINNER|     1  | DATE               | AGE     | LUNCH   | DINNER  |
 -- |----------+---+-----+------+     -- |----------+---------+---------+---------+---------|
 2  |2025-09-14| 14| $19 | $12  |     2  | 2025-09-14         | 13      | 26      | 34      |
 -- |----------+---------+------+     -- |--------------------+---------+---------+---------|
 3  |2025-09-14| 13| $26 | $34  |     3  | 2025-09-14         | 13      | 15      | 28      |
 -- |----------+---------+------+     -- |--------------------+---------+---------+---------|
 4  |2025-09-14| 13| $15 | $28  |     4  | 2025-09-14         | 14      | 22      | 22      |
 -- |----------+---------+------+     -- |----------+---------+---------+---------+---------|
 5  |2025-10-04| 14| $22 | $22  |     5  | Subtotal 2025-09-14|         | 63      | 84      |
 -- |----------+---------+------+     -- |--------------------+---------+---------+---------|
 6  |2025-10-04| 14| $13 | $32  |     6  | 2025-10-04         | 14      | 19      | 12      |
 -- |----------+---------+------+     -- |--------------------+---------+---------+---------|
 7  |2025-10-04| 12| $17 | $13  |     7  | 2025-10-04         | 14      | 13      | 32      |
 -- |----------+---------+------+     -- |--------------------+---------+---------+---------|
                                      8  | 2025-10-04         | 12      | 17      | 13      |
                                      -- |--------------------+---------+---------+---------|
                                      9  | Subtotal 2025-10-04|         | 49      | 57      |
                                      -- |--------------------+---------+---------+---------|
                                     10  | Total              |         | 112     | 141     |
                                      -- |--------------------+---------+---------+---------|

%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_; /*-- enable listing output and set options          --*/

proc datasets lib=work;   /*-- incase you rerun              --*/
 delete have;
run;quit;

data have;
  informat date $24.;
  input
    date age lunch dinner;
cards4;
2025-09-14 13 26 34
2025-09-14 13 15 28
2025-09-14 14 22 22
2025-10-04 14 19 12
2025-10-04 14 13 32
2025-10-04 12 17 13
;;;;
run;quit;

%utlfkil(d:/xls/input.xlsx);

ods excel file="d:/xls/input.xlsx" options(sheet_name="input");

proc print data=have noobs;
run;quit;

ods excel close;

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

408       ODS _ALL_ CLOSE;
409       FILENAME WPSWBHTM TEMP;
NOTE: Writing HTML(WBHTML) BODY file d:\wpswrk\_TD14852\#LN00026
410       ODS HTML(ID=WBHTML) BODY=WPSWBHTM GPATH="d:\wpswrk\_TD14852";
411       %utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
412       &_init_;
413
414       libname xlsinp excel "d:/xls/input.xlsx";
NOTE: Library xlsinp assigned as follows:
      Engine:        OLEDB
      Physical Name: d:/xls/input.xlsx

415
416       proc contents data=xlsinp._all_;
417       run;quit;
NOTE: Procedure contents step took :
      real time : 0.250
      cpu time  : 0.078


418
419       data have;
420         set xlsinp.'input$'n;
421       run;

NOTE: 6 observations were read from "XLSINP.input$"
NOTE: Data set "WORK.have" has 6 observation(s) and 4 variable(s)
NOTE: The data step took :
      real time : 0.147
      cpu time  : 0.062


421     !     quit;
422
423       proc print data=have;
424       run;quit;
NOTE: 6 observations were read from "WORK.have"
NOTE: Procedure print step took :
      real time : 0.039
      cpu time  : 0.000


NOTE: Libref XLSINP has been deassigned.
425
426       libname xlsinp clear;
427       quit; run;
428       ODS _ALL_ CLOSE;
429       FILENAME WPSWBHTM CLEAR;

/*                                                    _
/ |  _ __  _ __ ___   ___   _ __ ___ _ __   ___  _ __| |_
| | | `_ \| `__/ _ \ / __| | `__/ _ \ `_ \ / _ \| `__| __|
| | | |_) | | | (_) | (__  | | |  __/ |_) | (_) | |  | |_
|_| | .__/|_|  \___/ \___| |_|  \___| .__/ \___/|_|   \__|
    |_|                             |_|
*/

%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

libname xlsinp excel "d:/xls/input.xlsx";

%utlfkil(d:/xls/output.xlsx);
libname xlsout excel "d:/xls/output.xlsx";

utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

proc report data=xlsinp.have headline headskip out=clsout.output (drop=_break_);
 columns date age lunch dinner;
 define date / group;
 define age / display;
 define lunch / analysis sum ;
 define dinner / analysis sum;

 break after date / summarize;
 rbreak after / summarize;

 compute after date;
     date = catx(' ','Subtotal',date);
 endcomp;

 compute after;
     date = 'Total';
 endcomp;
run;quit;

libname xlsinp clear;
libname xlsout clear;

d:/xls/output.xlsx

------------------+
| A1| fx    | DATE|
--------------------------------------------------+
[_] |          A    |    B    |   C     |    D    |
--------------------------------------------------|
 1  | DATE          | AGE     | LUNCH   | DINNER  |
 -- |----------+----+---------+---------+---------|
 2  | 2025-09-14    | 13      | 26      | 34      |
 -- |---------------+---------+---------+---------|
 3  | 2025-09-14    | 13      | 15      | 28      |
 -- |---------------+---------+---------+---------|
 4  | 2025-09-14    | 14      | 22      | 22      |
 -- |----------+----+---------+---------+---------|
 5  | 2025-09-14    |         | 63      | 84      |
 -- |---------------+---------+---------+---------|
 6  | 2025-10-04    | 14      | 19      | 12      |
 -- |---------------+---------+---------+---------|
 7  | 2025-10-04    | 14      | 13      | 32      |
 -- |---------------+---------+---------+---------|
 8  | 2025-10-04    | 12      | 17      | 13      |
 -- |---------------+---------+---------+---------|
 9  | 2025-10-04    |         | 49      | 57      |
 -- |---------------+---------+---------+---------|
10  | Total         |         | 112     | 141     |
 -- |---------------+---------+---------+---------|


199      ODS _ALL_ CLOSE;
2200      FILENAME WPSWBHTM TEMP;
NOTE: Writing HTML(WBHTML) BODY file d:\wpswrk\_TD14852\#LN00120
2201      ODS HTML(ID=WBHTML) BODY=WPSWBHTM GPATH="d:\wpswrk\_TD14852";
2202
2203      %utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
2204      &_init_;
2205
2206      %utlfkil(d:/xls/output.xlsx);
The file d:/xls/output.xlsx does not exist
2207
2208      libname xlsinp excel "d:/xls/input.xlsx";
NOTE: Library xlsinp assigned as follows:
      Engine:        OLEDB
      Physical Name: d:/xls/input.xlsx

2209      libname xlsout excel "d:/xls/output.xlsx";
NOTE: Library xlsout assigned as follows:
      Engine:        OLEDB
      Physical Name: d:/xls/output.xlsx

2210
2211      proc datasets lib=xlsout;
NOTE: No matching members in directory
Altair SLC

The DATASETS Procedure

   Directory

Libref    XLSOUT
Engine    OLEDB
2212        delete output;
2213      run;quit;
NOTE: XLSOUT.OUTPUT (memtype="DATA") was not found, and has not been deleted
NOTE: Procedure datasets step took :
      real time : 0.220
      cpu time  : 0.046


2214
2215      data xlsout.output;
2216        length date $24 age 8;
2217        retain subtotal_lunch   total_lunch
2218                subtotal_dinner total_dinner 0;
2219        set xlsinp.'input$'n end=dne;
2220        by date;
2221        output;
2222        total_lunch = sum(total_lunch,lunch);
2223        total_dinner = sum(total_dinner,dinner);
2224        Subtotal_lunch=sum(Subtotal_lunch,lunch);
2225        Subtotal_dinner=sum(Subtotal_dinner,dinner);
2226        if last.date then do;
2227           date=catx(' ',"Subtotal",date);
2228           lunch=Subtotal_lunch;
2229           dinner=Subtotal_dinner;
2230           age=._;
2231           output;
2232           Subtotal_lunch = 0;
2233           Subtotal_dinner = 0;
2234        end;
2235        if dne then do;
2236           date='Total';
2237           lunch=total_lunch;
2238           dinner=total_dinner;
2239           age=._;
2240           output;
2241        end;
2242        drop total_lunch subtotal_lunch
2243             total_dinner subtotal_dinner dne;
2244      run;

NOTE: 6 observations were read from "XLSINP.input$"
NOTE: Data set "XLSOUT.output" has an unknown number of observation(s) and 4 variable(s)
NOTE: The data step took :
      real time : 0.240
      cpu time  : 0.093


2244    !     quit;
NOTE: Libref XLSINP has been deassigned.
2245
2246      libname xlsinp clear;
NOTE: Libref XLSOUT has been deassigned.
2247      libname xlsout clear;
2248
2249      quit; run;
2250      ODS _ALL_ CLOSE;
2251      FILENAME WPSWBHTM CLEAR;


/*___        _            _       _            _
|___ \   ___| | ___    __| | __ _| |_ __ _ ___| |_ ___ _ __
  __) | / __| |/ __|  / _` |/ _` | __/ _` / __| __/ _ \ `_ \
 / __/  \__ \ | (__  | (_| | (_| | || (_| \__ \ ||  __/ |_) |
|_____| |___/_|\___|  \__,_|\__,_|\__\__,_|___/\__\___| .__/
                                                      |_|
*/

%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

%utlfkil(d:/xls/output.xlsx);

libname xlsinp excel "d:/xls/input.xlsx";
libname xlsout excel "d:/xls/output.xlsx";

proc datasets lib=xlsout;
  delete output;
run;quit;

data xlsout.output;
  length date $24 age 8;
  retain subtotal_lunch   total_lunch
          subtotal_dinner total_dinner 0;
  set xlsinp.'input$'n end=dne;
  by date;
  output;
  total_lunch = sum(total_lunch,lunch);
  total_dinner = sum(total_dinner,dinner);
  Subtotal_lunch=sum(Subtotal_lunch,lunch);
  Subtotal_dinner=sum(Subtotal_dinner,dinner);
  if last.date then do;
     date=catx(' ',"Subtotal",date);
     lunch=Subtotal_lunch;
     dinner=Subtotal_dinner;
     age=.;
     output;
     Subtotal_lunch = 0;
     Subtotal_dinner = 0;
  end;
  if dne then do;
     date='Total';
     lunch=total_lunch;
     dinner=total_dinner;
     age=.;
     output;
  end;
  drop total_lunch subtotal_lunch
       total_dinner subtotal_dinner dne;
run;quit;

libname xlsinp clear;
libname xlsout clear;

d:/xls/output.xlsx

-----------------------+
| A1| fx    | DATE     |
-------------------------------------------------------+
[_] |          A         |    B    |   C     |    D    |
-------------------------------------------------------|
 1  | DATE               | AGE     | LUNCH   | DINNER  |
 -- |----------+---------+---------+---------+---------|
 2  | 2025-09-14         | 13      | 26      | 34      |
 -- |--------------------+---------+---------+---------|
 3  | 2025-09-14         | 13      | 15      | 28      |
 -- |--------------------+---------+---------+---------|
 4  | 2025-09-14         | 14      | 22      | 22      |
 -- |----------+---------+---------+---------+---------|
 5  | Subtotal 2025-09-14|         | 63      | 84      |
 -- |--------------------+---------+---------+---------|
 6  | 2025-10-04         | 14      | 19      | 12      |
 -- |--------------------+---------+---------+---------|
 7  | 2025-10-04         | 14      | 13      | 32      |
 -- |--------------------+---------+---------+---------|
 8  | 2025-10-04         | 12      | 17      | 13      |
 -- |--------------------+---------+---------+---------|
 9  | Subtotal 2025-10-04|         | 49      | 57      |
 -- |--------------------+---------+---------+---------|
10  | Total              |         | 112     | 141     |
 -- |--------------------+---------+---------+---------|

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

1406      ODS _ALL_ CLOSE;
1407      FILENAME WPSWBHTM TEMP;
NOTE: Writing HTML(WBHTML) BODY file d:\wpswrk\_TD14852\#LN00082
1408      ODS HTML(ID=WBHTML) BODY=WPSWBHTM GPATH="d:\wpswrk\_TD14852";
1409      %utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
1410      &_init_;
1411
1412      %utlfkil(d:/xls/output.xlsx);
1413
1414      libname xlsinp excel "d:/xls/input.xlsx";
NOTE: Library xlsinp assigned as follows:
      Engine:        OLEDB
      Physical Name: d:/xls/input.xlsx

1415      libname xlsout excel "d:/xls/output.xlsx";
NOTE: Library xlsout assigned as follows:
      Engine:        OLEDB
      Physical Name: d:/xls/output.xlsx

1416
1417      proc datasets lib=xlsout;
NOTE: No matching members in directory
Altair SLC

The DATASETS Procedure

   Directory

Libref    XLSOUT
Engine    OLEDB
1418        delete output;
1419      run;quit;
NOTE: XLSOUT.OUTPUT (memtype="DATA") was not found, and has not been deleted
NOTE: Procedure datasets step took :
      real time : 0.204
      cpu time  : 0.031


1420
1421      data xlsout.output;
1422        length date $24 age 8;
1423        retain subtotal_lunch   total_lunch
1424                subtotal_dinner total_dinner 0;
1425        set xlsinp.'input$'n end=dne;
1426        by date;
1427        output;
1428        total_lunch = sum(total_lunch,lunch);
1429        total_dinner = sum(total_dinner,dinner);
1430        Subtotal_lunch=sum(Subtotal_lunch,lunch);
1431        Subtotal_dinner=sum(Subtotal_dinner,dinner);
1432        if last.date then do;
1433           date=catx(' ',"Subtotal",date);
1434           lunch=Subtotal_lunch;
1435           dinner=Subtotal_dinner;
1436           output;
1437           Subtotal_lunch = 0;
1438           Subtotal_dinner = 0;
1439        end;
1440        if dne then do;
1441           date='Total';
1442           lunch=total_lunch;
1443           dinner=total_dinner;
1444           output;
1445        end;
1446        drop total_lunch subtotal_lunch
1447             total_dinner subtotal_dinner dne;
1448      run;

NOTE: 6 observations were read from "XLSINP.input$"
NOTE: Data set "XLSOUT.output" has an unknown number of observation(s) and 4 variable(s)
NOTE: The data step took :
      real time : 0.237
      cpu time  : 0.078


1448    !     quit;
NOTE: Libref XLSINP has been deassigned.
1449
1450      libname xlsinp clear;
NOTE: Libref XLSOUT has been deassigned.
1451      libname xlsout clear;
1452
1453      quit; run;
1454      ODS _ALL_ CLOSE;
1455      FILENAME WPSWBHTM CLEAR;

/*                                                    _
/ |  _ __  _ __ ___   ___   _ __ ___ _ __   ___  _ __| |_
| | | `_ \| `__/ _ \ / __| | `__/ _ \ `_ \ / _ \| `__| __|
| | | |_) | | | (_) | (__  | | |  __/ |_) | (_) | |  | |_
|_| | .__/|_|  \___/ \___| |_|  \___| .__/ \___/|_|   \__|
    |_|                             |_|
*/

%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

libname xlsinp excel "d:/xls/input.xlsx";

%utlfkil(d:/xls/output.xlsx);
libname xlsout excel "d:/xls/output.xlsx";

utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

proc report data=xlsinp.have headline headskip out=clsout.output (drop=_break_);
 columns date age lunch dinner;
 define date / group;
 define age / display;
 define lunch / analysis sum ;
 define dinner / analysis sum;

 break after date / summarize;
 rbreak after / summarize;

 compute after date;
     date = catx(' ','Subtotal',date);
 endcomp;

 compute after;
     date = 'Total';
 endcomp;
run;quit;

libname xlsinp clear;
libname xlsout clear;

d:/xls/output.xlsx

-----------------------+
| A1| fx    | DATE     |
-------------------------------------------------------+
[_] |          A         |    B    |   C     |    D    |
-------------------------------------------------------|
 1  | DATE               | AGE     | LUNCH   | DINNER  |
 -- |----------+---------+---------+---------+---------|
 2  | 2025-09-14         | 13      | 26      | 34      |
 -- |--------------------+---------+---------+---------|
 3  | 2025-09-14         | 13      | 15      | 28      |
 -- |--------------------+---------+---------+---------|
 4  | 2025-09-14         | 14      | 22      | 22      |
 -- |----------+---------+---------+---------+---------|
 5  | 2025-09-14         |         | 63      | 84      |
 -- |--------------------+---------+---------+---------|
 6  | 2025-10-04         | 14      | 19      | 12      |
 -- |--------------------+---------+---------+---------|
 7  | 2025-10-04         | 14      | 13      | 32      |
 -- |--------------------+---------+---------+---------|
 8  | 2025-10-04         | 12      | 17      | 13      |
 -- |--------------------+---------+---------+---------|
 9  | 2025-10-04         |         | 49      | 57      |
 -- |--------------------+---------+---------+---------|
10  | Total              |         | 112     | 141     |
 -- |--------------------+---------+---------+---------|


199      ODS _ALL_ CLOSE;
2200      FILENAME WPSWBHTM TEMP;
NOTE: Writing HTML(WBHTML) BODY file d:\wpswrk\_TD14852\#LN00120
2201      ODS HTML(ID=WBHTML) BODY=WPSWBHTM GPATH="d:\wpswrk\_TD14852";
2202
2203      %utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
2204      &_init_;
2205
2206      %utlfkil(d:/xls/output.xlsx);
The file d:/xls/output.xlsx does not exist
2207
2208      libname xlsinp excel "d:/xls/input.xlsx";
NOTE: Library xlsinp assigned as follows:
      Engine:        OLEDB
      Physical Name: d:/xls/input.xlsx

2209      libname xlsout excel "d:/xls/output.xlsx";
NOTE: Library xlsout assigned as follows:
      Engine:        OLEDB
      Physical Name: d:/xls/output.xlsx

2210
2211      proc datasets lib=xlsout;
NOTE: No matching members in directory
Altair SLC

The DATASETS Procedure

   Directory

Libref    XLSOUT
Engine    OLEDB
2212        delete output;
2213      run;quit;
NOTE: XLSOUT.OUTPUT (memtype="DATA") was not found, and has not been deleted
NOTE: Procedure datasets step took :
      real time : 0.220
      cpu time  : 0.046


2214
2215      data xlsout.output;
2216        length date $24 age 8;
2217        retain subtotal_lunch   total_lunch
2218                subtotal_dinner total_dinner 0;
2219        set xlsinp.'input$'n end=dne;
2220        by date;
2221        output;
2222        total_lunch = sum(total_lunch,lunch);
2223        total_dinner = sum(total_dinner,dinner);
2224        Subtotal_lunch=sum(Subtotal_lunch,lunch);
2225        Subtotal_dinner=sum(Subtotal_dinner,dinner);
2226        if last.date then do;
2227           date=catx(' ',"Subtotal",date);
2228           lunch=Subtotal_lunch;
2229           dinner=Subtotal_dinner;
2230           age=._;
2231           output;
2232           Subtotal_lunch = 0;
2233           Subtotal_dinner = 0;
2234        end;
2235        if dne then do;
2236           date='Total';
2237           lunch=total_lunch;
2238           dinner=total_dinner;
2239           age=._;
2240           output;
2241        end;
2242        drop total_lunch subtotal_lunch
2243             total_dinner subtotal_dinner dne;
2244      run;

NOTE: 6 observations were read from "XLSINP.input$"
NOTE: Data set "XLSOUT.output" has an unknown number of observation(s) and 4 variable(s)
NOTE: The data step took :
      real time : 0.240
      cpu time  : 0.093


2244    !     quit;
NOTE: Libref XLSINP has been deassigned.
2245
2246      libname xlsinp clear;
NOTE: Libref XLSOUT has been deassigned.
2247      libname xlsout clear;
2248
2249      quit; run;
2250      ODS _ALL_ CLOSE;
2251      FILENAME WPSWBHTM CLEAR;

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
