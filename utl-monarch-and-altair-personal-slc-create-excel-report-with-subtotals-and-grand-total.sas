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






































































































































proc report data=xlsinp.'input$'n headline headskip out=xlsout.'output$'n(rename=_break_=TOTALS) ;
 columns date age lunch dinner;
 define date / group;
 define age / group;
 define lunch / analysis sum ;
 define dinner / analysis sum;

 break after date / summarize;

 compute date;
    if not missing(date) then do;
      _break_ = ' ';
    end;
  endcomp;

 rbreak after / summarize;

 compute after date;
     date = catx(' ','Subtotal',date);
 endcomp;

 compute after;
     date = 'Total';
 endcomp;
run;quit;

libname xlsout clear; /*-- important --*/

With the sas access product does sas support, excel input to proc report see below like

libname xlsinp excel "d:/xls/input.xlsx";
proc report data=xlsinp.sheet1;
run;quit;










































%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

%utlfkil(d:/xls/output.xlsx);

libname xlsinp excel "d:/xls/input.xlsx";
libname xlsout excel "d:/xls/output.xlsx";

proc datasets lib=xlsout;
  delete output;
run;quit;

proc datasets lin=xlsinp._all_;
run;quit;

data xlsout.output;
  retain subtotal_lunch   total_lunch
          subtotal_dinner total_dinner 0;
  format lunch dinner dollar4.;
  length date $24;
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
     output;
     Subtotal_lunch = 0;
     Subtotal_dinner = 0;
  end;
  if dne then do;
     date='Total';
     lunch=total_lunch;
     dinner=total_dinner;
     output;
  end;
  drop total_lunch subtotal_lunch
       total_dinner subtotal_dinner dne;
run;quit;

libname xlsinp clear;
libname xlsout clear;



DATE                   AGE    LUNCH    DINNER

2025-09-14              13      26        34
2025-09-14              13      15        28
2025-09-14              14      22        22
Subtotal 2025-09-14     14      63        84
2025-10-04              14      19        12
2025-10-04              14      13        32
2025-10-04              12      17        13
Subtotal 2025-10-04     12      49        57
Total                   12     112       141




/*                       _       _
| |_ ___ _ __ ___  _ __ | | __ _| |_ ___
| __/ _ \ `_ ` _ \| `_ \| |/ _` | __/ _ \
| ||  __/ | | | | | |_) | | (_| | ||  __/
 \__\___|_| |_| |_| .__/|_|\__,_|\__\___|
                  |_|
*/
filename tmp temp;
data _null_;
 file tmp;
 do lyn=2 to 20;
  put @2 '-- |----------+---------+---------+---------+---------+---------+---------+---------+---------+---------|';
  row=put(lyn, 2. -r);
  put @1 row $2.
   @5 '|          |         |         |         |         |         |         |         |         |         |';
 end;
  put @2 '-- |----------+---------+---------+---------+---------+---------+---------+---------+---------+---------|';
 stop;
run;quit;

/*__ _ _ _            _ _
 / _(_) | |   ___ ___| | |___
| |_| | | |  / __/ _ \ | / __|
|  _| | | | | (_|  __/ | \__ \
|_| |_|_|_|  \___\___|_|_|___/

*/
data _null_;
  infile tmp sharebuffers;
  file tmp;
  input;
  put _infile_;
  input;
  set OUTPUT   point=_n_ nobs=numobs;
  put @7  DATE @28 AGE @38   LUNCH @48   DINNER;
  call symput('lines',2*_n_+1);
  if _n_=numobs then stop;
run;quit;
/*         _     _   _                    _
  __ _  __| | __| | | |__   ___  __ _  __| | ___ _ __
 / _` |/ _` |/ _` | | `_ \ / _ \/ _` |/ _` |/ _ \ `__|
| (_| | (_| | (_| | | | | |  __/ (_| | (_| |  __/ |
 \__,_|\__,_|\__,_| |_| |_|\___|\__,_|\__,_|\___|_|

*/
data _null_;
  infile tmp;
  input;
  if _n_=1 then do;
     put "-----------------------+";
     put "| A1| fx    |DAYNUM    |";
     put "---------------------------------------------------------------------------------------------------------+";
     put "[_] |    A     |    B    |    C    |   DE    |    E    |    F    |    G    |    H    |    I    |    K    |";
     put "---------------------------------------------------------------------------------------------------------|";
     PUT " 1  | NAME     |   SEX   |   AGE   | HEIGHT  | WEIGHT  |         |         |         |         |         |";
  end;
  putlog _infile_;
  if _n_=&lines then do;
     putlog '[CLASS}';
     stop;
  end;
run;quit;




-----------------------+
| A1| fx    | DATE     |
---------------------------------------------------------------------------------------------------------+
[_] |          A         |    B    |   C     |    D    |    F    |    G    |    H    |    I    |    K    |
---------------------------------------------------------------------------------------------------------|
 1  | NAME               | AGE     | LUNCHT  | DIMMER  |         |         |         |         |         |
 -- |----------+---------+---------+---------+---------+---------+---------+---------+---------+---------|
 2  | 2025-09-14         | 13      | 26      | 34      |         |         |         |         |         |
 -- |--------------------+---------+---------+---------+---------+---------+---------+---------+---------|
 3  | 2025-09-14         | 13      | 15      | 28      |         |         |         |         |         |
 -- |--------------------+---------+---------+---------+---------+---------+---------+---------+---------|
 4  | 2025-09-14         | 14      | 22      | 22      |         |         |         |         |         |
 -- |----------+---------+---------+---------+---------+---------+---------+---------+---------+---------|
 5  | Subtotal 2025-09-14| 14      | 63      | 84      |         |         |         |         |         |
 -- |--------------------+---------+---------+---------+---------+---------+---------+---------+---------|
 6  | 2025-10-04         | 14      | 19      | 12      |         |         |         |         |         |
 -- |--------------------+---------+---------+---------+---------+---------+---------+---------+---------|
 7  | 2025-10-04         | 14      | 13      | 32      |         |         |         |         |         |
 -- |--------------------+---------+---------+---------+---------+---------+---------+---------+---------|
 8  | 2025-10-04         | 12      | 17      | 13      |         |         |         |         |         |
 -- |--------------------+---------+---------+---------+---------+---------+---------+---------+---------|
 9  | Subtotal 2025-10-04| 12      | 49      | 57      |         |         |         |         |         |
 -- |--------------------+---------+---------+---------+---------+---------+---------+---------+---------|
10  | Total    |         | 12      | 112     | 141     |         |         |         |         |         |
 -- |----------+---------+---------+---------+---------+---------+---------+---------+---------+---------|













































%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

%utlfkil(d:/xls/output.xlsx);

libname xlsinp excel "d:/xls/input.xlsx";
libname xlsout excel "d:/xls/output.xlsx";

data xlsout.want;
  retain subtotal_lunch total_lunch 0;
  length date $24;
  set xlsinp.'input$'n;
  by date;
  output;
  total_lunch = sum(total_lunch,lunch);
  Subtotal_lunch=sum(Subtotal_lunch,lunch);
  if last.date then do;
     date=catx(' ',"Subtotal",date);
     lunch=Subtotal_lunch;
     output;
     Subtotal_lunch = 0;
  end;
  if dne then do;
     date='Total';
     lunch=total_lunch;
     output;
  end;
  drop total_lunch subtotal_lunch;
run;quit;

libname xlsinp clear;
libname xlsout clear;

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

430       ODS _ALL_ CLOSE;
431       FILENAME WPSWBHTM TEMP;
NOTE: Writing HTML(WBHTML) BODY file d:\wpswrk\_TD14852\#LN00028
432       ODS HTML(ID=WBHTML) BODY=WPSWBHTM GPATH="d:\wpswrk\_TD14852";
433
434
435       %utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
436       &_init_;
437
438       %utlfkil(d:/xls/output.xlsx);
The file d:/xls/output.xlsx does not exist
439
440       libname xlsinp excel "d:/xls/input.xlsx";
NOTE: Library xlsinp assigned as follows:
      Engine:        OLEDB
      Physical Name: d:/xls/input.xlsx

441       libname xlsout excel "d:/xls/output.xlsx";
NOTE: Library xlsout assigned as follows:
      Engine:        OLEDB
      Physical Name: d:/xls/output.xlsx

442
443       data xlsout.want;
444         retain subtotal_lunch total_lunch 0;
445         length date $24;
446         set xlsinp.'input$'n;
447         by date;
448         output;
449         total_lunch = sum(total_lunch,lunch);
450         Subtotal_lunch=sum(Subtotal_lunch,lunch);
451         if last.date then do;
452            date=catx(' ',"Subtotal",date);
453            lunch=Subtotal_lunch;
454            output;
455            Subtotal_lunch = 0;
456         end;
457         if dne then do;
458            date='Total';
459            lunch=total_lunch;
460            output;
461         end;
462         drop total_lunch subtotal_lunch;
463       run;quit;
NOTE: Variable "DNE" may not be initialized

NOTE: 6 observations were read from "XLSINP.input$"
NOTE: Data set "XLSOUT.want" has an unknown number of observation(s) and 5 variable(s)
NOTE: The data step took :
      real time : 0.239
      cpu time  : 0.046


NOTE: Libref XLSINP has been deassigned.
464
465       libname xlsinp clear;
NOTE: Libref XLSOUT has been deassigned.
466       libname xlsout clear;
467
468
469
470
471       quit; run;
472       ODS _ALL_ CLOSE;
473       FILENAME WPSWBHTM CLEAR;





 data want;
   retain subtotal_lunch total_lunch 0;
   retain subtotal_dinner total_ 0;
   length date $24;
   set have end=dne;
   by date;
   output;
   total_lunch = sum(total_lunch,lunch);
   Subtotal_lunch=sum(Subtotal_lunch,lunch);
   if last.date then do;
      date=catx(' ',"Subtotal",date);
      lunch=Subtotal_lunch;
      output;
      Subtotal_lunch = 0;
   end;
   if dne then do;
      date='Total';
      lunch=total_lunch;
      output;
   end;
   drop total_lunch subtotal_lunch dne;
 run;quit;



DATE      AGE      LUNCH      DINNER
2025-09-14      13      26      34
2025-09-14      13      15      28
2025-09-14      14      22      22
Subtotal 2025-09-14      14      63      22
2025-10-04      14      19      12
2025-10-04      14      13      32
2025-10-04      12      17      13
Subtotal 2025-10-04      12      49      13
































































%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

proc datasets lib=work;
 delete have;
run;quit;

libname xlsinp excel "d:/xls/input.xlsx";

proc contents data=xlsinp._all_;
run;quit;

data want;
  set xlsinp.'input$'n;
run;quit;

proc print want;

libname xlsinp clear;



%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;


%utlfkil(d:/xls/output.xlsx);

libname xlsinp excel "d:/xls/input.xlsx";
libname xlsout excel "d:/xls/output.xlsx";

data xlsout.want;
  length date $24 age 3.;;
  retain subtotal_lunch total_lunch 0;
  set xlsinp.'input$'n;
  by date;
  output;
  total_lunch = sum(total_lunch,lunch);
  Subtotal_lunch=sum(Subtotal_lunch,lunch);
  if last.date then do;
     date=catx(' ',"Subtotal",date);
     lunch=Subtotal_lunch;
     output;
     Subtotal_lunch = 0;
  end;
  if dne then do;
     date='Total';
     lunch=total_lunch;
     output;
  end;
  drop total_lunch subtotal_lunch;
run;quit;

libname xlsinp clear;
libname xlsout clear;




















































%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

libname xlsinp excel "d:/xls/input.xlsx";

data want;
  retain subtotal_lunch total_lunch 0;
  length date $24;
  set xlsinp.input;
  by date;
  output;
  total_lunch = sum(total_lunch,lunch);
  Subtotal_lunch=sum(Subtotal_lunch,lunch);
  if last.date then do;
     date=catx(' ',"Subtotal",date);
     lunch=Subtotal_lunch;
     output;
     Subtotal_lunch = 0;
  end;
  if dne then do;
     date='Total';
     lunch=total_lunch;
     output;
  end;
  drop total_lunch subtotal_lunch;
run;quit;














run;quit;

libname xlsinp clear;




      ;;;;%end;%mend;/*'*/ *);*};*];*/;/*"*/;run;quit;%end;end;run;endcomp;%utlfix;



%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;


libname xlsinp excel "d:/xls/input.xlsx";

data want;
  retain subtotal_lunch total_lunch 0;
  length date $24;
  set xlsinp.input;
  by date;
  output;
  total_lunch = sum(total_lunch,lunch);
  Subtotal_lunch=sum(Subtotal_lunch,lunch);
  if last.date then do;
     date=catx(' ',"Subtotal",date);
     lunch=Subtotal_lunch;
     output;
     Subtotal_lunch = 0;
  end;
  if dne then do;
     date='Total';
     lunch=total_lunch;
     output;
  end;
  drop total_lunch subtotal_lunch;
run;quit;

libname xlsinp clear;

































libname xlsinp odbc required="Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=d:/xls/input.xlsx";

ods excel file="d:/xls/output.xlsx" options(sheet_name="output");

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

data xlsout
libname xlsout excel "d:/xls/output.xlsx";

proc report data=have headline headskip out=xlsout(rename=_break_=TOTALS) ;
 columns date age lunch dinner;
 define date / group;
 define age / group;
 define lunch / analysis sum ;
 define dinner / analysis sum;

 break after date / summarize;

 compute date;
    if not missing(date) then do;
      _break_ = ' ';
    end;
  endcomp;

 rbreak after / summarize;

 compute after date;
     date = catx(' ','Subtotal',date);
 endcomp;

 compute after;
     date = 'Total';
 endcomp;
run;quit;

libname xlsinp clear; /*-- important --*/
ods excel close;

With the sas access product does sas support, excel input to proc report see below like

libname xlsinp excel "d:/xls/input.xlsx";
proc report data=xlsinp.sheet1;
run;quit;




































proc report data=have;
columns date age lunch dinnerr;
define date / order order=data;
define age / order;
define lunch / analysis format=dollar8.;
define dinr / analysis format=dollar8.;
break after date / skip;
run;

proc report data=have;
columns date age lunch dinr;
define date / order order=data;
define age / order;
define lunch / analysis format=dollar8.;
define dinr / analysis format=dollar8.;
break after date / skip;
run;

proc report data=have spanrows;
  column date date=date1 age lunch ;
  define date1 / display;
  define date / group;
  define age / group;
  define lunch / analysis sum format=dollar8.;

run;
























/*                       _       _
| |_ ___ _ __ ___  _ __ | | __ _| |_ ___
| __/ _ \ `_ ` _ \| `_ \| |/ _` | __/ _ \
| ||  __/ | | | | | |_) | | (_| | ||  __/
 \__\___|_| |_| |_| .__/|_|\__,_|\__\___|
                  |_|
*/
filename tmp temp;
data _null_;
 file tmp;
 do lyn=2 to 20;
  put @2 '-- |----------+---------+---------+---------+---------+---------+---------+---------+---------+---------|';
  row=put(lyn, 2. -r);
  put @1 row $2.
   @5 '|          |         |         |         |         |         |         |         |         |         |';
 end;
  put @2 '-- |----------+---------+---------+---------+---------+---------+---------+---------+---------+---------|';
 stop;
run;quit;

/*__ _ _ _            _ _
 / _(_) | |   ___ ___| | |___
| |_| | | |  / __/ _ \ | / __|
|  _| | | | | (_|  __/ | \__ \
|_| |_|_|_|  \___\___|_|_|___/

*/
data _null_;
  infile tmp sharebuffers;
  file tmp;
  input;
  put _infile_;
  input;
  set sd1.want   point=_n_ nobs=numobs;
  put @6 DAYNUM @16 DAY @26 BREAKFAST @36 LUNCH;
  call symput('lines',2*_n_+1);
  if _n_=numobs then stop;
run;quit;
/*         _     _   _                    _
  __ _  __| | __| | | |__   ___  __ _  __| | ___ _ __
 / _` |/ _` |/ _` | | `_ \ / _ \/ _` |/ _` |/ _ \ `__|
| (_| | (_| | (_| | | | | |  __/ (_| | (_| |  __/ |
 \__,_|\__,_|\__,_| |_| |_|\___|\__,_|\__,_|\___|_|

*/
data _null_;
  infile tmp;
  input;
  if _n_=1 then do;
     put "-----------------------+";
     put "| A1| fx    |DAYNUM    |";
     put "---------------------------------------------------------------------------------------------------------+";
     put "[_] |    A     |    B    |    C    |   DE    |    E    |    F    |    G    |    H    |    I    |    K    |";
     put "---------------------------------------------------------------------------------------------------------|";
     PUT " 1  | NAME     |   SEX   |   AGE   | HEIGHT  | WEIGHT  |         |         |         |         |         |";
  end;
  putlog _infile_;
  if _n_=&lines then do;
     putlog '[CLASS}';
     stop;
  end;
run;quit;











































Listing

Altair SLC

  DATE                            AGE      LUNCH     DINNER
  ---------------------------------------------------------

  2025-09-14                       13         41         62
                                   14         22         22
  Subtotal 2025-09-14                         63         84
  2025-10-04                       12         17         13
                                   14         32         44
  Subtotal 2025-10-04                         49         57
  Total                                      112        141

TOTS DATASET: REPORT OUTPUT DATASET (missing Sutotal and Total - should be present)

 Altair SLC

  Obs       date       age    lunch    dinner

   1     2025-09-14     13      41        62
   2     2025-09-14     14      22        22
   3     2025-09-14      .      63        84
   4     2025-10-04     12      17        13
   5     2025-10-04     14      32        44
   6     2025-10-04      .      49        57
   7                     .     112       141

/*--- lets add the total labels ---*/


%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_; /*-- enable listing output and set options          --*/

data want;
  set tots end=dne;
  by date notsorted;
  select;
     when (dne)       date=catx(" ","Total",date);
     when (last.date) date=catx(" ","SubTotal",date);
     otherwise;
  end;
run;quit;

ods excel file="d:/xls/want.xlsx" options(sheet_name="want");
proc print data=want;
run;
ods excel close;




proc tabulate data=have;
table DATE*AGE*LUNCH*DINNER ;
run;quit;

proc tabulate data=have;
table _all_;
run;quit;






                                                                                                                                                                                                                                                               Altair SLC


Obs           DATE            AGE    LUNCH    DINNER

 1     2025-09-14              13      41        62
 2     2025-09-14              14      22        22
 3     SubTotal 2025-09-14      .      63        84
 4     2025-10-04              12      17        13
 5     2025-10-04              14      32        44
 6     SubTotal 2025-10-04      .      49        57
 7     Total                    .     112       141

i have the following sas dataset, days,

days

   DATE       AGE    LUNCH    DINNER

2025-09-14     13      26       34
2025-09-14     13      15       28
2025-09-14     14      22       22
2025-10-04     14      19       12
2025-10-04     14      13       32
2025-10-04     12      17       13

i need a listing like this, does ot have
to look exacly like the below,just want gridded ascii output.
Please try to use proc report with FORMCHAR='|----|+|---+=|-/\<>*' .
I do not want a datastep.


 -----------------------------
 |DATE      |AGE|LUNCH|DINNER|
 |----------+---+-----+------+
 |2025-09-14| 14| $19 | $12  |
 |----------+---------+------+
 |2025-09-14| 13| $26 | $34  |
 |----------+---------+------+
 |2025-09-14| 13| $15 | $28  |
 |----------+---------+------+
 |2025-10-04| 14| $22 | $22  |
 |----------+---------+------+
 |2025-10-04| 14| $13 | $32  |
 |----------+---------+------+
 |2025-10-04| 12| $17 | $13  |
 |----------+---------+------+


proc report data=have nowd formchar='|----|+|---+=|-/<>*';
column DATE AGE LUNCH DINNER;
define DATE / group ;
define AGE / order ;
define LUNCH / display format=dollar5.;
define DINNER / display format=dollar5.;
compute after DATE;
line '----------+---------+------+';
endcomp;
run;

proc report data=have nowd
            formchar='|----|+|---+=|-/\<>*'
            out=_;
    column DATE AGE LUNCH DINNER;
    define DATE / display;
    define AGE / display;
    define LUNCH / display format=dollar5.;
    define DINNER / display format=dollar5.;
run;

ods listing
proc report data=have nowd
            formchar='|----|+|---+=|-/\<>*'
            style(report)=[rules=all frame=hsides]
            style(header)=[background=white]
            style(column)=[background=white];
    column DATE AGE LUNCH DINNER;
    define DATE / display;
    define AGE / display;
    define LUNCH / display format=dollar5.;
    define DINNER / display format=dollar5.;
run;

Gives
                                            DINNE
DATE                            AGE  LUNCH      R
2025-09-14                       13    $26    $34
2025-09-14                       13    $15    $28
2025-09-14                       14    $22    $22
2025-10-04                       14    $19    $12
2025-10-04                       14    $13    $32
2025-10-04                       12    $17    $13


options formchar='|----|+|---+=|-/<>*';
proc print data=have noobs formdlim='|';
format LUNCH dollar5. DINNER dollar5.;
run;

ods listing;
data hav1;
 retain seq;
 set have;
 seq=_n_;
run;quit;

data hav1;
  input
    date$11. age lunch dinner;
  seq=_n_;
cards4;
2025-09-14 13 26 34
2025-09-14 13 15 28
2025-09-14 14 22 22
2025-10-04 14 19 12
2025-10-04 14 13 32
2025-10-04 12 17 13
;;;;
run;quit;

ods _all_ close;
ods listing;
options ls=255;
options formchar='|----|+|---+=|-/\<>*' nodate nonumber;
proc tabulate data=hav1 format=8.;
    format _numeric_  12.;
    class seq date age dinner lunch;
    table SEQ=' ' * DATE=' ' * AGE=' '*
          LUNCH='LUNCH' *
          DINNER='DINNER'*N  / box={label='DATE   AGE'} rts=25;
run;
ods listing close;


How do i increase the global cell width in sas proc tabulate so the DINNER is not wrapped.
I want ascii listing text like output, not ps or pdf ...


How do i get sas proc report ascii printer listing destination to fill in the repeating dates
when using group and without call define, se below

-------------------------------------
|  DATE     |   AGE     |LUNCH|DINR |
|-----------------------+-----+-----|
|2025-09-14 |13         |  $41|  $62|
|-----------|-----------+-----+-----|
|           |14         |  $22|  $22|
|-----------+-----------+-----+-----|
|2025-10-04 |12         |  $17|  $13|
|-----------|-----------+-----+-----|
|           |14         |  $32|  $44|
-------------------------------------


ata want;
  set tots end=dne;
  by date notsorted;
  select;
     when (dne)       do; date=catx(" ","Total",date);    output; stop; end;
     when (last.date) do; date=catx(" ","SubTotal",date); output; end;
     otherwise output;
  end;
run;quit;











  2025-09-14                       13         41         62
                                   14         22         22
  Subtotal 2025-09-14                         63         84
  2025-10-04                       12         17         13
                                   14         32         44
  Subtotal 2025-10-04                         49         57
  Total                                      112        141











proc report data=have headline headskip out=tots (drop=_break_);
 columns date age lunch dinner;
 define date / group;
 define age / group;
 define lunch / summarize sum ;
 define dinner / summarize sum;

 break after date / summarize;
 rbreak after / summarize;

 compute after date;
     date = catx(' ','Subtotal',date);
 endcomp;

 compute after;
     date = 'Total';
 endcomp;
run;quit;








































                                                                                                                                                                                                                                                               Altair SLC


Obs       date       age    lunch    dinner

 1     2025-09-14     13      41        62
 2     2025-09-14     14      22        22
 3     2025-09-14      .      63        84
 4     2025-10-04     12      17        13
 5     2025-10-04     14      32        44
 6     2025-10-04      .      49        57
 7                     .     112       141


























libname xlsout excel "d:/xls/want.xlsx";































%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;


%utlfkil(d:/xls/input.xlsx); /*-- incase you rerun            --*/

libname xlsout excel "d:/xls/want.xlsx";




















%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

proc report data=have headline headskip out=tots (drop=_break_);
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


/*----
Altair SLC TOTS TABLE

Obs    sex    age    lunch    dinner

 1      F      13      41        62
 2      F      14      22        22
 3      F       .      63        84
 4      M      12      17        13
 5      M      14      32        44
 6      M       .      49        57
 7              .     112       141
----*/



ods excel file="d:/xls/input.xlsx" options(sheet_name="input");

data want;
  set tots end=dne;
  by sex notsorted;
  if dne then sex=catx(" ","Total",sex);
  else if last.sex then sex=catx(" ","SubTotal",sex);
run;quit;














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

%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

proc report data=have headline headskip out=tots (drop=_break_);
 columns date age lunch dinner;
 define date / group;
 define age / group;
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


/*----
Altair SLC

Obs    sex    age    lunch    dinner

 1      F      13      41        62
 2      F      14      22        22
 3      F       .      63        84
 4      M      12      17        13
 5      M      14      32        44
 6      M       .      49        57
 7              .     112       141
----*/

%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

data want;
  set tots end=dne;
  by sex notsorted;
  if dne then sex=catx(" ","Total",sex);
  else if last.sex then sex=catx(" ","SubTotal",sex);
run;quit;

ods excel file="d:/xls/wamt.xlsx" options(sheet_name="want");
proc print data=want noobs;
format lunch dinner dollar4.;
run;quit;
ods excel close;





2025-09-14
2025-09-14
2025-09-14
2025-10-04
2025-10-04
2025-10-04


















/*--- add sex to subtotal label ---*/

%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

proc print data=tots;
run;quit;

%utlfkil(d:/xls/want.xlsx); /*-- incase you rerun            --*/

libname xlsout excel "d:/xls/want.xlsx";

data xlsout.want;
 set tots;
 lagsex=lag(sex);
 if sex='Subtotal' then sex=catx(' ',sex,lagsex);
 drop lagsex;
run;quit;

libname xlsout clear;
















%utlfkil(d:/xls/want.xlsx); /*-- incase you rerun            --*/

libname xlsout excel "d:/xls/want.xlsx";

proc report data=have headline headskip out=xlsout.tots(drop=_break_);
    columns sex age lunch dinner;
    define sex / group 'SEX';
    define age / group 'AGE';
    define lunch / analysis  sum ;
    define dinner / analysis sum ;

    break after sex / summarize;
    rbreak after / summarize;

    compute after sex;

        sexlag=lag(sex);
        sex = catx(' ','Subtotal',sexlag);
        line ' ';
    endcomp;

    compute after;
        sex = 'Total';
        line ' ';
    endcomp;
run;quit;

libanme xlsout clear;


ods excel file="d:/xls/input.xlsx" options(sheet_name="input");

proc report data=have;
 cols sex age lunch dinner;
define lunch  / display f=dollar4.;
define dinner / display f=dollar4.;
run;quit;

ods excel close;



proc report data=have headline headskip out=tots (drop=_break_);
    columns sex age lunch;
    define sex / group;
    define age / group;
    define lunch / analysis sum format=8. 'LUNCH';

    break after sex / summarize;
    rbreak after / summarize;

    compute after sex;
        sex = 'Subtotal';
    endcomp;

    compute after;
        sex = 'Total';
    endcomp;
run;quit;



data have;
  format lunch dinner dollar4.;
  input
    sex$ age lunch dinner;
cards4;
M 14 19  12
F 13 26  34
F 13 15  28
F 14 22  22
M 14 13  32
M 12 17  13
;;;;
run;quit;

PROC REPORT data=HAVE spanrows;
columns SEX AGE LUNCH;
define SEX / group;
define AGE / display;
define LUNCH / analysis sum;
break after SEX / summarize suppress;
rbreak after / summarize suppress;
ompute SEX;
if break = 'SEX' then SEX = 'SUBTOT';
else if break = 'RBREAK' then SEX = 'TOTAL';
endcomp;
run;quit;

proc report data=HAVE;
columns SEX AGE LUNCH;
define SEX / group order=data;
define AGE / display;
define LUNCH / analysis sum;
compute before SEX;
line SEX $;
endcomp;
break after SEX / summarize;
rbreak after / summarize;
run;

data HAVE2;
set HAVE;
length SEX2 $8;
SEX2 = SEX;
run;

proc report data=HAVE2 spanrows;
columns SEX2 AGE LUNCH;
define SEX2 / group;
define AGE / display;
define LUNCH / analysis sum;
break after SEX2 / summarize;
rbreak after / summarize;
compute SEX2;
if break = 'SEX2' then SEX2 = 'SUBTOT';
else if break = 'RBREAK' then SEX2 = 'TOTAL';
endcomp;
run;

proc report data=HAVE2;
columns SEX2 AGE LUNCH SEX3;
define SEX2 / group noprint;
define AGE / display;
define LUNCH / analysis sum;
define SEX3 / computed length=8;
break after SEX2 / summarize;
rbreak after / summarize;
compute SEX3;
if break = 'SEX2' then SEX3 = 'SUBTOT';
else if break = 'RBREAK' then SEX3 = 'TOTAL';
else SEX3 = SEX2;
endcomp;
run;


proc report data=HAVE2;
columns SEX2 SEX3 AGE LUNCH;
define SEX2 / group noprint;
define SEX3 / computed character length=8;
define AGE / display;
define LUNCH / analysis sum;
break after SEX2 / summarize;
rbreak after / summarize;
compute SEX3;
if break = 'SEX2' then SEX3 = 'SUBTOT';
else if break = 'RBREAK' then SEX3 = 'TOTAL';
else SEX3 = SEX2;
endcomp;
run;

proc report data=HAVE2;
columns SEX2 SEX3 AGE LUNCH;
define SEX2 / group noprint;
define SEX3 / computed character length=8;
define AGE / display;
define LUNCH / analysis sum;
break after SEX2 / summarize;
rbreak after / summarize;
compute SEX3;
if break = 'SEX2' then SEX3 = 'SUBTOT';
else if break = 'RBREAK' then SEX3 = 'TOTAL';
else SEX3 = SEX2;
endcomp;
run;









































compute after SEX;
line 'SUBTOT';
endcomp;
compute after;
line 'TOTAL';
endcomp;
run;




































Given the sas datset have below

data have;
  format lunch dinner dollar4.;
  input
    sex$ age lunch dinner;
cards4;
M 14 19  12
F 13 26  34
F 13 15  28
F 14 22  22
M 14 13  32
M 12 17  13
;;;;
run;quit;

Can i use proc summary to create output, like this.
It does not need to have this exact format

 SEX  AGE      LUNCH

 F     13         26
 F     13         15
 F     14         22
 SUBTOTAL F       63
 M     14         19
 M     14         13
 M     12         17
 SUBTOTAL M       49
 GRAND TOTAL     112


PROC REPORT data=HAVE2 out=tots;
COLUMNS _SEX SEX AGE LUNCH;
DEFINE _SEX / GROUP NOPRINT;
DEFINE SEX / COMPUTED 'Sex';
DEFINE AGE / DISPLAY;
DEFINE LUNCH / ANALYSIS SUM;
BREAK AFTER _SEX / SUMMARIZE;
RBREAK AFTER / SUMMARIZE;
COMPUTE SEX / character length=8;
IF BREAK = '_SEX' THEN SEX = 'SUBTOT';
ELSE IF BREAK = 'RBREAK' THEN SEX = 'TOTAL';
ELSE SEX = _SEX;
endcomp;
run;quit;


proc report data=have out=tots completerows nowd;
    columns sex lunch;
    define sex / group;
    define lunch / analysis sum;
    break after sex / summarize;
    rbreak after / summarize;

    compute sex;
        if _break_ = '_BREAK_' then sex = 'TOTAL'; /* grand total row */
        else if _break_ = 'sex' then sex = 'SUBTOT'; /* subtotal row */
        else sex = sex; /* detail rows */
    endcomp;
run;












































This way, we use _SEX (which is the same as SEX) as the group variable and we don't print it. We then compute the SEX column, which in the break rows we set to 'SUBTOT' and 'TOTAL', and in the data rows we set to the actual sex.

But note: The break after _SEX will generate a break for each _SEX value, and the compute block for SEX will set the value to 'SUBTOT' for those breaks.

Also, the RBREAK will set the value to 'TOTAL'.

We have to be cautious because the compute block for SEX is executed for every row, including the break rows.

Also, we are using a computed variable for SEX, so we have to define its length (length=8) to accommodate 'SUBTOT' and 'TOTAL'.

Let's try this approach.

However, note that the original code also produces an output dataset. We are now using a computed variable for SEX, so the output dataset might have the computed SEX values (including 'SUBTOT' and 'TOTAL'). If that is not desired, we might need to adjust.

Alternatively, we can use the same technique but without modifying the SEX variable and instead use a separate column for the display? But the requirement is to have the SEX column filled and the break labels changed.

Given the complexity and the fact that the original report is being produced by the provided code, I think the following is a simpler solution that meets the requirement:

Use the COMPLETEROWS option to print the group variable in every row? But note: COMPLETEROWS is for completing the grid when there are across variables, and it doesn't force the group variable to print in every row.

After checking, I found that the option to print the group variable in every row is not directly available in PROC REPORT.

Alternatively, we can use a different procedure, such as PROC TABULATE, but the requirement is for PROC REPORT.

I think the


PROC REPORT data=HAVE nowd;
COLUMNS sex age lunch;
DEFINE sex / GROUP;
DEFINE age / DISPLAY;
DEFINE lunch / ANALYSIS SUM;

COMPUTE BEFORE _PAGE_;
   LINE ' SEX  AGE      LUNCH';
   LINE ' ';
ENDCOMP;

COMPUTE sex;
   IF _BREAK_ = ' ' THEN DO; /* Data rows */
      sex = sex; /* This ensures sex displays on every row */
   END;
ENDCOMP;

COMPUTE AFTER sex;
   sex = 'SUBTOT';
   age = .;
   lunch = lunch.sum;
   LINE @1 sex $6. @8 lunch 4.;
   sex = ' ';
ENDCOMP;

COMPUTE AFTER;
   sex = 'TOTAL';
   age = .;
   lunch = lunch.sum;
   LINE @1 sex $6. @8 lunch 4.;
ENDCOMP;
RUN;

PROC REPORT data=HAVE;
COLUMNS SEX AGE LUNCH;
DEFINE SEX / GROUP;
DEFINE AGE / DISPLAY;
DEFINE LUNCH / ANALYSIS SUM;
BREAK AFTER SEX / SUMMARIZE SKIP OL;
RBREAK AFTER / SUMMARIZE SKIP OL;

COMPUTE SEX;
   LENGTH _csex $6;
   IF _BREAK_ = 'SEX' THEN DO;
      _csex = 'SUBTOT';
   END;
   ELSE IF _BREAK_ = '_RBREAK_' THEN DO;
      _csex = 'TOTAL';
   END;
   ELSE DO;
      _csex = SEX;
   END;
   CALL DEFINE('SEX', 'style', 'style=[pretext="'||trim(_csex)||'"]');
ENDCOMP;
RUN;



 PROC REPORT data=HAVE;
COLUMNS SEX  LUNCH;
DEFINE SEX / GROUP;
DEFINE LUNCH / ANALYSIS SUM;
BREAK AFTER SEX / SUMMARIZE SKIP OL;
RBREAK AFTER / SUMMARIZE SKIP OL;

COMPUTE SEX;
   LENGTH _csex $6;
   IF _BREAK_ = 'SEX' THEN DO;
      _csex = 'SUBTOT';
   END;
   ELSE IF _BREAK_ = '_RBREAK_' THEN DO;
      _csex = 'TOTAL';
   END;
   ELSE DO;
      _csex = SEX;
   END;
   CALL DEFINE('SEX', 'style', 'style=[pretext="'||trim(_csex)||'"]');
ENDCOMP;
RUN;




 am working with a UDS data file (sample data below) for age group. The file has an ID (BHCMISID) and variables
for each age group i.e. each variable representing age group from under zero to 24 (

T3a_L1_Ca = under 1 for male;
T3a_L1_Cb = under 1 for female;
T3a_L2_Ca=age 1 for male,
T3a_L2_Cb=age 1 for female etc).

Beginning with variable
  T3a_L26_Ca  =25-29 (male),
  T3a_L26_Cb = 25-29 (female).

For better reporting, I would like to create age groups from these columns (variable),
  0-17,
18-24,
25-44,
45-65,
65+. How do I go about doing that? Here is sample data for one clinic. Thank you for your help.

data have;
input
BHCMISID,T3a_L6_Ca,T3a_L6_Cb,T3a_L7_Ca,T3a_L7_Cb,T3a_L8_Ca,T3a_L8_Cb,T3a_L9_Ca,T3a_L9_Cb,
T3a_L10_Ca,T3a_L10_Cb,T3a_L11_Ca,T3a_L11_Cb,T3a_L12_Ca,T3a_L12_Cb,T3a_L13_Ca,T3a_L13_Cb,
T3a_L14_Ca,T3a_L14_Cb,T3a_L15_Ca,T3a_L15_Cb,T3a_L16_Ca,T3a_L16_Cb,T3a_L17_Ca,T3a_L17_Cb,
T3a_L18_Ca,T3a_L18_Cb,T3a_L19_Ca,T3a_L19_Cb,T3a_L20_Ca,T3a_L20_Cb,T3a_L21_Ca,T3a_L21_Cb,
T3a_L22_Ca,T3a_L22_Cb,T3a_L23_a,T3a_L23_Cb,T3a_L24_Ca,T3a_L24_Cb,T3a_L25_Ca,T3a_L25_Cb,
T3a_L26_Ca,T3a_L26_Cb,T3a_L27_,T3a_L27_Cb,T3a_L28_Ca,T3a_L28_Cb,
T3a_L29_Ca,T3a_L29_Cb,T3a_L30_Ca,T3a_L30_Cb,T3a_L31_Ca,T3a_L31_Cb,T3a_L32_Ca,T3a_L32_Cb,
T3a_L33_Ca,T3a_L33_Cb,T3a_L34_Ca,T3a_L34_Cb,T3a_L35_Ca,T3a_L35_Cb,T3a_L36_Ca,T3a_L36_Cb,
T3a_L37_Ca,T3a_L37_Cb,T3a_L1_Ca,T3a_L1_Cb,T3a_L2_Ca,T3a_L2_Cb,T3a_L3_Ca,
T3a_L3_Cb,T3a_L4_Ca,T3a_L4_Cb,T3a_L5_Cb,T3a_L38_Ca,T3a_L38_Cb @@;
cards4;
090730,89,83,101,121,103,97,99,98,118,101,132,121,110,101,131,113,108,104,113,98,76,99,81,
103,91,81,86,82,59,71,51,60,58,56,58,57,73,65,60,59,349,370,509,460,496,399,455,
377,486,419,427,441,463,484,420,434,313,359,253,265,137,164,54,92,94,70,
85,105,72,62,83,88,85,31,83
;;;;
run;quit;

proc tanspose data=have;
by



data have;
input
BHCMISID T3a_L6_Ca T3a_L6_Cb T3a_L7_Ca T3a_L7_Cb T3a_L8_Ca T3a_L8_Cb T3a_L9_Ca T3a_L9_Cb
T3a_L10_Ca T3a_L10_Cb T3a_L11_Ca T3a_L11_Cb T3a_L12_Ca T3a_L12_Cb T3a_L13_Ca T3a_L13_Cb
T3a_L14_Ca T3a_L14_Cb T3a_L15_Ca T3a_L15_Cb T3a_L16_Ca T3a_L16_Cb T3a_L17_Ca T3a_L17_Cb
T3a_L18_Ca T3a_L18_Cb T3a_L19_Ca T3a_L19_Cb T3a_L20_Ca T3a_L20_Cb T3a_L21_Ca T3a_L21_Cb
T3a_L22_Ca T3a_L22_Cb T3a_L23_a T3a_L23_Cb T3a_L24_Ca T3a_L24_Cb T3a_L25_Ca T3a_L25_Cb
T3a_L26_Ca T3a_L26_Cb T3a_L27_ T3a_L27_Cb T3a_L28_Ca T3a_L28_Cb
T3a_L29_Ca T3a_L29_Cb T3a_L30_Ca T3a_L30_Cb T3a_L31_Ca T3a_L31_Cb T3a_L32_Ca T3a_L32_Cb
T3a_L33_Ca T3a_L33_Cb T3a_L34_Ca T3a_L34_Cb T3a_L35_Ca T3a_L35_Cb T3a_L36_Ca T3a_L36_Cb
T3a_L37_Ca T3a_L37_Cb T3a_L1_Ca T3a_L1_Cb T3a_L2_Ca T3a_L2_Cb T3a_L3_Ca
T3a_L3_Cb T3a_L4_Ca T3a_L4_Cb T3a_L5_Cb  T3a_L5_Ca T3a_L38_Ca T3a_L38_Cb @@;
cards4;
090730 89 83 101 121 103 97 99 98 118 101 132 121 110 101 131 113 108 104 113 98 76 99 81
103 91 81 86 82 59 71 51 60 58 56 58 57 73 65 60 59 349 370 509 460 496 399 455
377 486 419 427 441 463 484 420 434 313 359 253 265 137 164 54 92 94 70
85 105 72 62 83 88 85 31 83 66
;;;;
run;quit;

%array(tens,values=1-9);

/*-- not missing variables are ok with sum they act as zeroes ---*/
data sums;
 set have;
   age_1_10 =sum(%do_over(tens,phrase=T3A_L?_CA ,between=comma) + T3A_L10_CA );
   age_11_20=sum(%do_over(tens,phrase=T3A_L1?_CA,between=comma) + T3A_L20_CA );
   age_21_30=sum(%do_over(tens,phrase=T3A_L2?_CA,between=comma) + T3A_L30_CA );
   age_31_40=sum(%do_over(tens,phrase=T3A_L3?_CA,between=comma) + T3A_L40_CA );
   keep age:;
run;quit;

40 obs from SUMS total obs=1 06OCT2025:
              AGE_     AGE_     AGE_
 AGE_1_10    11_20    21_30    31_40

    875       987      2028     2150


/*--- lets look at just males ----*/
proc transpose data=have out=havxpo(where=(index(_name_,'_CA')>0));
by BHCMISID;
run;quit;

havxpo

90730     T3A_L1_CA       94
90730     T3A_L2_CA       85
90730     T3A_L3_CA       72
90730     T3A_L4_CA       83
90730     T3A_L5_CA       31
90730     T3A_L6_CA       89
90730     T3A_L7_CA      101
90730     T3A_L8_CA      103
90730     T3A_L9_CA       99
90730     T3A_L10_CA     118 875 agrees with code

90730     T3A_L11_CA     132
90730     T3A_L12_CA     110
90730     T3A_L13_CA     131
90730     T3A_L14_CA     108
90730     T3A_L15_CA     113
90730     T3A_L16_CA      76
90730     T3A_L17_CA      81
90730     T3A_L18_CA      91
90730     T3A_L19_CA      86
90730     T3A_L20_CA      59  987

90730     T3A_L21_CA      51
90730     T3A_L22_CA      58
90730     T3A_L24_CA      73
90730     T3A_L25_CA      60
90730     T3A_L26_CA     349
90730     T3A_L28_CA     496
90730     T3A_L29_CA     455
90730     T3A_L30_CA     486
90730     T3A_L31_CA     427
90730     T3A_L32_CA     463
90730     T3A_L33_CA     420
90730     T3A_L34_CA     313
90730     T3A_L35_CA     253
90730     T3A_L36_CA     137
90730     T3A_L37_CA      54
90730     T3A_L38_CA      83


3A_L1_CA , T3A_L2_CA , T3A_L3_CA , T3A_L4_CA , T3A_L5_CA , T3A_L6_CA , T3A_L7_CA , T3A_L8_CA , T3A_L9_CA , T3A_L10_CA










   age1_21_30=sum(%do_over(tens,phrase=T3A_L2?_CA,between=comma) );
   age1_31_40=sum(%do_over(tens,phrase=T3A_L3?_CA,between=comma) );
       AGE1_    AGE1_    AGE1_    AGE1_
Obs     1_10    11_20    21_30    31_40

 1      875      875      875      875


%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

%utlfkil(d:/xls/want.xlsx); /*-- incase you rerun            --*/

libname xlsout excel "d:/xls/want.xlsx";

proc report data=have headline headskip out=xlsout.tots(drop=_break_);
    columns sex age lunch dinner;
    define sex / group 'SEX';
    define age / group 'AGE';
    define lunch / analysis  sum ;
    define dinner / analysis sum ;

    break after sex / summarize;
    rbreak after / summarize;

    compute after sex;

        sexlag=lag(sex);
        sex = catx(' ','Subtotal',sexlag);
        line ' ';
    endcomp;

    compute after;
        sex = 'Total';
        line ' ';
    endcomp;
run;quit;

proc sort data=


 BHCMISID      _NAME_      COL1

   90730     T3A_L1_CA       94
   90730     T3A_L2_CA       85
   90730     T3A_L3_CA       72
   90730     T3A_L4_CA       83
   90730     T3A_L6_CA       89
   90730     T3A_L7_CA      101
   90730     T3A_L8_CA      103
   90730     T3A_L9_CA       99
   90730     T3A_L1_CA      118
   90730     T3A_L11_CA     132
   90730     T3A_L12_CA     110
   90730     T3A_L13_CA     131
   90730     T3A_L14_CA     108
   90730     T3A_L15_CA     113
   90730     T3A_L16_CA      76
   90730     T3A_L17_CA      81
   90730     T3A_L18_CA      91
   90730     T3A_L19_CA      86
   90730     T3A_L20_CA      59
   90730     T3A_L21_CA      51
   90730     T3A_L22_CA      58
   90730     T3A_L24_CA      73
   90730     T3A_L25_CA      60
   90730     T3A_L26_CA     349
   90730     T3A_L28_CA     496
   90730     T3A_L29_CA     455
   90730     T3A_L30_CA     486
   90730     T3A_L31_CA     427
   90730     T3A_L32_CA     463
   90730     T3A_L33_CA     420
   90730     T3A_L34_CA     313
   90730     T3A_L35_CA     253
   90730     T3A_L36_CA     137
   90730     T3A_L37_CA      54
   90730     T3A_L38_CA      31

Suppose
  ages 1  -10 years old are T3A_L1_CA-T3A_L10_CA
  ages 11 -20 years old are T3A_L11_CA-T3A_L20_CA
  ages 21 -30 years old are T3A_L21_CA-T3A_L30_CA
  ages 31 -38 years old are T3A_L31_CA-T3A_L38_CA


%array(bg.values=1 11 21 30);
%array(en.values=10 20 30 38);


Data sums;

 %do_ober(bg en,phrase=%str(
     range?=sum(




/*---
SEX         AGE    LUNCH    _BREAK_

F            13      41
F            14      22
Subtotal      .      63     SEX
M            12      17
M            14      32
Subtotal      .      49     SEX
Total         .     112     _RBREAK_
---*/

options missin=' ';
data want;
  length sex $16;
  set tots;
  lagsex=lag(sex);
  if sex='Subtotal' then sex=catx(' ',sex,lagsex);
  drop _break_ lagsex;
run;quit;

ods excel file="d:/xls/want.xlsx" options(sheet_name="want");
proc report;
 cols sex age






Given sas dataset have below.

HAVE

SEX    AGE    LUNCH

 M      14      19
 F      13      26
 F      13      15
 F      14      22
 M      14      13
 M      12      17

Create the report beloq using sas proc report.
You do not have to match the layout exactly but I want the totals.

SEX    AGE    LUNCH

 F      13      26
 F      13      15
 F      14      22
Subtotal        63

 M      14      13
 M      12      17
 M      14      19
Subtotal        49
Total          112










libname xlsinp clear;




proc datasets lib=work; /*-- incase you rerun              --*/
 delete have;
run;quit;















$
$
$
$
$
$
$

































%let pgm=utl-monarch-and-altair-percsonal-slc-automating-excel-data-preperation;

%stop_submission;

Monarch and altair percsonal slc automating excel data preperation

Too long to post in a listserve, see github

github
https://github.com/rogerjdeangelis/utl-automating-excel-data-preperation-monarch-and-altair-percsonal-slc

/community.altair
https://community.altair.com/discussion/comment/36489?tab=all#Comment_36489?utm_source=community-search&utm_medium=organic-search&utm_term=monarch+excel

Most listserves mangle my posts, except SAS-L, when viewed directly.
Most email systems also mangle my posts. Less is more, lets go back to recognizing fixed fonts?

A window may pop up saying there is an error with the local server, just ignore it.
See the clean log on the end.

WHAT OP WANTS

Example: load excel file>split the date column and filter only year value
in that column> Split the email address and keep only domain
address> export the file into excel sheets.

/*--- create mondays raw excel input. You get a new excel file every day ---*/
/*--- you need to prepare the excel for export to Monarch                ---*/
/*--- using the ops business rules above, add YEAR * DOMAIN to excel     ---*/

INPUT                                        OUTPUT (add year and domain)

d:/xls/monday.xlsx                           d:/xls/want.xlsx

-----------------------+                     -----------------+
| A1| fx        |NAME  |                     | A1| fx  |NAME  |
------------------------------------------   ---------------------------------------------------------
[_] |    A     |    B            | C | D |   [_] |  A |    B    |    C     |    C            | E | F |
------------------------------------------   ---------------------------------------------------------
 1  | NAME     |EMAIL            |SEX|AGE|    1  |YEAR|DOMAIN   | NAME     |EMAIL            |SEX|AGE|
 -- |----------+-----------------+---+---+    -- |----+---------+----------+-----------------+---+---+
 2  |2025-09-15|Alice.gmail.com  | F |   |    2  |2025|gmail.com|2025-09-15|Alice.gmail.com  | F |   |
 -- |--------+-------------------+---+---+    -- |--------------+--------+-------------------+---+---+
 3  |2025-09-26|Barbara.gmail.com| F | 17|    3  |2025|gmail.com|2025-09-26|Barbara.gmail.com| F | 17|
 -- |--------+-|-----------------+---+---+    -- |----|---------+--------+-|-----------------+---+---+
 4  |2025-10-14|Alfred.gmail.com | F | 11|    4  |2025|gmail.com|2025-10-14|Alfred.gmail.com | F | 11|
 -- |--------+-|-----------------+---+---+    -- |----|---------+--------+-|-----------------+---+---+
 5  |2025-10-28|Al.gmail.com     | F | 12|    5  |2025|gmail.com|2025-10-28|Al.gmail.com     | F | 12|
 -- |----------+--------------------------    -- |----+--------------------+--------------------------
[MONDAY]                                     [WANT]

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

d:/xls/monday.xlsx

-----------------------+
| A1| fx        |NAME  |
------------------------------------------
[_] |    A     |    B            | C | D |
------------------------------------------
 1  | NAME     |EMAIL            |SEX|AGE|
 -- |----------+-----------------+---+---+
 2  |2025-09-15|Alice.gmail.com  | F |   |
 -- |--------+-------------------+---+---+
 3  |2025-09-26|Barbara.gmail.com| F | 17|
 -- |--------+-|-----------------+---+---+
 4  |2025-10-14|Alfred.gmail.com | F | 11|
 -- |--------+-|-----------------+---+---+
 5  |2025-10-28|Al.gmail.com     | F | 12|
 -- |----------+--------------------------
[MONDAY]


%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

%utlfkil(d:/xls/monday.xlsx); /*-- incase you rerun          --*/

libname xlsinp excel "d:/xls/monday.xlsx";

proc datasets lib=xlsinp; /*-- incase you rerun              --*/
 delete monday;
run;quit;

data xlsinp.monday;
 informat date $10. email $19. sex $2.;
 input date email sex age;
cards4;
2025-09-15 Alice@gmail.com  F 15
2025-09-26 Barbara@gmail.com F 17
2025-10-14 Alfred@gmail.com F 11
2025-10-28 Al@gmail.com F 12
;;;;
run;quit;

libname xlsinp clear;

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

862
5863      proc datasets lib=xlsinp; /*-- incase you rerun              --*/
NOTE: No matching members in directory
CREATE INDENTED SEX COLUMN
START READING AT ROW 4

The DATASETS Procedure

   Directory

Libref    XLSINP
Engine    OLEDB
5864       delete monday;
5865      run;quit;
NOTE: XLSINP.MONDAY (memtype="DATA") was not found, and has not been deleted
NOTE: Procedure datasets step took :
      real time : 0.206
      cpu time  : 0.062


5866
5867      data xlsinp.monday;
5868       informat date $10. email $19. sex $2.;
5869       input date email sex age;
5870      cards4;

NOTE: Data set "XLSINP.monday" has an unknown number of o  ervation(s) and 4 variable(s)
NOTE: The data step took :
      real time : 0.219
      cpu time  : 0.046


5871      2025-09-15 Alice@gmail.com  F 15
5872      2025-09-26 Barbara@gmail.com F 17
5873      2025-10-14 Alfred@gmail.com F 11
5874      2025-10-28 Al.gmail@com F 12
5875      ;;;;
5876      run;quit;
NOTE: Libref XLSINP has been deassigned.
5877
5878      libname xlsinp clear;
5879      quit; run;
5880      ODS _ALL_ CLOSE;
5881      FILENAME WPSWBHTM CLEAR;

/*---
 _ __  _ __ ___   ___ ___  ___ ___
| `_ \| `__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|

We create a routine, prepxls with input and output arguments.
Put the macro in your autocall library and call using your
new inputs and outputs.

%prepxls(
   inp_workbook = d:/xls/monday.xlsx
  ,inp_sheet    = monday
  ,out_workbook = d:/xls/want.xlsx
  ,out_sheet    = want
  );
---*/

%utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
&_init_;

%macro prepxls(
   inp_workbook = d:/xls/monday.xlsx
  ,inp_sheet    = monday
  ,out_workbook = d:/xls/want.xlsx
  ,out_sheet    = want
  );

   %utlfkil(&out_workbook); /*-- incase you rerun --*/

   libname xlsinp excel "&inp_workbook";
   libname xlsout excel "&out_workbook";

   proc datasets lib=xlsout;
    delete want;
   run;quit;

   data xlsout.&outsheett;
     retain year domain;
     set xlsinp.&inp_sheet;
     domain = scan(email,2,'@');
     year   = scan(date,1,'.');
   run;quit;

   libname xlsinp clear;
   libname xlsout clear;

%mend prepxls;

%prepxls(
   inp_workbook = d:/xls/monday.xlsx
  ,inp_sheet    = monday
  ,out_workbook = d:/xls/want.xlsx
  ,out_sheet    = want
  );

OUTPUT
======

COLUMNS YEAR AND DOMAIN ADDED

d:/xls/want.xlsx

-----------------+
| A1| fx  |NAME  |
---------------------------------------------------------
[_] |  A |    B    |    C     |    C            | E | F |
---------------------------------------------------------
 1  |YEAR|DOMAIN   | NAME     |EMAIL            |SEX|AGE|
 -- |----+---------+----------+-----------------+---+---+
 2  |2025|gmail.com|2025-09-15|Alice.gmail.com  | F |   |
 -- |--------------+--------+-------------------+---+---+
 3  |2025|gmail.com|2025-09-26|Barbara.gmail.com| F | 17|
 -- |----|---------+--------+-|-----------------+---+---+
 4  |2025|gmail.com|2025-10-14|Alfred.gmail.com | F | 11|
 -- |----|---------+--------+-|-----------------+---+---+
 5  |2025|gmail.com|2025-10-28|Al.gmail.com     | F | 12|
 -- |----+--------------------+--------------------------
[WANT}

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

233      ODS _ALL_ CLOSE;
6234      FILENAME WPSWBHTM TEMP;
NOTE: Writing HTML(WBHTML) BODY file d:\wpswrk\_TD9416\#LN00279
6235      ODS HTML(ID=WBHTML) BODY=WPSWBHTM GPATH="d:\wpswrk\_TD9416";
6236      %utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
6237      &_init_;
6238
6239      %macro prepxls(
6240         inp_workbook = d:/xls/monday.xlsx
6241        ,inp_sheet    = monday
6242        ,out_workbook = d:/xls/want.xlsx
6243        ,out_sheet    = want
6244        );
6245
6246         %utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
6247         &_init_;
6248
6249         %utlfkil(d:/xls/want.xlsx); /*-- incase you rerun --*/
6250
6251         libname xlsinp excel "d:/xls/monday.xlsx";
6252         libname xlsout excel "d:/xls/want.xlsx";
6253
6254         proc datasets lib=xlsout;
6255          delete want;
6256         run;quit;
6257
6258         data xlsout.want;
6259           retain year domain;
6260           set xlsinp.monday;
6261           domain = scan(email,2,'@');
6262           year   = scan(date,1,'.');
6263         run;quit;
6264
6265         libname xlsinp clear;
6266         libname xlsout clear;
6267
6268      %mend prepxls;
6269
6270      %prepxls(
6271         inp_workbook = d:/xls/monday.xlsx
6272        ,inp_sheet    = monday
6273        ,out_workbook = d:/xls/want.xlsx
6274        ,out_sheet    = want
6275        );
The file d:/xls/want.xlsx does not exist
NOTE: Library xlsinp assigned as follows:
      Engine:        OLEDB
      Physical Name: d:/xls/monday.xlsx

NOTE: Library xlsout assigned as follows:
      Engine:        OLEDB
      Physical Name: d:/xls/want.xlsx

NOTE: No matching members in directory
CREATE INDENTED SEX COLUMN
START READING AT ROW 4

The DATASETS Procedure

   Directory

Libref    XLSOUT
Engine    OLEDB
NOTE: XLSOUT.WANT (memtype="DATA") was not found, and has not been deleted
NOTE: Procedure datasets step took :
      real time : 0.205
      cpu time  : 0.078



NOTE: 4 observations were read from "XLSINP.monday"
NOTE: Data set "XLSOUT.want" has an unknown number of observation(s) and 6 variable(s)
NOTE: The data step took :
      real time : 0.330
      cpu time  : 0.187


NOTE: Libref XLSINP has been deassigned.
NOTE: Libref XLSOUT has been deassigned.
6276      quit; run;
6277      ODS _ALL_ CLOSE;
6278      FILENAME WPSWBHTM CLEAR;

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
