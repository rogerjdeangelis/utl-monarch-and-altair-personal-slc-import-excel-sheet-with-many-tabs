%let pgm=utl-monarch-and-altair-personal-slc-import-excel-sheet-with-many-tabs;

%top_submission;

Monarch and altair personal slc import excel sheet with many tabs

too long to post to a listserv see github
github
https://github.com/rogerjdeangelis/utl-monarch-and-altair-personal-slc-import-excel-sheet-with-many-tabs

Too long for listserves, see github

CONTENTS
  1 slc excel engine
  2 do_over macro (in this repor and im many macros)
    https://github.com/rogerjdeangelis/utl-monarch-and-altair-personal-slc-import-excel-sheet-with-many-tabs/blob/main/DO_OVER.SAS

community.altair.com
https://community.altair.com/discussion/comment/68603?tab=all#Comment_68603?utm_source=community-search&utm_medium=organic-search&utm_term=monarch

macros
https://github.com/rogerjdeangelis/utl-macros-used-in-many-of-rogerjdeangelis-repositories

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

d:/xls/manytabs.xlsx

Three tabs

------------------+
| A1| fx  NAME    |
-------------------
[_]  A  | B  | C  |
-------------------
1 |NAME | SEX| AGE|
--|-----+----+----+
2 |Alex | F  | 14 |
--|-----+----+----+
3 |Alice| F  | 13 |
--|-----+----+----+
[F]

------------------+
| A1| fx  NAME    |
-------------------
[_]   A  | B  | C |
-------------------
1 |NAME | SEX| AGE|
--|-----+----+----+
4 |Al   | M  | 13 |
--|-----+----+----+
5 |Jack | M  | 14 |
--|-----+----+----+
[M]

------------------+
| A1| fx  NAME    |
-------------------
[_]  A  | B  | C  |
--|-----+----+----+
6 |Henry| X  | 14 |
--|-----+----+----+
7 |James| X  | 12 |
--|-----+----+----+
[X]

&_init_;
data class;
  input
    name$
    sex$ age;
cards4;
Alex    F 14
Alice   F 13
Al      M 13
George  M 14
Henry   X 14
James   X 12
;;;;
run;quit;

/*--- CREATE WORKBOOK WITH 3 TABS ---*/
%utlfkil(d:/xls/manytabs.xlsx);

ods excel
 file="d:/xls/manytabs.xlsx"
  options(
  sheet_interval='bygroup'
  sheet_name= "#byval1");

proc report data=class;
  by sex;
run;quit;
ods excel close;

/*       _                          _                    _
/ |  ___| | ___    _____  _____ ___| |   ___ _ __   __ _(_)_ __   ___
| | / __| |/ __|  / _ \ \/ / __/ _ \ |  / _ \ `_ \ / _` | | `_ \ / _ \
| | \__ \ | (__  |  __/>  < (_|  __/ | |  __/ | | | (_| | | | | |  __/
|_| |___/_|\___|  \___/_/\_\___\___|_|  \___|_| |_|\__, |_|_| |_|\___|
                                                   |___/
*/

&_init_;

libname myxls xlsx "d:/xls/manytabs.xlsx";

proc datasets nolist nodetails;
 delete f m x;
run;quit;

/*--- create macro array ----*/
proc sql;
  select
    memname
  into
    :tabs1-
  from
    dictionary.tables
  where
    libname = "MYXLS"
;quit;

%let tabsn=&sqlobs;

%put &=tabs1;  /* F */
%put &=tabs2;  /* M */
%put &=tabs3;  /* X */
%put &=tabsn;  /* 3 */

%do_over(tabs,phrase=%str(
 data ?;
   set myxls.?(firstobs=2);
 run;quit;
 ));

&_init_;
proc print data=f;
proc print data=m;
proc print data=x;
run;quit;

* OUTPUT ;
* ====== ;

work.F

Obs    SEX_F    VAR2       VAR3

 1      1       Alice      13
 2      2       Barbara    13

work.M

Obs    SEX_M     VAR2    VAR3

 1     Al        M       13
 2     George    M       14

work.X

Obs    SEX_X    VAR2    VAR3

 1     Henry    X       14
 2     James    X       12

/*___                                         _
|___ \   _ __ ___   __ _  ___ _ __ ___     __| |    _____   _____ _ __
  __) | | `_ ` _ \ / _` |/ __| `__/ _ \   / _` |   / _ \ \ / / _ \ `__|
 / __/  | | | | | | (_| | (__| | | (_) | | (_| |  | (_) \ V /  __/ |
|_____| |_| |_| |_|\__,_|\___|_|  \___/   \__,_|___\___/ \_/ \___|_|
                                              |_____|
*/

https://github.com/rogerjdeangelis/utl-monarch-and-altair-personal-slc-import-excel-sheet-with-many-tabs/blob/main/DO_OVER.SAS

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
