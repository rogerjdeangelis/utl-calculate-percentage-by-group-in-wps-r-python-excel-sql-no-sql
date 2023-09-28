# utl-calculate-percentage-by-group-in-wps-r-python-excel-sql-no-sql
Calculate percentage by group in wps r python excel sql no sql 
    %let pgm=utl-calculate-percentage-by-group-in-wps-r-python-excel-sql-no-sql;

    Calculate percentage by group in wps r python excel sql no sql
    
    https://tinyurl.com/2r8ffh37
    https://github.com/rogerjdeangelis/utl-calculate-percentage-by-group-in-wps-r-python-excel-sql-no-sql

    https://stackoverflow.com/questions/77189587/scale-values-in-one-column-based-on-separate-column

       EIGHT  SOLUTIONS

            1 wps proc freq
              proc freq data=sd1.have;
                by project;
                tables reaction/ out=sd1.want;
                weight cnt;
            2 wps DOW loopl
            3 wps proc sql
              select
                project
               ,reaction
               ,cnt
               ,100*cnt/sum(cnt) as percentage
              from
                sd1.have
              group
                 by project
            4 wps excel sql passthru
              Inplace update to workbook. MS access passthru sql to excel
              Interesting sql column functions like 'isnumeric'
              Check workbook before importing?
              python and R excel packages more functional then either SAS or WPS?
            5 wps r sql
            6 wps r
              https://stackoverflow.com/users/13321647/tarjae
              have %>% ungroup() %>% mutate(percentage = (100*CNT / sum(CNT)), .by=PROJECT);
            7 wps python no sql
              https://stackoverflow.com/users/6287308/vaishali
              have["PERCENT"] = have.groupby("PROJECT")["CNT"].apply(lambda x: x*100/x.sum());
            8 wps python sql

      MS access sql reference on the end, useful for access or excel?

    /*                   _                                                _
    (_)_ __  _ __  _   _| |_ ___  __      ___ __  ___    _____  _____ ___| |
    | | `_ \| `_ \| | | | __/ __| \ \ /\ / / `_ \/ __|  / _ \ \/ / __/ _ \ |
    | | | | | |_) | |_| | |_\__ \  \ V  V /| |_) \__ \ |  __/>  < (_|  __/ |
    |_|_| |_| .__/ \__,_|\__|___/   \_/\_/ | .__/|___/  \___/_/\_\___\___|_|
            |_|                            |_|
    */

    /**************************************************************************************************************************/
    /*                                                        |                          |                                    */
    /*                                                        |                          |                                    */
    /* WPS INPUT SD1.HAVE total obs=20                        |     PROCESS              |   OUTPUT  ADD THIS COLUMN          */
    /*                                                        |                          |                                    */
    /* Obs  PROJECT  REACTION                            CNT  | percentage=cnt/sum(cnt)  |   PERCENT                          */
    /*                                                        |                          |                                    */
    /*   1    G39    sulfite-reduction                     9  | select                   |    3.1450                          */
    /*   2    G39    selenate-reduction                   38  |   project                |    0.3359                          */
    /*   3    G39    iron-oxidation                       85  |  ,reaction               |    4.7939                          */
    /*   4    G39    thiosulfate-disproportionation       13  |  ,n                      |    2.7481                          */
    /*   5    G39    methanol-oxidation                   17  |  ,n/sum(n) as normalize  |   70.1069                          */
    /*   6    G39    hydrogen-oxidation                 2296  | from                     |    2.5954                          */
    /*   7    G39    sulfide-oxidation                     8  |   sd1.have               |    1.0076                          */
    /*   8    G39    halogenated-compounds-breakdown      90  | group                    |    4.7634                          */
    /*   9    G39    sulfur-oxidation                    259  |    by project            |    0.5191                          */
    /*  10    G39    formaldehyde-oxidation              157  |                          |    1.1603                          */
    /*  11    G39    iron-reduction                       33  |                          |    0.2443                          */
    /*  12    G39    arsenate-reduction                  103  |                          |    0.2748                          */
    /*  13    G39    carbon-fixation                      11  |                          |    7.9084                          */
    /*  14    G39    manganese-oxidation                 156  |                          |    0.3969                          */
    /*                                                        |                          |  100.0000  Total Percent           */
    /*                                                        |                          |                                    */
    /*                                                        |                          |    3.5197                          */
    /*                                                        |                          |    0.8213                          */
    /*  15    G40    arsenate-reduction                   90  |                          |    7.8999                          */
    /*  16    G40    iron-oxidation                       73  |                          |    3.1678                          */
    /*  17    G40    hydrogen-oxidation                 2090  |                          |   81.7364                          */
    /*  18    G40    halogenated-compounds-breakdown      81  |                          |    2.8549                          */
    /*  19    G40    formaldehyde-oxidation              202  |                          |  100.0000  Total Percent           */
    /*  20    G40    carbon-fixation                      21  |                          |                                    */
    /*                                                        |                          |                                    */
    /**************************************************************************************************************************/
    /*                                       |                           |                                                    */
    /* EXCEL INPUT                           | PROCESS                   |  OUTPUT NAMED RANGE PERCENTAGE IN SHEET 2          */
    /*                                       |                           |                                                    */
    /* EXCEL WORKBOOKd:/xls/have.xlsx        | Create ODBC connection    |  EXCEL WORKBOOKd:/xls/have.xlsx (named ranges)     */
    /*                                       | Create passthru sql query |                                                    */
    /*   +-------------                      | and add percentage column |    +-------------                                  */
    /* 1 |  PROJECT   | =>named range        | Send output back to excel |  1 | PERCENTAGE | =>named range                    */
    /*   +------------+                      |                           |    +------------+                                  */
    /*                                       |                           |                                                    */
    /*   +---------------------------------+ |                           |    +--------------------------------------------+  */
    /*   |     A  |         B         |  C | |                           |    |     A  |         B         |   C   |   D   |  */
    /*   +---------------------------------+ |                           |    +--------------------------------------------+  */
    /* 1 | PROJECT|      REACTION     | CNT| |                           |  1 | PROJECT|      REACTION     |  CNT  |PERCENT|  */
    /*   +--------+-----+-------------+----+ |                           |    +--------+-----+-------------+-------+-------+  */
    /* 2 | G39    |selenate-reduction | 22 | |                           |  2 | G39    |selenate-reduction |  22   | 1.03  |  */
    /*   +--------+-----+-------------+----+ |                           |    +--------+-----+-------------+-------+-------+  */
    /* 3 | G39    |iron-oxidation     | 31 | |                           |  3 | G39    |iron-oxidation     |  31   | 1.50  |  */
    /*   +--------+-----+-------------+----+ |                           |    +--------+-----+-------------+-------+-------+  */
    /* 4 | G40    |methanol-oxidation | 44 | |                           |  4 | G40    |methanol-oxidation |  44   | 3/33  |  */
    /*   +--------+-----+-------------+----+ |                           |    +--------+-----+-------------+-------+-------+  */
    /*                                       |                           |                                                    */
    /* [SHEET 1]                             |                           |   [SHEET 2]                                        */
    /*                                       |                           |                                                    */
    /**************************************************************************************************************************/

    /*                   _
    (_)_ __  _ __  _   _| |_ ___
    | | `_ \| `_ \| | | | __/ __|
    | | | | | |_) | |_| | |_\__ \
    |_|_| |_| .__/ \__,_|\__|___/
            |_|               _       _                 _
    __      ___ __  ___    __| | __ _| |_ __ _ ___  ___| |_
    \ \ /\ / / `_ \/ __|  / _` |/ _` | __/ _` / __|/ _ \ __|
     \ V  V /| |_) \__ \ | (_| | (_| | || (_| \__ \  __/ |_
      \_/\_/ | .__/|___/  \__,_|\__,_|\__\__,_|___/\___|\__|
             |_|
    */

    data sd1.have;informat
    PROJECT $9.
    REACTION $31.
    CNT 8.
    ;input
    PROJECT REACTION CNT;
    cards4;
    G39 arsenate-reduction 103
    G39 carbon-fixation 11
    G39 formaldehyde-oxidation 157
    G39 halogenated-compounds-breakdown 90
    G39 hydrogen-oxidation 2296
    G39 iron-oxidation 85
    G39 iron-reduction 33
    G39 manganese-oxidation 156
    G39 methanol-oxidation 17
    G39 selenate-reduction 38
    G39 sulfide-oxidation 8
    G39 sulfite-reduction 9
    G39 sulfur-oxidation 259
    G39 thiosulfate-disproportionation 13
    G40 arsenate-reduction 90
    G40 carbon-fixation 21
    G40 formaldehyde-oxidation 202
    G40 halogenated-compounds-breakdown 81
    G40 hydrogen-oxidation 2090
    G40 iron-oxidation 73
    ;;;;
    run;quit;

    /*                                      _   _        _     _
    __      ___ __  ___    _____  _____ ___| | | |_ __ _| |__ | | ___
    \ \ /\ / / `_ \/ __|  / _ \ \/ / __/ _ \ | | __/ _` | `_ \| |/ _ \
     \ V  V /| |_) \__ \ |  __/>  < (_|  __/ | | || (_| | |_) | |  __/
      \_/\_/ | .__/|___/  \___/_/\_\___\___|_|  \__\__,_|_.__/|_|\___|
             |_|
    */

    /*----                                                                   ----*/
    /*---- Create odbc connection (if exist remove previous dsn )            ----*/
    /*---- Longline macro variable not needed if editor suppport the length   ---*/
    /*----                                                                   ----*/

    %let longline=%str(Add-OdbcDsn -Name 'have' -DriverName 'Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)'
     -DsnType 'User' -Platform '64-bit' -SetPropertyValue 'Dbq=d:\xls\have.xlsx');

    options ls=255;
    %put "&=longline";

    %utl_submit_ps64("
    Remove-OdbcDsn -Name 'have' -DsnType 'User' -Platform '64-bit';
    &longline;
    Get-OdbcDsn;
    ");


    /*----                                                                   ----*/
    /*---- Create named ranges "have" in excel workbook d:/xls/have.xlsx     ----*/
    /*----                                                                   ----*/

    %utlfkil(d:/xls/have.xlsx); * delete if exist - it works with an existing workbook;

    %utl_submit_wps64x('
    libname sd1 "d:/sd1";
    proc r;
     export data=sd1.have  r=have;
    submit;
    library(openxlsx);
    wb <- createWorkbook("d:/xls/have.xlsx");
    addWorksheet(wb, "sheet 1");
    writeData(wb, sheet = 1, x = have, startCol = 1, startRow = 1);
    createNamedRegion(
      wb = wb,
      sheet = 1,
      name = "have",
      rows = 1:(nrow(have) + 1),
      cols = 1:ncol(have)
    );
    addWorksheet(wb, "sheet 2");
    saveWorkbook(wb,"d:/xls/have.xlsx", overwrite = TRUE);
    endsubmit;
    ');

    /*         _       _   _
     ___  ___ | |_   _| |_(_) ___  _ __  ___
    / __|/ _ \| | | | | __| |/ _ \| `_ \/ __|
    \__ \ (_) | | |_| | |_| | (_) | | | \__ \
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|___/
     _                                                __
    / | __      ___ __  ___   _ __  _ __ ___   ___   / _|_ __ ___  __ _
    | | \ \ /\ / / `_ \/ __| | `_ \| `__/ _ \ / __| | |_| `__/ _ \/ _` |
    | |  \ V  V /| |_) \__ \ | |_) | | | (_) | (__  |  _| | |  __/ (_| |
    |_|   \_/\_/ | .__/|___/ | .__/|_|  \___/ \___| |_| |_|  \___|\__, |
                 |_|         |_|                                     |_|
    */

    proc datasets lib=sd1 nolist nodetails;delete want; run;quit;

    %utl_submit_wps64x('
    libname sd1 "d:/sd1";
    proc freq data=sd1.have;
      by project;
      tables reaction/ out=sd1.want;
      weight cnt;
    run;quit;
    ');

    proc print data=sd1.want;
    run;quit;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /* WPS PrOC FREQ                                                                                                          */
    /*                                                                                                                        */
    /* Obs    PROJECT    REACTION                           COUNT    PERCENT                                                  */
    /*                                                                                                                        */
    /*   1      G39      arsenate-reduction                   103     3.1450                                                  */
    /*   2      G39      carbon-fixation                       11     0.3359                                                  */
    /*   3      G39      formaldehyde-oxidation               157     4.7939                                                  */
    /*   4      G39      halogenated-compounds-breakdown       90     2.7481                                                  */
    /*   5      G39      hydrogen-oxidation                  2296    70.1069                                                  */
    /*   6      G39      iron-oxidation                        85     2.5954                                                  */
    /*   7      G39      iron-reduction                        33     1.0076                                                  */
    /*   8      G39      manganese-oxidation                  156     4.7634                                                  */
    /*   9      G39      methanol-oxidation                    17     0.5191                                                  */
    /*  10      G39      selenate-reduction                    38     1.1603                                                  */
    /*  11      G39      sulfide-oxidation                      8     0.2443                                                  */
    /*  12      G39      sulfite-reduction                      9     0.2748                                                  */
    /*  13      G39      sulfur-oxidation                     259     7.9084                                                  */
    /*  14      G39      thiosulfate-disproportionation        13     0.3969                                                  */
    /*                                                                                                                        */
    /*  15      G40      arsenate-reduction                    90     3.5197                                                  */
    /*  16      G40      carbon-fixation                       21     0.8213                                                  */
    /*  17      G40      formaldehyde-oxidation               202     7.8999                                                  */
    /*  18      G40      halogenated-compounds-breakdown       81     3.1678                                                  */
    /*  19      G40      hydrogen-oxidation                  2090    81.7364                                                  */
    /*  20      G40      iron-oxidation                        73     2.8549                                                  */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*___                             _                 _
    |___ \  __      ___ __  ___    __| | _____      __ | | ___   ___  _ __
      __) | \ \ /\ / / `_ \/ __|  / _` |/ _ \ \ /\ / / | |/ _ \ / _ \| `_ \
     / __/   \ V  V /| |_) \__ \ | (_| | (_) \ V  V /  | | (_) | (_) | |_) |
    |_____|   \_/\_/ | .__/|___/  \__,_|\___/ \_/\_/   |_|\___/ \___/| .__/
                     |_|                                             |_|
    */
    proc datasets lib=sd1 nolist nodetails;delete want; run;quit;

    %utl_submit_wps64x('
    libname sd1 "d:/sd1";

    data sd1.want;

     retain tot 0;

     do until (last.project);
       set sd1.have;
       by project;
       tot=tot+cnt;
     end;

     do until (last.project);
       set sd1.have;
       by project;
       percent=100*cnt/tot;
       output;
     end;
     tot =0;
     drop tot;
    run;quit;
    ');

    proc print data=sd1.want;
    run;quit;


    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /* WPS DOW LOOP                                                                                                           */
    /*                                                                                                                        */
    /* Obs    PROJECT    REACTION                           COUNT    PERCENT                                                  */
    /*                                                                                                                        */
    /*   1      G39      arsenate-reduction                   103     3.1450                                                  */
    /*   2      G39      carbon-fixation                       11     0.3359                                                  */
    /*   3      G39      formaldehyde-oxidation               157     4.7939                                                  */
    /*   4      G39      halogenated-compounds-breakdown       90     2.7481                                                  */
    /*   5      G39      hydrogen-oxidation                  2296    70.1069                                                  */
    /*   6      G39      iron-oxidation                        85     2.5954                                                  */
    /*   7      G39      iron-reduction                        33     1.0076                                                  */
    /*   8      G39      manganese-oxidation                  156     4.7634                                                  */
    /*   9      G39      methanol-oxidation                    17     0.5191                                                  */
    /*  10      G39      selenate-reduction                    38     1.1603                                                  */
    /*  11      G39      sulfide-oxidation                      8     0.2443                                                  */
    /*  12      G39      sulfite-reduction                      9     0.2748                                                  */
    /*  13      G39      sulfur-oxidation                     259     7.9084                                                  */
    /*  14      G39      thiosulfate-disproportionation        13     0.3969                                                  */
    /*                                                                                                                        */
    /*  15      G40      arsenate-reduction                    90     3.5197                                                  */
    /*  16      G40      carbon-fixation                       21     0.8213                                                  */
    /*  17      G40      formaldehyde-oxidation               202     7.8999                                                  */
    /*  18      G40      halogenated-compounds-breakdown       81     3.1678                                                  */
    /*  19      G40      hydrogen-oxidation                  2090    81.7364                                                  */
    /*  20      G40      iron-oxidation                        73     2.8549                                                  */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*____                                  _
    |___ /  __      ___ __  ___   ___  __ _| |
      |_ \  \ \ /\ / / `_ \/ __| / __|/ _` | |
     ___) |  \ V  V /| |_) \__ \ \__ \ (_| | |
    |____/    \_/\_/ | .__/|___/ |___/\__, |_|
                     |_|                 |_|
    */

    proc datasets lib=sd1 nolist nodetails;delete want; run;quit;

    %utl_submit_wps54x('
    options validvarname=any;
    libname sd1 "d:/sd1";
    proc sql;
      create
        table sd1.want as
      select
        project
       ,reaction
       ,cnt
       ,100*cnt/sum(cnt) as percentage
      from
        sd1.have
      group
         by project
    ;quit;
    ');

    proc print data=sd1.want;
    run;quit;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*  Obs    PROJECT    REACTION                            CNT   PERCENTAGE                                                */
    /*                                                                                                                        */
    /*    1      G39      arsenate-reduction                  103      3.1450                                                 */
    /*    2      G39      carbon-fixation                      11      0.3359                                                 */
    /*    3      G39      formaldehyde-oxidation              157      4.7939                                                 */
    /*    4      G39      halogenated-compounds-breakdown      90      2.7481                                                 */
    /*    5      G39      hydrogen-oxidation                 2296     70.1069                                                 */
    /*    6      G39      iron-oxidation                       85      2.5954                                                 */
    /*    7      G39      iron-reduction                       33      1.0076                                                 */
    /*    8      G39      manganese-oxidation                 156      4.7634                                                 */
    /*    9      G39      methanol-oxidation                   17      0.5191                                                 */
    /*   10      G39      selenate-reduction                   38      1.1603                                                 */
    /*   11      G39      sulfide-oxidation                     8      0.2443                                                 */
    /*   12      G39      sulfite-reduction                     9      0.2748                                                 */
    /*   13      G39      sulfur-oxidation                    259      7.9084                                                 */
    /*   14      G39      thiosulfate-disproportionation       13      0.3969                                                 */
    /*                                                                                                                        */
    /*   15      G40      arsenate-reduction                   90      3.5197                                                 */
    /*   16      G40      carbon-fixation                      21      0.8213                                                 */
    /*   17      G40      formaldehyde-oxidation              202      7.8999                                                 */
    /*   18      G40      halogenated-compounds-breakdown      81      3.1678                                                 */
    /*   19      G40      hydrogen-oxidation                 2090     81.7364                                                 */
    /*   20      G40      iron-oxidation                       73      2.8549                                                 */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*  _                                            _                       _   _
    | || |   __      ___ __  ___    _____  _____ ___| |  _ __   __ _ ___ ___| |_| |__  _ __ _   _
    | || |_  \ \ /\ / / `_ \/ __|  / _ \ \/ / __/ _ \ | | `_ \ / _` / __/ __| __| `_ \| `__| | | |
    |__   _|  \ V  V /| |_) \__ \ |  __/>  < (_|  __/ | | |_) | (_| \__ \__ \ |_| | | | |  | |_| |
       |_|     \_/\_/ | .__/|___/  \___/_/\_\___\___|_| | .__/ \__,_|___/___/\__|_| |_|_|   \__,_|
                      |_|                               |_|
    */


    /*----                                                                   ----*/
    /*---- Using ms sql excel query inside excel                             ----*/
    /*---- ODBC connection created above on input                            ----*/
    /*----                                                                   ----*/

    %utl_submit_wps64x('
    libname sd1 "d:/sd1";
    proc r;
    submit;
    library(RODBC);
    library(openxlsx);
    ch <- odbcConnect("have");
    want<-sqlQuery(ch,"
        select
           l.project
          ,l.reaction
          ,l.cnt
          ,100*l.cnt/r.tot as percentage
        from
           (select project, reaction, cnt from have) as l
        left join
           (select project, sum(cnt) as tot from have group by project) as r
        on
           l.project = r.project
        ");
    want;
    odbcClose(ch);
    wb <- loadWorkbook("d:/xls/have.xlsx");
    writeData(wb, sheet = 2, x =want , startCol = 1, startRow = 1);
    createNamedRegion(
      wb = wb,
      sheet = 2,
      name = "percentage",
      rows = 1:(nrow(want) + 1),
      cols = 1:ncol(want)
    );
    saveWorkbook(wb,"d:/xls/have.xlsx",overwrite=TRUE);
    endsubmit;
    ');

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /* OUTPUT NAMED RANGE PERCENTAGE IN SHEET2                                                                                */
    /*                                                                                                                        */
    /* EXCEL WORKBOOKd:/xls/have.xlsx (named ranges)                                                                          */
    /*                                                                                                                        */
    /*   +-------------                                                                                                       */
    /* 1 | PERCENTAGE | =>named range                                                                                         */
    /*   +------------+                                                                                                       */
    /*                                                                                                                        */
    /*   +------------------------------------------------+                                                                   */
    /*   |     A  |         B         |   C   |   D       |                                                                   */
    /*   +------------------------------------------------+                                                                   */
    /* 1 | PROJECT|      REACTION     |  CNT  |PERCENTAGE |                                                                   */
    /*   +--------+-----+-------------+-------+-----------+                                                                   */
    /* 2 | G39    |selenate-reduction |  22   | 1.03      |                                                                   */
    /*   +--------+-----+-------------+-------+-----------+                                                                   */
    /* 3 | G39    |iron-oxidation     |  31   | 1.50      |                                                                   */
    /*   +--------+-----+-------------+-------+-----------+                                                                   */
    /* 4 | G40    |methanol-oxidation |  44   | 3/33      |                                                                   */
    /*   +--------+-----+-------------+-------+-----------+                                                                   */
    /*                                                                                                                        */
    /*  [SHEET 2]                                                                                                             */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*___                                          _
    | ___|  __      ___ __  ___   _ __   ___  __ _| |
    |___ \  \ \ /\ / / `_ \/ __| | `__| / __|/ _` | |
     ___) |  \ V  V /| |_) \__ \ | |    \__ \ (_| | |
    |____/    \_/\_/ | .__/|___/ |_|    |___/\__, |_|
                     |_|                        |_|
    */

    proc datasets lib=sd1 nolist nodetails;delete want; run;quit;

    %utl_submit_wps64x('
    libname sd1 "d:/sd1";
    proc r;
    export data=sd1.have r=have;
    submit;
    library(sqldf);
    want <- sqldf("
        select
           l.project
          ,l.reaction
          ,l.cnt
          ,100*l.cnt/r.tot as percentage
        from
           (select project, reaction, cnt from have) as l
        left join
           (select project, sum(cnt) as tot from have group by project) as r
        on
           l.project = r.project
        ");
    want;
    endsubmit;
    import data=sd1.want r=want;
    ');

    proc print data=sd1.want;
    run;quit;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /* PROJECT    REACTION                            CNT    PERCENTAGE                                                       */
    /*                                                                                                                        */
    /*   G39      arsenate-reduction                  103       3.1450                                                        */
    /*   G39      carbon-fixation                      11       0.3359                                                        */
    /*   G39      formaldehyde-oxidation              157       4.7939                                                        */
    /*   G39      halogenated-compounds-breakdown      90       2.7481                                                        */
    /*   G39      hydrogen-oxidation                 2296      70.1069                                                        */
    /*   G39      iron-oxidation                       85       2.5954                                                        */
    /*   G39      iron-reduction                       33       1.0076                                                        */
    /*   G39      manganese-oxidation                 156       4.7634                                                        */
    /*   G39      methanol-oxidation                   17       0.5191                                                        */
    /*   G39      selenate-reduction                   38       1.1603                                                        */
    /*   G39      sulfide-oxidation                     8       0.2443                                                        */
    /*   G39      sulfite-reduction                     9       0.2748                                                        */
    /*   G39      sulfur-oxidation                    259       7.9084                                                        */
    /*   G39      thiosulfate-disproportionation       13       0.3969                                                        */
    /*   G40      arsenate-reduction                   90       3.5197                                                        */
    /*   G40      carbon-fixation                      21       0.8213                                                        */
    /*   G40      formaldehyde-oxidation              202       7.8999                                                        */
    /*   G40      halogenated-compounds-breakdown      81       3.1678                                                        */
    /*   G40      hydrogen-oxidation                 2090      81.7364                                                        */
    /*   G40      iron-oxidation                       73       2.8549                                                        */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    proc datasets lib=sd1 nolist nodetails;delete want; run;quit;

    %utl_submit_wps64x('
    libname sd1 "d:/sd1";
    proc r;
    export data=sd1.have r=have;
    submit;
    library(dplyr);
    want<-have %>%
      ungroup() %>%
      mutate(percentage = (100*CNT / sum(CNT)), .by=PROJECT);
    want;
    endsubmit;
    import data=sd1.want r=want;
    run;quit;
    ');

    proc print data=sd1.want;
    run;quit;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*  The WPS Pro RSystem                                                                                                   */
    /*                                                                                                                        */
    /*     PROJECT                        REACTION  CNT percentage                                                            */
    /*  1      G39              arsenate-reduction  103  3.1450382                                                            */
    /*  2      G39                 carbon-fixation   11  0.3358779                                                            */
    /*  3      G39          formaldehyde-oxidation  157  4.7938931                                                            */
    /*  4      G39 halogenated-compounds-breakdown   90  2.7480916                                                            */
    /*  5      G39              hydrogen-oxidation 2296 70.1068702                                                            */
    /*  6      G39                  iron-oxidation   85  2.5954198                                                            */
    /*  7      G39                  iron-reduction   33  1.0076336                                                            */
    /*  8      G39             manganese-oxidation  156  4.7633588                                                            */
    /*  9      G39              methanol-oxidation   17  0.5190840                                                            */
    /*  10     G39              selenate-reduction   38  1.1603053                                                            */
    /*  11     G39               sulfide-oxidation    8  0.2442748                                                            */
    /*  12     G39               sulfite-reduction    9  0.2748092                                                            */
    /*  13     G39                sulfur-oxidation  259  7.9083969                                                            */
    /*  14     G39  thiosulfate-disproportionation   13  0.3969466                                                            */
    /*  15     G40              arsenate-reduction   90  3.5197497                                                            */
    /*  16     G40                 carbon-fixation   21  0.8212749                                                            */
    /*  17     G40          formaldehyde-oxidation  202  7.8998827                                                            */
    /*  18     G40 halogenated-compounds-breakdown   81  3.1677747                                                            */
    /*  19     G40              hydrogen-oxidation 2090 81.7364099                                                            */
    /*  20     G40                  iron-oxidation   73  2.8549081                                                            */
    /*                                                                                                                        */
    /* WPS                                                                                                                    */
    /*                                                                                                                        */
    /* Obs    PROJECT    REACTION                            CNT    PERCENTAGE                                                */
    /*                                                                                                                        */
    /*   1      G39      arsenate-reduction                  103       3.1450                                                 */
    /*   2      G39      carbon-fixation                      11       0.3359                                                 */
    /*   3      G39      formaldehyde-oxidation              157       4.7939                                                 */
    /*   4      G39      halogenated-compounds-breakdown      90       2.7481                                                 */
    /*   5      G39      hydrogen-oxidation                 2296      70.1069                                                 */
    /*   6      G39      iron-oxidation                       85       2.5954                                                 */
    /*   7      G39      iron-reduction                       33       1.0076                                                 */
    /*   8      G39      manganese-oxidation                 156       4.7634                                                 */
    /*   9      G39      methanol-oxidation                   17       0.5191                                                 */
    /*  10      G39      selenate-reduction                   38       1.1603                                                 */
    /*  11      G39      sulfide-oxidation                     8       0.2443                                                 */
    /*  12      G39      sulfite-reduction                     9       0.2748                                                 */
    /*  13      G39      sulfur-oxidation                    259       7.9084                                                 */
    /*  14      G39      thiosulfate-disproportionation       13       0.3969                                                 */
    /*  15      G40      arsenate-reduction                   90       3.5197                                                 */
    /*  16      G40      carbon-fixation                      21       0.8213                                                 */
    /*  17      G40      formaldehyde-oxidation              202       7.8999                                                 */
    /*  18      G40      halogenated-compounds-breakdown      81       3.1678                                                 */
    /*  19      G40      hydrogen-oxidation                 2090      81.7364                                                 */
    /*  20      G40      iron-oxidation                       73       2.8549                                                 */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*____                                    _   _                                          _
    |___  | __      ___ __  ___   _ __  _   _| |_| |__   ___  _ __    _ __   ___   ___  __ _| |
       / /  \ \ /\ / / `_ \/ __| | `_ \| | | | __| `_ \ / _ \| `_ \  | `_ \ / _ \ / __|/ _` | |
      / /    \ V  V /| |_) \__ \ | |_) | |_| | |_| | | | (_) | | | | | | | | (_) |\__ \ (_| | |
     /_/      \_/\_/ | .__/|___/ | .__/ \__, |\__|_| |_|\___/|_| |_| |_| |_|\___/ |___/\__, |_|
                     |_|         |_|    |___/                                             |_|
    */

    proc datasets lib=sd1 nolist nodetails;delete want; run;quit;

    %utl_submit_wps64x('
    libname sd1 "d:/sd1";
    proc python;
    export data=sd1.have python=have;
    submit;
    import pandas as pd;
    have["PERCENT"] = have.groupby("PROJECT")["CNT"].apply(lambda x: x*100/x.sum());
    print (have);
    endsubmit;
    import data=sd1.want python=have;
    ');

    proc print data=sd1.want;
    run;quit;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /* The PYTHON Procedure                                                                                                   */
    /*                                                                                                                        */
    /*       PROJECT                         REACTION     CNT    PERCENT                                                      */
    /* 0   G39        arsenate-reduction                103.0   3.145038                                                      */
    /* 1   G39        carbon-fixation                    11.0   0.335878                                                      */
    /* 2   G39        formaldehyde-oxidation            157.0   4.793893                                                      */
    /* 3   G39        halogenated-compounds-breakdown    90.0   2.748092                                                      */
    /* 4   G39        hydrogen-oxidation               2296.0  70.106870                                                      */
    /* 5   G39        iron-oxidation                     85.0   2.595420                                                      */
    /* 6   G39        iron-reduction                     33.0   1.007634                                                      */
    /* 7   G39        manganese-oxidation               156.0   4.763359                                                      */
    /* 8   G39        methanol-oxidation                 17.0   0.519084                                                      */
    /* 9   G39        selenate-reduction                 38.0   1.160305                                                      */
    /* 10  G39        sulfide-oxidation                   8.0   0.244275                                                      */
    /* 11  G39        sulfite-reduction                   9.0   0.274809                                                      */
    /* 12  G39        sulfur-oxidation                  259.0   7.908397                                                      */
    /* 13  G39        thiosulfate-disproportionation     13.0   0.396947                                                      */
    /* 14  G40        arsenate-reduction                 90.0   3.519750                                                      */
    /* 15  G40        carbon-fixation                    21.0   0.821275                                                      */
    /* 16  G40        formaldehyde-oxidation            202.0   7.899883                                                      */
    /* 17  G40        halogenated-compounds-breakdown    81.0   3.167775                                                      */
    /* 18  G40        hydrogen-oxidation               2090.0  81.736410                                                      */
    /* 19  G40        iron-oxidation                     73.0   2.854908                                                      */
    /*                                                                                                                        */
    /* WPS                                                                                                                    */
    /*                                                                                                                        */
    /*   PROJECT    REACTION                            CNT    PERCENT                                                        */
    /*                                                                                                                        */
    /*     G39      arsenate-reduction                  103     3.1450                                                        */
    /*     G39      carbon-fixation                      11     0.3359                                                        */
    /*     G39      formaldehyde-oxidation              157     4.7939                                                        */
    /*     G39      halogenated-compounds-breakdown      90     2.7481                                                        */
    /*     G39      hydrogen-oxidation                 2296    70.1069                                                        */
    /*     G39      iron-oxidation                       85     2.5954                                                        */
    /*     G39      iron-reduction                       33     1.0076                                                        */
    /*     G39      manganese-oxidation                 156     4.7634                                                        */
    /*     G39      methanol-oxidation                   17     0.5191                                                        */
    /*     G39      selenate-reduction                   38     1.1603                                                        */
    /*     G39      sulfide-oxidation                     8     0.2443                                                        */
    /*     G39      sulfite-reduction                     9     0.2748                                                        */
    /*     G39      sulfur-oxidation                    259     7.9084                                                        */
    /*     G39      thiosulfate-disproportionation       13     0.3969                                                        */
    /*     G40      arsenate-reduction                   90     3.5197                                                        */
    /*     G40      carbon-fixation                      21     0.8213                                                        */
    /*     G40      formaldehyde-oxidation              202     7.8999                                                        */
    /*     G40      halogenated-compounds-breakdown      81     3.1678                                                        */
    /*     G40      hydrogen-oxidation                 2090    81.7364                                                        */
    /*     G40      iron-oxidation                       73     2.8549                                                        */
    /*                                                                                                                        */
    /**************************************************************************************************************************/


    /*___                                     _   _                             _
     ( _ )  __      ___ __  ___   _ __  _   _| |_| |__   ___  _ __    ___  __ _| |
     / _ \  \ \ /\ / / `_ \/ __| | `_ \| | | | __| `_ \ / _ \| `_ \  / __|/ _` | |
    | (_) |  \ V  V /| |_) \__ \ | |_) | |_| | |_| | | | (_) | | | | \__ \ (_| | |
     \___/    \_/\_/ | .__/|___/ | .__/ \__, |\__|_| |_|\___/|_| |_| |___/\__, |_|
                     |_|         |_|    |___/                                |_|
    */

    proc datasets lib=sd1 nolist nodetails;delete want; run;quit;

    %utl_submit_wps64x("
    options validvarname=any ;
    libname sd1 'd:/sd1';
    proc python;
    export data=sd1.have python=have;
    submit;
    print(have);
    from os import path;
    import pandas as pd;
    import numpy as np;
    from pandasql import sqldf;
    mysql = lambda q: sqldf(q, globals());
    from pandasql import PandaSQL;
    pdsql = PandaSQL(persist=True);
    sqlite3conn = next(pdsql.conn.gen).connection.connection;
    sqlite3conn.enable_load_extension(True);
    sqlite3conn.load_extension('c:/temp/libsqlitefunctions.dll');
    mysql = lambda q: sqldf(q, globals());
    want = pdsql('''
        select
           l.project
          ,l.reaction
          ,l.cnt
          ,100*l.cnt/r.tot as percentage
        from
           (select project, reaction, cnt from have) as l
        left join
           (select project, sum(cnt) as tot from have group by project) as r
        on
           l.project = r.project
    ''');
    print(want);
    endsubmit;
    import data=sd1.want python=want;
    ');
    run;quit;
    "));

    proc print data=sd1.want;
    run;quit;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /* The PYTHON Procedure                                                                                                   */
    /*                                                                                                                        */
    /*       project                         reaction     cnt  percentage                                                     */
    /* 0   G39        arsenate-reduction                103.0    3.145038                                                     */
    /* 1   G39        carbon-fixation                    11.0    0.335878                                                     */
    /* 2   G39        formaldehyde-oxidation            157.0    4.793893                                                     */
    /* 3   G39        halogenated-compounds-breakdown    90.0    2.748092                                                     */
    /* 4   G39        hydrogen-oxidation               2296.0   70.106870                                                     */
    /* 5   G39        iron-oxidation                     85.0    2.595420                                                     */
    /* 6   G39        iron-reduction                     33.0    1.007634                                                     */
    /* 7   G39        manganese-oxidation               156.0    4.763359                                                     */
    /* 8   G39        methanol-oxidation                 17.0    0.519084                                                     */
    /* 9   G39        selenate-reduction                 38.0    1.160305                                                     */
    /* 10  G39        sulfide-oxidation                   8.0    0.244275                                                     */
    /* 11  G39        sulfite-reduction                   9.0    0.274809                                                     */
    /* 12  G39        sulfur-oxidation                  259.0    7.908397                                                     */
    /* 13  G39        thiosulfate-disproportionation     13.0    0.396947                                                     */
    /* 14  G40        arsenate-reduction                 90.0    3.519750                                                     */
    /* 15  G40        carbon-fixation                    21.0    0.821275                                                     */
    /* 16  G40        formaldehyde-oxidation            202.0    7.899883                                                     */
    /* 17  G40        halogenated-compounds-breakdown    81.0    3.167775                                                     */
    /* 18  G40        hydrogen-oxidation               2090.0   81.736410                                                     */
    /* 19  G40        iron-oxidation                     73.0    2.854908                                                     */
    /*                                                                                                                        */
    /* WPS                                                                                                                    */
    /*                                                                                                                        */
    /* Obs    project    reaction                            cnt    percentage                                                */
    /*                                                                                                                        */
    /*   1      G39      arsenate-reduction                  103       3.1450                                                 */
    /*   2      G39      carbon-fixation                      11       0.3359                                                 */
    /*   3      G39      formaldehyde-oxidation              157       4.7939                                                 */
    /*   4      G39      halogenated-compounds-breakdown      90       2.7481                                                 */
    /*   5      G39      hydrogen-oxidation                 2296      70.1069                                                 */
    /*   6      G39      iron-oxidation                       85       2.5954                                                 */
    /*   7      G39      iron-reduction                       33       1.0076                                                 */
    /*   8      G39      manganese-oxidation                 156       4.7634                                                 */
    /*   9      G39      methanol-oxidation                   17       0.5191                                                 */
    /*  10      G39      selenate-reduction                   38       1.1603                                                 */
    /*  11      G39      sulfide-oxidation                     8       0.2443                                                 */
    /*  12      G39      sulfite-reduction                     9       0.2748                                                 */
    /*  13      G39      sulfur-oxidation                    259       7.9084                                                 */
    /*  14      G39      thiosulfate-disproportionation       13       0.3969                                                 */
    /*  15      G40      arsenate-reduction                   90       3.5197                                                 */
    /*  16      G40      carbon-fixation                      21       0.8213                                                 */
    /*  17      G40      formaldehyde-oxidation              202       7.8999                                                 */
    /*  18      G40      halogenated-compounds-breakdown      81       3.1678                                                 */
    /*  19      G40      hydrogen-oxidation                 2090      81.7364                                                 */
    /*  20      G40      iron-oxidation                       73       2.8549                                                 */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    *                          _
     _ __ ___  ___   ___  __ _| |
    | '_ ` _ \/ __| / __|/ _` | |
    | | | | | \__ \ \__ \ (_| | |
    |_| |_| |_|___/ |___/\__, |_|
                            |_|
    ;


    https://ss64.com/access/

    a
      Abs             The absolute value of a number (nore negative sn).
     .AddMenu         Add a custom menu bar/shortcut bar.
     .AddNew          Add a new record to a recordset.
     .ApplyFilter     Apply a filter clause to a table, form, or report.
      Array           Create an Array.
      Asc             The Ascii code of a character.
      AscW            The Unicode of a character.
      Atn             Display the ArcTan of an angle.
      Avg (SQL)       Average.
    b
     .Beep (DoCmd)    Sound a tone.
     .BrowseTo(DoCmd) Navate between objects.
    c
      Call            Call a procedure.
     .CancelEvent (DoCmd) Cancel an event.
     .CancelUpdate    Cancel recordset changes.
      Case            If Then Else.
      CBool           Convert to boolean.
      CByte           Convert to byte.
      CCur            Convert to currency (number)
      CDate           Convert to Date.
      CVDate          Convert to Date.
      CDbl            Convert to Double (number)
      CDec            Convert to Decimal (number)
      Choose          Return a value from a list based on position.
      ChDir           Change the current directory or folder.
      ChDrive         Change the current drive.
      Chr             Return a character based on an ASCII code.
     .ClearMacroError (DoCmd) Clear MacroError.
     .Close (DoCmd)           Close a form/report/window.
     .CloseDatabase (DoCmd)   Close the database.
      CInt                    Convert to Integer (number)
      CLng                    Convert to Long (number)
      Command                 Return command line option string.
     .CopyDatabaseFile(DoCmd) Copy to an SQL .mdf file.
     .CopyObject (DoCmd)      Copy an Access database object.
      Cos                     Display Cosine of an angle.
      Count (SQL)             Count records.
      CSng             Convert to Single (number.)
      CStr             Convert to String.
      CurDir           Return the current path.
      CurrentDb        Return an object variable for the current database.
      CurrentUser      Return the current user.
      CVar             Convert to a Variant.
    d
      Date             The current date.
      DateAdd          Add a time interval to a date.
      DateDiff         The time difference between two dates.
      DatePart         Return part of a given date.
      DateSerial       Return a date given a year, month, and day.
      DateValue        Convert a string to a date.
      DAvg             Average from a set of records.
      Day              Return the day of the month.
      DCount           Count the number of records in a table/query.
      Delete (SQL)          Delete records.
     .DeleteObject (DoCmd)  Delete an object.
      DeleteSetting         Delete a value from the users registry
     .DoMenuItem (DoCmd)    Display a menu or toolbar command.
      DFirst           The first value from a set of records.
      Dir              List the files in a folder.
      DLast            The last value from a set of records.
      DLookup          Get the value of a particular field.
      DMax             Return the maximum value from a set of records.
      DMin             Return the minimum value from a set of records.
      DoEvents         Allow the operating system to process other events.
      DStDev           Estimate Standard deviation for domain (subset of records)
      DStDevP          Estimate Standard deviation for population (subset of records)
      DSum             Return the sum of values from a set of records.
      DVar             Estimate variance for domain (subset of records)
      DVarP            Estimate variance for population (subset of records)
    e
     .Echo             Turn screen updating on or off.
      Environ          Return the value of an OS environment variable.
      EOF              End of file input.
      Error            Return the error message for an error No.
      Eval             Evaluate an expression.
      Execute(SQL/VBA) Execute a procedure or run SQL.
      Exp              Exponential e raised to the nth power.
    f
      FileDateTime      Filename last modified date/time.
      FileLen           The size of a file in bytes.
     .FindFirst/Last/Next/Previous Record.
     .FindRecord(DoCmd) Find a specific record.
      First (SQL)       Return the first value from a query.
      Fix               Return the integer portion of a number.
      For               Loop.
      Format            Format a Number/Date/Time.
      FreeFile          The next file No. available to open.
      From              Specify the table(s) to be used in an .
      FV                Future Value of an annuity.
    g
      GetAllSettings    List the settings saved in the registry.
      GetAttr           Get file/folder attributes.
      GetObject         Return a reference to an ActiveX object
      GetSetting        Retrieve a value from the users registry.
      form.GoToPage     Move to a page on specific form.
     .GoToRecord (DoCmd)Move to a specific record in a dataset.
    h
      Hex               Convert a number to Hex.
      Hour              Return the hour of the day.
     .Hourglass (DoCmd) Display the hourglass icon.
      HyperlinkPart     Return information about data stored as a hyperlink.
    i
      If Then Else      If-Then-Else
      IIf               If-Then-Else function.
      Input             Return characters from a file.
      InputBox          Prompt for user input.
      Insert (SQL)      Add records to a table (append query).
      InStr             Return the position of one string within another.
      InstrRev          Return the position of one string within another.
      Int               Return the integer portion of a number.
      IPmt              Interest payment for an annuity
      IsArray           Test if an expression is an array
      IsDate            Test if an expression is a date.
      IsEmpty           Test if an expression is Empty (unassned).
      IsError           Test if an expression is returning an error.
      IsMissing         Test if a missing expression.
      IsNull            Test for a NULL expression or Zero Length string.
      IsNumeric         Test for a valid Number.
      IsObject          Test if an expression is an Object.
    L
      Last (SQL)        Return the last value from a query.
      LBound            Return the smallest subscript from an array.
      LCase             Convert a string to lower-case.
      Left              Extract a substring from a string.
      Len               Return the length of a string.
      LoadPicture       Load a picture into an ActiveX control.
      Loc               The current position within an open file.
     .LockNavationPane(DoCmd) Lock the Navation Pane.
      LOF               The length of a file opened with Open()
      Log               Return the natural logarithm of a number.
      LTrim             Remove leading spaces from a string.
    m
      Max (SQL)         Return the maximum value from a query.
     .Maximize (DoCmd)  Enlarge the active window.
      Mid               Extract a substring from a string.
      Min (SQL)         Return the minimum value from a query.
     .Minimize (DoCmd)  Minimise a window.
      Minute            Return the minute of the hour.
      MkDir             Create directory.
      Month             Return the month for a given date.
      MonthName         Return  a string representing the month.
     .Move              Move through a Recordset.
     .MoveFirst/Last/Next/Previous Record
     .MoveSize (DoCmd)  Move or Resize a Window.
      MsgBox            Display a message in a dialogue box.
    n
      Next              Continue a for loop.
      Now               Return the current date and time.
      Nz                Detect a NULL value or a Zero Length string.
    o
      Oct               Convert an integer to Octal.
      OnClick, OnOpen   Events.
     .OpenForm (DoCmd)  Open a form.
     .OpenQuery (DoCmd) Open a .
     .OpenRecordset         Create a new Recordset.
     .OpenReport (DoCmd)    Open a report.
     .OutputTo (DoCmd)      Export to a Text/CSV/Spreadsheet file.
    p
      Partition (SQL)       Locate a number within a range.
     .PrintOut (DoCmd)      Print the active object (form/report etc.)
    q
      Quit                  Quit Microsoft Access
    r
     .RefreshRecord (DoCmd) Refresh the data in a form.
     .Rename (DoCmd)        Rename an object.
     .RepaintObject (DoCmd) Complete any pending screen updates.
      Replace               Replace a sequence of characters in a string.
     .Re               Re the data in a form or a control.
     .Restore (DoCmd)       Restore a maximized or minimized window.
      RGB                   Convert an RGB color to a number.
      Rht                 Extract a substring from a string.
      Rnd                   Generate a random number.
      Round                 Round a number to n decimal places.
      RTrim                 Remove trailing spaces from a string.
     .RunCommand            Run an Access menu or toolbar command.
     .RunDataMacro (DoCmd)  Run a named data macro.
     .RunMacro (DoCmd)      Run a macro.
     .RunSavedImportExport (DoCmd) Run a saved import or export specification.
     .RunSQL (DoCmd)        Run an SQL .
    s
     .Save (DoCmd)          Save a database object.
      SaveSetting           Store a value in the users registry
     .SearchForRecord(DoCmd) Search for a specific record.
      Second                Return the seconds of the minute.
      Seek                  The position within a file opened with Open.
      Select (SQL)          Retrieve data from one or more tables or queries.
      Select Into (SQL)     Make-table .
      Select-Sub (SQL) Sub.
     .SelectObject (DoCmd)  Select a specific database object.
     .SendObject (DoCmd)    Send an email with a database object attached.
      SendKeys              Send keystrokes to the active window.
      SetAttr               Set the attributes of a file.
     .SetDisplayedCategories (DoCmd)  Change Navation Pane display options.
     .SetFilter (DoCmd)     Apply a filter to the records being displayed.
      SetFocus              Move focus to a specified field or control.
     .SetMenuItem (DoCmd)   Set the state of menubar items (enabled /checked)
     .SetOrderBy (DoCmd)    Apply a sort to the active datasheet, form or report.
     .SetParameter (DoCmd)  Set a parameter before opening a Form or Report.
     .SetWarnings (DoCmd)   Turn system messages on or off.
      Sgn                   Return the sn of a number.
     .ShowAllRecords(DoCmd) Remove any applied filter.
     .ShowToolbar (DoCmd)   Display or hide a custom toolbar.
      Shell                 Run an executable program.
      Sin                   Display Sine of an angle.
      SLN                   Straht Line Depreciation.
      Space                 Return a number of spaces.
      Sqr                   Return the square root of a number.
      StDev (SQL)           Estimate the standard deviation for a population.
      Str                   Return a string representation of a number.
      StrComp               Compare two strings.
      StrConv               Convert a string to Upper/lower case or Unicode.
      String                Repeat a character n times.
      Sum (SQL)             Add up the values in a  result set.
      Switch                Return one of several values.
      SysCmd                Display a progress meter.
    t
      Top 1 *               Get first rpw
      Tan                   Display Tangent of an angle.
      Time                  Return the current system time.
      Timer                 Return a number (single) of seconds since midnht.
      TimeSerial            Return a time given an hour, minute, and second.
      TimeValue             Convert a string to a Time.
     .TransferDatabase (DoCmd)      Import or export data to/from another database.
     .TransferSharePointList(DoCmd) Import or link data from a SharePoint Foundation site.
     .TransferSpreadsheet (DoCmd)   Import or export data to/from a spreadsheet file.
     .TransferSQLDatabase (DoCmd)   Copy an entire SQL Server database.
     .TransferText (DoCmd)          Import or export data to/from a text file.
      Transform (SQL)       Create a crosstab .
      Trim                  Remove leading and trailing spaces from a string.
      TypeName              Return the data type of a variable.
    u
      UBound                Return the largest subscript from an array.
      UCase                 Convert a string to upper-case.
      Undo                  Undo the last data edit.
      Union (SQL)           Combine the results of two SQL queries.
      Update (SQL)          Update existing field values in a table.
     .Update                Save a recordset.
    v
      Val                   Extract a numeric value from a string.
      Var (SQL)             Estimate variance for sample (all records)
      VarP (SQL)            Estimate variance for population (all records)
      VarType               Return a number indicating the data type of a variable.
    w
      Weekday               Return the weekday (1-7) from a date.
      WeekdayName           Return the day of the week.
    y
      Year                  Return the year for a given date.



    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
