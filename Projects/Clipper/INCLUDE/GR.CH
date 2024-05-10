/*
   GR.CH
   Crucial directives
   Written March 1991 through March 1992
   Greg Lief
*/

//ÄÄÄÄÄ version #
#define GR_REVNO  "1.10"

//ÄÄÄÄÄ default name for data dictionary
#define DD_DEFNAME   "NEW.GRD"

//ÄÄÄÄÄ default name for printer driver file
#define PRINTER_FILE "grprint.dbf"

//ÄÄÄÄÄ command-line switch for developer... enables data dict section
#define DEVELOPER_SWITCH  "/DEV"

//ÄÄÄÄÄ if you don't need spreadsheet output, please comment out the next
//ÄÄÄÄÄ line and remove CLIPWKS5.LIB from the link file
// #define SPREADSHEET_OUTPUT

//ÄÄÄÄÄ if you don't need context-specific help, comment out the next line
#define USING_HELP

//ÄÄÄÄÄ if you don't need the DOS shell, please comment out the next
//ÄÄÄÄÄ line and remove all reference to OVERCL from the link file
// #define DOS_SHELL

//ÄÄÄÄÄ if you don't need Flexfile support, please comment out the next
//ÄÄÄÄÄ line and remove FLEXFILE.LIB from the link file
#define FLEXFILE_SUPPORT

//ÄÄÄÄÄ master key for encrypting data dictionary passwords
#define SKELETON_KEY    "’†—Šr†“‰—ŠŠ"

//ÄÄÄÄÄ next item is used in DD.PRG and VIEWDATA.PRG for pseudo-fields
#define PSEUDO_FIELD    "PSEUDO_FLD"  

#define READONLY   .t.

//ÄÄÄÄÄ directives to return either path or filename
#xtranslate StripPath( <f> ) => ;
           if("\" $ <f>, substr(<f>, rat("\", <f>) + 1), <f>)

#xtranslate KeepPath( <f> ) => ;
           if("\" $ <f>, substr(<f>, 1, rat("\", <f>)), '')

//ÄÄÄÄÄ return an array and a logical value from the same function
#xtranslate load_array(<array>, <file>) => ;
            (len(<array> := gr_ldarray(<file>)) > 0)

//ÄÄÄÄÄ manifest constants for use with data dictionary array
#define DD_DESCRIP  1
#define DD_PASSWORD 2
#define DD_SECURITY 3
#define DD_FILES    4
#define DESCRIP_LENGTH  60


//ÄÄÄÄÄ manifest constants for use within nested file arrays in data dictionary
#define PATH             1
#define FILENAME         2
#define PSEUDONAME       3
#define OPEN_ALIAS       4
#define PARENTFLAG       5
#define FLEXFILEFLAG     6
#define INDEXES          7
#define RELATIONS        8
#define FIELDS           9


//ÄÄÄÄÄ manifest constants for use with fields subarrays in data dict.
#define FIELDNAME     1
#define FIELDTYPE     2
#define WIDTH         4
#define PSEUDOWIDTH   5
#define PICTURE       6
#define FORMULA       7
#define SORT_ORDER    8
#define HELPTEXT      9
#define ALIAS         10
#define PSEUDOALIAS   11
#define DECIMALS      12
#define N_FORMULAS    13         // relevant only for numeric fields
   #define N_TOTAL        1      // ditto
   #define N_AVERAGE      2      // ditto
   #define N_HIGHEST      3      // ditto
   #define N_LOWEST       4      // ditto
   #define N_VARIANCE     5      // ditto
   #define N_DEVIATION    6      // ditto
   #define N_SUBTOTALS    7      // ditto
   #define TOTAL_FORMULAS 7      // sets loop maximum in GR_FIELD.PRG
#define FLEXFILENAME  14         // relevant only for Flexfile pointers
#define HIDDEN        15         // generally only for linking pseudo-fields


//ÄÄÄÄÄ manifest constants for use with relations subarrays in data dict.
#define R_FILENAME    1
#define R_PARENT_KEY  2
#define R_CHILD_KEY   3


//ÄÄÄÄÄ manifest constants for use with GR_SET() function
#define SECURITY             1      // security level (numeric)
#define DATADICT             2      // data dictionary array
#define DATADICT_NAME        3      // name of data dictionary file
#define MASTER_FIELDS        4      // array of master field info
#define ACTIVE_FIELDS        5      // array of fields to be displayed
#define FIELDNAMES           6      // array of fieldnames for selection
#define FILES_ARRAY          7      // array containing DBF file info
#define QUERY_CONDITION      8      // array containing query condition
#define SORT_FIELDS          9      // array containing fields for index
#define CAN_PROCESS_QUERY    10     // (.T. == yes, .F. == no)
#define DATADICT_SAVED       11     // was data dictionary saved (.T. = yes)
#define BROWSE               12     // TBrowse object for browsing data
#define REBUILD_BROWSE       13     // rebuild TBrowse? (.T. == yes, .F. == no)
#define P_LOCAL_OPTIONS      14     // array of print options (on report level)
#define REPORT_NAME          15     // name of active report
#define REPORT_DESCRIP       16     // description of active report
#define QUERY_NAME           17     // name of active query
#define QUERY_DESCRIP        18     // description of active query
#define FORM_OBJECTS         19     // array containing form objects
#define FORM_NAME            20     // name of active form
#define FORM_DESCRIP         21     // description of active form
#define FORM_PRINT_OPTIONS   22     // array containing form print options
   #define FORM_PAGEBREAK     1     // page break between forms?
   #define FORM_PAGELENGTH    2     // page length for forms?
   #define FORM_LINES         3     // lines between forms?
   #define FORM_PAUSE         4     // pause between forms?
   #define FORM_COPIES        5     // # of copies to print
#define REBUILD_FILES_ARRAY  23     // .T. if there was DBF output
#define SUBSET_ACTIVE        24     // .T. if active data subset in GR_VIEW()
#define PROGRAM_NAME         25     // self-explanatory
#define VIEW_NAME            26     // name of view (.GRV) file (if any)


//ÄÄÄÄÄ manifest constants for use with GR_GLOBAL() function
#define CASE_SENSITIVE     1        // case sensitivity (.T. on, .F. off)
        //ÄÄÄÄÄ NOTE: case sensitivity affects both query conditions
        //ÄÄÄÄÄ and data filter conditions in the interactive data browser
#define SHOW_ALIAS            2     // show alias? (.T. == yes, .F. == no)
#define SHOW_DELETED          3     // show deleted recs?  (.T. == yes)
#define SHOW_INDEXBAR         4     // show index status bar?  (.T. == yes)
#define PUGET_SOUND           5     // use sound for wait/error msgs (.T. == on)
#define DATEFORMAT            6     // date format
#define COMPANY_NAME          7     // company name (used in output report)
#define EXPANDED_FORMS        8     // switch to 43/50 line mode for forms?
#define P_NAME                9     // printer name
#define P_NUMBER              10    // record # in GRPRINT.DBF (for speed)
#define P_GLOBAL_OPTIONS      11
  #define P_PAGELENGTH            1  // page length
  #define P_SHOW_TOTAL_RECS       2  // show # total recs at end of report?
  #define P_LEFTMARGIN            3  // obvious
  #define P_TOPMARGIN             4  // obvious
  #define P_COLUMNSPACING         5  // spacing between columns
  #define P_HEADER                6  // report header
  #define P_LINES_AFTER_HEADER    7  // lines between header & column headings
  #define P_FOOTER                8  // report footer
  #define P_LINES_BEFORE_FOOTER   9  // lines between data and report footer
  #define P_ORIENTATION           10 // orientation
  #define P_LPI                   11 // lines per inch
  #define P_NO_DUPE_KEYS          12 // suppress duplicate key values
  #define P_SUBGROUP_COUNT        13 // show record count for subgroups
  #define P_SUBGROUP_PERCENT      14 // show % of total for subgroups
  #define P_DOUBLE_SPACING        15 // double-spaced output?
  #define P_SUMMARY_REPORT        16 // summary report only? .T. == yes
  #define P_PORT                  17 // printer port
  #define P_AUTO_COMPRESS         18 // auto-compress if >80 columns?
  #define P_EJECT_AFTER_SUBGROUP  19 // eject after each subgroup?
  #define P_USE_LEGAL             20 // eject after each subgroup?
// note: item #12 in array is unused (for compatibility w/ versions >1.03)
#define P_COMPRESSED          13
#define P_WIDTH               14    // page width
#define P_WIDE_ON             15
#define P_WIDE_OFF            16
#define P_BOLD_ON             17
#define P_BOLD_OFF            18
#define P_UNDERSCORE_ON       19
#define P_UNDERSCORE_OFF      20
#define P_ITALIC_ON           21
#define P_ITALIC_OFF          22
#define P_COMPRESSED_ON       23
#define P_COMPRESSED_OFF      24
#define P_SIX_LPI             25
#define P_EIGHT_LPI           26
#define P_PORTRAIT            27
#define P_LANDSCAPE           28
#define P_INITIALIZE          29
#define P_RESET               30
#define VIDEO_ROWS            31
#define VIDEO_COLS            32
#define FILESPEC              33     // filespec for searches (e.g., "*.DBF")
#define P_IBMEXT              34
#define INPUT_PATH            35
#define OUTPUT_PATH           36
#define P_LEGAL               37     // escape sequence for legal-size paper
#define P_LETTER              38     // escape sequence for letter-size paper


//ÄÄÄÄÄ manifest constants for security levels throughout
#define S_AVERAGE_USER      1
#define S_POWER_USER        5
#define S_ADMINISTRATOR    10
#define S_DEVELOPER       100


//ÄÄÄÄÄ manifest constants for use with StruMatch() return array
#define PARENT_KEY keyfields_[1]
#define CHILD_KEY  keyfields_[2]


//ÄÄÄÄÄ used by GR_Manual(), DD(), & GR_PrtDD() for nested arrays of indexes
#define INDEXNAME         1
#define INDEX_PSEUDONAME  2
#define INDEXKEY          3


//ÄÄÄÄÄ manifest constant for nested arrays in sortfields array
#define SORT_FIELDNAME   1
#define SORT_ASCENDING   2


//ÄÄÄÄÄ manifest constant for nested arrays in display fields array
#define DISP_FIELDNAME   1
#define DISP_FIELDINFO   2


//ÄÄÄÄÄ since I must use a nasty internal, I might as well protect myself...
#xtranslate StuffKey( <key> ) => __Keyboard( <key> )

#define FORMULA_NAMES { "Total", "Average", "Maximum", "Minimum", ;
                        "Variance", "Deviation", "Subtotals" }

* eof GR.CH
