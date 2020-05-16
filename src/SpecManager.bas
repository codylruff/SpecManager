Attribute VB_Name = "SpecManager"
Option Explicit
' String Constants
Public Const PUBLIC_DIR               As String = "S:\Data Manager"
Public Const SM_PATH                  As String = "C:\Users\"
Public Const DATABASE_PATH            As String = "C:\Users\cruff\Documents\Projects\source\Spec-Manager\Database\SAATI_Spec_Manager.db3"
Public Const TEST_DATABASE_PATH       As String = "S:\Data Manager\Database\SAATI_Spec_Manager.db3"
Public Const nullstr                  As String = vbNullString

' SPEC MANAGER ERROR DESCRIPTIONS:
' Error Id + vbObjectError (ie. -2147221504)
Public Const SM_SEARCH_SUCCESS          As Long = -2147221064
Public Const SEARCH_ERR                 As Long = -2147221063
Public Const REGEX_ERR                  As Long = -2147221062
Public Const MATERIAL_EXISTS_ERR        As Long = -2147221061
Public Const INTERNAL_ERR               As Long = -2147221060
Public Const DB_PUSH_SUCCESS            As Long = -2147220604
Public Const DB_PUSH_ERR                As Long = -2147220603
Public Const DB_DELETE_SUCCESS          As Long = -2147220602
Public Const DB_DELETE_ERR              As Long = -2147220601
Public Const DB_SELECT_SUCCESS          As Long = -2147220600
Public Const DB_SELECT_ERR              As Long = -2147220599
Public Const DB_PUSH_DENIED_ERR_ERR     As Long = -2147220598
Public Const DB_DELETE_DENIED_ERR_ERR   As Long = -2147220597
Public Const DB_TRANSACTION_ERR         As Long = -2147220596
Public Const DB_TRANSACTION_SUCCESS     As Long = -2147220595

' ACCOUNT PRIVLEDGE LEVELS
Public Const USER_MANAGER          As Long = 21
Public Const USER_READONLY         As Long = 20
Public Const USER_ADMIN            As Long = 25

' Returned from SQLite3Initialize
Public Const SQLITE_INIT_OK     As Long = 0
Public Const SQLITE_INIT_ERROR  As Long = 1

' SQLite data types
Public Const SQLITE_INTEGER  As Long = 1
Public Const SQLITE_FLOAT    As Long = 2
Public Const SQLITE_TEXT     As Long = 3
Public Const SQLITE_BLOB     As Long = 4
Public Const SQLITE_NULL     As Long = 5

' SQLite atandard return
Public Const SQLITE_OK         As Long = 0   ' Successful result
Public Const SQLITE_ERROR      As Long = 1   ' SQL error or missing database
Public Const SQLITE_INTERNAL   As Long = 2   ' Internal logic error in SQLite
Public Const SQLITE_PERM       As Long = 3   ' Access permission denied
Public Const SQLITE_ABORT      As Long = 4   ' Callback routine requested an abort
Public Const SQLITE_BUSY       As Long = 5   ' The database file is locked
Public Const SQLITE_LOCKED     As Long = 6   ' A table in the database is locked
Public Const SQLITE_NOMEM      As Long = 7   ' A malloc() failed
Public Const SQLITE_READONLY   As Long = 8   ' Attempt to write a readonly database
Public Const SQLITE_INTERRUPT  As Long = 9   ' Operation terminated by sqlite3_interrupt()
Public Const SQLITE_IOERR      As Long = 10  ' Some kind of disk I/O error occurred
Public Const SQLITE_CORRUPT    As Long = 11  ' The database disk image is malformed
Public Const SQLITE_NOTFOUND   As Long = 12  ' NOT USED. Table or record not found
Public Const SQLITE_FULL       As Long = 13  ' Insertion failed because database is full
Public Const SQLITE_CANTOPEN   As Long = 14  ' Unable to open the database file
Public Const SQLITE_PROTOCOL   As Long = 15  ' NOT USED. Database lock protocol error
Public Const SQLITE_EMPTY      As Long = 16  ' Database is empty
Public Const SQLITE_SCHEMA     As Long = 17  ' The database schema changed
Public Const SQLITE_TOOBIG     As Long = 18  ' String or BLOB exceeds size limit
Public Const SQLITE_CONSTRAINT As Long = 19  ' Abort due to constraint violation
Public Const SQLITE_MISMATCH   As Long = 20  ' Data type mismatch
Public Const SQLITE_MISUSE     As Long = 21  ' Library used incorrectly
Public Const SQLITE_NOLFS      As Long = 22  ' Uses OS features not supported on host
Public Const SQLITE_AUTH       As Long = 23  ' Authorization denied
Public Const SQLITE_FORMAT     As Long = 24  ' Auxiliary database format error
Public Const SQLITE_RANGE      As Long = 25  ' 2nd parameter to sqlite3_bind out of range
Public Const SQLITE_NOTADB     As Long = 26  ' File opened that is not a database file
Public Const SQLITE_ROW        As Long = 100 ' sqlite3_step() has another row ready
Public Const SQLITE_DONE       As Long = 101 ' sqlite3_step() has finished executing

' Extended error codes
Public Const SQLITE_IOERR_READ               As Long = 266  '(SQLITE_IOERR | (1<<8))
Public Const SQLITE_IOERR_SHORT_READ         As Long = 522  '(SQLITE_IOERR | (2<<8))
Public Const SQLITE_IOERR_WRITE              As Long = 778  '(SQLITE_IOERR | (3<<8))
Public Const SQLITE_IOERR_FSYNC              As Long = 1034 '(SQLITE_IOERR | (4<<8))
Public Const SQLITE_IOERR_DIR_FSYNC          As Long = 1290 '(SQLITE_IOERR | (5<<8))
Public Const SQLITE_IOERR_TRUNCATE           As Long = 1546 '(SQLITE_IOERR | (6<<8))
Public Const SQLITE_IOERR_FSTAT              As Long = 1802 '(SQLITE_IOERR | (7<<8))
Public Const SQLITE_IOERR_UNLOCK             As Long = 2058 '(SQLITE_IOERR | (8<<8))
Public Const SQLITE_IOERR_RDLOCK             As Long = 2314 '(SQLITE_IOERR | (9<<8))
Public Const SQLITE_IOERR_DELETE             As Long = 2570 '(SQLITE_IOERR | (10<<8))
Public Const SQLITE_IOERR_BLOCKED            As Long = 2826 '(SQLITE_IOERR | (11<<8))
Public Const SQLITE_IOERR_NOMEM              As Long = 3082 '(SQLITE_IOERR | (12<<8))
Public Const SQLITE_IOERR_ACCESS             As Long = 3338 '(SQLITE_IOERR | (13<<8))
Public Const SQLITE_IOERR_CHECKRESERVEDLOCK  As Long = 3594 '(SQLITE_IOERR | (14<<8))
Public Const SQLITE_IOERR_LOCK               As Long = 3850 '(SQLITE_IOERR | (15<<8))
Public Const SQLITE_IOERR_CLOSE              As Long = 4106 '(SQLITE_IOERR | (16<<8))
Public Const SQLITE_IOERR_DIR_CLOSE          As Long = 4362 '(SQLITE_IOERR | (17<<8))
Public Const SQLITE_LOCKED_SHAREDCACHE       As Long = 265  '(SQLITE_LOCKED | (1<<8) )

' Flags For File Open Operations
Public Const SQLITE_OPEN_READONLY           As Long = 1       ' Ok for sqlite3_open_v2()
Public Const SQLITE_OPEN_READWRITE          As Long = 2       ' Ok for sqlite3_open_v2()
Public Const SQLITE_OPEN_CREATE             As Long = 4       ' Ok for sqlite3_open_v2()
Public Const SQLITE_OPEN_DELETEONCLOSE      As Long = 8       ' VFS only
Public Const SQLITE_OPEN_EXCLUSIVE          As Long = 16      ' VFS only
Public Const SQLITE_OPEN_AUTOPROXY          As Long = 32      ' VFS only
Public Const SQLITE_OPEN_URI                As Long = 64      ' Ok for sqlite3_open_v2()
Public Const SQLITE_OPEN_MEMORY             As Long = 128     ' Ok for sqlite3_open_v2()
Public Const SQLITE_OPEN_MAIN_DB            As Long = 256     ' VFS only
Public Const SQLITE_OPEN_TEMP_DB            As Long = 512     ' VFS only
Public Const SQLITE_OPEN_TRANSIENT_DB       As Long = 1024    ' VFS only
Public Const SQLITE_OPEN_MAIN_JOURNAL       As Long = 2048    ' VFS only
Public Const SQLITE_OPEN_TEMP_JOURNAL       As Long = 4096    ' VFS only
Public Const SQLITE_OPEN_SUBJOURNAL         As Long = 8192    ' VFS only
Public Const SQLITE_OPEN_MASTER_JOURNAL     As Long = 16384   ' VFS only
Public Const SQLITE_OPEN_NOMUTEX            As Long = 32768   ' Ok for sqlite3_open_v2()
Public Const SQLITE_OPEN_FULLMUTEX          As Long = 65536   ' Ok for sqlite3_open_v2()
Public Const SQLITE_OPEN_SHAREDCACHE        As Long = 131072  ' Ok for sqlite3_open_v2()
Public Const SQLITE_OPEN_PRIVATECACHE       As Long = 262144  ' Ok for sqlite3_open_v2()
Public Const SQLITE_OPEN_WAL                As Long = 524288  ' VFS only

' VBA-TOOLS COLOR CODES
Public Const SAATI_BLUE             As String = "#153fd6"
Public Const default_color          As String = "#2BBBAD"
Public Const default_color_dark     As String = "#00695c"
Public Const primary_color          As String = "#4285F4"
Public Const primary_color_dark     As String = "#0d47a1"
Public Const secondary_color        As String = "#aa66cc"
Public Const secondary_color_dark   As String = "#9933CC"
Public Const DANGER_COLOR           As String = "#FF4444"
Public Const DANGER_COLOR_DARK      As String = "#CC0000"
Public Const WARNING_COLOR          As String = "#FFBB33"
Public Const WARNING_COLOR_DARK     As String = "#FF8800"
Public Const SUCCESS_COLOR          As String = "#00C851"
Public Const SUCCESS_COLOR_DARK     As String = "#007E33"
Public Const INFO_COLOR             As String = "#33B5E5"
Public Const INFO_COLOR_DARK        As String = "#0099CC"
Public Const PURPLE_LIGHTEN_1       As String = "#ab47bc"
Public Const AMBER                  As String = "#ffc107"
Public Const AMBER_ACCENT_1         As String = "#ffe57f"
Public Const AMBER_ACCENT_2         As String = "#ffd740"
Public Const AMBER_ACCENT_3         As String = "#ffc400"
Public Const AMBER_ACCENT_4         As String = "#ffab00"
Public Const AMBER_DARKEN_1         As String = "#ffb300"
Public Const AMBER_DARKEN_2         As String = "#ffa000"
Public Const AMBER_DARKEN_3         As String = "#ff8f00"
Public Const AMBER_DARKEN_4         As String = "#ff6f00"
Public Const AMBER_LIGHTEN_1        As String = "#ffca28"
Public Const AMBER_LIGHTEN_2        As String = "#ffd54f"
Public Const AMBER_LIGHTEN_3        As String = "#ffe082"
Public Const AMBER_LIGHTEN_4        As String = "#ffecb3"
Public Const AMBER_LIGHTEN_5        As String = "#fff8e1"
Public Const BLACK                  As String = "#000000"
Public Const BLUE                   As String = "#2196f3"
Public Const BLUE_ACCENT_1          As String = "#82b1ff"
Public Const BLUE_ACCENT_2          As String = "#448aff"
Public Const BLUE_ACCENT_3          As String = "#2979ff"
Public Const BLUE_ACCENT_4          As String = "#2962ff"
Public Const BLUE_DARKEN_1          As String = "#1e88e5"
Public Const BLUE_DARKEN_2          As String = "#1976d2"
Public Const BLUE_DARKEN_3          As String = "#1565c0"
Public Const BLUE_DARKEN_4          As String = "#0d47a1"
Public Const BLUE_GREY              As String = "#607d8b"
Public Const BLUE_GREY_DARKEN_1     As String = "#546e7a"
Public Const BLUE_GREY_DARKEN_2     As String = "#455a64"
Public Const BLUE_GREY_DARKEN_3     As String = "#37474f"
Public Const BLUE_GREY_DARKEN_4     As String = "#263238"
Public Const BLUE_GREY_LIGHTEN_1    As String = "#78909c"
Public Const BLUE_GREY_LIGHTEN_2    As String = "#90a4ae"
Public Const BLUE_GREY_LIGHTEN_3    As String = "#b0bec5"
Public Const BLUE_GREY_LIGHTEN_4    As String = "#cfd8dc"
Public Const BLUE_GREY_LIGHTEN_5    As String = "#eceff1"
Public Const BLUE_LIGHTEN_1         As String = "#42a5f5"
Public Const BLUE_LIGHTEN_2         As String = "#64b5f6"
Public Const BLUE_LIGHTEN_3         As String = "#90caf9"
Public Const BLUE_LIGHTEN_4         As String = "#bbdefb"
Public Const BLUE_LIGHTEN_5         As String = "#e3f2fd"
Public Const BROWN                  As String = "#795548"
Public Const BROWN_DARKEN_1         As String = "#6d4c41"
Public Const BROWN_DARKEN_2         As String = "#5d4037"
Public Const BROWN_DARKEN_3         As String = "#4e342e"
Public Const BROWN_DARKEN_4         As String = "#3e2723"
Public Const BROWN_LIGHTEN_1        As String = "#8d6e63"
Public Const BROWN_LIGHTEN_2        As String = "#a1887f"
Public Const BROWN_LIGHTEN_3        As String = "#bcaaa4"
Public Const BROWN_LIGHTEN_4        As String = "#d7ccc8"
Public Const BROWN_LIGHTEN_5        As String = "#efebe9"
Public Const CYAN                   As String = "#00bcd4"
Public Const CYAN_ACCENT_1          As String = "#84ffff"
Public Const CYAN_ACCENT_2          As String = "#18ffff"
Public Const CYAN_ACCENT_3          As String = "#00e5ff"
Public Const CYAN_ACCENT_4          As String = "#00b8d4"
Public Const CYAN_DARKEN_1          As String = "#00acc1"
Public Const CYAN_DARKEN_2          As String = "#0097a7"
Public Const CYAN_DARKEN_3          As String = "#00838f"
Public Const CYAN_DARKEN_4          As String = "#006064"
Public Const CYAN_LIGHTEN_1         As String = "#26c6da"
Public Const CYAN_LIGHTEN_2         As String = "#4dd0e1"
Public Const CYAN_LIGHTEN_3         As String = "#80deea"
Public Const CYAN_LIGHTEN_4         As String = "#b2ebf2"
Public Const CYAN_LIGHTEN_5         As String = "#e0f7fa"
Public Const DEEP_ORANGE            As String = "#ff5722"
Public Const DEEP_ORANGE_DARKEN_1   As String = "#f4511e"
Public Const DEEP_ORANGE_DARKEN_2   As String = "#e64a19"
Public Const DEEP_ORANGE_DARKEN_3   As String = "#d84315"
Public Const DEEP_ORANGE_DARKEN_4   As String = "#bf360c"
Public Const DEEP_ORANGE_LIGHTEN_1  As String = "#ff7043"
Public Const DEEP_ORANGE_LIGHTEN_2  As String = "#ff8a65"
Public Const DEEP_ORANGE_LIGHTEN_3  As String = "#ffab91"
Public Const DEEP_ORANGE_LIGHTEN_4  As String = "#ffccbc"
Public Const DEEP_ORANGE_LIGHTEN_5  As String = "#fbe9e7"
Public Const DEEP_PURPLE            As String = "#673ab7"
Public Const DEEP_PURPLE_ACCENT_1   As String = "#b388ff"
Public Const DEEP_PURPLE_ACCENT_2   As String = "#7c4dff"
Public Const DEEP_PURPLE_ACCENT_3   As String = "#651fff"
Public Const DEEP_PURPLE_ACCENT_4   As String = "#6200ea"
Public Const DEEP_PURPLE_DARKEN_1   As String = "#5e35b1"
Public Const DEEP_PURPLE_DARKEN_2   As String = "#512da8"
Public Const DEEP_PURPLE_DARKEN_3   As String = "#4527a0"
Public Const DEEP_PURPLE_DARKEN_4   As String = "#311b92"
Public Const DEEP_PURPLE_LIGHTEN_1  As String = "#7e57c2"
Public Const DEEP_PURPLE_LIGHTEN_2  As String = "#9575cd"
Public Const DEEP_PURPLE_LIGHTEN_3  As String = "#b39ddb"
Public Const DEEP_PURPLE_LIGHTEN_4  As String = "#d1c4e9"
Public Const DEEP_PURPLE_LIGHTEN_5  As String = "#ede7f6"
Public Const GREEN                  As String = "#4caf50"
Public Const GREEN_ACCENT_1         As String = "#b9f6ca"
Public Const GREEN_ACCENT_2         As String = "#69f0ae"
Public Const GREEN_ACCENT_3         As String = "#00e676"
Public Const GREEN_ACCENT_4         As String = "#00c853"
Public Const GREEN_DARKEN_1         As String = "#43a047"
Public Const GREEN_DARKEN_2         As String = "#388e3c"
Public Const GREEN_DARKEN_3         As String = "#2e7d32"
Public Const GREEN_DARKEN_4         As String = "#1b5e20"
Public Const GREEN_LIGHTEN_1        As String = "#66bb6a"
Public Const GREEN_LIGHTEN_2        As String = "#81c784"
Public Const GREEN_LIGHTEN_3        As String = "#a5d6a7"
Public Const GREEN_LIGHTEN_4        As String = "#c8e6c9"
Public Const GREEN_LIGHTEN_5        As String = "#e8f5e9"
Public Const GREY                   As String = "#9e9e9e"
Public Const GREY_DARKEN_1          As String = "#757575"
Public Const GREY_DARKEN_2          As String = "#616161"
Public Const GREY_DARKEN_3          As String = "#424242"
Public Const GREY_DARKEN_4          As String = "#212121"
Public Const GREY_LIGHTEN_1         As String = "#bdbdbd"
Public Const GREY_LIGHTEN_2         As String = "#e0e0e0"
Public Const GREY_LIGHTEN_3         As String = "#eeeeee"
Public Const GREY_LIGHTEN_4         As String = "#f5f5f5"
Public Const GREY_LIGHTEN_5         As String = "#fafafa"
Public Const INDIGO                 As String = "#3f51b5"
Public Const INDIGO_ACCENT_1        As String = "#8c9eff"
Public Const INDIGO_ACCENT_2        As String = "#536dfe"
Public Const INDIGO_ACCENT_3        As String = "#3d5afe"
Public Const INDIGO_ACCENT_4        As String = "#304ffe"
Public Const INDIGO_DARKEN_1        As String = "#3949ab"
Public Const INDIGO_DARKEN_2        As String = "#303f9f"
Public Const INDIGO_DARKEN_3        As String = "#283593"
Public Const INDIGO_DARKEN_4        As String = "#1a237e"
Public Const INDIGO_LIGHTEN_1       As String = "#5c6bc0"
Public Const INDIGO_LIGHTEN_2       As String = "#7986cb"
Public Const INDIGO_LIGHTEN_3       As String = "#9fa8da"
Public Const INDIGO_LIGHTEN_4       As String = "#c5cae9"
Public Const INDIGO_LIGHTEN_5       As String = "#e8eaf6"
Public Const LIGHT_BLUE             As String = "#03a9f4"
Public Const LIGHT_BLUE_ACCENT_1    As String = "#l80d8ff"
Public Const LIGHT_BLUE_ACCENT_2    As String = "#40c4ff"
Public Const LIGHT_BLUE_ACCENT_3    As String = "#00b0ff"
Public Const LIGHT_BLUE_ACCENT_4    As String = "#0091ea"
Public Const LIGHT_BLUE_DARKEN_1    As String = "#039be5"
Public Const LIGHT_BLUE_DARKEN_2    As String = "#0288d1"
Public Const LIGHT_BLUE_DARKEN_3    As String = "#0277bd"
Public Const LIGHT_BLUE_DARKEN_4    As String = "#01579b"
Public Const LIGHT_BLUE_LIGHTEN_1   As String = "#29b6f6"
Public Const LIGHT_BLUE_LIGHTEN_2   As String = "#4fc3f7"
Public Const LIGHT_BLUE_LIGHTEN_3   As String = "#81d4fa"
Public Const LIGHT_BLUE_LIGHTEN_4   As String = "#b3e5fc"
Public Const LIGHT_BLUE_LIGHTEN_5   As String = "#e1f5fe"
Public Const LIGHT_GREEN            As String = "#8bc34a"
Public Const LIGHT_GREEN_ACCENT_1   As String = "#ccff90"
Public Const LIGHT_GREEN_ACCENT_2   As String = "#b2ff59"
Public Const LIGHT_GREEN_ACCENT_3   As String = "#76ff03"
Public Const LIGHT_GREEN_ACCENT_4   As String = "#64dd17"
Public Const LIGHT_GREEN_DARKEN_1   As String = "#7cb342"
Public Const LIGHT_GREEN_DARKEN_2   As String = "#689f38"
Public Const LIGHT_GREEN_DARKEN_3   As String = "#558b2f"
Public Const LIGHT_GREEN_DARKEN_4   As String = "#33691e"
Public Const LIGHT_GREEN_LIGHTEN_1  As String = "#9ccc65"
Public Const LIGHT_GREEN_LIGHTEN_2  As String = "#aed581"
Public Const LIGHT_GREEN_LIGHTEN_3  As String = "#c5e1a5"
Public Const LIGHT_GREEN_LIGHTEN_4  As String = "#dcedc8"
Public Const LIGHT_GREEN_LIGHTEN_5  As String = "#f1f8e9"
Public Const LIME                   As String = "#cddc39"
Public Const LIME_ACCENT_1          As String = "#f4ff81"
Public Const LIME_ACCENT_2          As String = "#eeff41"
Public Const LIME_ACCENT_3          As String = "#c6ff00"
Public Const LIME_ACCENT_4          As String = "#aeea00"
Public Const LIME_DARKEN_1          As String = "#c0ca33"
Public Const LIME_DARKEN_2          As String = "#afb42b"
Public Const LIME_DARKEN_3          As String = "#9e9d24"
Public Const LIME_DARKEN_4          As String = "#827717"
Public Const LIME_LIGHTEN_1         As String = "#d4e157"
Public Const LIME_LIGHTEN_2         As String = "#dce775"
Public Const LIME_LIGHTEN_3         As String = "#e6ee9c"
Public Const LIME_LIGHTEN_4         As String = "#f0f4c3"
Public Const LIME_LIGHTEN_5         As String = "#f9fbe7"
Public Const MDB_COLOR              As String = "#45526e"
Public Const MDB_COLOR_DARKEN_1     As String = "#3b465e"
Public Const MDB_COLOR_DARKEN_2     As String = "#2e3951"
Public Const MDB_COLOR_DARKEN_3     As String = "#1c2a48"
Public Const MDB_COLOR_DARKEN_4     As String = "#1c2331"
Public Const MDB_COLOR_LIGHTEN_1    As String = "#59698d"
Public Const MDB_COLOR_LIGHTEN_2    As String = "#7283a7"
Public Const MDB_COLOR_LIGHTEN_3    As String = "#929fba"
Public Const MDB_COLOR_LIGHTEN_4    As String = "#b1bace"
Public Const MDB_COLOR_LIGHTEN_5    As String = "#d0d6e2"
Public Const ORANGE                 As String = "#ff9800"
Public Const ORANGE_ACCENT_1        As String = "#ffd180"
Public Const ORANGE_ACCENT_2        As String = "#ffab40"
Public Const ORANGE_ACCENT_3        As String = "#ff9100"
Public Const ORANGE_ACCENT_4        As String = "#ff6d00"
Public Const ORANGE_DARKEN_1        As String = "#fb8c00"
Public Const ORANGE_DARKEN_2        As String = "#f57c00"
Public Const ORANGE_DARKEN_3        As String = "#ef6c00"
Public Const ORANGE_DARKEN_4        As String = "#e65100"
Public Const ORANGE_LIGHTEN_1       As String = "#ffa726"
Public Const ORANGE_LIGHTEN_2       As String = "#ffb74d"
Public Const ORANGE_LIGHTEN_3       As String = "#ffcc80"
Public Const ORANGE_LIGHTEN_4       As String = "#ffe0b2"
Public Const ORANGE_LIGHTEN_5       As String = "#fff3e0"
Public Const PINK                   As String = "#e91e63"
Public Const PINK_ACCENT_1          As String = "#ff80ab"
Public Const PINK_ACCENT_2          As String = "#ff4081"
Public Const PINK_ACCENT_3          As String = "#f50057"
Public Const PINK_ACCENT_4          As String = "#c51162"
Public Const PINK_DARKEN_1          As String = "#d81b60"
Public Const PINK_DARKEN_2          As String = "#c2185b"
Public Const PINK_DARKEN_3          As String = "#ad1457"
Public Const PINK_DARKEN_4          As String = "#880e4f"
Public Const PINK_LIGHTEN_1         As String = "#ec407a"
Public Const PINK_LIGHTEN_2         As String = "#f06292"
Public Const PINK_LIGHTEN_3         As String = "#f48fb1"
Public Const PINK_LIGHTEN_4         As String = "#f8bbd0"
Public Const PINK_LIGHTEN_5         As String = "#fce4ec"
Public Const PURPLE                 As String = "#9c27b0"
Public Const PURPLE_ACCENT_4        As String = "#aa00ff"
Public Const PURPLE_ACCENT_1        As String = "#d500f9"
Public Const PURPLE_ACCENT_2        As String = "#e040fb"
Public Const PURPLE_ACCENT_3        As String = "#ea80fc"
Public Const PURPLE_DARKEN_1        As String = "#8e24aa"
Public Const PURPLE_DARKEN_2        As String = "#7b1fa2"
Public Const PURPLE_DARKEN_3        As String = "#6a1b9a"
Public Const PURPLE_DARKEN_4        As String = "#4a148c"
Public Const PURPLE_LIGHTEN_2       As String = "#ba68c8"
Public Const PURPLE_LIGHTEN_3       As String = "#ce93d8"
Public Const PURPLE_LIGHTEN_4       As String = "#e1bee7"
Public Const PURPLE_LIGHTEN_5       As String = "#f3e5f5"
Public Const RED                    As String = "#f44336"
Public Const RED_ACCENT_1           As String = "#ff8a80"
Public Const RED_ACCENT_2           As String = "#ff5252"
Public Const RED_ACCENT_3           As String = "#ff1744"
Public Const RED_ACCENT_4           As String = "#d50000"
Public Const RED_DARKEN_1           As String = "#e53935"
Public Const RED_DARKEN_2           As String = "#d32f2f"
Public Const RED_DARKEN_3           As String = "#c62828"
Public Const RED_DARKEN_4           As String = "#b71c1c"
Public Const RED_LIGHTEN_1          As String = "#ef5350"
Public Const RED_LIGHTEN_2          As String = "#e57373"
Public Const RED_LIGHTEN_3          As String = "#ef9a9a"
Public Const RED_LIGHTEN_4          As String = "#ffcdd2"
Public Const RED_LIGHTEN_5          As String = "#ffebee"
Public Const TEAL                   As String = "#009688"
Public Const TEAL_ACCENT_1          As String = "#a7ffeb"
Public Const TEAL_ACCENT_2          As String = "#64ffda"
Public Const TEAL_ACCENT_3          As String = "#1de9b6"
Public Const TEAL_ACCENT_4          As String = "#00bfa5"
Public Const TEAL_DARKEN_1          As String = "#00897b"
Public Const TEAL_DARKEN_2          As String = "#00796b"
Public Const TEAL_DARKEN_3          As String = "#00695c"
Public Const TEAL_DARKEN_4          As String = "#004d40"
Public Const TEAL_LIGHTEN_1         As String = "#26a69a"
Public Const TEAL_LIGHTEN_2         As String = "#4db6ac"
Public Const TEAL_LIGHTEN_3         As String = "#80cbc4"
Public Const TEAL_LIGHTEN_4         As String = "#b2dfdb"
Public Const TEAL_LIGHTEN_5         As String = "#e0f2f1"
Public Const WHITE                  As String = "#ffffff"
Public Const YELLOW                 As String = "#ffeb3b"
Public Const YELLOW_ACCENT_1        As String = "#ffff8d"
Public Const YELLOW_ACCENT_2        As String = "#ffff00"
Public Const YELLOW_ACCENT_3        As String = "#ffea00"
Public Const YELLOW_ACCENT_4        As String = "#ffd600"
Public Const YELLOW_DARKEN_1        As String = "#fdd835"
Public Const YELLOW_DARKEN_2        As String = "#fbc02d"
Public Const YELLOW_DARKEN_3        As String = "#f9a825"
Public Const YELLOW_DARKEN_4        As String = "#f57f17"
Public Const YELLOW_LIGHTEN_1       As String = "#ffee58"
Public Const YELLOW_LIGHTEN_2       As String = "#fff176"
Public Const YELLOW_LIGHTEN_3       As String = "#fff59d"
Public Const YELLOW_LIGHTEN_4       As String = "#fff9c4"
Public Const YELLOW_LIGHTEN_5       As String = "#fffde7"

' LOG TYPES
Public Enum LogType
    RuntimeLog = 0
    UserLog = 1
    ErrorLog = 3
    TestLog = 4
    DebugLog = 5
    SqlLog = 6
    ExportLog = 7
    RevisionLog = 8
End Enum

' LOG VERBOSITY LEVELS
Public Enum LogLevel
    LOG_LOW = 3
    LOG_TEST = 4
    LOG_DEBUG = 6
    LOG_ALL = 8
End Enum

Public Enum InputBoxType
    Password = 1 ' Masked using systempassword mask
    SingleLineText = 2 ' Single line text
    MultiLineText = 32 ' Multi line text
    Number = 4 ' Numbers only
    ShortDate = 4 ' Masked dd/mm/yyyy. Dates are validated upon exit
    LongDate = 16 ' asked using dd/Month/yyyy
    DateTime = 48 ' masked using dd/mm/yyyy hh:mm:ss
End Enum

Public Enum DocumentPackageVariant
    Default = 0
    WeavingStyleChange = 1
    WeavingTieBack = 2
    FinishingWithQC = 3
    FinishingNoQC = 4
    Isotex = 5
End Enum

Public Enum PixelDirection
    Horizontal
    Vertical
End Enum

Public Enum symbology
        '/// <summary>
        '/// Code One 2D symbol.
        '/// </summary>
        CodeOne = 0
        '/// <summary>
        '/// Code 39 (ISO 16388)
        '/// </summary>
        Code39 = 1
        '/// <summary>
        '/// Code 39 extended ASCII.
        '/// </summary>
        Code39Extended = 2
        '/// <summary>
        '/// Logistics Applications of Automated Marking and Reading Symbol.
        '/// </summary>
        LOGMARS = 3
        '/// <summary>
        '/// Code 32 (Italian Pharmacode)
        '/// </summary>
        Code32 = 4
        '/// <summary>
        '/// Pharmazentralnummer (PZN - German Pharmaceutical Code)
        '/// </summary>
        PharmaZentralNummer = 5
        '/// <summary>
        '/// Pharmaceutical Binary Code.
        '/// </summary>
        Pharmacode = 6
        '/// <summary>
        '/// Pharmaceutical Binary Code (2 Track)
        '/// </summary>
        Pharmacode2Track = 7
        '/// <summary>
        '/// Code 93
        '/// </summary>
        Code93 = 8
        '/// <summary>
        '/// Channel Code.
        '/// </summary>
        ChannelCode = 9
        '/// <summary>
        '/// Telepen Code.
        '/// </summary>
        Telepen = 10
        '/// <summary>
        '/// Telepen Numeric Code.
        '/// </summary>
        TelepenNumeric = 11
        '/// <summary>
        '/// Code 128/GS1-128 (ISO 15417)
        '/// </summary>
        Code128 = 12
        '/// <summary>
        '/// European Article Number (14)
        '/// </summary>
        EAN14 = 13
        '/// <summary>
        '/// Serial Shipping Container Code.
        '/// </summary>
        SSCC18 = 14
        '/// <summary>
        '/// Standard 2 of 5 Code.
        '/// </summary>
        Standard2of5 = 15
        '/// <summary>
        '/// Interleaved 2 of 5 Code.
        '/// </summary>
        Interleaved2of5 = 16
        '/// <summary>
        '/// Matrix 2 of 5 Code.
        '/// </summary>
        Matrix2of5 = 17
        '/// <summary>
        '/// IATA 2 of 5 Code.
        '/// </summary>
        IATA2of5 = 18
        '/// <summary>
        '/// Datalogic 2 of 5 Code.
        '/// </summary>
        DataLogic2of5 = 19
        '/// <summary>
        '/// ITF 14 (GS1 2 of 5 Code)
        '/// </summary>
        ITF14 = 20
        '/// <summary>
        '/// Deutsche Post Identcode (DHL)
        '/// </summary>
        DeutschePostIdentCode = 21
        '/// <summary>
        '/// Deutsche Post Leitcode (DHL)
        '/// </summary>
        DeutshePostLeitCode = 22
        '/// <summary>
        '/// Codabar Code.
        '/// </summary>
        Codabar = 23
        '/// <summary>
        '/// MSI Plessey Code.
        '/// </summary>
        MSIPlessey = 24
        '/// <summary>
        '/// UK Plessey Code.
        '/// </summary>
        UKPlessey = 25
        '/// <summary>
        '/// Code 11.
        '/// </summary>
        Code11 = 26
        '/// <summary>
        '/// International Standard Book Number.
        '/// </summary>
        ISBN = 27
        '/// <summary>
        '/// European Article Number (13)
        '/// </summary>
        EAN13 = 28
        '/// <summary>
        '/// European Article Number (8)
        '/// </summary>
        EAN8 = 29
        '/// <summary>
        '/// Universal Product Code (A)
        '/// </summary>
        UPCA = 30
        '/// <summary>
        '/// Universal Product Code (E)
        '/// </summary>
        UPCE = 31
        '/// <summary>
        '/// GS1 Databar Omnidirectional.
        '/// </summary>
        DatabarOmni = 32
        '/// <summary>
        '/// GS1 Databar Omnidirectional Stacked.
        '/// </summary>
        DatabarOmniStacked = 33
        '/// <summary>
        '/// GS1 Databar Stacked.
        '/// </summary>
        DatabarStacked = 34
        '/// <summary>
        '/// GS1 Databar Omnidirectional Truncated.
        '/// </summary>
        DatabarTruncated = 35
        '/// <summary>
        '/// GS1 Databar Limited.
        '/// </summary>
        DatabarLimited = 36
        '/// <summary>
        '/// GS1 Databar Expanded.
        '/// </summary>
        DatabarExpanded = 37
        '/// <summary>
        '/// GS1 Databar Expanded Stacked.
        '/// </summary>
        DatabarExpandedStacked = 38
        '/// <summary>
        '/// Data Matrix (ISO 16022)
        '/// </summary>
        DataMatrix = 39
        '/// <summary>
        '/// QR Code (ISO 18004)
        '/// </summary>
        QRCode = 40
        '/// <summary>
        '/// Micro variation of QR Code.
        '/// </summary>
        MicroQRCode = 41
        '/// <summary>
        '/// UPN variation of QR Code.
        '/// </summary>
        UPNQR = 42
        '/// <summary>
        '/// Aztec (ISO 24778)
        '/// </summary>
        Aztec = 43
        '/// <summary>
        '/// Aztec Runes.
        '/// </summary>
        AztecRunes = 44
        '/// <summary>
        '/// Maxicode (ISO 16023)
        '/// </summary>
        MaxiCode = 45
        '/// <summary>
        '/// PDF417 (ISO 15438)
        '/// </summary>
        PDF417 = 46
        '/// <summary>
        '/// PDF417 Truncated.
        '/// </summary>
        PDF417Truncated = 47
        '/// <summary>
        '/// Micro PDF417 (ISO 24728)
        '/// </summary>
        MicroPDF417 = 48
        '/// <summary>
        '/// Australia Post Standard.
        '/// </summary>
        AusPostStandard = 49
        '/// <summary>
        '/// Australia Post Reply Paid.
        '/// </summary>
        AusPostReplyPaid = 50
        '/// <summary>
        '/// Australia Post Redirect.
        '/// </summary>
        AusPostRedirect = 51
        '/// <summary>
        '/// Australia Post Routing.
        '/// </summary>
        AusPostRouting = 52
        '/// <summary>
        '/// United States Postal Service Intelligent Mail.
        '/// </summary>
        USPS = 53
        '/// <summary>
        '/// PostNET (Postal Numeric Encoding Technique)
        '/// </summary>
        PostNet = 54
        '/// <summary>
        '/// Planet (Postal Alpha Numeric Encoding Technique)
        '/// </summary>
        Planet = 55
        '/// <summary>
        '/// Korean Post.
        '/// </summary>
        KoreaPost = 56
        '/// <summary>
        '/// Facing Identification Mark (FIM)
        '/// </summary>
        FIM = 57
        '/// <summary>
        '/// UK Royal Mail 4 State Code.
        '/// </summary>
        RoyalMail = 58
        '/// <summary>
        '/// KIX Dutch 4 State Code.
        '/// </summary>
        KixCode = 59
        '/// <summary>
        '/// DAFT Code (Generic 4 State Code)
        '/// </summary>
        DaftCode = 60
        '/// <summary>
        '/// Flattermarken (Markup Code)
        '/// </summary>
        Flattermarken = 61
        '/// <summary>
        '/// Japanese Post.
        '/// </summary>
        JapanPost = 62
        '/// <summary>
        '/// Codablock-F 2D symbol.
        '/// </summary>
        CodablockF = 63
        '/// <summary>
        '/// Code 16K 2D symbol.
        '/// </summary>
        Code16K = 64
        '/// <summary>
        '/// Dot Code 2D symbol.
        '/// </summary>
        DotCode = 65
        '/// <summary>
        '/// Grid Matrix 2D symbol.
        '/// </summary>
        GridMatrix = 66
        '/// <summary>
        '/// Code 49 2D symbol.
        '/// </summary>
        Code49 = 67
        '/// <summary>
        '/// Han Xin 2D symbol.
        '/// </summary>
        HanXin = 68

        '/// <summary>
        '/// VIN code symbol.
        '/// </summary>
        VINCode = 69

        '/// <summary>
        '/// Mailmark 4 state postal.
        '/// </summary>
        RoyalMailMailmark = 70

        '/// <summary>
        '/// Not a valid Symbol ID.
        '/// </summary>
        Invalid = -1
End Enum

'FTP enums
Public Enum ProtocolEnum
    Sftp = 0
    Scp = 1
    ftp = 2
    Webdav = 3
    S3 = 4
End Enum

Public Enum FtpSecureEnum
    None = 0
    Implicit = 1
    Explicit = 3
End Enum

'end of ftp enums

Public Enum JsonFormatting
    '
    ' Summary:
    '     Specifies formatting options for the Newtonsoft.Json.JsonTextWriter.
    '
    ' Summary:
    '     No special formatting is applied. This is the default.
    None = 0
    '
    ' Summary:
    '     Causes child objects to be indented according to the Newtonsoft.Json.JsonTextWriter.Indentation
    '     and Newtonsoft.Json.JsonTextWriter.IndentChar settings.
    Indented = 1
End Enum


Public Enum AcImportXMLOption
    'Creates a new table based on the structure of the specified XML file.
    acStructureOnly = 0
    'Imports the data into a new table based on the structure of the specified XML file.
    acStructureAndData = 1
    
    acAppendData = 2      'Imports the data into an existing table.
End Enum

Public Sub StartApp()
    App.Start
    GUI.Start
    GuiCommands.ResetExcelGUI
    Logger.Trace "Starting Application"
    'App.current_user.ListenTo App.printer
End Sub

Public Sub RestartApp()
    Logger.Trace "Restarting Application"
    GuiCommands.ResetExcelGUI
    App.RefreshObjects
End Sub

Public Sub StopApp()
    On Error GoTo ResumeShutdown
    Logger.Trace "Stopping Application"
    Logger.SaveLog
    GUI.Shutdown
    App.Shutdown
ResumeShutdown:
    GuiCommands.ResetExcelGUI
End Sub

Public Sub LoadExistingTemplate(template_type As String)
    With App
        Set .current_template = SpecManager.GetTemplate(template_type)
        .current_template.SpecType = template_type
    End With

End Sub

Function NewDocumentInput(template_type As String, spec_name As String, machine_id As String) As String
    If template_type <> nullstr Then
        LoadExistingTemplate template_type
        With App
        Set .current_doc = New Document
        .current_doc.SpecType = .current_template.SpecType
        .current_doc.Revision = "1"
        .current_doc.MaterialId = spec_name
        .current_doc.MachineId = machine_id
        End With
        NewDocumentInput = spec_name
    Else
        NewDocumentInput = nullstr
    End If
End Function

Function TemplateInput(template_type As String) As String
    Set App.current_template = Factory.CreateNewTemplate(template_type)
    TemplateInput = template_type
End Function

Sub MaterialInput(material_id As String)
' Takes user input for material search
    Dim ret_val As Long
    If material_id = nullstr Then
        ' You must enter a material id before clicking search
        Prompt.Error "Document not found!"
        Exit Sub
    End If
    ret_val = SpecManager.SearchForDocuments(material_id)
    If ret_val = SEARCH_ERR Then
        ' Let the user know that the specifcation could not be found.
        Prompt.Error "Document not found!"
        Exit Sub
    ElseIf ret_val = SM_SEARCH_AGAIN Then
        ret_val = SpecManager.SearchForDocuments(material_id)
        If ret_val = SEARCH_ERR Then
            ' Let the user know that the specifcation could not be found.
            Prompt.Error "Document not found!"
            Exit Sub
        End If
    End If

End Sub

Function SearchForDocuments(material_id As String) As Long
' Manages the search procedure
    Dim specs_dict As Object
    Dim itms
    Set specs_dict = SpecManager.GetDocuments(material_id)
    If specs_dict Is Nothing Then
        Logger.Log "Could not find a specifaction for : " & material_id
        SearchForDocuments = SEARCH_ERR
    Else
        Set App.specs = specs_dict
        itms = App.specs.Items
        Set App.current_doc = itms(0)
        Logger.Log "Succesfully retrieved specifications for : " & material_id
        ' If SpecManager.UpdateTemplateChanges Then
        '     Logger.Log "Specs updated"
        ' End If
        SearchForDocuments = SM_SEARCH_SUCCESS
    End If
End Function

Function GetTemplate(template_type As String) As Template
    Dim df As DataFrame
    Set df = DataAccess.GetTemplateRecord(template_type)
    If Not df Is Nothing Then
        Logger.Log "Succesfully retrieved template for : " & template_type
        Set GetTemplate = Factory.CreateTemplateFromRecord(df)
    Else
        Logger.Log "Could not find a template for : " & template_type
        Set GetTemplate = Nothing
    End If

End Function

Function GetAllTemplates() As VBA.Collection
    Dim df As DataFrame
    Dim dict As Object
    Dim coll As VBA.Collection
    Set coll = New VBA.Collection
    Set df = DataAccess.GetTemplateTypes
    ' obsoleted
    Logger.Log "Listing all template types (spec Types) . . . "
    For Each dict In df.records
        coll.Add item:=Factory.CreateTemplateFromDict(dict), Key:=dict.item("Spec_Type")
    Next dict
    Set GetAllTemplates = coll
End Function

Private Function UpdateTemplateChanges() As Boolean
    ' Apply any changes to material specs that happened since the previous template was revised.
    Dim Key, T As Variant
    Dim ret_val As Long
    Dim Updated As Boolean
    Dim doc As Document
    Dim Template As Template
    Dim old_spec As Document
    Logger.Log "Checking specifications for any template updates . . ."
    For Each T In App.specs
    Updated = False
        Set spec = App.specs.item(T)
        Set old_spec = Factory.CopyDocument(spec)
        Set App.current_template = GetTemplate(doc.SpecType)
        For Each Key In App.current_template.Properties
            ' Checks for existence current template properites in previous spec
            If Not doc.Properties.Exists(Key) Then
                ' Missing properties are added.
                Logger.Log "Adding : " & Key & " to " & doc.MaterialId & " properties list."
                doc.Properties.Add Key:=Key, item:=nullstr
                Updated = True
            End If
        Next Key
        For Each Key In doc.Properties
            ' Checks for existance of current_doc Properties in current_template.
            If Not App.current_template.Properties.Exists(Key) Then
                ' Old properties are removed
                Logger.Log "Removing : " & Key & " from " & doc.MaterialId & " properties list."
                doc.Properties.Remove Key
                Updated = True
            End If
        Next Key
        If Updated = True Then
            doc.Revision = CStr(CDbl(doc.Revision) + 1#)
            ret_val = SpecManager.SaveDocument(spec, old_spec)
            If ret_val <> DB_PUSH_SUCCESS Then
                Logger.Log "Data Access returned: " & ret_val, DebugLog
                Logger.Log "New Document Was Not Saved. Contact Admin."
            Else
                Logger.Log "Data Access returned: " & ret_val, DebugLog
                Logger.Log "New Document Succesfully Saved."
            End If
        End If
    Next T

    UpdateTemplateChanges = Updated
End Function

Function GetDocuments(material_id As String) As Object
    Dim json_dict As Object
    Dim specs_dict As Object
    Dim doc As Document
    Dim rev As String
    Dim Key As Variant
    Dim df As DataFrame

    On Error GoTo NullSpecException

    Set df = DataAccess.GetDocumentRecords(MaterialInputValidation(material_id))
    Set specs_dict = Factory.CreateDictionary
    
    If df.records.Count = 0 Then
        Set GetDocuments = Nothing
        Exit Function
    Else
        For Each json_dict In df.records
            Set spec = Factory.CreateSpecFromDict(json_dict)
            specs_dict.Add doc.UID, spec
        Next json_dict
        Set GetDocuments = specs_dict
    End If
    Exit Function
NullSpecException:
    Logger.Log "SpecManager.GetDocuments()"
    Set GetDocuments = Nothing
End Function

Sub ListDocuments(frm As MSForms.UserForm)
' Lists the specifications currently selected in txtConsole for the given form
    Logger.Log "Listing Documents . . . "
    Set App.printer = Factory.CreateDocumentPrinter(frm)
    If Not App.specs Is Nothing Then
        App.printer.ListObjects App.DocumentsByUID
    Else
        App.printer.WriteLine "No specifications are available for this code."
    End If
End Sub

Sub PrintDocument(frm As MSForms.UserForm)
    Logger.Log "Writing Document to Console. . . "
    Set App.printer.FormId = frm
    If Not App.current_doc Is Nothing Then
        App.printer.PrintObjectToConsole App.current_doc
    End If
End Sub

Sub PrintTemplate(frm As MSForms.UserForm)
    Logger.Log "Writing Template to Console . . . "
    Set App.printer.FormId = frm
    App.printer.PrintObjectToConsole App.current_template
End Sub

Public Sub UpdateSingleProperty(property_name As String, property_value As Variant, material_id As String)
' Updates the value of a single property without the use of the UI. This should make Admin easier.
End Sub

Public Sub ApplyTemplateChangesToDocuments(spec_type As String, changes As Variant)
' Apply template changes to all existing specs of that type
    Dim specifications As VBA.Collection
    Dim doc As Document
    Dim old_spec As Document
    Dim i As Long
    Dim transaction As SqlTransaction
    Set specifications = SelectAllDocumentsByType(spec_type)
    Set transaction = DataAccess.BeginTransaction
    For Each spec In specifications
        Set old_spec = Factory.CopyDocument(spec)
        For i = LBound(changes) To UBound(changes)
            doc.AddProperty CStr(changes(i))
        Next i
        doc.Revision = CStr(CDbl(old_doc.Revision) + 1)
        Logger.Log SpecManager.SaveDocument(spec, old_spec, transaction), DebugLog
    Next spec
    'Logger.Log transaction.Commit, DebugLog
End Sub

Private Function SelectAllDocumentsByType(spec_type As String) As VBA.Collection
    Dim record_coll As VBA.Collection
    Dim record_dict As Object
    Dim specifications As New VBA.Collection
    Set record_dict = Factory.CreateDictionary
    Set record_coll = DataAccess.SelectAllDocuments(spec_type)
    For Each record_dict In record_coll
        specifications.Add Factory.CreateSpecFromDict(record_dict)
    Next record_dict
    Set SelectAllDocumentsByType = specifications
End Function

Function CreateDocumentFromCopy(doc As Document, material_id As String) As Long
' Takes a material and makes a copy of it under a new material id
    Dim spec_copy As Document
    Set spec_copy = Factory.CopyDocument(spec)
    spec_copy.MaterialId = material_id
    spec_copy.Revision = 1
    CreateDocumentFromCopy = SaveNewDocument(spec_copy)
End Function

Function GetMaterialDescription(material_id As String) As Variant
' Retrieve material description from the database.
    GetMaterialDescription = DataAccess.GetColumn("Material_Id", material_id, "Description", "materials").Data
End Function

Function AddNewMaterialDescription(material_id As String, description As String, process_id As String) As Long
' If a material description does not exist create it.
    Dim ret_val As Long
    If GetMaterialDescription(material_id) = Empty Then
        ret_val = DataAccess.PushValue("Material_Id", material_id, "Description", description, "materials")
        AddNewMaterialDescription = DataAccess.UpdateValue("Material_Id", material_id, "Process_Id", process_id, "materials")
    Else
        AddNewMaterialDescription = MATERIAL_EXISTS_ERR
    End If
End Function

Function SaveNewDocument(doc As Document, Optional material_description As String) As Long
    Dim ret_val As Long
    If ManagerOrAdmin Then
        If DataAccess.GetDocument(doc.MaterialId, doc.SpecType, doc.MachineId).records.Count = 0 Then
            ret_val = IIf(DataAccess.PushIQueryable(spec, "standard_specifications") = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_ERR)
            ActionLog.CrudOnDocument spec, "Created New Document"
            If IsEmpty(GetMaterialDescription(doc.MaterialId)) Then
                ' Use material_description param or prompt the user to enter one.
                If material_description = nullstr Then
                    material_description = CStr(Prompt.UserInput(SingleLineText, "Material Description: " & doc.MaterialId, _
                                "Enter Material Description :"))
                End If
                ' Add the new Material Description to the materials table.
                SaveNewDocument = AddNewMaterialDescription(doc.MaterialId, material_description, doc.ProcessId)
                ActionLog.CrudOnDocument spec, "Created New Material"
                Exit Function
            End If
            SaveNewDocument = ret_val
        Else
            SaveNewDocument = MATERIAL_EXISTS_ERR
        End If
    Else
        SaveNewDocument = DB_PUSH_DENIED_ERR
    End If

End Function

Function SaveDocument(doc As Document, old_spec As Document, Optional transaction As SqlTransaction) As Long
    If ManagerOrAdmin Then
        If Utils.IsNothing(transaction) Then
            If ArchiveDocument(old_spec) = DB_DELETE_SUCCESS Then
                SaveDocument = IIf(DataAccess.PushIQueryable(spec, "standard_specifications") = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_ERR)
            Else
                SaveDocument = DB_PUSH_DENIED_ERR
            End If
        Else
            If ArchiveDocument(old_spec, transaction) = DB_DELETE_SUCCESS Then
                SaveDocument = IIf(DataAccess.PushIQueryable(spec, "standard_specifications", transaction) = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_ERR)
            Else
                SaveDocument = DB_PUSH_DENIED_ERR
            End If
        End If
    Else
        SaveDocument = DB_PUSH_DENIED_ERR
    End If
    ActionLog.CrudOnDocument spec, "Revised Document"
End Function

Function ArchiveDocument(old_spec As Document, Optional transaction As SqlTransaction) As Long
' Archives the last spec in order to make room for the new one.
    Dim ret_val As Long
    If Utils.IsNothing(transaction) Then
        ' 1. Insert old version into archived_specifications
        ret_val = IIf(DataAccess.PushIQueryable(old_spec, "archived_specifications") = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_ERR)
        ' 2. Delete old version from standard_specifications
        If ret_val = DB_PUSH_SUCCESS Then
            ArchiveDocument = IIf(DeleteDocument(old_spec) = DB_DELETE_SUCCESS, DB_DELETE_SUCCESS, DB_DELETE_ERR)
        End If
    Else
        ' 1. Insert old version into archived_specifications
        ret_val = IIf(DataAccess.PushIQueryable(old_spec, "archived_specifications", transaction) = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_ERR)
        ' 2. Delete old version from standard_specifications
        If ret_val = DB_PUSH_SUCCESS Then
            ArchiveDocument = IIf(DeleteDocument(old_spec, "standard_specifications", transaction) = DB_DELETE_SUCCESS, DB_DELETE_SUCCESS, DB_DELETE_FAILURE)
        End If
    End If
    'ActionLog.CrudOnDocument old_spec, "Archived Document"
End Function

Function SaveTemplate(Template As Template) As Long
    If ManagerOrAdmin Then
        SaveTemplate = IIf(DataAccess.PushIQueryable(Template, "template_specifications") = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_ERR)
    Else
        SaveTemplate = DB_PUSH_DENIED_ERR
    End If
    ActionLog.CrudOnTemplate Template, "Created New Template"
End Function

Function UpdateTemplate(Template As Template) As Long
    If ManagerOrAdmin Then
        UpdateTemplate = IIf(DataAccess.UpdateTemplate(Template) = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_ERR)
    Else
        UpdateTemplate = DB_PUSH_DENIED_ERR
    End If
    ActionLog.CrudOnTemplate Template, "Revised Template"
End Function

Function DeleteTemplate(Template As Template) As Long
    If App.current_user.PrivledgeLevel = USER_ADMIN Then
        DeleteTemplate = IIf(DataAccess.DeleteTemplate(Template) = DB_DELETE_SUCCESS, DB_DELETE_SUCCESS, DB_DELETE_ERR)
    Else
        DeleteTemplate = DB_DELETE_DENIED_ERR
    End If
    ActionLog.CrudOnTemplate Template, "Deleted Template"
End Function

Function DeleteDocument(doc As Document, Optional tbl As String = "standard_specifications", Optional trans As SqlTransaction) As Long
    If App.current_user.PrivledgeLevel = USER_ADMIN Then
        If IsNothing(trans) Then
            DeleteDocument = IIf(DataAccess.DeleteSpec(spec, doc.MachineId, tbl) = DB_DELETE_SUCCESS, DB_DELETE_SUCCESS, DB_DELETE_ERR)
        Else
            DeleteDocument = IIf(DataAccess.DeleteSpec(spec, doc.MachineId, tbl, trans) = DB_DELETE_SUCCESS, DB_DELETE_SUCCESS, DB_DELETE_ERR)
        End If
    Else
        DeleteDocument = DB_DELETE_DENIED_ERR
    End If
    ActionLog.CrudOnDocument spec, "Deleted Document"
End Function

Private Function ManagerOrAdmin() As Boolean
' Test to see if the current account has the manager privledges.
    On Error GoTo ErrorHandler
    If App.current_user.ProductLine = App.current_template.ProductLine Or App.current_user.ProductLine = "Admin" Then
        ManagerOrAdmin = True
    Else
        ManagerOrAdmin = False
    End If
    ManagerOrAdmin = True
    Exit Function
ErrorHandler:
    Dim Account As Account
    Set Account = AccessControl.Account_Initialize
    ManagerOrAdmin = IIf(Account.ProductLine = "Admin", True, False)
End Function

Private Function MaterialInputValidation(material_id As String) As String
' Ensures that the material id input by the user is parseable.
    ' PASS
    MaterialInputValidation = material_id
    
End Function

Function InitializeNewDocument()
    With App
        Set App.current_doc = New Document
        .current_doc.SpecType = .current_template.SpecType
        .current_doc.Revision = "1"
        Set .current_doc.Properties = .current_template.Properties
    End With
End Function

Public Sub DumpAllSpecsToWorksheet(spec_type As String)
    Dim ws As Worksheet
    Dim dicts As Collection
    Dim dict As Object
    Dim props As Variant
    'RestartApp
    App.Start
    ' Turn on Performance Mode
    If Not GUI.PerformanceModeEnabled Then GUI.PerformanceMode (True)

    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = Utils.CreateNewSheet(spec_type & " Dump " & Format(CStr(Now()), "dd-mm-yy"), True)
    Set dicts = DataAccess.SelectAllDocuments(spec_type)
    i = 2
    For Each dict In dicts
        Set App.current_doc = Factory.CreateSpecFromDict(dict)
        props = App.current_doc.ToArray
        If i = 2 Then ws.Range(Cells(1, 1), Cells(1, ArrayLength(props))).value = App.current_doc.header
        ws.Range(Cells(i, 1), Cells(i, ArrayLength(props))).value = props
        i = i + 1
    Next dict
    ws.Range(Cells(1, 1), Cells(1, ArrayLength(props))).Columns.AutoFit
    
    ' Turn off Performance Mode
    If GUI.PerformanceModeEnabled Then GUI.PerformanceMode (False)
    App.Shutdown
End Sub

Public Sub MassCreateDocuments(num_rows As Integer, num_cols As Integer, ws As Worksheet, Optional start_row As Integer = 2, Optional start_col As Integer = 1, Optional print_json_column As Boolean = True, Optional write_to_live As Boolean = False)
' Create a column at the end of a table and fill it with a json string represent each row.
    Dim dict As Object
    Dim i, k As Integer
    Dim json_string As String
    Dim new_spec As Document
    Dim spec_dict As Object
    App.Start
    With ws
        For i = start_row To num_rows - start_row + 1
            Set dict = Factory.CreateDictionary
            Set spec_dict = Factory.CreateDictionary
            For k = start_col To num_cols
                dict.Add .Cells(1, k), .Cells(i, k)
            Next k
            json_string = JsonVBA.ConvertToJson(dict)
            ' If requested, print the json string in a new column.
            If print_json_column Then
                .Cells(i, num_cols + start_col).value = json_string
            End If
            If write_to_live Then
                spec_dict.Add "Properties_Json", json_string
                spec_dict.Add "Material_Id", .Cells(i, 1).value
                spec_dict.Add "Spec_Type", .Cells(i, 2).value
                spec_dict.Add "Revision", 1
                spec_dict.Add "Machine_Id", "BASE"
                Set new_spec = Factory.CreateSpecFromDict(spec_dict)
                ret_val = SpecManager.SaveNewDocument(new_spec, .Cells(i, 3))
                If ret_val = DB_PUSH_SUCCESS Then
                    Logger.Log new_doc.GetName & " Created."
                    ActionLog.CrudOnDocument new_spec, "Created New Document"
                ElseIf ret_val = MATERIAL_EXISTS_ERR Then
                    Logger.Log new_doc.GetName & " Already Exists."
                Else
                    Logger.Log new_doc.GetName & " Was Not Saved."
                End If
            End If
        Next i
        
    End With
    App.Shutdown
End Sub

Public Sub UpdateRbas()
    ParseSpecsTable "Weaving RBA Dump 19-01-20", "WeavingRbaDump", False, True
End Sub

Public Sub ParseSpecsTable(ws_name As String, table_name As String, Optional print_json_column As Boolean = True, Optional write_to_live As Boolean = False)
' Converts each row in the table to json format, then loads it into the specs db
    Dim tbl As Table

    Set tbl = Factory.CreateTable(ActiveWorkbook.Sheets(ws_name), table_name)
    ' Validate table column headers
    If tbl.HeaderRowRange(1) <> "Material_Id" Then ' The first column must be the material_id
        Logger.Log "The first column must be the 'Material_Id'"
    ElseIf tbl.HeaderRowRange(2) <> "Spec_Type" Then ' The second column must be the spec_type
        Logger.Log "The second column must be the 'Spec_Type'"
    ElseIf tbl.HeaderRowRange(3) <> "Description" Then ' The third column must be the material_description
        Logger.Log "The second column must be the 'Description'"
    Else
        MassCreateDocuments num_rows:=tbl.Rows.Count, _
                    num_cols:=CInt(tbl.Columns.Count), _
                    ws:=tbl.Worksheet, _
                    start_col:=4, _
                    print_json_column:=print_json_column, _
                    write_to_live:=write_to_live
    End If

End Sub

Public Sub CopyPropertiesFromFile()
    ' Get range of material ids
    Dim ws As Worksheet
    Dim style_number As String
    Dim json_string As String
    Dim json_file_path As String
    Dim r As Long
    Set ws = Sheet4
    For r = 2 To 58
        style_number = Mid(ws.Cells(r, 1), 6, 3)
        json_file_path = ThisWorkbook.path & "\RBAs\" & style_number & ".json"
        json_string = Replace(JsonVBA.ReadJsonFileToString(json_file_path), "NaN", nullstr)
        ws.Cells(r, 2).value = json_string
    Next r
End Sub

Public Function BuildBallisticTestSpec(material_id As String, package_length_inches As Double, fabric_width_inches As Double, conditioned_weight_gsm As Double, target_psf As Double, machine_id As String, Optional is_test As Boolean = True) As Long
    Dim doc As Document
    Dim package As BallisticPackage
    Dim ret_val As Long

    Set package = Factory.CreateBallisticPackage(package_length_inches, fabric_width_inches, conditioned_weight_gsm, target_psf)
    Set spec = Factory.CreateDocument
    With spec
        .MaterialId = material_id
        .SpecType = "Ballistic Testing Requirements"
        .Revision = 1
        .CreateFromTestingPlan package
        .MachineId = machine_id
    End With
    If is_test Then
        ret_val = 0
    Else
        ret_val = SaveNewDocument(spec)
    End If

    BuildBallisticTestSpec = ret_val
End Function

Public Function LoadBlankWeavingRba(material_id As String, loom_number As String)
' Selects the BlankRBA spec from the database.
    Dim blank_rba As Document
    ' Retrieve the blank rba spec

    Set blank_rba = Factory.CreateDocumentFromRecord(DataAccess.GetDocument("BLANKRBA", "Weaving RBA", "NONE"))
    ' Add material/loom ids to the blank rba
    blank_rba.MaterialId = material_id
    blank_rba.MachineId = loom_number
    ' Load the blank rba into App.specs
    App.specs.Add "Weaving RBA", blank_rba
End Function

Public Function LoadBaseWeavingRba(material_id As String, loom_number As String)
' Selects the BlankRBA spec from the database.
    Dim base_rba As Document
    ' Retrieve the base rba spec
    Set base_rba = Factory.CreateDocumentFromRecord(DataAccess.GetDocument(material_id, "Weaving RBA", "BASE"))
    ' Add material/loom ids to the blank rba
    base_rba.MaterialId = material_id
    base_rba.MachineId = loom_number
    ' Load the base rba into App.specs
    App.specs.Add "Weaving RBA", base_rba
End Function

Public Sub FilterByMachineId(selected_machine_id As String)
' This routine removes all but the specifications associated with the machine_id selected. (From App.specs)
' TODO: Validate input based on list of loom numbers
    Dim machine_id As String
    Dim spec_id As Variant
    
    machine_id = selected_machine_id
    If Not App.specs.Exists("Weaving RBA(" & machine_id & ")") Then
        If Not App.specs.Exists("Weaving RBA(BASE)") Then
            ' If this loom has no RBA print a blank one.
            LoadBlankWeavingRba Utils.RemoveWhiteSpace(App.current_doc.MaterialId), selected_machine_id
        Else
            With App.specs("Weaving RBA(BASE)")
                .MachineId = machine_id
            End With
        End If
    End If
    ' Remove all but selected loom
    For Each spec_id In App.specs
        With App.specs(spec_id)
            If .SpecType = "Weaving RBA" Then
                If .MachineId <> machine_id Then
                    App.specs.Remove spec_id
                End If
            End If
        End With
    Next spec_id
End Sub
