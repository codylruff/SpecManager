Attribute VB_Name = "Symbols"
Option Explicit
' Test Comment
Public Const PUBLIC_DIR               As String = "S:\Data Manager"
Public Const SM_PATH                  As String = "C:\Users\"
Public Const DATABASE_PATH            As String = "S:\Data Manager\Database\SAATI_Spec_Manager.db3"

' SPEC MANAGER ERROR DESCRIPTIONS:
Public Const SM_SEARCH_SUCCESS     As Long = 0
Public Const SM_SEARCH_FAILURE     As Long = 1
Public Const SM_SEARCH_AGAIN       As Long = 3

' DATABASE ERROR DESCRIPTIONS:
Public Const DB_PUSH_SUCCESS       As Long = 0
Public Const DB_PUSH_FAILURE       As Long = 1
Public Const DB_DELETE_SUCCESS     As Long = 2
Public Const DB_DELETE_FAILURE     As Long = 3
Public Const DB_SELECT_SUCCESS     As Long = 4
Public Const DB_SELECT_FAILURE     As Long = 5
Public Const DB_PUSH_DENIED        As Long = 6
Public Const DB_DELETE_DENIED      As Long = 7

' ACCOUNT PRIVLEDGE LEVELS
Public Const USER_MANAGER          As Long = 21
Public Const USER_READONLY         As Long = 20
Public Const USER_ADMIN            As Long = 25

' LOG VERBOSITY LEVELS
Public Const LOG_CRIT              As Long = 0
Public Const LOG_DEBUG             As Long = 1
Public Const LOG_ALL               As Long = 2

' Returned from SQLite3Initialize
Public Const SQLITE_INIT_OK     As Long = 0
Public Const SQLITE_INIT_ERROR  As Long = 1

' SQLite data types
Public Const SQLITE_INTEGER  As Long = 1
Public Const SQLITE_FLOAT    As Long = 2
Public Const SQLITE_TEXT     As Long = 3
Public Const SQLITE_BLOB     As Long = 4
Public Const SQLITE_NULL     As Long = 5

' SQLite atandard return value
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
