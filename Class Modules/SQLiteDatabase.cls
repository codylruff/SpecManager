VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("Classes")
Implements IDatabase
'----------------------------------
Private m_header As Variant
Private m_data As Variant
Private m_dbPath As String
Private m_qtyRows As Long
Private m_qtyColumns As Long
'================================================================
'init type when querying
Private Enum initType
  onlyInit = 0
  getRows = 1
  getColumns = 2
  executa = 3
End Enum
'================================================================
' Notes:
' Microsoft uses UTF-16, little endian byte order.
Private Const JULIANDAY_OFFSET As Double = 2415018.5
'----------------------------------
' Returned from SQLite3Initialize
Private Const SQLITE_INIT_OK     As Long = 0
Private Const SQLITE_INIT_ERROR  As Long = 1
'----------------------------------
' SQLite data types
Private Const SQLITE_INTEGER  As Long = 1
Private Const SQLITE_FLOAT    As Long = 2
Private Const SQLITE_TEXT     As Long = 3
Private Const SQLITE_BLOB     As Long = 4
Private Const SQLITE_NULL     As Long = 5
'----------------------------------
' SQLite atandard return value
Private Const SQLITE_OK          As Long = 0   ' Successful result
Private Const SQLITE_ERROR       As Long = 1   ' SQL error or missing database
Private Const SQLITE_INTERNAL    As Long = 2   ' Internal logic error in SQLite
Private Const SQLITE_PERM        As Long = 3   ' Access permission denied
Private Const SQLITE_ABORT       As Long = 4   ' Callback routine requested an abort
Private Const SQLITE_BUSY        As Long = 5   ' The database file is locked
Private Const SQLITE_LOCKED      As Long = 6   ' A table in the database is locked
Private Const SQLITE_NOMEM       As Long = 7   ' A malloc() failed
Private Const SQLITE_READONLY    As Long = 8   ' Attempt to write a readonly database
Private Const SQLITE_INTERRUPT   As Long = 9   ' Operation terminated by sqlite3_interrupt()
Private Const SQLITE_IOERR      As Long = 10   ' Some kind of disk I/O error occurred
Private Const SQLITE_CORRUPT    As Long = 11   ' The database disk image is malformed
Private Const SQLITE_NOTFOUND   As Long = 12   ' NOT USED. Table or record not found
Private Const SQLITE_FULL       As Long = 13   ' Insertion failed because database is full
Private Const SQLITE_CANTOPEN   As Long = 14   ' Unable to open the database file
Private Const SQLITE_PROTOCOL   As Long = 15   ' NOT USED. Database lock protocol error
Private Const SQLITE_EMPTY      As Long = 16   ' Database is empty
Private Const SQLITE_SCHEMA     As Long = 17   ' The database schema changed
Private Const SQLITE_TOOBIG     As Long = 18   ' String or BLOB exceeds size limit
Private Const SQLITE_CONSTRAINT As Long = 19   ' Abort due to constraint violation
Private Const SQLITE_MISMATCH   As Long = 20   ' Data type mismatch
Private Const SQLITE_MISUSE     As Long = 21   ' Library used incorrectly
Private Const SQLITE_NOLFS      As Long = 22   ' Uses OS features not supported on host
Private Const SQLITE_AUTH       As Long = 23   ' Authorization denied
Private Const SQLITE_FORMAT     As Long = 24   ' Auxiliary database format error
Private Const SQLITE_RANGE      As Long = 25   ' 2nd parameter to sqlite3_bind out of range
Private Const SQLITE_NOTADB     As Long = 26   ' File opened that is not a database file
Private Const SQLITE_ROW        As Long = 100  ' sqlite3_step() has another row ready
Private Const SQLITE_DONE       As Long = 101  ' sqlite3_step() has finished executing
'----------------------------------
' Extended error codes
Private Const SQLITE_IOERR_READ               As Long = 266  '(SQLITE_IOERR | (1<<8))
Private Const SQLITE_IOERR_SHORT_READ         As Long = 522  '(SQLITE_IOERR | (2<<8))
Private Const SQLITE_IOERR_WRITE              As Long = 778  '(SQLITE_IOERR | (3<<8))
Private Const SQLITE_IOERR_FSYNC              As Long = 1034 '(SQLITE_IOERR | (4<<8))
Private Const SQLITE_IOERR_DIR_FSYNC          As Long = 1290 '(SQLITE_IOERR | (5<<8))
Private Const SQLITE_IOERR_TRUNCATE           As Long = 1546 '(SQLITE_IOERR | (6<<8))
Private Const SQLITE_IOERR_FSTAT              As Long = 1802 '(SQLITE_IOERR | (7<<8))
Private Const SQLITE_IOERR_UNLOCK             As Long = 2058 '(SQLITE_IOERR | (8<<8))
Private Const SQLITE_IOERR_RDLOCK             As Long = 2314 '(SQLITE_IOERR | (9<<8))
Private Const SQLITE_IOERR_DELETE             As Long = 2570 '(SQLITE_IOERR | (10<<8))
Private Const SQLITE_IOERR_BLOCKED            As Long = 2826 '(SQLITE_IOERR | (11<<8))
Private Const SQLITE_IOERR_NOMEM              As Long = 3082 '(SQLITE_IOERR | (12<<8))
Private Const SQLITE_IOERR_ACCESS             As Long = 3338 '(SQLITE_IOERR | (13<<8))
Private Const SQLITE_IOERR_CHECKRESERVEDLOCK  As Long = 3594 '(SQLITE_IOERR | (14<<8))
Private Const SQLITE_IOERR_LOCK               As Long = 3850 '(SQLITE_IOERR | (15<<8))
Private Const SQLITE_IOERR_CLOSE              As Long = 4106 '(SQLITE_IOERR | (16<<8))
Private Const SQLITE_IOERR_DIR_CLOSE          As Long = 4362 '(SQLITE_IOERR | (17<<8))
Private Const SQLITE_LOCKED_SHAREDCACHE       As Long = 265  '(SQLITE_LOCKED | (1<<8) )
'----------------------------------
' Flags For File Open Operations
Private Const SQLITE_OPEN_READONLY           As Long = 1       ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_READWRITE          As Long = 2       ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_CREATE             As Long = 4       ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_DELETEONCLOSE      As Long = 8       ' VFS only
Private Const SQLITE_OPEN_EXCLUSIVE          As Long = 16      ' VFS only
Private Const SQLITE_OPEN_AUTOPROXY          As Long = 32      ' VFS only
Private Const SQLITE_OPEN_URI                As Long = 64      ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_MEMORY             As Long = 128     ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_MAIN_DB            As Long = 256     ' VFS only
Private Const SQLITE_OPEN_TEMP_DB            As Long = 512     ' VFS only
Private Const SQLITE_OPEN_TRANSIENT_DB       As Long = 1024    ' VFS only
Private Const SQLITE_OPEN_MAIN_JOURNAL       As Long = 2048    ' VFS only
Private Const SQLITE_OPEN_TEMP_JOURNAL       As Long = 4096    ' VFS only
Private Const SQLITE_OPEN_SUBJOURNAL         As Long = 8192    ' VFS only
Private Const SQLITE_OPEN_MASTER_JOURNAL     As Long = 16384   ' VFS only
Private Const SQLITE_OPEN_NOMUTEX            As Long = 32768   ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_FULLMUTEX          As Long = 65536   ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_SHAREDCACHE        As Long = 131072  ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_PRIVATECACHE       As Long = 262144  ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_WAL                As Long = 524288  ' VFS only
'----------------------------------
' Options for Text and Blob binding
Private Const SQLITE_STATIC      As Long = 0
Private Const SQLITE_TRANSIENT   As Long = -1
'----------------------------------
' System calls
Private Const CP_UTF8 As Long = 65001
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As Long, ByVal pSource As Long, ByVal length As Long)
Private Declare Function lstrcpynW Lib "kernel32" (ByVal pwsDest As Long, ByVal pwsSource As Long, ByVal cchCount As Long) As Long
Private Declare Function lstrcpyW Lib "kernel32" (ByVal pwsDest As Long, ByVal pwsSource As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal pwsString As Long) As Long
Private Declare Function SysAllocString Lib "OleAut32" (ByRef pwsString As Long) As Long
Private Declare Function SysStringLen Lib "OleAut32" (ByVal bstrString As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
'=====================================================================================
' SQLite StdCall Imports
'-----------------------
' SQLite library version
Private Declare Function sqlite3_stdcall_libversion Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_libversion@0" () As Long ' PtrUtf8String
' Database connections
Private Declare Function sqlite3_stdcall_open16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_open16@8" (ByVal pwsFileName As Long, ByRef hDb As Long) As Long ' PtrDb
Private Declare Function sqlite3_stdcall_open_v2 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_open_v2@16" (ByVal pwsFileName As Long, ByRef hDb As Long, ByVal iFlags As Long, ByVal zVfs As Long) As Long ' PtrDb
Private Declare Function sqlite3_stdcall_close Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_close@4" (ByVal hDb As Long) As Long
' Database connection error info
Private Declare Function sqlite3_stdcall_errmsg Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_errmsg@4" (ByVal hDb As Long) As Long ' PtrUtf8String
Private Declare Function sqlite3_stdcall_errmsg16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_errmsg16@4" (ByVal hDb As Long) As Long ' PtrUtf16String
Private Declare Function sqlite3_stdcall_errcode Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_errcode@4" (ByVal hDb As Long) As Long
Private Declare Function sqlite3_stdcall_extended_errcode Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_extended_errcode@4" (ByVal hDb As Long) As Long
' Database connection change counts
Private Declare Function sqlite3_stdcall_changes Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_changes@4" (ByVal hDb As Long) As Long
Private Declare Function sqlite3_stdcall_total_changes Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_total_changes@4" (ByVal hDb As Long) As Long
'----------------------------------
' Statements
Private Declare Function sqlite3_stdcall_prepare16_v2 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_prepare16_v2@20" (ByVal hDb As Long, ByVal pwsSql As Long, ByVal nSqlLength As Long, ByRef hStmt As Long, ByVal ppwsTailOut As Long) As Long
Private Declare Function sqlite3_stdcall_step Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_step@4" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_stdcall_reset Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_reset@4" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_stdcall_finalize Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_finalize@4" (ByVal hStmt As Long) As Long
'----------------------------------
' Statement column access (0-based indices)
Private Declare Function sqlite3_stdcall_column_count Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_count@4" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_stdcall_column_type Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_type@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_stdcall_column_name Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_name@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_stdcall_column_name16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_name16@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrWString
'----------------------------------
Private Declare Function sqlite3_stdcall_column_blob Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_blob@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrData
Private Declare Function sqlite3_stdcall_column_bytes Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_bytes@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_stdcall_column_bytes16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_bytes16@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_stdcall_column_double Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_double@8" (ByVal hStmt As Long, ByVal iCol As Long) As Double
Private Declare Function sqlite3_stdcall_column_int Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_int@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_stdcall_column_int64 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_int64@8" (ByVal hStmt As Long, ByVal iCol As Long) As Currency ' UNTESTED ....?
Private Declare Function sqlite3_stdcall_column_text Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_text@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_stdcall_column_text16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_text16@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrWString
Private Declare Function sqlite3_stdcall_column_value Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_value@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrSqlite3Value
'----------------------------------
' Statement parameter binding (1-based indices!)
Private Declare Function sqlite3_stdcall_bind_parameter_count Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_parameter_count@4" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_stdcall_bind_parameter_name Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_parameter_name@8" (ByVal hStmt As Long, ByVal paramIndex As Long) As Long
Private Declare Function sqlite3_stdcall_bind_parameter_index Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_parameter_index@8" (ByVal hStmt As Long, ByVal paramName As Long) As Long
Private Declare Function sqlite3_stdcall_bind_null Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_null@8" (ByVal hStmt As Long, ByVal paramIndex As Long) As Long
Private Declare Function sqlite3_stdcall_bind_blob Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_blob@20" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_stdcall_bind_zeroblob Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_zeroblob@12" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal nBytes As Long) As Long
Private Declare Function sqlite3_stdcall_bind_double Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_double@16" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal value As Double) As Long
Private Declare Function sqlite3_stdcall_bind_int Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_int@12" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal value As Long) As Long
Private Declare Function sqlite3_stdcall_bind_int64 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_int64@16" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal value As Currency) As Long ' UNTESTED ....?
Private Declare Function sqlite3_stdcall_bind_text Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_text@20" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal psValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_stdcall_bind_text16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_text16@20" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pswValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_stdcall_bind_value Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_value@12" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pSqlite3Value As Long) As Long
Private Declare Function sqlite3_stdcall_clear_bindings Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_clear_bindings@4" (ByVal hStmt As Long) As Long
'----------------------------------
'Backup
Private Declare Function sqlite3_stdcall_sleep Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_sleep@4" (ByVal msToSleep As Long) As Long
Private Declare Function sqlite3_stdcall_backup_init Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_init@16" (ByVal hDbDest As Long, ByVal zDestName As Long, ByVal hDbSource As Long, ByVal zSourceName As Long) As Long
Private Declare Function sqlite3_stdcall_backup_step Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_step@8" (ByVal hBackup As Long, ByVal nPage As Long) As Long
Private Declare Function sqlite3_stdcall_backup_finish Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_finish@4" (ByVal hBackup As Long) As Long
Private Declare Function sqlite3_stdcall_backup_remaining Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_remaining@4" (ByVal hBackup As Long) As Long
Private Declare Function sqlite3_stdcall_backup_pagecount Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_pagecount@4" (ByVal hBackup As Long) As Long
'=====================================================================================
' Initialize - load libraries explicitly
Private hSQLiteLibrary As Long
Private hSQLiteStdCallLibrary As Long
'================================================================
'  database properties
Public Property Get dbPath() As Variant:  dbPath = m_dbPath: End Property
Private Property Let dbPath(ByVal vNewValue As Variant): m_dbPath = vNewValue: End Property
'================================================================
' public properties
Public Property Get IDatabase_header() As Variant: IDatabase_header = m_header: End Property
Public Property Get IDatabase_data() As Variant: IDatabase_data = m_data: End Property
Public Property Get qtyRows() As Long: qtyRows = m_qtyRows: End Property
Public Property Get qtyColumns() As Long: qtyColumns = m_qtyColumns: End Property
Public Property Get IDatabase_NumRows() As Long: IDatabase_NumRows = m_qtyRows: End Property
Public Property Get IDatabase_NumColumns() As Long: IDatabase_NumColumns = m_qtyColumns: End Property
'================================================================
' opens a database
Public Sub IDatabase_openDb(sqliPath As String)
  '---------------------------------
  If sqliPath = vbNullString Or Len(Dir(sqliPath)) = 0 Then
    Err.Raise 1, "Database not found", "Please select a valid file address"
    Exit Sub
  End If
  '---------------------------------
  dbPath = sqliPath 'seta o endereco do bd na base
End Sub

' Execute sql statement
Public Sub IDatabase_execute(ssql As String)
  Dim myStmtHandle As Long
  '----------------------------------------------'
  Call initQuery(executa, myStmtHandle, Me.dbPath, ssql)     'manda o comando execute
End Sub

' Execute select statement
Public Sub IDatabase_selectQry(ssql As String)
  Dim myStmtHandle As Long
  '----------------------------------------------'
  Call initQuery(getRows, myStmtHandle, Me.dbPath, "select count(*) from (" & ssql & ")") 'get rows
  Call initQuery(getColumns, myStmtHandle, Me.dbPath, ssql) 'get columns
  '----------------------------------------------'
  Call getQryValue(myStmtHandle)
End Sub

'=====================================================================================
'initiate a query
Private Function initQuery(q_comand As initType, _
                          ByRef myStmtHandle As Long, _
                          dbPath As String, _
                          ssql As String) As Double
  '----------------------------------------------'
  Dim RetVal As Long
  Dim stepMsg As String, InitReturn As Long, msgErro As Long
  Dim myDbHandle As Long
  InitReturn = SQLite3Initialize ' Default path is ThisWorkbook.Path but can specify other path where the .dlls reside.
  '----------------------------------------------'
  RetVal = SQLite3Finalize(myStmtHandle)
  RetVal = SQLite3Open(dbPath, myDbHandle) 'abre o db
  RetVal = SQLite3PrepareV2(myDbHandle, ssql, myStmtHandle)
  RetVal = SQLite3Step(myStmtHandle) 'vai pra primeiro registro
  msgErro = RetVal '<--- pega o returns da msg da execucao
  '----------------------------------------------'
  If q_comand = executa Then 'quando excuta o comando, valida se foi executado
    RetVal = SQLite3PrepareV2(myDbHandle, ssql, myStmtHandle) 'executa a query
    If msgErro <> SQLITE_DONE Then Err.Raise 1, "Database not found", "Database command error " & msgErro
  '--------------------------------
  ElseIf q_comand = getRows Then
    m_qtyRows = getRowsCount(myStmtHandle) ' pega o numero de linhas
  '--------------------------------
  ElseIf q_comand = getColumns Then
    m_qtyColumns = SQLite3ColumnCount(myStmtHandle)
  End If
  '----------------------------------------------'
End Function

'returns array of the table
Private Sub getQryValue(ByVal stmtHandle As Long)
    Dim colName As String
    Dim colType As Long
    Dim colTypeName As String
    Dim colValue As Variant
    Dim arrdata() As Variant, arrheader As Variant
    Dim rowsCount As Long
    '--------------------------
    ReDim arrdata(1 To qtyRows, 1 To Me.qtyColumns) 'redimensao dos data
    ReDim arrheader(1 To 1, 1 To Me.qtyColumns) 'redimensao do header
    '--------------------------
    Dim i As Long, j As Long, k As Long, m As Long
    '--------------------------------------
    colType = SQLite3ColumnType(stmtHandle, 0)
    '--------------------------------------
    For m = 1 To Me.qtyRows 'itera pelas linhas
      k = k + 1 'adiciona uma nova linha
      '///////////////
      For i = 0 To Me.qtyColumns - 1 'itera pelas colunas
        '=============================
        colName = SQLite3ColumnName(stmtHandle, i)
        If arrheader(1, i + 1) = vbNullString Then arrheader(1, i + 1) = colName
        '=============================
        colType = SQLite3ColumnType(stmtHandle, i)
        colTypeName = TypeName(colType)
        colValue = ColumnValue(stmtHandle, i, colType)
        '------------------
        arrdata(k, i + 1) = colValue 'adiciona os registros no array
        '------------------
      Next i
      '///////////////
      Call SQLite3Step(stmtHandle) 'vai para o proximo registro
      '--------------------------------------
    Next m
    m_data = arrdata 'contem os data
    m_header = arrheader 'contem o header
End Sub

'returns the row count
Private Function getRowsCount(ByVal stmtHandle As Long) As Long
  Dim ystmtHandle As Long, colType As Long
  Dim k As Long
  ystmtHandle = stmtHandle
  '------------------------------
  colType = SQLite3ColumnType(ystmtHandle, 0)
  getRowsCount = ColumnValue(ystmtHandle, 0, colType) 'return
End Function

Function TypeName(ByVal SQLiteType As Long) As String
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            TypeName = "INTEGER"
        Case SQLITE_FLOAT:
            TypeName = "FLOAT"
        Case SQLITE_TEXT:
            TypeName = "TEXT"
        Case SQLITE_BLOB:
            TypeName = "BLOB"
        Case SQLITE_NULL:
            TypeName = "NULL"
    End Select
End Function

Private Function ColumnValue(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            ColumnValue = SQLite3ColumnInt32(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_FLOAT:
            ColumnValue = SQLite3ColumnDouble(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_TEXT:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_BLOB:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_NULL:
            ColumnValue = Null
    End Select
End Function

Private Function SQLite3Initialize(Optional ByVal libDir As String) As Long
    ' A nice option here is to call SetDllDirectory, but that API is only available since Windows XP SP1.
    If libDir = vbNullString Then libDir = ThisWorkbook.path
    If Right(libDir, 1) <> "\" Then libDir = libDir & "\"
    
    If hSQLiteLibrary = 0 Then
        hSQLiteLibrary = LoadLibrary(libDir + "SQLite3.dll")
        If hSQLiteLibrary = 0 Then
            Logger.Log "SQLite3Initialize Error Loading " + libDir + "SQLite3.dll:" & Err.LastDllError
            SQLite3Initialize = SQLITE_INIT_ERROR
            Exit Function
        End If
    End If
        
    If hSQLiteStdCallLibrary = 0 Then
        hSQLiteStdCallLibrary = LoadLibrary(libDir + "SQLite3_StdCall.dll")
        If hSQLiteStdCallLibrary = 0 Then
            Logger.Log "SQLite3Initialize Error Loading " + libDir + "SQLite3_StdCall.dll:" & Err.LastDllError
            SQLite3Initialize = SQLITE_INIT_ERROR
            Exit Function
        End If
    End If
    SQLite3Initialize = SQLITE_INIT_OK
End Function

Private Sub SQLite3Free()
   Dim refCount As Long
   If hSQLiteStdCallLibrary <> 0 Then
        refCount = FreeLibrary(hSQLiteStdCallLibrary)
        hSQLiteStdCallLibrary = 0
        If refCount = 0 Then
            Logger.Log "SQLite3Free Error Freeing SQLite3_StdCall.dll:", refCount, Err.LastDllError
        End If
    End If
    If hSQLiteLibrary <> 0 Then
        refCount = FreeLibrary(hSQLiteLibrary)
        hSQLiteLibrary = 0
        If refCount = 0 Then
            Logger.Log "SQLite3Free Error Freeing SQLite3.dll:", refCount, Err.LastDllError
        End If
    End If
End Sub

'=====================================================================================
' SQLite library version
Private Function SQLite3LibVersion() As String
    SQLite3LibVersion = Utf8PtrToString(sqlite3_stdcall_libversion())
End Function

'=====================================================================================
' Database connections
Private Function SQLite3Open(ByVal fileName As String, ByRef dbHandle As Long) As Long
    SQLite3Open = sqlite3_stdcall_open16(StrPtr(fileName), dbHandle)
End Function

Private Function SQLite3OpenV2(ByVal fileName As String, ByRef dbHandle As Long, ByVal flags As Long, ByVal vfsName As String) As Long
    Dim bufFileName() As Byte
    Dim bufVfsName() As Byte
    bufFileName = StringToUtf8Bytes(fileName)
    If vfsName = Empty Then
        SQLite3OpenV2 = sqlite3_stdcall_open_v2(VarPtr(bufFileName(0)), dbHandle, flags, 0)
    Else
        bufVfsName = StringToUtf8Bytes(vfsName)
        SQLite3OpenV2 = sqlite3_stdcall_open_v2(VarPtr(bufFileName(0)), dbHandle, flags, VarPtr(bufVfsName(0)))
    End If
End Function

Private Function SQLite3Close(ByVal dbHandle As Long) As Long
    SQLite3Close = sqlite3_stdcall_close(dbHandle)
End Function

'=====================================================================================
' Error information
Private Function SQLite3ErrMsg(ByVal dbHandle As Long) As String
    SQLite3ErrMsg = Utf8PtrToString(sqlite3_stdcall_errmsg(dbHandle))
End Function

Private Function SQLite3ErrCode(ByVal dbHandle As Long) As Long
    SQLite3ErrCode = sqlite3_stdcall_errcode(dbHandle)
End Function

Private Function SQLite3ExtendedErrCode(ByVal dbHandle As Long) As Long
    SQLite3ExtendedErrCode = sqlite3_stdcall_extended_errcode(dbHandle)
End Function

'=====================================================================================
' Change Counts
Private Function SQLite3Changes(ByVal dbHandle As Long) As Long
    SQLite3Changes = sqlite3_stdcall_changes(dbHandle)
End Function

Private Function SQLite3TotalChanges(ByVal dbHandle As Long) As Long
    SQLite3TotalChanges = sqlite3_stdcall_total_changes(dbHandle)
End Function

'=====================================================================================
' Statements
Private Function SQLite3PrepareV2(ByVal dbHandle As Long, ByVal sql As String, ByRef stmtHandle As Long) As Long
    ' Only the first statement (up to ';') is prepared. Currently we don't retrieve the 'tail' pointer.
    SQLite3PrepareV2 = sqlite3_stdcall_prepare16_v2(dbHandle, StrPtr(sql), Len(sql) * 2, stmtHandle, 0)
End Function

Private Function SQLite3Step(ByVal stmtHandle As Long) As Long
    SQLite3Step = sqlite3_stdcall_step(stmtHandle)
End Function

Private Function SQLite3Reset(ByVal stmtHandle As Long) As Long
    SQLite3Reset = sqlite3_stdcall_reset(stmtHandle)
End Function

Private Function SQLite3Finalize(ByVal stmtHandle As Long) As Long
    SQLite3Finalize = sqlite3_stdcall_finalize(stmtHandle)
End Function

'=====================================================================================
' Statement column access (0-based indices)
Private Function SQLite3ColumnCount(ByVal stmtHandle As Long) As Long
    SQLite3ColumnCount = sqlite3_stdcall_column_count(stmtHandle)
End Function

Private Function SQLite3ColumnType(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Long
    SQLite3ColumnType = sqlite3_stdcall_column_type(stmtHandle, ZeroBasedColIndex)
End Function

Private Function SQLite3ColumnName(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As String
    SQLite3ColumnName = Utf8PtrToString(sqlite3_stdcall_column_name(stmtHandle, ZeroBasedColIndex))
End Function

Private Function SQLite3ColumnDouble(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Double
    SQLite3ColumnDouble = sqlite3_stdcall_column_double(stmtHandle, ZeroBasedColIndex)
End Function

Private Function SQLite3ColumnInt32(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Long
    SQLite3ColumnInt32 = sqlite3_stdcall_column_int(stmtHandle, ZeroBasedColIndex)
End Function

Private Function SQLite3ColumnText(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As String
    SQLite3ColumnText = Utf8PtrToString(sqlite3_stdcall_column_text(stmtHandle, ZeroBasedColIndex))
End Function

Private Function SQLite3ColumnDate(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Date
    SQLite3ColumnDate = FromJulianDay(sqlite3_stdcall_column_double(stmtHandle, ZeroBasedColIndex))
End Function

Private Function SQLite3ColumnBlob(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Byte()
    Dim ptr As Long
    Dim length As Long
    Dim buf() As Byte
    '------------------
    ptr = sqlite3_stdcall_column_blob(stmtHandle, ZeroBasedColIndex)
    length = sqlite3_stdcall_column_bytes(stmtHandle, ZeroBasedColIndex)
    ReDim buf(length - 1)
    RtlMoveMemory VarPtr(buf(0)), ptr, length
    SQLite3ColumnBlob = buf
End Function

'=====================================================================================
' Statement bindings
Private Function SQLite3BindText(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal value As String) As Long
    SQLite3BindText = sqlite3_stdcall_bind_text16(stmtHandle, OneBasedParamIndex, StrPtr(value), -1, SQLITE_TRANSIENT)
End Function

Private Function SQLite3BindDouble(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal value As Double) As Long
    SQLite3BindDouble = sqlite3_stdcall_bind_double(stmtHandle, OneBasedParamIndex, value)
End Function

Private Function SQLite3BindInt32(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal value As Long) As Long
    SQLite3BindInt32 = sqlite3_stdcall_bind_int(stmtHandle, OneBasedParamIndex, value)
End Function

Private Function SQLite3BindDate(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal value As Date) As Long
    SQLite3BindDate = sqlite3_stdcall_bind_double(stmtHandle, OneBasedParamIndex, ToJulianDay(value))
End Function

Private Function SQLite3BindBlob(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByRef value() As Byte) As Long
    Dim length As Long
    length = UBound(value) - LBound(value) + 1
    SQLite3BindBlob = sqlite3_stdcall_bind_blob(stmtHandle, OneBasedParamIndex, VarPtr(value(0)), length, SQLITE_TRANSIENT)
End Function

Private Function SQLite3BindNull(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long) As Long
    SQLite3BindNull = sqlite3_stdcall_bind_null(stmtHandle, OneBasedParamIndex)
End Function

Private Function SQLite3BindParameterCount(ByVal stmtHandle As Long) As Long
    SQLite3BindParameterCount = sqlite3_stdcall_bind_parameter_count(stmtHandle)
End Function

Private Function SQLite3BindParameterName(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long) As String
    SQLite3BindParameterName = Utf8PtrToString(sqlite3_stdcall_bind_parameter_name(stmtHandle, OneBasedParamIndex))
End Function

Private Function SQLite3BindParameterIndex(ByVal stmtHandle As Long, ByVal paramName As String) As Long
    Dim buf() As Byte
    buf = StringToUtf8Bytes(paramName)
    SQLite3BindParameterIndex = sqlite3_stdcall_bind_parameter_index(stmtHandle, VarPtr(buf(0)))
End Function

Private Function SQLite3ClearBindings(ByVal stmtHandle As Long) As Long
    SQLite3ClearBindings = sqlite3_stdcall_clear_bindings(stmtHandle)
End Function

'=====================================================================================
' Backup
Private Function SQLite3Sleep(ByVal timeToSleepInMs As Long) As Long
    SQLite3Sleep = sqlite3_stdcall_sleep(timeToSleepInMs)
End Function

Private Function SQLite3BackupInit(ByVal dbHandleDestination As Long, ByVal destinationName As String, ByVal dbHandleSource As Long, ByVal sourceName As String) As Long
    Dim bufDestinationName() As Byte
    Dim bufSourceName() As Byte
    bufDestinationName = StringToUtf8Bytes(destinationName)
    bufSourceName = StringToUtf8Bytes(sourceName)
    SQLite3BackupInit = sqlite3_stdcall_backup_init(dbHandleDestination, VarPtr(bufDestinationName(0)), dbHandleSource, VarPtr(bufSourceName(0)))
End Function

Private Function SQLite3BackupFinish(ByVal backupHandle As Long) As Long
    SQLite3BackupFinish = sqlite3_stdcall_backup_finish(backupHandle)
End Function

Private Function SQLite3BackupStep(ByVal backupHandle As Long, ByVal numberOfPages) As Long
    SQLite3BackupStep = sqlite3_stdcall_backup_step(backupHandle, numberOfPages)
End Function

Private Function SQLite3BackupPageCount(ByVal backupHandle As Long) As Long
    SQLite3BackupPageCount = sqlite3_stdcall_backup_pagecount(backupHandle)
End Function

Private Function SQLite3BackupRemaining(ByVal backupHandle As Long) As Long
    SQLite3BackupRemaining = sqlite3_stdcall_backup_remaining(backupHandle)
End Function

' =========================================='
' String Helpers
Private Function Utf8PtrToString(ByVal pUtf8String As Long) As String
    Dim buf As String
    Dim cSize As Long
    Dim RetVal As Long
    
    cSize = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, 0, 0)
    ' cSize includes the terminating null character
    If cSize <= 1 Then
        Utf8PtrToString = vbNullString
        Exit Function
    End If
    
    Utf8PtrToString = String(cSize - 1, "*") ' and a termintating null char.
    RetVal = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, StrPtr(Utf8PtrToString), cSize)
    If RetVal = 0 Then
        Logger.Log "Utf8PtrToString Error:" & Err.LastDllError
        Exit Function
    End If
End Function

Private Function StringToUtf8Bytes(ByVal str As String) As Variant
    Dim bSize As Long
    Dim RetVal As Long
    Dim buf() As Byte
    
    bSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(str), -1, 0, 0, 0, 0)
    If bSize = 0 Then
        Exit Function
    End If
    
    ReDim buf(bSize)
    RetVal = WideCharToMultiByte(CP_UTF8, 0, StrPtr(str), -1, VarPtr(buf(0)), bSize, 0, 0)
    If RetVal = 0 Then
        Logger.Log "StringToUtf8Bytes Error:", Err.LastDllError
        Exit Function
    End If
    StringToUtf8Bytes = buf
End Function

Private Function Utf16PtrToString(ByVal pUtf16String As Long) As String
    Dim StrLen As Long
    Dim RetVal As Long
    
    StrLen = lstrlenW(pUtf16String)
    Utf16PtrToString = String(StrLen, "*")
    lstrcpynW StrPtr(Utf16PtrToString), pUtf16String, StrLen
End Function

' Date Helpers
Private Function ToJulianDay(oleDate As Date) As Double
    ToJulianDay = CDbl(oleDate) + JULIANDAY_OFFSET
End Function

Private Function FromJulianDay(julianDay As Double) As Date
    FromJulianDay = CDate(julianDay - JULIANDAY_OFFSET)
End Function

'/////////////////////////////////////////////////////////////


