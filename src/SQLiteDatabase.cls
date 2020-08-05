VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=src/lib
Option Explicit

' Microsoft uses UTF-16, little endian byte order.
Private Const JULIANDAY_OFFSET As Double = 2415018.5
' Options for Text and Blob binding
Private Const SQLITE_STATIC      As Long = 0
Private Const SQLITE_TRANSIENT   As Long = -1
' System calls
Private Const CP_UTF8 As Long = 65001
#If Win64 Then
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As LongPtr, ByVal pSource As LongPtr, ByVal length As Long)
Private Declare PtrSafe Function lstrcpynW Lib "kernel32" (ByVal pwsDest As LongPtr, ByVal pwsSource As LongPtr, ByVal cchCount As Long) As LongPtr
Private Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal pwsDest As LongPtr, ByVal pwsSource As LongPtr) As LongPtr
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal pwsString As LongPtr) As Long
Private Declare PtrSafe Function SysAllocString Lib "OleAut32" (ByRef pwsString As LongPtr) As LongPtr
Private Declare PtrSafe Function SysStringLen Lib "OleAut32" (ByVal bstrString As LongPtr) As Long
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
#Else
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
#End If
'=====================================================================================
' SQLite StdCall Imports
'-----------------------
#If Win64 Then
' SQLite library version
Private Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As LongPtr ' PtrUtf8String
' Database connections
Private Declare PtrSafe Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr, ByVal iFlags As Long, ByVal zVfs As LongPtr) As Long ' PtrDb
Private Declare PtrSafe Function sqlite3_close Lib "SQLite3" (ByVal hDb As LongPtr) As Long
' Database connection error info
Private Declare PtrSafe Function sqlite3_errmsg Lib "SQLite3" (ByVal hDb As LongPtr) As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_errmsg16 Lib "SQLite3" (ByVal hDb As LongPtr) As LongPtr ' PtrUtf16String
Private Declare PtrSafe Function sqlite3_errcode Lib "SQLite3" (ByVal hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_extended_errcode Lib "SQLite3" (ByVal hDb As LongPtr) As Long
' Database connection change counts
Private Declare PtrSafe Function sqlite3_changes Lib "SQLite3" (ByVal hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_total_changes Lib "SQLite3" (ByVal hDb As LongPtr) As Long

' Statements
Private Declare PtrSafe Function sqlite3_prepare16_v2 Lib "SQLite3" _
    (ByVal hDb As LongPtr, ByVal pwsSql As LongPtr, ByVal nSqlLength As Long, ByRef hStmt As LongPtr, ByVal ppwsTailOut As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_step Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_reset Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_finalize Lib "SQLite3" (ByVal hStmt As LongPtr) As Long

' Statement column access (0-based indices)
Private Declare PtrSafe Function sqlite3_column_count Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_column_type Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_name16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrWString

Private Declare PtrSafe Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrData
Private Declare PtrSafe Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_bytes16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_double Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Double
Private Declare PtrSafe Function sqlite3_column_int Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongLong
Private Declare PtrSafe Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_text16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrWString
Private Declare PtrSafe Function sqlite3_column_value Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrSqlite3Value

' Statement parameter binding (1-based indices!)
Private Declare PtrSafe Function sqlite3_bind_parameter_count Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_parameter_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long) As LongPtr
Private Declare PtrSafe Function sqlite3_bind_parameter_index Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramName As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_null Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long) As Long
Private Declare PtrSafe Function sqlite3_bind_blob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_zeroblob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal nBytes As Long) As Long
Private Declare PtrSafe Function sqlite3_bind_double Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal value As Double) As Long
Private Declare PtrSafe Function sqlite3_bind_int Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal value As Long) As Long
Private Declare PtrSafe Function sqlite3_bind_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal value As LongLong) As Long
Private Declare PtrSafe Function sqlite3_bind_text Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal psValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_text16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pswValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_value Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pSqlite3Value As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_clear_bindings Lib "SQLite3" (ByVal hStmt As LongPtr) As Long

'Backup
Private Declare PtrSafe Function sqlite3_sleep Lib "SQLite3" (ByVal msToSleep As Long) As Long
Private Declare PtrSafe Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As LongPtr, ByVal zDestName As LongPtr, ByVal hDbSource As LongPtr, ByVal zSourceName As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As LongPtr, ByVal nPage As Long) As Long
Private Declare PtrSafe Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
#Else

' SQLite library version
Private Declare Function sqlite3_libversion Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_libversion@0" () As Long ' PtrUtf8String
' Database connections
Private Declare Function sqlite3_open16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_open16@8" (ByVal pwsFileName As Long, ByRef hDb As Long) As Long ' PtrDb
Private Declare Function sqlite3_open_v2 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_open_v2@16" (ByVal pwsFileName As Long, ByRef hDb As Long, ByVal iFlags As Long, ByVal zVfs As Long) As Long ' PtrDb
Private Declare Function sqlite3_close Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_close@4" (ByVal hDb As Long) As Long
' Database connection error info
Private Declare Function sqlite3_errmsg Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_errmsg@4" (ByVal hDb As Long) As Long ' PtrUtf8String
Private Declare Function sqlite3_errmsg16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_errmsg16@4" (ByVal hDb As Long) As Long ' PtrUtf16String
Private Declare Function sqlite3_errcode Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_errcode@4" (ByVal hDb As Long) As Long
Private Declare Function sqlite3_extended_errcode Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_extended_errcode@4" (ByVal hDb As Long) As Long
' Database connection change counts
Private Declare Function sqlite3_changes Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_changes@4" (ByVal hDb As Long) As Long
Private Declare Function sqlite3_total_changes Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_total_changes@4" (ByVal hDb As Long) As Long

' Statements
Private Declare Function sqlite3_prepare16_v2 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_prepare16_v2@20" _
    (ByVal hDb As Long, ByVal pwsSql As Long, ByVal nSqlLength As Long, ByRef hStmt As Long, ByVal ppwsTailOut As Long) As Long
Private Declare Function sqlite3_step Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_step@4" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_reset Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_reset@4" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_finalize Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_finalize@4" (ByVal hStmt As Long) As Long

' Statement column access (0-based indices)
Private Declare Function sqlite3_column_count Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_count@4" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_column_type Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_type@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_name Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_name@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_name16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_name16@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrWString

Private Declare Function sqlite3_column_blob Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_blob@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrData
Private Declare Function sqlite3_column_bytes Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_bytes@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_bytes16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_bytes16@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_double Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_double@8" (ByVal hStmt As Long, ByVal iCol As Long) As Double
Private Declare Function sqlite3_column_int Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_int@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_int64 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_int64@8" (ByVal hStmt As Long, ByVal iCol As Long) As Currency ' UNTESTED ....?
Private Declare Function sqlite3_column_text Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_text@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_text16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_text16@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrWString
Private Declare Function sqlite3_column_value Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_value@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrSqlite3Value

' Statement parameter binding (1-based indices!)
Private Declare Function sqlite3_bind_parameter_count Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_parameter_count@4" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_bind_parameter_name Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_parameter_name@8" (ByVal hStmt As Long, ByVal paramIndex As Long) As Long
Private Declare Function sqlite3_bind_parameter_index Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_parameter_index@8" (ByVal hStmt As Long, ByVal paramName As Long) As Long
Private Declare Function sqlite3_bind_null Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_null@8" (ByVal hStmt As Long, ByVal paramIndex As Long) As Long
Private Declare Function sqlite3_bind_blob Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_blob@20" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_bind_zeroblob Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_zeroblob@12" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal nBytes As Long) As Long
Private Declare Function sqlite3_bind_double Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_double@16" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal value As Double) As Long
Private Declare Function sqlite3_bind_int Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_int@12" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal value As Long) As Long
Private Declare Function sqlite3_bind_int64 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_int64@16" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal value As Currency) As Long ' UNTESTED ....?
Private Declare Function sqlite3_bind_text Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_text@20" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal psValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_bind_text16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_text16@20" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pswValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_bind_value Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_value@12" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pSqlite3Value As Long) As Long
Private Declare Function sqlite3_clear_bindings Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_clear_bindings@4" (ByVal hStmt As Long) As Long

'Backup
Private Declare Function sqlite3_sleep Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_sleep@4" (ByVal msToSleep As Long) As Long
Private Declare Function sqlite3_backup_init Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_init@16" (ByVal hDbDest As Long, ByVal zDestName As Long, ByVal hDbSource As Long, ByVal zSourceName As Long) As Long
Private Declare Function sqlite3_backup_step Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_step@8" (ByVal hBackup As Long, ByVal nPage As Long) As Long
Private Declare Function sqlite3_backup_finish Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_finish@4" (ByVal hBackup As Long) As Long
Private Declare Function sqlite3_backup_remaining Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_remaining@4" (ByVal hBackup As Long) As Long
Private Declare Function sqlite3_backup_pagecount Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_pagecount@4" (ByVal hBackup As Long) As Long
#End If
'=====================================================================================
' Initialize - load libraries explicitly
#If Win64 Then
Private hSQLiteLibrary As LongPtr
Private hSQLiteStdCallLibrary As LongPtr
#Else
Private hSQLiteLibrary As Long
Private hSQLiteStdCallLibrary As Long
#End If
' -------------------------------------------------------------------------------------
' -------------------------------------------------------------------------------------
' -------------------------------------------------------------------------------------
'@Folder("Classes")
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
'  database properties
Public Property Get dbPath() As Variant:  dbPath = m_dbPath: End Property
Private Property Let dbPath(ByVal vNewValue As Variant): m_dbPath = vNewValue: End Property
'================================================================
' public properties
Public Property Get header() As Variant: header = m_header: End Property
Public Property Get Data() As Variant: Data = m_data: End Property
Public Property Get qtyRows() As Long: qtyRows = m_qtyRows: End Property
Public Property Get qtyColumns() As Long: qtyColumns = m_qtyColumns: End Property
Public Property Get NumRows() As Long: NumRows = m_qtyRows: End Property
Public Property Get NumColumns() As Long: NumColumns = m_qtyColumns: End Property
'================================================================
' opens a database x64
Public Sub openDb(sqliPath As String)
  If sqliPath = nullstr Or Len(dir(sqliPath)) = 0 Then
    err.Raise 1, "Database not found", "Please select a valid file address"
    Exit Sub
  End If
  dbPath = sqliPath
End Sub

' Execute sql statement x64
Public Sub execute(ssql As String)
    #If Win64 Then
    Dim myStmtHandle As LongPtr
    #Else
    Dim myStmtHandle As Long
    #End If
    Call initQuery(executa, myStmtHandle, Me.dbPath, ssql)
End Sub

' Execute select statement x64
Public Sub selectQry(ssql As String)
    #If Win64 Then
    Dim myStmtHandle As LongPtr
    #Else
    Dim myStmtHandle As Long
    #End If

    Call initQuery(getRows, myStmtHandle, Me.dbPath, "select count(*) from (" & ssql & ")") 'get rows
    Call initQuery(getColumns, myStmtHandle, Me.dbPath, ssql) 'get columns
    Call getQryValue(myStmtHandle)
End Sub

' Initialize a query x64
#If Win64 Then
Private Function initQuery(q_comand As initType, ByRef myStmtHandle As LongPtr, dbPath As String, ssql As String) As Double
    Dim myDbHandle As LongPtr
#Else
Private Function initQuery(q_comand As initType, ByRef myStmtHandle As Long, dbPath As String, ssql As String) As Double
    Dim myDbHandle As Long
#End If
    Dim RetVal As Long
    Dim stepMsg As String, InitReturn As Long, msgErro As Long
    
    ' Default path is ThisWorkbook.Path but can specify other path where the .dlls reside.
    InitReturn = SQLite3Initialize(ThisWorkbook.path & "\libs\")
    RetVal = SQLite3Finalize(myStmtHandle)
    RetVal = SQLite3Open(dbPath, myDbHandle)
    RetVal = SQLite3PrepareV2(myDbHandle, ssql, myStmtHandle)
    RetVal = SQLite3Step(myStmtHandle)
    msgErro = RetVal

    If q_comand = executa Then
    RetVal = SQLite3PrepareV2(myDbHandle, ssql, myStmtHandle)
    If msgErro <> SQLITE_DONE Then err.Raise 1, "Database not found", "Database command error " & msgErro
    ElseIf q_comand = getRows Then
    m_qtyRows = getRowsCount(myStmtHandle)
    ElseIf q_comand = getColumns Then
    m_qtyColumns = SQLite3ColumnCount(myStmtHandle)
    End If
End Function

'returns array of the table x64
#If Win64 Then
Private Sub getQryValue(ByVal stmtHandle As LongPtr)
#Else
Private Sub getQryValue(ByVal stmtHandle As Long)
#End If
    Dim colName As String
    Dim colType As Long
    Dim colTypeName As String
    Dim colValue As Variant
    Dim arrdata() As Variant, arrheader As Variant
    Dim rowsCount As Long
    Dim i As Long, j As Long, k As Long, m As Long

    ReDim arrdata(1 To qtyRows, 1 To Me.qtyColumns)
    ReDim arrheader(1 To 1, 1 To Me.qtyColumns)

    colType = SQLite3ColumnType(stmtHandle, 0)
    For m = 1 To Me.qtyRows
      k = k + 1
      For i = 0 To Me.qtyColumns - 1
        colName = SQLite3ColumnName(stmtHandle, i)
        If arrheader(1, i + 1) = nullstr Then arrheader(1, i + 1) = colName
        colType = SQLite3ColumnType(stmtHandle, i)
        colTypeName = TypeName(colType)
        colValue = ColumnValue(stmtHandle, i, colType)
        arrdata(k, i + 1) = colValue
      Next i
      Call SQLite3Step(stmtHandle)
    Next m
    m_data = arrdata
    m_header = arrheader
End Sub

' returns the row count x64
#If Win64 Then
Private Function getRowsCount(ByVal stmtHandle As LongPtr) As Long
#Else
Private Function getRowsCount(ByVal stmtHandle As Long) As Long
#End If
  Dim colType As Long
  getRowsCount = ColumnValue(stmtHandle, 0, SQLite3ColumnType(stmtHandle, 0)) 'return
End Function

' Converts sqlitetype constant to string representation x64
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

' Returns the value of a particular column in a record x64
#If Win64 Then
Private Function ColumnValue(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
#Else
Private Function ColumnValue(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
#End If
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
' -------------------------------------------------------------------------------------
' -------------------------------------------------------------------------------------
' -------------------------------------------------------------------------------------
Public Function SQLite3Initialize(Optional ByVal libDir As String) As Long
    ' A nice option here is to call SetDllDirectory, but that API is only available since Windows XP SP1.
    If libDir = "" Then libDir = ThisWorkbook.path
    If Right(libDir, 1) <> "\" Then libDir = libDir & "\"
    
    If hSQLiteLibrary = 0 Then
        hSQLiteLibrary = LoadLibrary(libDir + "SQLite3.dll")
        If hSQLiteLibrary = 0 Then
            Debug.Print "SQLite3Initialize Error Loading " + libDir + "SQLite3.dll : ", err.LastDllError
            SQLite3Initialize = SQLITE_INIT_ERROR
            Exit Function
        End If
    End If
        
    #If Win64 Then
    #Else
    If hSQLiteStdCallLibrary = 0 Then
        hSQLiteStdCallLibrary = LoadLibrary(libDir + "SQLite3_StdCall.dll")
        If hSQLiteStdCallLibrary = 0 Then
            Debug.Print "SQLite3Initialize Error Loading " + libDir + "SQLite3_StdCall.dll : ", err.LastDllError
            SQLite3Initialize = SQLITE_INIT_ERROR
            Exit Function
        End If
    End If
    #End If
    SQLite3Initialize = SQLITE_INIT_OK
End Function
Public Sub SQLite3Free()
   Dim refCount As Long
   If hSQLiteStdCallLibrary <> 0 Then
        refCount = FreeLibrary(hSQLiteStdCallLibrary)
        hSQLiteStdCallLibrary = 0
        If refCount = 0 Then
            Debug.Print "SQLite3Free Error Freeing SQLite3_StdCall.dll:", refCount, err.LastDllError
        End If
    End If
    If hSQLiteLibrary <> 0 Then
        refCount = FreeLibrary(hSQLiteLibrary)
        hSQLiteLibrary = 0
        If refCount = 0 Then
            Debug.Print "SQLite3Free Error Freeing SQLite3.dll:", refCount, err.LastDllError
        End If
    End If
End Sub
'=====================================================================================
' SQLite library version
Public Function SQLite3LibVersion() As String
    SQLite3LibVersion = Utf8PtrToString(sqlite3_libversion())
End Function
'=====================================================================================
' Database connections
#If Win64 Then
Public Function SQLite3Open(ByVal fileName As String, ByRef dbHandle As LongPtr) As Long
#Else
Public Function SQLite3Open(ByVal fileName As String, ByRef dbHandle As Long) As Long
#End If
    SQLite3Open = sqlite3_open16(StrPtr(fileName), dbHandle)
End Function

#If Win64 Then
Public Function SQLite3OpenV2(ByVal fileName As String, ByRef dbHandle As LongPtr, ByVal flags As Long, ByVal vfsName As String) As Long
#Else
Public Function SQLite3OpenV2(ByVal fileName As String, ByRef dbHandle As Long, ByVal flags As Long, ByVal vfsName As String) As Long
#End If

    Dim bufFileName() As Byte
    Dim bufVfsName() As Byte
    bufFileName = StringToUtf8Bytes(fileName)
    If vfsName = Empty Then
        SQLite3OpenV2 = sqlite3_open_v2(VarPtr(bufFileName(0)), dbHandle, flags, 0)
    Else
        bufVfsName = StringToUtf8Bytes(vfsName)
        SQLite3OpenV2 = sqlite3_open_v2(VarPtr(bufFileName(0)), dbHandle, flags, VarPtr(bufVfsName(0)))
    End If

End Function

#If Win64 Then
Public Function SQLite3Close(ByVal dbHandle As LongPtr) As Long
#Else
Public Function SQLite3Close(ByVal dbHandle As Long) As Long
#End If
    SQLite3Close = sqlite3_close(dbHandle)
End Function

'=====================================================================================
' Error information

#If Win64 Then
Public Function SQLite3ErrMsg(ByVal dbHandle As LongPtr) As String
#Else
Public Function SQLite3ErrMsg(ByVal dbHandle As Long) As String
#End If
    SQLite3ErrMsg = Utf8PtrToString(sqlite3_errmsg(dbHandle))
End Function

#If Win64 Then
Public Function SQLite3ErrCode(ByVal dbHandle As LongPtr) As Long
#Else
Public Function SQLite3ErrCode(ByVal dbHandle As Long) As Long
#End If
    SQLite3ErrCode = sqlite3_errcode(dbHandle)
End Function

#If Win64 Then
Public Function SQLite3ExtendedErrCode(ByVal dbHandle As LongPtr) As Long
#Else
Public Function SQLite3ExtendedErrCode(ByVal dbHandle As Long) As Long
#End If
    SQLite3ExtendedErrCode = sqlite3_extended_errcode(dbHandle)
End Function

'=====================================================================================
' Change Counts

#If Win64 Then
Public Function SQLite3Changes(ByVal dbHandle As LongPtr) As Long
#Else
Public Function SQLite3Changes(ByVal dbHandle As Long) As Long
#End If
    SQLite3Changes = sqlite3_changes(dbHandle)
End Function

#If Win64 Then
Public Function SQLite3TotalChanges(ByVal dbHandle As LongPtr) As Long
#Else
Public Function SQLite3TotalChanges(ByVal dbHandle As Long) As Long
#End If
    SQLite3TotalChanges = sqlite3_total_changes(dbHandle)
End Function

'=====================================================================================
' Statements

#If Win64 Then
Public Function SQLite3PrepareV2(ByVal dbHandle As LongPtr, ByVal sql As String, ByRef stmtHandle As LongPtr) As Long
#Else
Public Function SQLite3PrepareV2(ByVal dbHandle As Long, ByVal sql As String, ByRef stmtHandle As Long) As Long
#End If
    ' Only the first statement (up to ';') is prepared. Currently we don't retrieve the 'tail' pointer.
    SQLite3PrepareV2 = sqlite3_prepare16_v2(dbHandle, StrPtr(sql), Len(sql) * 2, stmtHandle, 0)
End Function

#If Win64 Then
Public Function SQLite3Step(ByVal stmtHandle As LongPtr) As Long
#Else
Public Function SQLite3Step(ByVal stmtHandle As Long) As Long
#End If
    SQLite3Step = sqlite3_step(stmtHandle)
End Function

#If Win64 Then
Public Function SQLite3Reset(ByVal stmtHandle As LongPtr) As Long
#Else
Public Function SQLite3Reset(ByVal stmtHandle As Long) As Long
#End If
    SQLite3Reset = sqlite3_reset(stmtHandle)
End Function

#If Win64 Then
Public Function SQLite3Finalize(ByVal stmtHandle As LongPtr) As Long
#Else
Public Function SQLite3Finalize(ByVal stmtHandle As Long) As Long
#End If
    SQLite3Finalize = sqlite3_finalize(stmtHandle)
End Function

'=====================================================================================
' Statement column access (0-based indices)

#If Win64 Then
Public Function SQLite3ColumnCount(ByVal stmtHandle As LongPtr) As Long
#Else
Public Function SQLite3ColumnCount(ByVal stmtHandle As Long) As Long
#End If
    SQLite3ColumnCount = sqlite3_column_count(stmtHandle)
End Function

#If Win64 Then
Public Function SQLite3ColumnType(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Long
#Else
Public Function SQLite3ColumnType(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Long
#End If
    SQLite3ColumnType = sqlite3_column_type(stmtHandle, ZeroBasedColIndex)
End Function

#If Win64 Then
Public Function SQLite3ColumnName(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As String
#Else
Public Function SQLite3ColumnName(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As String
#End If
    SQLite3ColumnName = Utf8PtrToString(sqlite3_column_name(stmtHandle, ZeroBasedColIndex))
End Function

#If Win64 Then
Public Function SQLite3ColumnDouble(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Double
#Else
Public Function SQLite3ColumnDouble(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Double
#End If
    SQLite3ColumnDouble = sqlite3_column_double(stmtHandle, ZeroBasedColIndex)
End Function

#If Win64 Then
Public Function SQLite3ColumnInt32(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Long
#Else
Public Function SQLite3ColumnInt32(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Long
#End If
    SQLite3ColumnInt32 = sqlite3_column_int(stmtHandle, ZeroBasedColIndex)
End Function

#If Win64 Then
Public Function SQLite3ColumnText(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As String
#Else
Public Function SQLite3ColumnText(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As String
#End If
    SQLite3ColumnText = Utf8PtrToString(sqlite3_column_text(stmtHandle, ZeroBasedColIndex))
End Function

#If Win64 Then
Public Function SQLite3ColumnDate(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Date
#Else
Public Function SQLite3ColumnDate(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Date
#End If
    SQLite3ColumnDate = FromJulianDay(sqlite3_column_double(stmtHandle, ZeroBasedColIndex))
End Function

#If Win64 Then
Public Function SQLite3ColumnBlob(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Byte()
    Dim ptr As LongPtr
#Else
Public Function SQLite3ColumnBlob(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Byte()
    Dim ptr As Long
#End If

    Dim length As Long
    Dim buf() As Byte
    
    ptr = sqlite3_column_blob(stmtHandle, ZeroBasedColIndex)
    length = sqlite3_column_bytes(stmtHandle, ZeroBasedColIndex)
    ReDim buf(length - 1)
    RtlMoveMemory VarPtr(buf(0)), ptr, length
    SQLite3ColumnBlob = buf
End Function
'=====================================================================================
' Statement bindings

#If Win64 Then
Public Function SQLite3BindText(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal value As String) As Long
#Else
Public Function SQLite3BindText(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal value As String) As Long
#End If
    SQLite3BindText = sqlite3_bind_text16(stmtHandle, OneBasedParamIndex, StrPtr(value), -1, SQLITE_TRANSIENT)
End Function

#If Win64 Then
Public Function SQLite3BindDouble(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal value As Double) As Long
#Else
Public Function SQLite3BindDouble(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal value As Double) As Long
#End If
    SQLite3BindDouble = sqlite3_bind_double(stmtHandle, OneBasedParamIndex, value)
End Function

#If Win64 Then
Public Function SQLite3BindInt32(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal value As Long) As Long
#Else
Public Function SQLite3BindInt32(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal value As Long) As Long
#End If
    SQLite3BindInt32 = sqlite3_bind_int(stmtHandle, OneBasedParamIndex, value)
End Function

#If Win64 Then
Public Function SQLite3BindDate(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal value As Date) As Long
#Else
Public Function SQLite3BindDate(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal value As Date) As Long
#End If
    SQLite3BindDate = sqlite3_bind_double(stmtHandle, OneBasedParamIndex, ToJulianDay(value))
End Function

#If Win64 Then
Public Function SQLite3BindBlob(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByRef value() As Byte) As Long
#Else
Public Function SQLite3BindBlob(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByRef value() As Byte) As Long
#End If
    Dim length As Long
    length = UBound(value) - LBound(value) + 1
    SQLite3BindBlob = sqlite3_bind_blob(stmtHandle, OneBasedParamIndex, VarPtr(value(0)), length, SQLITE_TRANSIENT)
End Function

#If Win64 Then
Public Function SQLite3BindNull(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long) As Long
#Else
Public Function SQLite3BindNull(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long) As Long
#End If
    SQLite3BindNull = sqlite3_bind_null(stmtHandle, OneBasedParamIndex)
End Function

#If Win64 Then
Public Function SQLite3BindParameterCount(ByVal stmtHandle As LongPtr) As Long
#Else
Public Function SQLite3BindParameterCount(ByVal stmtHandle As Long) As Long
#End If
    SQLite3BindParameterCount = sqlite3_bind_parameter_count(stmtHandle)
End Function

#If Win64 Then
Public Function SQLite3BindParameterName(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long) As String
#Else
Public Function SQLite3BindParameterName(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long) As String
#End If
    SQLite3BindParameterName = Utf8PtrToString(sqlite3_bind_parameter_name(stmtHandle, OneBasedParamIndex))
End Function

#If Win64 Then
Public Function SQLite3BindParameterIndex(ByVal stmtHandle As LongPtr, ByVal paramName As String) As Long
#Else
Public Function SQLite3BindParameterIndex(ByVal stmtHandle As Long, ByVal paramName As String) As Long
#End If
    Dim buf() As Byte
    buf = StringToUtf8Bytes(paramName)
    SQLite3BindParameterIndex = sqlite3_bind_parameter_index(stmtHandle, VarPtr(buf(0)))
End Function

#If Win64 Then
Public Function SQLite3ClearBindings(ByVal stmtHandle As LongPtr) As Long
#Else
Public Function SQLite3ClearBindings(ByVal stmtHandle As Long) As Long
#End If
    SQLite3ClearBindings = sqlite3_clear_bindings(stmtHandle)
End Function


'=====================================================================================
' Backup
Public Function SQLite3Sleep(ByVal timeToSleepInMs As Long) As Long
    SQLite3Sleep = sqlite3_sleep(timeToSleepInMs)
End Function

#If Win64 Then
Public Function SQLite3BackupInit(ByVal dbHandleDestination As LongPtr, ByVal destinationName As String, ByVal dbHandleSource As LongPtr, ByVal sourceName As String) As LongPtr
#Else
Public Function SQLite3BackupInit(ByVal dbHandleDestination As Long, ByVal destinationName As String, ByVal dbHandleSource As Long, ByVal sourceName As String) As Long
#End If
    Dim bufDestinationName() As Byte
    Dim bufSourceName() As Byte
    bufDestinationName = StringToUtf8Bytes(destinationName)
    bufSourceName = StringToUtf8Bytes(sourceName)
    SQLite3BackupInit = sqlite3_backup_init(dbHandleDestination, VarPtr(bufDestinationName(0)), dbHandleSource, VarPtr(bufSourceName(0)))
End Function

#If Win64 Then
Public Function SQLite3BackupFinish(ByVal backupHandle As LongPtr) As Long
#Else
Public Function SQLite3BackupFinish(ByVal backupHandle As Long) As Long
#End If
    SQLite3BackupFinish = sqlite3_backup_finish(backupHandle)
End Function

#If Win64 Then
Public Function SQLite3BackupStep(ByVal backupHandle As LongPtr, ByVal numberOfPages) As Long
#Else
Public Function SQLite3BackupStep(ByVal backupHandle As Long, ByVal numberOfPages) As Long
#End If
    SQLite3BackupStep = sqlite3_backup_step(backupHandle, numberOfPages)
End Function

#If Win64 Then
Public Function SQLite3BackupPageCount(ByVal backupHandle As LongPtr) As Long
#Else
Public Function SQLite3BackupPageCount(ByVal backupHandle As Long) As Long
#End If
    SQLite3BackupPageCount = sqlite3_backup_pagecount(backupHandle)
End Function

#If Win64 Then
Public Function SQLite3BackupRemaining(ByVal backupHandle As LongPtr) As Long
#Else
Public Function SQLite3BackupRemaining(ByVal backupHandle As Long) As Long
#End If
    SQLite3BackupRemaining = sqlite3_backup_remaining(backupHandle)
End Function

' String Helpers
#If Win64 Then
Function Utf8PtrToString(ByVal pUtf8String As LongPtr) As String
#Else
Function Utf8PtrToString(ByVal pUtf8String As Long) As String
#End If
    Dim buf As String
    Dim cSize As Long
    Dim RetVal As Long
    
    cSize = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, 0, 0)
    ' cSize includes the terminating null character
    If cSize <= 1 Then
        Utf8PtrToString = ""
        Exit Function
    End If
    
    Utf8PtrToString = String(cSize - 1, "*") ' and a termintating null char.
    RetVal = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, StrPtr(Utf8PtrToString), cSize)
    If RetVal = 0 Then
        Debug.Print "Utf8PtrToString Error:", err.LastDllError
        Exit Function
    End If
End Function

Function StringToUtf8Bytes(ByVal str As String) As Variant
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
        Debug.Print "StringToUtf8Bytes Error:", err.LastDllError
        Exit Function
    End If
    StringToUtf8Bytes = buf
End Function

#If Win64 Then
Function Utf16PtrToString(ByVal pUtf16String As LongPtr) As String
#Else
Function Utf16PtrToString(ByVal pUtf16String As Long) As String
#End If
    Dim StrLen As Long
    
    StrLen = lstrlenW(pUtf16String)
    Utf16PtrToString = String(StrLen, "*")
    lstrcpynW StrPtr(Utf16PtrToString), pUtf16String, StrLen
End Function

' Date Helpers
Public Function ToJulianDay(oleDate As Date) As Double
    ToJulianDay = CDbl(oleDate) + JULIANDAY_OFFSET
End Function

Public Function FromJulianDay(julianDay As Double) As Date
    FromJulianDay = CDate(julianDay - JULIANDAY_OFFSET)
End Function

