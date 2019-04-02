Attribute VB_Name = "Constants"
Option Explicit
'@Folder("Modules")

Public Const PublicDir          As String = "S:\Data Manager"
Public Const LocalDir           As String = "W:\App Development\Spec Manager"
Public Const SQLITE_PATH        As String = "C:\Users\cruff\source\SM - Final\Database\SAATI_Spec_Manager.db3"
Public Const GitBashExe         As String = "C:\Users\cruff\AppData\Local\Programs\Git\git-bash.exe"
Public Const GitRepo            As String = "C:\Users\cruff\source\SM - Final"
Public Const SlitterPath        As String = "S:\Public\04 Division - Filtration\Slitter set up parameters"

' COM SERVER ERROR DESCRIPTIONS:
Public Const COM_PUSH_FAILURE      As Long = -1
Public Const COM_PUSH_COMPLETE     As Long = 0
Public Const COM_GET_FAILURE       As Long = 1 ' DM.Net could not find a spec matching your query

' SPEC MANAGER ERROR DESCRIPTIONS:
Public Const SM_SEARCH_SUCCESS     As Long = 0
Public Const SM_SEARCH_FAILURE     As Long = 1

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
