Attribute VB_Name = "Constants"
Option Explicit
'@Folder("Modules")

Public Const PublicDir          As String = "S:\Data Manager"
Public Const LocalDir           As String = "W:\App Development\Spec Manager"
Public Const SQLITE_PATH        As String = "C:\Users\cruff\source\SM - Final\Database\SAATI_Spec_Manager.db3"
Public Const GitBashExe         As String = "C:\Users\cruff\AppData\Local\Programs\Git\git-bash.exe"
Public Const GitRepo            As String = "C:\Users\cruff\source\SM - Final"
Public Const SlitterPath        As String = "S:\Public\04 Division - Filtration\Slitter set up parameters"

Public Const ISPEC_WARPING      As Long = 1 ' warper spec identifer for ispec factory
Public Const ISPEC_STYLE        As Long = 2 ' fabric style spec identifier for ispec factory
Public Const ISPEC_SLITTER      As Long = 3 ' heat slitter spec identifier for ispec factory
Public Const ISPEC_ULTRASONIC   As Long = 4 ' ultra sonic welder spec identifier for ispec factory

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
