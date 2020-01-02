Attribute VB_Name = "AccessControl"
Option Explicit

Function Account_Initialize(Optional user_name As String) As Account
    Dim User As Account
    If user_name = nullstr Then
        Set User = GetPrivledges(UCase(VBA.Environ("Username")))
        If User Is Nothing Then
            Logger.Log "Creating new user " & VBA.Environ("Username"), UserLog
            Set Account_Initialize = AccessControl.AutoAddNewUser
        Else
            Logger.Log "Selected User : {" & User.ToString & "}", UserLog
            Set Account_Initialize = User
        End If
    Else
        Set Account_Initialize = GetPrivledges(UCase(user_name))
    End If
End Function

Function GetPrivledges(Name As String) As Account
    Dim User As Account
    Dim df As DataFrame

    Set df = DataAccess.GetUser(Name)
    If df.Rows = 0 Then
        Set GetPrivledges = Nothing
    Else
        Set User = New Account
        User.Name = df.records(1).item("Name")
        User.PrivledgeLevel = df.records(1).item("Privledge_Level")
        User.ProductLine = df.records(1).item("Product_Line")
        User.SecretSHA1 = df.records(1).item("Secret")
        User.FlaggedForPasswordChange = IIf(df.records(1).item("New_Secret_Required") = 0, False, True)
        Set GetPrivledges = User
    End If
End Function

Function AutoAddNewUser() As Account
    Dim new_user As Account
    Set new_user = New Account
    new_user.Name = Environ("Username")
    new_user.ChangeSetting "name", Environ("Username")
    ' Users have read only access by default
    new_user.PrivledgeLevel = USER_READONLY
    new_user.ChangeSetting "privledge_level", USER_READONLY
    new_user.ProductLine = "User"
    new_user.ChangeSetting "product_line", "User"
    new_user.SaveUserJson
    Logger.Log DataAccess.PushIQueryable(new_user, "user_privledges"), UserLog
    Set AutoAddNewUser = new_user
End Function

Public Sub ConfigControl()
'Initializes the password form for config access.
    Dim w As Window
    
    If Open_Config(App.current_user) Then
        ' Turn on Performance Mode
        App.PerformanceMode True
        If Windows.Count <> 1 Then
            For Each w In Windows
                If w.Parent.Name = ThisWorkbook.Name Then w.Visible = True
            Next w
        Else
            Application.Visible = True
        End If
        ' Show all worksheets
        GuiCommands.ShowAllSheets SAATI_Data_Manager.ThisWorkbook
        ThisWorkbook.Sheets("Administrator").Activate
        ' Turn off Performance Mode
        App.PerformanceMode False
    End If
End Sub

Private Function Open_Config(User As Account) As Boolean
'Performs a password check and opens config.
    Dim w As Window
    If CheckSecret(User) Then
        Open_Config = True
    Else
        PromptHandler.AccessDenied
        Open_Config = False
    End If
End Function

Private Function CheckSecret(User As Account) As Boolean
' Compares password hashes for match
    CheckSecret = IIf(User.GetSecret = GetSHA1Hash(PromptHandler.GetPassword), True, False)
End Function

Public Sub CreateNewAdmin()
' Creates or changes the admin password.
    Dim new_admin As String
    SpecManager.StartApp
    If CheckSecret(App.current_user) Then
        new_admin = PromptHandler.UserInput( _
            SingleLineText, "Access Control", "Enter user-name for new admin account :")
    End If
    DataAccess.FlagUserForSecretChange new_admin
    SpecManager.StopApp
End Sub

Public Sub ChangeSecret(User As Account)
' Changes a password in the db (SHA1 hash)
    If DataAccess.ChangeUserSecret(User.Name, GetSHA1Hash(PromptHandler.ChangePassword)) = DB_PUSH_FAILURE Then
        PromptHandler.Error "Password was not changed! Contact Admin"
    Else
        PromptHandler.Success "Password was changed succesfully!"
    End If
End Sub

Public Sub Hashing_Test()
    Dim input_text As String
    input_text = App.GUI.CreateInputBox(Password, "Password Hash Test", "Password : ")
    Debug.Print input_text & " : " & GetSHA1Hash(input_text)
End Sub

' Convert the string into bytes so we can use the above functions
' From Chris Hulbert: http://splinter.com.au/blog

Public Function GetSHA1Hash(str)
  Dim i As Integer
  Dim arr() As Byte
  ReDim arr(0 To Len(str) - 1) As Byte
  For i = 0 To Len(str) - 1
   arr(i) = Asc(Mid(str, i + 1, 1))
  Next i
  GetSHA1Hash = Replace(LCase(HexDefaultSHA1(arr)), " ", "")
End Function

Function HexDefaultSHA1(message() As Byte) As String
    Dim H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long
    DefaultSHA1 message, H1, H2, H3, H4, H5
    HexDefaultSHA1 = DecToHex5(H1, H2, H3, H4, H5)
End Function

Function HexSHA1(message() As Byte, ByVal Key1 As Long, ByVal Key2 As Long, ByVal Key3 As Long, ByVal Key4 As Long) As String
    Dim H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long
    xSHA1 message, Key1, Key2, Key3, Key4, H1, H2, H3, H4, H5
    HexSHA1 = DecToHex5(H1, H2, H3, H4, H5)
End Function

Sub DefaultSHA1(message() As Byte, H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long)
    xSHA1 message, &H5A827999, &H6ED9EBA1, &H8F1BBCDC, &HCA62C1D6, H1, H2, H3, H4, H5
End Sub

Sub xSHA1(message() As Byte, ByVal Key1 As Long, ByVal Key2 As Long, ByVal Key3 As Long, ByVal Key4 As Long, H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long)
 'CA62C1D68F1BBCDC6ED9EBA15A827999 + "abc" = "A9993E36 4706816A BA3E2571 7850C26C 9CD0D89D"
 '"abc" = "A9993E36 4706816A BA3E2571 7850C26C 9CD0D89D"

    Dim U As Long, P As Long
    Dim FB As FourBytes, OL As OneLong
    Dim i As Integer
    Dim w(80) As Long
    Dim A As Long, B As Long, C As Long, D As Long, E As Long
    Dim T As Long

    H1 = &H67452301: H2 = &HEFCDAB89: H3 = &H98BADCFE: H4 = &H10325476: H5 = &HC3D2E1F0

    U = UBound(message) + 1: OL.L = U32ShiftLeft3(U): A = U \ &H20000000: LSet FB = OL 'U32ShiftRight29(U)

    ReDim Preserve message(0 To (U + 8 And -64) + 63)
    message(U) = 128

    U = UBound(message)
    message(U - 4) = A
    message(U - 3) = FB.D
    message(U - 2) = FB.C
    message(U - 1) = FB.B
    message(U) = FB.A

    While P < U
        For i = 0 To 15
            FB.D = message(P)
            FB.C = message(P + 1)
            FB.B = message(P + 2)
            FB.A = message(P + 3)
            LSet OL = FB
            w(i) = OL.L
            P = P + 4
        Next i

        For i = 16 To 79
            w(i) = U32RotateLeft1(w(i - 3) Xor w(i - 8) Xor w(i - 14) Xor w(i - 16))
        Next i

        A = H1: B = H2: C = H3: D = H4: E = H5

        For i = 0 To 19
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(A), E), w(i)), Key1), ((B And C) Or ((Not B) And D)))
            E = D: D = C: C = U32RotateLeft30(B): B = A: A = T
        Next i
        For i = 20 To 39
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(A), E), w(i)), Key2), (B Xor C Xor D))
            E = D: D = C: C = U32RotateLeft30(B): B = A: A = T
        Next i
        For i = 40 To 59
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(A), E), w(i)), Key3), ((B And C) Or (B And D) Or (C And D)))
            E = D: D = C: C = U32RotateLeft30(B): B = A: A = T
        Next i
        For i = 60 To 79
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(A), E), w(i)), Key4), (B Xor C Xor D))
            E = D: D = C: C = U32RotateLeft30(B): B = A: A = T
        Next i

        H1 = U32Add(H1, A): H2 = U32Add(H2, B): H3 = U32Add(H3, C): H4 = U32Add(H4, D): H5 = U32Add(H5, E)
    Wend
End Sub

Function U32Add(ByVal A As Long, ByVal B As Long) As Long
    If (A Xor B) < 0 Then
        U32Add = A + B
    Else
        U32Add = (A Xor &H80000000) + B Xor &H80000000
    End If
End Function

Function U32ShiftLeft3(ByVal A As Long) As Long
    U32ShiftLeft3 = (A And &HFFFFFFF) * 8
    If A And &H10000000 Then U32ShiftLeft3 = U32ShiftLeft3 Or &H80000000
End Function

Function U32ShiftRight29(ByVal A As Long) As Long
    U32ShiftRight29 = (A And &HE0000000) \ &H20000000 And 7
End Function

Function U32RotateLeft1(ByVal A As Long) As Long
    U32RotateLeft1 = (A And &H3FFFFFFF) * 2
    If A And &H40000000 Then U32RotateLeft1 = U32RotateLeft1 Or &H80000000
    If A And &H80000000 Then U32RotateLeft1 = U32RotateLeft1 Or 1
End Function

Function U32RotateLeft5(ByVal A As Long) As Long
    U32RotateLeft5 = (A And &H3FFFFFF) * 32 Or (A And &HF8000000) \ &H8000000 And 31
    If A And &H4000000 Then U32RotateLeft5 = U32RotateLeft5 Or &H80000000
End Function

Function U32RotateLeft30(ByVal A As Long) As Long
    U32RotateLeft30 = (A And 1) * &H40000000 Or (A And &HFFFC) \ 4 And &H3FFFFFFF
    If A And 2 Then U32RotateLeft30 = U32RotateLeft30 Or &H80000000
End Function

Function DecToHex5(ByVal H1 As Long, ByVal H2 As Long, ByVal H3 As Long, ByVal H4 As Long, ByVal H5 As Long) As String
    Dim H As String, L As Long
    DecToHex5 = "00000000 00000000 00000000 00000000 00000000"
    H = Hex(H1): L = Len(H): Mid(DecToHex5, 9 - L, L) = H
    H = Hex(H2): L = Len(H): Mid(DecToHex5, 18 - L, L) = H
    H = Hex(H3): L = Len(H): Mid(DecToHex5, 27 - L, L) = H
    H = Hex(H4): L = Len(H): Mid(DecToHex5, 36 - L, L) = H
    H = Hex(H5): L = Len(H): Mid(DecToHex5, 45 - L, L) = H
End Function
