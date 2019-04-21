Attribute VB_Name = "ONSModule"
Public YesNOStatus As Boolean
Public ONSMain As New ONSClass
Public cOn  As New ADODB.Connection

Private Const MAX_PATH = 260
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_STATUSTEXT = 4

Private Const ODBC_ADD_SYS_DSN = 4
Private Const ODBC_REMOVE_SYS_DSN = 6

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXTA = (WM_USER + 100)
Private Const BFFM_SETSELECTIONA = (WM_USER + 102)


Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Private Type BrowseInfo
        hwndOwner As Long
        pIDLRoot As Long
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfn As Long
        lParam As Long
        iImage As Long
End Type

Public Enum MassageType
       ONSInformation
       ONSCritical
       ONSYesNo
End Enum
 
Public Enum YesNO
       ONSYes
       ONSNo
End Enum

Public Sub Main()
       Call NewConnection(cOn)
       'App.HelpFile = App.Path & "\HELP.chm"
       'FrmOpening.Show
End Sub

Public Function PassChar() As String
       Dim Password As String
       Open App.Path & "\ONSPASS" For Input As #1
       Input #1, Password
       PassChar = dbzCodeSystemObject.DecryptText(Password, "DBZ")
       Close #1
End Function

Public Function NewConnection(cOn As ADODB.Connection)
       Dim strNewSQl As String
       strNewSQl = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & ONSMain.MachName & "';Mode=ReadWrite|Share Deny None;Persist Security Info=False;Jet OLEDB:Database Password='" & PassChar & "'"
       Set cOn = ONSMain.ONSConnection(cOn, strNewSQl)
End Function

Public Function SavePictureToDB(rss As ADODB.Recordset, sFileName As String, RowN As Integer, ImageBox As Object)
       On Error GoTo Ext
       Dim oPict As StdPicture
       Dim strStream  As ADODB.Stream
       Set oPict = LoadPicture(sFileName)
       If oPict Is Nothing Then
          DBZMsgbox "Invalid Picture File!", ONSInformation, "Oops!"
          SavePictureToDB = False
          GoTo procExitSub
       End If
    
       Set strStream = New ADODB.Stream
       strStream.Type = adTypeBinary
       strStream.Open
       strStream.LoadFromFile sFileName
       rss.Fields(RowN).Value = strStream.Read
       rss.Update
       ImageBox.Picture = LoadPicture(sFileName)
       SavePictureToDB = True
       Exit Function
Ext:
       MsgBox Err.Description
       SavePictureToDB = False
End Function

Public Function LoadPictureFromDB(rsL As ADODB.Recordset, RowsI As Integer)
    On Error GoTo Ext
    Dim strStream As ADODB.Stream
    
    If rsL Is Nothing Then
        GoTo procNoPicture
    End If
    
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
    strStream.Write rsL.Fields(RowsI).Value
    strStream.SaveToFile "C:\Temp.bmp", adSaveCreateOverWrite
    LoadPictureFromDB = True
    Exit Function
Ext:
    LoadPictureFromDB = False
End Function

Public Sub loadGroupList(ImageBox1 As Object, GroupListView As ListView, strSQL As String, Optional ImageBox2 As Object, Optional Forth As Boolean)
       On Error Resume Next
       Dim rs As New ADODB.Recordset
       GroupListView.Icons = ImageBox2
       Set rs = dbzCodeSystemObject.SQL(strSQL, rs)
       GroupListView.ListItems.Clear
       If rs.RecordCount > 0 Then
          Dim imgX As ListImage
          rs.MoveFirst
          ImageBox1.ListImages.Clear
          ImageBox1.ImageHeight = 100
          ImageBox1.ImageWidth = 100
          While Not rs.EOF
                If LoadPictureFromDB(rs, 2) = True Then
                   Set imgX = ImageBox1.ListImages.Add(, , LoadPicture("C:\Temp.bmp"))
                   Kill ("C:\Temp.bmp")
                Else
                   Set imgX = ImageBox1.ListImages.Add(, , LoadPicture(App.Path & "\Image1.bmp"))
                End If
                rs.MoveNext
          Wend
          GroupListView.ListItems.Clear
          GroupListView.Icons = ImageBox1
          Dim lvItem As ListItem
          rs.MoveFirst
          i = 1
          Dim str As String
          str = ""
          While Not rs.EOF And Not rs.BOF
                If Forth = True Then
                   str = "," & rs.Fields(3)
                End If
                Set lvItem = GroupListView.ListItems.Add(, CStr(rs.Fields(0)) & "<DBZ>", rs.Fields(1) & str, i)
                i = i + 1
                rs.MoveNext
          Wend
       End If
End Sub

Public Sub ListOption(ListViewName As Object, Optional ListViewName1 As Object, Optional ComboBoxName As Integer)
       ListViewName.GridLines = False
       ListViewName.View = ComboBoxName
       If ComboBoxName = 3 Then
          ListViewName.GridLines = True
       End If
       ListViewName.Refresh
       If ListViewName1 Is Nothing Then
       Else
          ListViewName1.GridLines = False
          ListViewName1.View = ComboBoxName
          If ComboBoxName = 3 Then
             ListViewName1.GridLines = True
          End If
          ListViewName1.Refresh
       End If
End Sub

Public Function CreateAccessDSN(DSNName As String, DatabaseFullPath As String) As Boolean
       Dim sAttributes As String
       DatabaseFullPath = dbzCodeSystemObject.MachName
       
       If Dir(DatabaseFullPath) = "" Then Exit Function
       sAttributes = "DSN=" & DSNName & Chr(0) & "UID=Admin;pwd=" & PassChar & Chr(0)
       sAttributes = sAttributes & "DBQ=" & DatabaseFullPath & Chr(0)
       CreateAccessDSN = CreateDSN("Microsoft Access Driver (*.mdb)", sAttributes)
End Function

Public Function CreateDSN(Driver As String, Attributes As String) As Boolean
       CreateDSN = SQLConfigDataSource(0&, ODBC_ADD_SYS_DSN, Driver, Attributes)
End Function

Public Function RemoveAccessDSN(DSNName As String, DatabaseFullPath As String) As Boolean
       Dim sAttributes As String
       If Dir(DatabaseFullPath) = "" Then Exit Function
       sAttributes = "DSN=" & DSNName & Chr(0) & "UID=Admin;pwd=" & PassChar & Chr(0)
       sAttributes = sAttributes & "DBQ=" & DatabaseFullPath & Chr(0)
       RemoveAccessDSN = RemoveDSN("Microsoft Access Driver (*.mdb)", sAttributes)
End Function

Public Function RemoveDSN(Driver As String, Attributes As String) As Boolean
       RemoveDSN = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, Driver, Attributes)
End Function

Public Function BrowseForFolder(DefaultFolder As String, Optional Parent As Long = 0, Optional Caption As String = "") As String
       Dim bi As BrowseInfo
       Dim sResult As String, nResult As Long
       bi.hwndOwner = Parent
       bi.pIDLRoot = 0
       bi.pszDisplayName = String$(MAX_PATH, Chr$(0))
       If Len(Caption) > 0 Then
          bi.lpszTitle = Caption
       End If
       bi.ulFlags = BIF_RETURNONLYFSDIRS   'Or BIF_STATUSTEXT
       bi.lpfn = GetAddress(AddressOf BrowseCallbackProc)
       bi.lParam = 0
       bi.iImage = 0
         'Set local default folder string
         '(will be set in callback after dialog initializes)
       m_sDefaultFolder = DefaultFolder
         'Call API
       nResult = SHBrowseForFolder(bi)
         'Get result if successful
       If nResult <> 0 Then
          sResult = String(MAX_PATH, 0)
          If SHGetPathFromIDList(nResult, sResult) Then
             BrowseForFolder = Left$(sResult, InStr(sResult, Chr$(0)) - 1)
          End If
         'Free memory allocated by SHBrowseForFolder
          CoTaskMemFree nResult
       End If
End Function

Private Function GetAddress(nAddress As Long) As Long
        GetAddress = nAddress
End Function

Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
        Select Case uMsg
        Case BFFM_INITIALIZED
             'Note: This code was crashing VB when the default folder was empty
             If Len(m_sDefaultFolder) > 0 Then
                'Set default folder when dialog has initialized
                SendMessage hWnd, BFFM_SETSELECTIONA, True, ByVal m_sDefaultFolder
             End If
        End Select
End Function

Public Function CheckSecurityForButton(ByVal FormName As Form, ByVal username As String)
       Dim rscheckUserName As New ADODB.Recordset
       Dim strLevel As String
       Dim rsCheck As New ADODB.Recordset
       With rsCheck
            If .State = adStateOpen Then
               .Close
            End If
            .CursorType = adOpenStatic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open "select * from SecurityUserControl where SucUserName='" & username & "'and SucFormName='" & FormName.Name & "' and SucControlStatus='2'", conn
            If .RecordCount > 0 Then
            Else
               Set rscheckUserName = conn.Execute("select * from securityuser where UsrUserName='" & username & "'")
               If Not rscheckUserName.EOF And Not rscheckUserName.BOF Then
                  strLevel = rscheckUserName!UsrDepartment & "<SR>" & rscheckUserName!usrlevel
                  Set rscheckUserName = conn.Execute("select * from SecurityUserControl where SucUserName='" & strLevel & "' and SucFormName='" & FormName.Name & "' and  SucControlStatus='2'")
                  If Not rscheckUserName.EOF And Not rscheckUserName.BOF Then
                     Set rsCheck = rscheckUserName
                  Else
                     Exit Function
                  End If
               Else
                  Exit Function
               End If
            End If
            While Not rsCheck.EOF And Not rsCheck.BOF
                  If rsCheck.Fields("SucIndex") <> "" Then
                     FormName.Controls(rsCheck.Fields("SucFormControl"))(rsCheck.Fields("SucIndex")).Enabled = False
                  Else
                     FormName.Controls(rsCheck.Fields("SucFormControl")).Enabled = False
                  End If
                  rsCheck.MoveNext
            Wend
       End With
End Function

Public Function CheckSecurityForMenu(ByVal FormName As Form, ByVal username As String)
       Dim rsCheck As New ADODB.Recordset
       Dim rscheckUserName As New ADODB.Recordset
       Dim strLevel As String
       With rsCheck
            If .State = adStateOpen Then
               .Close
            End If
            .CursorType = adOpenStatic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open "select * from MenuUserSecurity where MusUserName='" & username & "'and  MusControlStatus='2'", conn
            If .RecordCount > 0 Then
            Else
                Set rscheckUserName = conn.Execute("select * from securityuser where UsrUserName='" & username & "'")
                If Not rscheckUserName.EOF And Not rscheckUserName.BOF Then
                   strLevel = rscheckUserName!UsrDepartment & "<SR>" & rscheckUserName!usrlevel
                   Set rscheckUserName = conn.Execute("select * from MenuUserSecurity where MusUserName='" & strLevel & "' and  MusControlStatus='2'")
                   If Not rscheckUserName.EOF And Not rscheckUserName.BOF Then
                      Set rsCheck = rscheckUserName
                   Else
                      Exit Function
                   End If
                Else
                   Exit Function
                End If
             End If
             
             While Not rsCheck.EOF And Not rsCheck.BOF
                   Dim Rsmnu As New ADODB.Recordset
                   If Rsmnu.State = adStateOpen Then
                      Rsmnu.Close
                   End If
                   With Rsmnu
                        .ActiveConnection = conn
                        .Source = "Select * From Menu where MnuFormName='" & FormName.Name & "'and MnuRoot='" & rsCheck!MusRoot & "'and MnuId='" & rsCheck!MusId & "'"
                        .CursorLocation = adUseClient
                        .CursorType = adOpenDynamic
                        .LockType = adLockOptimistic
                        .Open
                        .Sort = "mnuID"
                   End With
                   If Rsmnu.RecordCount > 0 Then
                      'FormName.ActiveBar21.Controls(Rsmnu.Fields("MnuName")).Enabled = False
                   End If
                   rsCheck.MoveNext
             Wend
       End With
End Function

Public Function ONSMsgbox(Massage As String, Optional msgType As MassageType = ONSInformation, Optional MsgTitle As String) As YesNO
       frmMsgbox.Caption = MsgTitle
       frmMsgbox.lblMassage.Caption = Massage
       Call frmMsgbox.MsgFormCotrol(msgType)
       frmMsgbox.Show vbModal
       If msgType = dbzYesNo Then
          If YesNOStatus = True Then
             DBZMsgbox = dbzYes
          Else
             DBZMsgbox = dbzNo
          End If
       End If
End Function

Public Function ChangeDatabasePassword(DBPath As String, oldPassword As String) As Boolean
       If Dir(DBPath) = "" Then Exit Function
       If Dir(App.Path & "\JTR4B11X") = "" Then Exit Function
       Call RemoveAccessDSN("HighLights", dbzCodeSystemObject.MachName)
       Dim newPassword As String
       Dim Password As String
       Open App.Path & "\JTR4B11X" For Input As #1
       Input #1, Password
       newPassword = dbzCodeSystemObject.DecryptText(Password, "DBZ")
       Close #1
       
       Dim db As DAO.Database
       Set db = OpenDatabase(DBPath, True, False, ";pwd=" & oldPassword)
       If Err.Number <> 0 Then Exit Function
       db.newPassword oldPassword, newPassword
       ChangeDatabasePassword = Err.Number = 0
       Open App.Path & "\HRGT9H4" For Output As #1
       Print #1, dbzCodeSystemObject.EncryptText(newPassword, "DBZ")
       Close #1
       Kill (App.Path & "\JTR4B11X")
       db.Close
End Function

