VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ONSClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function SysGetComputerName() As String
       Dim Computer As String
       Dim BufSize As Long
       Dim RetCode As Double
       Dim NullCharPos As Double

       Computer = Space(80)
       BufSize = Len(Computer)
       '---Call WINAPI
       RetCode = GetComputerName(Computer, BufSize)
       '---Search for the end of the string
       NullCharPos = InStr(Computer, Chr(0))
       If NullCharPos > 0 Then
          Computer = Left(Computer, NullCharPos - 1)
       Else
          Computer = "NO"
       End If
       SysGetComputerName = Computer
End Function

Public Function NZ(FieldValue, Optional Fvalue As String = "") As String
       If IsNull(FieldValue) Then
          NZ = Fvalue
       Else
          NZ = FieldValue
       End If
End Function

Public Function SQL(strSQL As String, rs As Recordset) As Recordset
       Debug.Print strSQL
       If rs.State = adStateOpen Then
          rs.Close
       End If
       rs.ActiveConnection = conn
       rs.CursorLocation = adUseClient
       rs.CursorType = adOpenDynamic
       rs.LockType = adLockOptimistic
       rs.Source = strSQL
       rs.Open
       Set SQL = rs
End Function

Public Sub LoadCombo(strSQL As String, LoadedComdo As Object, Optional LoadedListBox As Object)
       Dim rsLoadCombo As New ADODB.Recordset
       Set rsLoadCombo = SQL(strSQL, rsLoadCombo)
           LoadedComdo.Clear
       If LoadedListBox Is Nothing Then
       Else
          LoadedListBox.Clear
       End If
       If rsLoadCombo.RecordCount > 0 Then
          rsLoadCombo.MoveFirst
          While Not rsLoadCombo.EOF And Not rsLoadCombo.BOF
                LoadedComdo.AddItem IIf(IsNull(rsLoadCombo.Fields(0)), "", rsLoadCombo.Fields(0))
                If LoadedListBox Is Nothing Then
                Else
                   LoadedListBox.AddItem rsLoadCombo.Fields(1)
                End If
                rsLoadCombo.MoveNext
          Wend
          LoadedComdo.ListIndex = 0
       End If
End Sub

Public Function ONSConnection(ConObj As Object, ConnectionString As String) As Object
       With ConObj
            If .State = 1 Then .Close
            .ConnectionString = ConnectionString
            .Open
           Set ONSConnection = ConObj
       End With
End Function

Public Function DBname() As String
       On Error GoTo Ext
       Dim DBn As String, fn As Integer
       DBname = ""
       fn = FreeFile
       Open App.Path & "\Dbn.ons" For Input As #fn
       Input #fn, DBn
       DBname = DBn
       Close #1
       Exit Function
Ext:
       If Err.Number = 53 Then
          DBZMsgbox "Dbn.dbz file is not found", dbzCritical, "Sorry.......": Exit Sub
       End If
       If Err.Number = 62 Then
          DBZMsgbox "Please insert Database Name in Dbn.dbz file.", dbzCritical, "Sorry.......": Exit Sub
       End If
End Function

Public Function MachName() As String
       On Error GoTo Ext
       Dim servname As String
       MachName = ""
       Open App.Path & "\mac.ons" For Input As #1
       Input #1, servname
       MachName = servname
       Close #1
       Exit Function
Ext:
       If Err.Number = 53 Then
          DBZMsgbox "Dbn.dbz file is not found", dbzCritical, "Sorry.......": Exit Function
       End If
       If Err.Number = 62 Then
          DBZMsgbox "Please insert Database Name in Dbn.dbz file.", dbzCritical, "Sorry.......": Exit Function
       End If
End Function

Public Sub SetIcon(frm As Form)
       On Error Resume Next
       frm.Icon = LoadPicture(App.Path & "\icon.ico")
End Sub

Private Function IsAlphaNumeric(CharCode As Integer) As Boolean
        If (CharCode >= 65 And CharCode <= 90) Then
           IsAlphaNumeric = True
        ElseIf (CharCode >= 97 And CharCode <= 122) Then
           IsAlphaNumeric = True
        ElseIf (CharCode >= 48 And CharCode <= 57) Then
           IsAlphaNumeric = True
        Else
           IsAlphaNumeric = False
        End If
End Function

Public Function RandString(strLen As Integer) As String
       Dim xChar As Integer
       Randomize
       While strLen <> 0
             xChar = Rnd * 255
             If IsAlphaNumeric(xChar) = True Then
                RandString = RandString & Chr(xChar)
                strLen = strLen - 1
             End If
       Wend
 End Function
 
 Public Function EncripPassword(strPassword As String) As String
        s = Len(strPassword)
        Dim login, pass1 As String
        Dim passch(10) As String
        Dim ctr
        Dim passnew As String
            pass1 = Len(strPassword)
            ctr = 1
            Do While ctr <= Val(pass1)
               passch(ctr) = Mid(strPassword, ctr, 1)
               passtemp = Asc(passch(ctr)) + 17
               passnew = CStr(passnew) & Chr(passtemp)
               passtemp = ""
               ctr = ctr + 1
            Loop
        EncripPassword = passnew
 End Function
 
 Public Function EncryptText(strText As String, ByVal strPwd As String) As String
        Dim i As Integer, c As Integer
        Dim strBuff As String
        If Len(strPwd) Then
           For i = 1 To Len(strText)
               c = Asc(Mid$(strText, i, 1))
               c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
               strBuff = strBuff & Chr$(c And &HFF)
           Next i
        Else
           strBuff = strText
        End If
        EncryptText = strBuff
End Function

Public Function ReportTemp(CrystalReportName As CrystalReport, strReprotFileName As String, strSelectionFormula As String, strWindowsTitle As String, Optional strParameterfields As String, Optional strFormulafields As String)
       On Error GoTo Ext
       CrystalReportName.Reset
       CrystalReportName.ReportFileName = strReprotFileName
       CrystalReportName.SelectionFormula = strSelectionFormula
       CrystalReportName.WindowState = crptMaximized
       CrystalReportName.WindowTitle = strWindowsTitle
       CrystalReportName.WindowAllowDrillDown = True
       CrystalReportName.WindowShowGroupTree = True
       If strParameterfields <> "" Then
          Dim strT
          strT = Split(strParameterfields, "<DBZ>")
          For i = 0 To UBound(strT)
              CrystalReportName.ParameterFields(i) = strT(i)
          Next
       End If
       If strFormulafields <> "" Then
          strT = Split(strFormulafields, "<DBZ>")
          For i = 0 To UBound(strT)
              CrystalReportName.Formulas(i) = strT(i)
          Next
       End If
       CrystalReportName.Action = 1
       Exit Function
Ext:
End Function

Public Function ShowReport(ReportViewer As CRViewer, ReportName As Object, SelectedFormaula As String, Optional optFormula As String)
       'On error GoTo Ext
       Screen.MousePointer = vbHourglass
   
       'ReportViewer.Refresh
       ReportViewer.ReportSource = ReportName
       ReportName.RecordSelectionFormula = SelectedFormaula
       ReportViewer.EnableExportButton = True
       ReportViewer.EnableRefreshButton = True
       ReportViewer.EnableSearchControl = True
       ReportViewer.EnableSearchExpertButton = True
        
       If Not optFormula = "" Then
          Call Curr(optFormula, ReportName)
       End If
    
       ReportName.PaperSize = crPaperA4
       ReportViewer.Zoom (100)
       ReportViewer.ViewReport
       Screen.MousePointer = vbDefault
       Exit Function
Label:
       Screen.MousePointer = vbDefault
End Function

Public Function ReturnListIndex(lstBox As ListBox, strItem As String) As Integer
       ReturnListIndex = -1
       Dim i As Integer
       For i = 0 To lstBox.ListCount - 1
           If UCase(lstBox.List(i)) = UCase(strItem) Then
              ReturnListIndex = i
              Exit Function
           End If
       Next
End Function

Public Function DecryptText(strText As String, ByVal strPwd As String) As String
       Dim strBuff As String
       Dim i As Integer, c As Integer
       If Len(strPwd) Then
          For i = 1 To Len(strText)
              c = Asc(Mid$(strText, i, 1))
              c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
              strBuff = strBuff & Chr$(c And &HFF)
          Next i
       Else
          strBuff = strText
       End If
       DecryptText = strBuff
       Exit Function
End Function
