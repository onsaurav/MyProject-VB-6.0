Attribute VB_Name = "ONSMAC"
'Option Explicit
'
'---------------------------------------------------------------------------
' Used to get the MAC address.
'---------------------------------------------------------------------------
'
Private Const NCBNAMSZ As Long = 16
Private Const NCBENUM As Long = &H37
Private Const NCBRESET As Long = &H32
Private Const NCBASTAT As Long = &H33
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const HEAP_GENERATE_EXCEPTIONS As Long = &H4

Private Type NET_CONTROL_BLOCK  'NCB
        ncb_command    As Byte
        ncb_retcode    As Byte
        ncb_lsn        As Byte
        ncb_num        As Byte
        ncb_buffer     As Long
        ncb_length     As Integer
        ncb_callname   As String * NCBNAMSZ
        ncb_name       As String * NCBNAMSZ
        ncb_rto        As Byte
        ncb_sto        As Byte
        ncb_post       As Long
        ncb_lana_num   As Byte
        ncb_cmd_cplt   As Byte
        ncb_reserve(9) As Byte 'Reserved, must be 0
        ncb_event      As Long
End Type

Private Type ADAPTER_STATUS
        adapter_address(5) As Byte
        rev_major          As Byte
        reserved0          As Byte
        adapter_type       As Byte
        rev_minor          As Byte
        duration           As Integer
        frmr_recv          As Integer
        frmr_xmit          As Integer
        iframe_recv_err    As Integer
        xmit_aborts        As Integer
        xmit_success       As Long
        recv_success       As Long
        iframe_xmit_err    As Integer
        recv_buff_unavail  As Integer
        t1_timeouts        As Integer
        ti_timeouts        As Integer
        Reserved1          As Long
        free_ncbs          As Integer
        max_cfg_ncbs       As Integer
        max_ncbs           As Integer
        xmit_buf_unavail   As Integer
        max_dgram_size     As Integer
        pending_sess       As Integer
        max_cfg_sess       As Integer
        max_sess           As Integer
        max_sess_pkt_size  As Integer
        name_count         As Integer
End Type

Private Type NAME_BUFFER
        name_(0 To NCBNAMSZ - 1) As Byte
        name_num                 As Byte
        name_flags               As Byte
End Type

Private Type ASTAT
        adapt             As ADAPTER_STATUS
        NameBuff(0 To 29) As NAME_BUFFER
End Type

Private Declare Function Netbios Lib "netapi32" _
        (pncb As NET_CONTROL_BLOCK) As Byte

Private Declare Sub CopyMemory Lib "kernel32" _
        Alias "RtlMoveMemory" (hpvDest As Any, ByVal _
        hpvSource As Long, ByVal cbCopy As Long)

Private Declare Function GetProcessHeap Lib "kernel32" () As Long

Private Declare Function HeapAlloc Lib "kernel32" _
        (ByVal hHeap As Long, ByVal dwFlags As Long, _
        ByVal dwBytes As Long) As Long
     
Private Declare Function HeapFree Lib "kernel32" _
        (ByVal hHeap As Long, ByVal dwFlags As Long, _
        lpMem As Any) As Long

Public Function GetMacAddress() As String
        Dim l As Long
        Dim lngError As Long
        Dim lngSize As Long
        Dim pAdapt As Long
        Dim pAddrStr As Long
        Dim pASTAT As Long
        Dim strTemp As String
        Dim strAddress As String
        Dim strMACAddress As String
        Dim AST As ASTAT
        Dim NCB As NET_CONTROL_BLOCK
    
        '---------------------------------------------------------------------------
        ' Get the network interface card's MAC address.
        '----------------------------------------------------------------------------
        On Error GoTo ErrorHandler
        GetMacAddress = ""
        strMACAddress = ""
        
        ' Try to get MAC address from NetBios. Requires NetBios installed.
        '
        ' Supported on 95, 98, ME, NT, 2K, XP
        '
        ' Results Connected Disconnected
        ' ------- --------- ------------
        '   XP       OK         Fail (Fail after reboot)
        '   NT       OK         OK   (OK after reboot)
        '   98       OK         OK   (OK after reboot)
        '   95       OK         OK   (OK after reboot)
        
        NCB.ncb_command = NCBRESET
        Call Netbios(NCB)
    
        NCB.ncb_callname = "*               "
        NCB.ncb_command = NCBASTAT
        NCB.ncb_lana_num = 0
        NCB.ncb_length = Len(AST)
    
        pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or _
                           HEAP_ZERO_MEMORY, NCB.ncb_length)
        If pASTAT = 0 Then GoTo ErrorHandler
        NCB.ncb_buffer = pASTAT
        Call Netbios(NCB)
        Call CopyMemory(AST, NCB.ncb_buffer, Len(AST))
    
        strMACAddress = Right$("00" & Hex(AST.adapt.adapter_address(0)), 2) & _
                        Right$("00" & Hex(AST.adapt.adapter_address(1)), 2) & _
                        Right$("00" & Hex(AST.adapt.adapter_address(2)), 2) & _
                        Right$("00" & Hex(AST.adapt.adapter_address(3)), 2) & _
                        Right$("00" & Hex(AST.adapt.adapter_address(4)), 2) & _
                        Right$("00" & Hex(AST.adapt.adapter_address(5)), 2)
    
        Call HeapFree(GetProcessHeap(), 0, pASTAT)
    
        GetMacAddress = strMACAddress
        GoTo NormalExit
ErrorHandler:
        Call MsgBox(Err.Description, vbCritical, "Error")
NormalExit:
End Function

''''For checking registration, machine and date sequence
'''28/01/2008 by Rafayel

Public Function CheckDateSequence() As Boolean
       Dim TotalDays As Integer
       Dim NewDate As Date, PreviousDate As Date, PresentDate As Date
       Dim rsDateAdd As New ADODB.Recordset
       Dim rsDateCheck As New ADODB.Recordset
       Dim rsDateSequence As New ADODB.Recordset
       CheckDateSequence = True
       Set rsDateCheck = dbzCodeSystemObject.SQL("Select * from Registration", rsDateCheck)
       If rsDateCheck.RecordCount > 0 Then
          If IsNull(rsDateCheck!RgsDateFrom) = True Then
             frmfirstuse.Show vbModal
             CheckDateSequence = True
             Exit Function
'             rsDateCheck!RgsDateFrom = FormatDateTime(Date, vbShortDate)
'             rsDateCheck.Update
          End If
       End If
       Set rsDateCheck = dbzCodeSystemObject.SQL("Select * from Registration", rsDateCheck)
       If rsDateCheck.RecordCount > 0 Then
          If Trim(rsDateCheck!RgsMode) = "T" Then
             If DateDiff("D", Date, FormatDateTime(DateAdd("d", (rsDateCheck!rgstotalDays) - 1, rsDateCheck!RgsDateFrom), vbShortDate)) < 0 Then
                DBZMsgbox "Your Trial Time Already Finished. " & Chr(13) & " Please Complete The Registration.", dbzInformation, "Sorry..."
                conn.Execute "Update Registration set RgsTotalDays=0"
                frmfirstuse.Show vbModal
                CheckDateSequence = False

                Exit Function
             Else
                Set rsDateSequence = dbzCodeSystemObject.SQL("Select * from DateSequence Order by Date", rsDateSequence)
                If rsDateSequence.RecordCount > 0 Then
                   rsDateSequence.MoveLast
                   If DateDiff("D", FormatDateTime(rsDateSequence!Date, vbShortDate), Date) < 0 Then
                      rsDateCheck!rgstotalDays = 0
                      rsDateCheck.Update
                      frmfirstuse.Show vbModal
                      CheckDateSequence = False
                      Exit Function
                   End If
                End If
                Set rsDateAdd = dbzCodeSystemObject.SQL("Select * from DateSequence Where Date = #" & Date & "#", rsDateAdd)
                If rsDateAdd.RecordCount = 0 Then
                   conn.Execute "Insert Into DateSequence Values (#" & Date & "#)"
                End If
             End If
          End If
       End If
End Function

Public Function CheckMechine() As Boolean
       Dim strMACAddress As String
       Dim rsMachine As New ADODB.Recordset
       Dim rsMachineCheck As New ADODB.Recordset
       Dim rsRegistration As New ADODB.Recordset
       CheckMechine = True
       strMACAddress = GetMacAddress
       Set rsRegistration = dbzCodeSystemObject.SQL("Select * from Registration", rsRegistration)
       If rsRegistration.RecordCount > 0 Then
          Set rsMachine = dbzCodeSystemObject.SQL("Select * from Machine", rsMachine)
          If rsMachine.RecordCount <= rsRegistration!RgsUser Then
             Set rsMachineCheck = dbzCodeSystemObject.SQL("Select * from Machine Where MachineName='" & strMACAddress & "'", rsMachineCheck)
             If rsMachineCheck.RecordCount = 0 Then
                conn.Execute "Insert into Machine (MachineName) Values ('" & strMACAddress & "')"
             End If
          Else
             DBZMsgbox "Number of User Limit Over" & "Please Complete The Registration.", dbzInformation, "Sorry..."
             CheckMechine = False
             Open App.Path & "/msscc.ti" For Output As #1
             Print #1, "0"
             Close #1
             Exit Function
          End If
       End If
End Function

Public Function CheckRegistered() As Boolean
       Dim rsRegistration As New ADODB.Recordset
       Dim rsCheck  As New ADODB.Recordset
       CheckRegistered = True
       Set rsRegistration = dbzCodeSystemObject.SQL("Select * from Registration", rsRegistration)
       If rsRegistration.RecordCount > 0 Then
           If Trim(rsRegistration!RgsMode) = "T" Then
              If rsRegistration!rgstotalDays = 0 Then
                 frmfirstuse.Show vbModal
                 rsRegistration.Requery
                 If Trim(rsRegistration!RgsMode) = "R" Then
                    CheckRegistered = True
                    Exit Function
                 End If
              End If
              If CheckDateSequence = False Then
                 CheckRegistered = False
                 Exit Function
              End If
              bTrialOrRegist = "T"
           ElseIf Trim(rsRegistration!RgsMode) = "R" Then
              'If frmMode.optClient.Value = True Then
                 If CheckMechine = False Then
                    CheckRegistered = False
                    Exit Function
                 End If
                 bTrialOrRegist = "B"
              'End If
           End If
       End If
End Function


'============ Kartik/13-01-2007/Check the LogIn Mode
Public Function CheckMode() As Boolean
        'On Error GoTo Ext
        CheckMode = True
        Open App.Path & "/msscc.ti" For Input As #1
        Input #1, getstats
        Close #1
        If Val(Trim(getstats)) = 0 Then
           frmMode.Show 1
           Call ChangeDatabasePassword(dbzCodeSystemObject.machName, PassChar)
           Call NewConnection(conn)
           Call CreateAccessDSN("HighLights", dbzCodeSystemObject.machName)
        ElseIf Val(Trim(getstats)) = 111 Then
            Call ChangeDatabasePassword(dbzCodeSystemObject.machName, PassChar)
           Call NewConnection(conn)
           Call CreateAccessDSN("HighLights", dbzCodeSystemObject.machName)
        End If
        If App.PrevInstance Then
           DBZMsgbox "The Program is aleady running in your machine.", dbzInformation, "Sorry..."
           CheckMode = False
           Exit Function
        End If
        Open App.Path & "/msscc.ti" For Output As #1
        Print #1, "111"
        Close #1
        Exit Function
Ext:
        DBZMsgbox "You don't have the authority to use this System.", dbzInformation, "Sorry..."
        CheckMode = False
End Function


Public Function DecryptText(strText As String, ByVal strPwd As String) As String

3     Dim i As Integer, c As Integer
4     Dim strBuff As String
5     If Len(strPwd) Then
6         For i = 1 To Len(strText)
7             c = Asc(Mid$(strText, i, 1))
8             c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
9             strBuff = strBuff & Chr$(c And &HFF)
10         Next i
11     Else
12         strBuff = strText
13     End If
14     DecryptText = strBuff
15 Exit Function

End Function

