VERSION 5.00
Object = "{84AF4DF3-4B59-4D87-85BA-FA878460F831}#6.4#0"; "StormDoor.ocx"
Begin VB.Form frmMain 
   Caption         =   "StormBase"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin StormDoorX.StormDoor door 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8916
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub door_ConnectionClosed()
    door.Quit
    End
End Sub

Private Sub door_LostFocus()
    door.SetFocus
End Sub

Private Sub door_UserInput(data As String)
Dim bKeepGoing As Boolean
Dim totrecs As Long

    Select Case Menu
        Case "MainMenu"
            Select Case UCase(data)
                Case "I"
                    Shell "C:\BBS\msyncfos\run.bat " & door.ThisNode & " " & door.SocketHandle & " 50 C:\BBS\StormBase\IceEdit\iceedit.bat " & Command
                Case "R"
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open "SELECT PostID FROM Posts WHERE BaseID = " & BaseID, cn, adOpenStatic
                    If rs.RecordCount > 0 Then
                        iOnMsg = 1
                        door.MoveToPos 4, 3
                        rs.Close
                        DisplayNextMsg
                    Else
                        rs.Close
                        door.ClearDisplay
                        door.DisplayANSI App.Path & "\TITLE.ANS"
                        door.MoveToPos 4, 3
                        door.ChangeColor fRed, bWhite
                        door.Display "NO MESSAGES IN THIS BASE"
                        door.ChangeColor fYellow
                        door.MoveToPos 6, 3
                        door.Display "::PRESS ANY KEY::"
                        Menu = "WaitForAnyKeyMain"
                    End If
                Case "N"
                    If rs.State = adStateOpen Then rs.Close
                    'rs.Open "SELECT Posts.PostID, PostRead.ReadID FROM Posts LEFT JOIN PostRead ON Posts.PostID = PostRead.PostID WHERE Posts.BaseID=" & BaseID & " AND (" & UserID & " OR PostRead.UserID Is Null) ORDER BY Posts.PostID", cn, adOpenStatic
                    rs.Open "SELECT PostID FROM Posts WHERE BaseID = " & BaseID, cn, adOpenStatic
                    iOnMsg = 0
                    For i = 1 To rs.RecordCount
                        If rs2.State = adStateOpen Then rs2.Close
                        rs2.Open "SELECT ReadID FROM PostRead WHERE UserID = " & UserID & " AND PostID = " & rs.Fields("PostID"), cn, adOpenStatic
                        If rs2.RecordCount = 0 Then
                            iOnMsg = i
                            Exit For
                        End If
                        rs.MoveNext
                    Next
                        
                    If iOnMsg > 0 Then
                        door.MoveToPos 4, 3
                        rs.Close
                        DisplayNextMsg
                    Else
                        rs.Close
                        door.ClearDisplay
                        door.DisplayANSI App.Path & "\TITLE.ANS"
                        door.MoveToPos 4, 3
                        door.ChangeColor fRed, bWhite
                        door.Display "NO NEW MESSAGES IN THIS BASE"
                        door.ChangeColor fYellow
                        door.MoveToPos 6, 3
                        door.Display "::PRESS ANY KEY::"
                        Menu = "WaitForAnyKeyMain"
                    End If
                Case "A"
                    door.ClearDisplay
                    door.DisplayANSI App.Path & "\ABOUT.ANS"
                    
                    door.MoveToPos 21, 3
                    door.ChangeColor fGray, bBlack
                    door.Display "version   "
                    door.ChangeColor fDarkGray, bBlack
                    door.Display "0.1beta   "
                    door.ChangeColor fGray, bBlack
                    door.Display "http://www.vaper.org   "
                    door.ChangeColor fDarkGray, bBlack
                    door.Display "nullvalue@bbsmates.com   "
                    door.ChangeColor fGray, bBlack
                    door.Display "oddware"
                    
                    door.MoveToPos 24, 2 '62
                    door.Display ":: PRESS ANY KEY ::"
                    Menu = "WaitForAnyKeyMain"
                Case "P"
                    If BaseID > 0 Then
                        door.ClearDisplay
                        door.DisplayANSI App.Path & "\TITLE.ANS"
                        door.MoveToPos 4, 1
                        door.ChangeColor fLighCyan, bBlack
                        door.Display "  Message Base: " & BaseName & "\n\n"
                        door.ChangeColor fWhite
                        door.Display "       Subject: "
                        door.ChangeColor fBlue, bWhite
                        door.Display Space(50)
                        door.MoveToPos 6, 17
                        sInput = ""
                        Menu = "AskNewPostSubject"
                    End If
                Case "S"
                    Dim iLine As Integer, iChar As Integer
                    door.ClearDisplay
                    door.DisplayANSI App.Path & "\SELBASE.ANS"
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open "SELECT * FROM Bases ORDER BY BaseOrder", cn, adOpenStatic
                    iLine = 7
                    iChar = 65
                    For i = 1 To 26
                        SelectName(iChar - 65).Letter = ""
                        SelectName(iChar - 65).BaseID = 0
                        SelectName(iChar - 65).BaseName = ""
                    Next
                    For i = 1 To rs.RecordCount
                        If rs.EOF = True Then Exit For
                        door.MoveToPos iLine, 10
                        door.ChangeColor fLightCyan, bCyan
                        door.Display " " & Chr(iChar) & " "
                        door.ChangeColor fCyan
                        door.Display " " & rs("BaseName")
                        SelectName(iChar - 64).Letter = Chr(iChar)
                        SelectName(iChar - 64).BaseID = rs("BaseID")
                        SelectName(iChar - 64).BaseName = rs("BaseName")
                        rs.MoveNext
                        If rs.EOF = True Then Exit For
                        iChar = iChar + 1
                        door.MoveToPos iLine, 42
                        door.ChangeColor fLightCyan, bCyan
                        door.Display " " & Chr(iChar) & " "
                        door.ChangeColor fCyan
                        door.Display " " & rs("BaseName")
                        SelectName(iChar - 64).Letter = Chr(iChar)
                        SelectName(iChar - 64).BaseID = rs("BaseID")
                        SelectName(iChar - 64).BaseName = rs("BaseName")
                        rs.MoveNext
                        iChar = iChar + 1
                        iLine = iLine + 1
                    Next
                    rs.Close
                    door.MoveToPos 21, 47
                    'door.ChangeColor fWhite, bBlack
                    Menu = "SelectBaseName"
                Case "]"
                    bKeepGoing = True
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open "SELECT * FROM Bases ORDER BY BaseOrder", cn, adOpenStatic
                    For i = 1 To rs.RecordCount + 1
                        If bKeepGoing = True Then
                            If rs("BaseID") = BaseID Then
                                If rs.AbsolutePosition = rs.RecordCount Then
                                    rs.MoveFirst
                                Else
                                    rs.MoveNext
                                End If
                                    
                                bKeepGoing = False
                            Else
                                rs.MoveNext
                            End If
                        Else
                            BaseName = rs("BaseName")
                            BaseID = rs("BaseID")
                            door.ChangeColor fLighCyan, bBlack
                            door.MoveToPos 7, 39
                            door.Display Space(26)
                            door.MoveToPos 7, 39
                            door.Display BaseName
                            door.MoveToPos 5, 17
                            door.ChangeColor fLighCyan
                            door.Display "[" & Pad_String(rs.AbsolutePosition, 2, " ", 0) & "/" & Pad_String(rs.RecordCount, 2, " ", 0) & "]"
                            rs.Close
                            
                            rs.Open "SELECT Count(*) as RecCnt FROM Posts WHERE BaseID=" & BaseID, cn, adOpenStatic
                            
                            door.MoveToPos 5, 35
                            totrecs = rs("RecCnt")
                            door.Display Pad_String(CStr(totrecs), 5, " ", 0) & " POSTS"
                            rs.Close
                            rs.Open "SELECT Count(*) as RecCnt FROM PostRead WHERE UserID=" & UserID & " AND BaseID = " & BaseID, cn, adOpenStatic
                            door.MoveToPos 5, 56
                            door.Display Pad_String(CStr(totrecs - rs("RecCnt")), 5, " ", 0) & " NEW"

                            door.MoveToPos 21, 50
                            rs.Close
                            Exit For
                        End If
                    Next
                Case "["
                    bKeepGoing = True
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open "SELECT * FROM Bases ORDER BY BaseOrder", cn, adOpenStatic
                    If rs.RecordCount > 0 Then rs.MoveLast
                    For i = 1 To rs.RecordCount + 1
                        If bKeepGoing = True Then
                            If rs("BaseID") = BaseID Then
                                If rs.AbsolutePosition = 1 Then
                                    rs.MoveLast
                                Else
                                    rs.MovePrevious
                                End If
                                    
                                bKeepGoing = False
                            Else
                                rs.MovePrevious
                            End If
                            
                        Else
                            BaseName = rs("BaseName")
                            BaseID = rs("BaseID")
                            door.ChangeColor fLighCyan, bBlack
                            door.MoveToPos 7, 39
                            door.Display Space(26)
                            door.MoveToPos 7, 39
                            door.Display BaseName
                            door.MoveToPos 5, 17
                            door.ChangeColor fLighCyan
                            door.Display "[" & Pad_String(rs.AbsolutePosition, 2, " ", 0) & "/" & Pad_String(rs.RecordCount, 2, " ", 0) & "]"
                            rs.Close
                            rs.Open "SELECT Count(*) as RecCnt FROM Posts WHERE BaseID=" & BaseID, cn, adOpenStatic
                            
                            door.MoveToPos 5, 35
                            totrecs = rs("RecCnt")
                            door.Display Pad_String(CStr(totrecs), 5, " ", 0) & " POSTS"
                            rs.Close
                            rs.Open "SELECT Count(*) as RecCnt FROM PostRead WHERE UserID=" & UserID & " AND BaseID = " & BaseID, cn, adOpenStatic
                            door.MoveToPos 5, 56
                            door.Display Pad_String(CStr(totrecs - rs("RecCnt")), 5, " ", 0) & " NEW"

                            door.MoveToPos 21, 50
                            rs.Close
                            Exit For
                        End If
                    Next
                Case "E"
                    door.ClearDisplay
                    door.DisplayANSI App.Path & "\TITLE.ANS"
                    door.MoveToPos 4, 1
                    door.ChangeColor fDarkGray
                    door.Display "  WELCOME TO STORM-BASE\n\n"
                    door.ChangeColor fWhite
                    door.Display "  How would you like your name to be displayed? "
                    door.ChangeColor fBlue, bWhite
                    door.Display Space(25)
                    door.MoveToPos 6, 49
                    door.Display DisplayName
                    sInput = DisplayName
                    Menu = "EditDisplayName"
                Case "Q"
                    door.Quit
                    End
            End Select
        Case "SelectBaseName"
            If data = "0" Then
                DisplayMain
            Else
                For i = 1 To 26
                    If SelectName(i).Letter = UCase(data) Then
                        BaseID = SelectName(i).BaseID
                        BaseName = SelectName(i).BaseName
                        DisplayMain
                    End If
                Next
            End If
        Case "AskDisplayName"
            If Left(data, 2) <> Chr(27) & "[" Then
                If data = vbCr Or data = vbCrLf Then
                    If Trim(sInput) <> "" Then
                        If rs.State = adStateOpen Then rs.Close
                        rs.Open "INSERT INTO Users (UserName, DisplayName) VALUES ('" & sFix(door.Alias) & "','" & sFix(sInput) & "')", cn
                        Username = door.Alias
                        DisplayName = sInput
                        If rs.State = adStateOpen Then rs.Close
                        rs.Open "SELECT * FROM Users WHERE UserName = '" & sFix(Username) & "'", cn, adOpenStatic
                        UserID = rs("UserID")
                        rs.Close
                        DisplayMain
                    End If
                ElseIf data = Chr(8) Then
                    If sInput <> "" Then
                        door.Display Chr(27) & "[D " & Chr(27) & "[D"
                        sInput = Left(sInput, Len(sInput) - 1)
                    End If
                Else
                    If Len(sInput) < 25 Then
                        sInput = sInput & data
                        door.Display data
                    End If
                End If
            End If
        Case "EditDisplayName"
            If Left(data, 2) <> Chr(27) & "[" Then
                If data = vbCr Or data = vbCrLf Then
                    If Trim(sInput) <> "" Then
                        If rs.State = adStateOpen Then rs.Close
                        rs.Open "UPDATE Users SET DisplayName = '" & sFix(sInput) & "' WHERE UserID = " & UserID, cn
                        DisplayName = sInput
                        DisplayMain
                    End If
                ElseIf data = Chr(8) Then
                    If sInput <> "" Then
                        door.Display Chr(27) & "[D " & Chr(27) & "[D"
                        sInput = Left(sInput, Len(sInput) - 1)
                    End If
                Else
                    If Len(sInput) < 25 Then
                        sInput = sInput & data
                        door.Display data
                    End If
                End If
            End If
        Case "AskNewPostSubject"
            If Left(data, 2) <> Chr(27) & "[" Then
                If data = vbCr Or data = vbCrLf Then
                    If Trim(sInput) = "" Then
                        door.ChangeColor fYellow
                        door.Display "\n\n::MESSAGE CANCELED::"
                        Menu = "WaitForAnyKeyMain"
                    Else
                        sSubject = sInput
                        door.MoveToPos 8, 1
                        door.ChangeColor fLighCyan, bBlack
                        door.Display "    Press ""/"" on a blank line for commands\n\n"
                        door.ChangeColor fLightBlue, bBlack
                        door.Display "  >"
                        door.ChangeColor fWhite, bBlack
                        sInput = ""
                        Menu = "NewPostBody"
                    End If
                ElseIf data = Chr(8) Then
                    If sInput <> "" Then
                        door.Display Chr(27) & "[D " & Chr(27) & "[D"
                        sInput = Left(sInput, Len(sInput) - 1)
                    End If
                ElseIf data = Chr(27) Then
                    door.ChangeColor fYellow
                    door.Display "\n\n::MESSAGE CANCELED::"
                    Menu = "WaitForAnyKeyMain"
                Else
                    If Len(sInput) < 50 Then
                        sInput = sInput & data
                        door.Display data
                    End If
                End If
            End If
        Case "NewPostBody"
            If Left(data, 2) <> Chr(27) & "[" Then
                If data = vbCr Or data = vbCrLf Then
                    door.ChangeColor fLightBlue, bBlack
                    door.Display "\n  >"
                    door.ChangeColor fWhite
                    sInput = sInput & vbCr
                    iCharsTyped = 0
                ElseIf Asc(data) = 19 Then  'save
                    If Trim(sInput) <> "" Then
                        sBody = sInput
                        If rs.State = adStateOpen Then rs.Close
                        rs.Open "INSERT INTO Posts (UserID,BaseID,Subject,Body,PostDate) VALUES (" & UserID & "," & BaseID & ",'" & sFix(sSubject) & "','" & sFix(sBody) & "','" & Now & "')", cn
                        DisplayMain
                    End If
                ElseIf Asc(data) = 3 Then  'cancel
                    door.ChangeColor fYellow
                    door.Display "\n\n::MESSAGE CANCELED::"
                    Menu = "WaitForAnyKeyMain"
                ElseIf data = Chr(8) Then
                    If sInput <> "" And Right(sInput, 1) <> vbCr Then
                        door.Display Chr(27) & "[D " & Chr(27) & "[D"
                        sInput = Left(sInput, Len(sInput) - 1)
                        iCharsTyped = iCharsTyped - 1
                    ElseIf sInput <> "" And Right(sInput, 1) = vbCr Then
                        door.Display Chr(27) & "[D " & Chr(27) & "[D"
                        door.Display Chr(27) & "[A"
                        sInput = Left(sInput, Len(sInput) - 1)
                        a = InStrRev(sInput, vbCr)
                        a = a + 1
                        door.ChangeColor fLightBlue, bBlack
                        door.Display ">"
                        door.ChangeColor fWhite, bBlack
                        door.Display Mid(sInput, a)
                        iCharsTyped = Len(Mid(sInput, a))
                    End If
                Else
                    If iCharsTyped = 74 Then
                        If data <> " " Then
                            lastspace = InStrRev(sInput, " ")
                            lastcr = InStrRev(sInput, vbCr)
                            If lastspace > lastcr Then
                                For i = 1 To Len(sInput) - lastspace
                                    door.Display Chr(27) & "[D " & Chr(27) & "[D"
                                Next
                                data = Mid(sInput, lastspace + 1) & data
                                sInput = Left(sInput, lastspace)
                                sInput = sInput & vbCr & data
                                iCharsTyped = Len(data)
                            Else
                                iCharsTyped = 0
                            End If
                            door.ChangeColor fLightBlue, bBlack
                            door.Display "\n  >"
                            door.ChangeColor fWhite
                            door.Display data
                        Else
                            sInput = sInput & vbCr
                            iCharsTyped = 0
                            door.ChangeColor fLightBlue, bBlack
                            door.Display "\n  >"
                            door.ChangeColor fWhite
                        End If
                    ElseIf iCharsTyped = 0 And data = "/" Then
                        door.ChangeColor fLightGreen, bBlack
                        door.Display Chr(27) & "[D " & Chr(27) & "[D" & "Command: "
                        door.ChangeColor fGreen, bBlack
                        door.Display "(? for help) "
                        Menu = "BodyCommand"
                    Else
                        sInput = sInput & data
                        door.Display data
                        iCharsTyped = iCharsTyped + Len(data)
                    End If
                End If
            End If
        Case "BodyCommand"
            Select Case UCase(data)
                Case "?"
                    door.ChangeColor fWhite, bBlack
                    door.Display UCase(data)
                    door.ChangeColor fYellow, bBlack
                    door.Display "\n\n  - Command Help -\n"
                    door.Display "\n  S        Send Message"
                    door.Display "\n  C        Cancel Message"
                    door.Display "\n  L        Line Editor"
                    door.Display "\n\n"
                    door.ChangeColor fLightGreen, bBlack
                    door.Display "  Command: "
                    door.ChangeColor fGreen, bBlack
                    door.Display "(? for help) "
                Case "C"
                    door.ChangeColor fWhite, bBlack
                    door.Display UCase(data)
                    door.ChangeColor fYellow
                    door.Display "\n\n::MESSAGE CANCELED::"
                    Menu = "WaitForAnyKeyMain"
                Case "L"
                    door.ChangeColor fWhite, bBlack
                    door.Display UCase(data)
                    door.ChangeColor fYellow
                    door.Display "\n  not functional yet\n\n"
                    door.ChangeColor fLightGreen, bBlack
                    door.Display "  Command: "
                    door.ChangeColor fGreen, bBlack
                    door.Display "(? for help) "
                Case "S"
                    door.ChangeColor fWhite, bBlack
                    door.Display UCase(data)
                    If Trim(sInput) <> "" Then
                        sBody = sInput
                        If rs.State = adStateOpen Then rs.Close
                        rs.Open "INSERT INTO Posts (UserID,BaseID,Subject,Body,PostDate) VALUES (" & UserID & "," & BaseID & ",'" & sFix(sSubject) & "','" & sFix(sBody) & "','" & Now & "')", cn
                        DisplayMain
                    End If
                Case Else
                    door.Display Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & "                      " & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D" & Chr(27) & "[D"
                    door.ChangeColor fLightBlue, bBlack
                    door.Display ">"
                    door.ChangeColor fWhite, bBlack
                    Menu = "NewPostBody"
            End Select
        Case "WaitForAnyKeyMain"
            DisplayMain
        Case "DisplayedMessageWait"
            Select Case UCase(data)
                Case "R"
                    door.ChangeColor fDarkGray, bBlack
                    door.ClearDisplay
                    door.DisplayANSI App.Path & "\TITLE.ANS"
                    door.MoveToPos 4, 1
                    door.ChangeColor fDarkGray, bBlack
                    door.Display "  Message Base: " & BaseName & "\n\n"
                    door.ChangeColor fWhite
                    door.Display "       Subject: "
                    door.ChangeColor fBlue, bWhite
                    door.Display Space(50)
                    door.MoveToPos 6, 15
                    sSubject = sLastSubject
                    If Left(sSubject, 4) <> "RE: " Then sSubject = "RE: " & sSubject
                    door.Display sSubject
                    sInput = sSubject
                    Menu = "AskNewPostSubject"
                Case "N", "]"
                    iOnMsg = iOnMsg + 1
                    If iOnMsg > iTotalMsgs Then iOnMsg = iTotalMsgs
                    
                    DisplayNextMsg
                Case "P", "["
                    iOnMsg = iOnMsg - 1
                    If iOnMsg < 1 Then iOnMsg = 1
                    door.Display "\n"
                    DisplayNextMsg
                Case "Q"
                    DisplayMain
            End Select
        Case "DisplayedMessageMore"
            Select Case UCase(data)
                Case "Y", vbCr, vbLf, vbCrLf
                    DisplayMsgMore
                Case "N"
                    door.MoveToPos 24, 2
                    door.ChangeColor fBlack, bWhite
                    door.Display "  ð   (R)eply   ð   (N)ext Message   ð   (P)revious Message   ð   (Q)uit   ð  "
                    door.MoveToPos 24, 80
                    Menu = "DisplayedMessageWait"
                Case "Q"
                    DisplayMain
            End Select
    End Select
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs2 = New ADODB.Recordset

Dim sDBPath As String
sDBPath = App.Path
'sDBPath = "C:\Personal VB\StormBase"

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDBPath & "\StormBase.mdb;Persist Security Info=False"
    Me.Show
    DoEvents
    door.OpenDropFile Command
    'Load user details
    door.ClearDisplay
    door.DisplayANSI App.Path & "\TITLE.ANS"
    door.MoveToPos 4, 1
GetUserID:
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM Users WHERE UserName = '" & door.Alias & "'", cn, adOpenStatic
    If rs.RecordCount = 0 Then
        rs.Close
        door.MoveToPos 4, 1
        door.ChangeColor fDarkGray
        door.Display "  WELCOME TO STORM-BASE\n\n"
        door.ChangeColor fWhite
        door.Display "  How would you like your name to be displayed? "
        door.ChangeColor fBlue, bWhite
        door.Display Space(25)
        door.MoveToPos 6, 49
        door.Display door.Alias
        sInput = door.Alias
        Menu = "AskDisplayName"
    Else
        Username = rs("Username")
        UserID = rs("UserID")
        DisplayName = rs("DisplayName")
        rs.Close
        DisplayMain
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    door.Quit
    End
End Sub

Private Sub DisplayMain()
    door.ChangeColor fWhite, bBlack
    door.ClearDisplay
    door.DisplayANSI App.Path & "\main.ans"
    door.MoveToPos 7, 39
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM Bases ORDER BY BaseOrder", cn, adOpenStatic
    If BaseName = "" Then
        If rs.RecordCount > 0 Then
            door.ChangeColor fLighCyan, bBlack
            BaseID = rs("BaseID")
            BaseName = rs("BaseName")
            door.Display BaseName
        Else
            door.ChangeColor fLightRed
            door.Display "No Message Bases Defined"
        End If
    Else
        For i = 1 To rs.RecordCount
            If BaseID = rs("BaseID") Then
                Exit For
            Else
                rs.MoveNext
            End If
        Next
        door.ChangeColor fLighCyan, bBlack
        door.Display BaseName
    End If
    door.ChangeColor fLighCyan
    door.MoveToPos 5, 17
    door.Display "[" & Pad_String(rs.AbsolutePosition, 2, " ", 0) & "/" & Pad_String(rs.RecordCount, 2, " ", 0) & "]"
    rs.Close
    rs.Open "SELECT Count(*) as RecCnt FROM Posts WHERE BaseID=" & BaseID, cn, adOpenStatic
    Dim totrecs As Long
    door.MoveToPos 5, 35
    totrecs = rs("RecCnt")
    door.Display Pad_String(CStr(totrecs), 5, " ", 0) & " POSTS"
    rs.Close
    rs.Open "SELECT Count(*) as RecCnt FROM PostRead WHERE UserID=" & UserID & " AND BaseID = " & BaseID, cn, adOpenStatic
    door.MoveToPos 5, 56
    door.Display Pad_String(CStr(totrecs - rs("RecCnt")), 5, " ", 0) & " NEW"
    rs.Close
    door.ChangeColor fWhite, bBlack
    door.MoveToPos 21, 50
    Menu = "MainMenu"
End Sub

Private Sub DisplayNextMsg()
    door.ChangeColor fWhite, bBlack
    door.ClearDisplay
    door.DisplayANSI App.Path & "\TITLE.ANS"
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM Posts WHERE BaseID = " & BaseID, cn, adOpenStatic
    iTotalMsgs = rs.RecordCount
    If iOnMsg > rs.RecordCount Then
        door.ChangeColor fWhite, bBlack
        door.Display "\n\nNo More Messages in Base\n\n"
        door.ChangeColor fBlack, bWhite
        door.Display "  -ð-  (R)eply  -ð-  (N)ext Message  -ð-  (P)revious Message  -ð-  (Q)uit  -ð-  "
    Else
        For i = 1 To iOnMsg - 1
            rs.MoveNext
        Next
        ReadMessage UserID, rs("PostID"), BaseID
        door.MoveToPos 2, 3
        door.ChangeColor fLighCyan, bBlack
        door.Display "[" & iOnMsg & "/" & rs.RecordCount & "]"
        door.ChangeColor fDarkGray, bBlack
        door.MoveToPos 3, 1
        door.Display "\n  Message Base: " & BaseName
        door.Display "\n          From: " & GetUserName(rs("UserID"))
        door.Display "\n       Subject: " & rs("Subject")
        door.Display "\n          Date: " & rs("PostDate")
        sLastSubject = rs("Subject")
        door.ChangeColor fGray, bBlack
        sBody = rs("Body")
        If GetNumberOfLines(sBody) > 15 Then
            ss = Split(sBody, vbCr)
            door.Display "\n\n  "
            i = 0
            For Each s In ss
                If i < 15 Then
                    s = s & vbCrLf & "  "
                    door.Display CStr(s)
                    i = i + 1
                End If
            Next
            iBodyLine = 15
            door.MoveToPos 24, 2
            door.ChangeColor fBlack, bWhite
            door.Display " ð MORE? ([Y],N,Q) ð "
            door.MoveToPos 24, 80
            Menu = "DisplayedMessageMore"
        Else
            door.Display "\n\n  " & Replace(sBody, vbCr, vbCrLf & "  ")
            door.MoveToPos 24, 2
            door.ChangeColor fBlack, bWhite
            door.Display "  ð   (R)eply   ð   (N)ext Message   ð   (P)revious Message   ð   (Q)uit   ð  "
            door.MoveToPos 24, 80
            Menu = "DisplayedMessageWait"
        End If
        
    End If
    rs.Close
End Sub

Private Function GetUserName(iID As Integer)
    Dim rsG As ADODB.Recordset
    Set rsG = New ADODB.Recordset
    rsG.Open "SELECT DisplayName FROM Users WHERE UserID=" & iID, cn, adOpenStatic
    If rsG.RecordCount = 0 Then
        GetUserName = "--UNKNOWN--"
    Else
        GetUserName = rsG("DisplayName")
    End If
    rsG.Close
    Set rsG = Nothing
End Function

Private Sub DisplayMsgMore()
    ss = Split(sBody, vbCr)
    g = GetNumberOfLines(sBody)
    
    door.ChangeColor fGray, bBlack
    For i = 9 To 23
        door.MoveToPos CInt(i), 1
        door.ClearToEndOfLine
    Next
    door.MoveToPos 9, 3
    door.ChangeColor fGray, bBlack
    
    i = 0
    If g < iBodyLine + 15 Then
        For Each s In ss
            If i >= iBodyLine And i < iBodyLine + 15 Then
                s = s & vbCrLf & "  "
                door.Display CStr(s)
            End If
            i = i + 1
        Next
        door.MoveToPos 24, 2
        door.ChangeColor fBlack, bWhite
        door.Display "  ð   (R)eply   ð   (N)ext Message   ð   (P)revious Message   ð   (Q)uit   ð  "
        door.MoveToPos 24, 80
        Menu = "DisplayedMessageWait"
    Else
        For Each s In ss
            If i >= iBodyLine And i < iBodyLine + 15 Then
                s = s & vbCrLf & "  "
                door.Display CStr(s)
            End If
            i = i + 1
        Next
        iBodyLine = iBodyLine + 15
        door.MoveToPos 24, 2
        door.ChangeColor fBlack, bWhite
        door.Display " ð MORE? ([Y],N,Q) ð "
        door.MoveToPos 24, 80
        Menu = "DisplayedMessageMore"
    End If
End Sub

