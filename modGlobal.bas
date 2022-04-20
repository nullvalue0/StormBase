Attribute VB_Name = "modGlobal"
Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Public rs2 As ADODB.Recordset

Public BaseID As Integer
Public BaseName As String
Public Menu As String
Public DisplayName As String
Public Username As String
Public UserID As Integer
Public sInput As String
Public sSubject As String
Public sBody As String
Public iBodyLine As Integer
Public iOnMsg As Integer
Public iTotalMsgs As Integer
Public iCharsTyped As Integer
Public sLastSubject As String

Public Type sn
    Letter As String
    BaseName As String
    BaseID As Integer
End Type

Public SelectName(26) As sn

Function Pad_String(sPadString As String, iNumChar As Integer, sPadChar As String, sAlignment0L1R As Integer)
    'PAD A STRING WITH A SPECIFIED CHARACTER, SO THAT IT RETURNS AS A SPECIFIED LENGTH
    Dim iLen As Integer
    iLen = Len(sPadString)
    If iLen >= iNumChar Then
        Pad_String = Left(sPadString, iNumChar)
    Else
        Do Until Len(sPadString) >= iNumChar
            If sAlignment0L1R = 0 Then
                sPadString = sPadChar & sPadString
            Else
                sPadString = sPadString & sPadChar
            End If
        Loop
        Pad_String = sPadString
    End If
End Function

Function sFix(s As String)
    sFix = Replace(s, "'", "''")
End Function

Function GetNumberOfLines(s As String)
    Dim c, cc, i
    cc = Split(s, vbCr)
    For Each c In cc
        i = i + 1
    Next
    GetNumberOfLines = i
End Function

Function ReadMessage(uid As Integer, pid As Integer, bid As Integer)
    Dim rsrm As ADODB.Recordset
    Set rsrm = New ADODB.Recordset
    rsrm.Open "SELECT UserID FROM PostRead WHERE UserID = " & uid & " AND PostID = " & pid, cn, adOpenStatic
    If rsrm.RecordCount = 0 Then
        rsrm.Close
        rsrm.Open "INSERT INTO PostRead (UserID, PostID,BaseID) VALUES (" & uid & "," & pid & "," & bid & ")", cn
    Else
        rsrm.Close
    End If
    Set rsrm = Nothing
End Function
