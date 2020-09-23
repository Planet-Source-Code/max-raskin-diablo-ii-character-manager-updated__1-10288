Attribute VB_Name = "modMain"
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function EnumFilesByExt(Path As String, ListBox As ListBox, Extension As String)
ListBox.Clear
    Dim XDir() As String
    Dim TmpDir As String
    If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If


    DoEvents
        TmpDir = Dir(Path, vbDirectory + vbHidden + vbSystem + vbArchive + vbReadOnly)


        Do While TmpDir <> ""


            If TmpDir <> "." And TmpDir <> ".." Then


                If (GetAttr(Path & TmpDir)) <> vbDirectory Then
                    If Right(TmpDir, Len(Extension)) = Extension Then ListBox.AddItem TmpDir
                    ReDim Preserve XDir(DirCount) As String
                End If
            End If
            TmpDir = Dir
            
        Loop
End Function

Public Sub MoveForm(Form As Form)
    ReleaseCapture
    Call SendMessage(Form.hwnd, &HA1, 2, 0&)
End Sub

Public Sub SaveFile(Filename As String, Text As String)
On Error GoTo ErrorHandler
    FileNumber = FreeFile
    Open Filename For Binary As FileNumber Len = Len(Text)
    Put #FileNumber, , Text
    Close FileNumber
ErrorHandler:
        If Err.Number <> 0 Then
            Exit Sub
        End If
End Sub

Function OpenFile(Filename As String, TextBox As TextBox)
On Error Resume Next
    Open Filename For Binary As #1
    TextBox.Text = Input(LOF(1), #1)
    Close #1
End Function

'Gets character level
Function GetLevel(SaveFile As String) As String
    Dim vRetVal, nLVL As Integer, lPos As Long
    lPos = 37 'The position where the value stands
    Open SaveFile For Binary As #1  'open a save file as binary
    Get #1, lPos, nLVL 'Now, get the value
    Close #1   'Close the file
    vRetVal = Hex(nLVL)
    vRetVal = "&H" & CStr(vRetVal) 'convert it to a vb hex value because the clng function does not know the diffrence between a number and a hex without the &H
    If vRetVal = 0 Then
        GetLevel = "1" 'Get the level, and we're all done ! :-)
    Else
        GetLevel = CStr(CLng(vRetVal)) 'Get the level, and we're all done ! :-)
    End If
End Function

'Gets character's status (title)
Function GetStatus(SaveFile As String) As String
    Dim nStatus As Integer, str As String
    Open SaveFile For Binary As #1  'open a save file a binary
    Get #1, 26, nStatus 'now, get the value
    Close #1   'close the file
    str = GetClass(SaveFile)
    If str = "Barbarian" Then GoTo SetMan
    If str = "Necromancer" Then GoTo SetMan
    If str = "Paladin" Then GoTo SetMan
    If str = "Amazon" Then GoTo SetWomen
    If str = "Sorceress" Then GoTo SetWomen
SetMan:         If Hex(nStatus) = 7 Then GetStatus = "Sir"
                If Hex(nStatus) = 5 Then GetStatus = "Sir"
                If Hex(nStatus) = 9 Then GetStatus = "Lord"
                If CStr(Hex(nStatus)) = "C" Then GetStatus = "Baron"
                Exit Function
SetWomen:   If Hex(nStatus) = 7 Then GetStatus = "Dame"
            If Hex(nStatus) = 5 Then GetStatus = "Dame"
            If Hex(nStatus) = 9 Then GetStatus = "Lady"
            If CStr(Hex(nStatus)) = "C" Then GetStatus = "Baroness"
    If Hex(nStatus) = 0 Then GetStatus = "" 'None (Not killed Diablo yet)
End Function

'Gets the character class out of a save file
Function GetClass(SaveFile As String) As String
    Dim vRetVal As Integer, nClass As Integer
    Open SaveFile For Binary As #1  'open a save file as binary
    Get #1, 35, nClass 'now, get the value
    Close #1   'close the file
    Select Case nClass 'Returned cases:
    Case 0
        GetClass = "Amazon"
    Case 1
        GetClass = "Sorceress"
    Case 2
        GetClass = "Necromancer"
    Case 3
        GetClass = "Paladin"
    Case 4
        GetClass = "Barbarian"
    End Select
End Function

