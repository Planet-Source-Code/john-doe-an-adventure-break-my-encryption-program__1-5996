Attribute VB_Name = "Functions"
Global Saved As Boolean 'Did the document has been saved?
Global OpenFilename As String   'Filename that open now

Function GetFileWithPath() As String
If MainFrm.Files.Filename = "" Then
    GetFileWithPath = ""
    Exit Function
Else
GetFileWithPath = GetPath(MainFrm.Files.Path) & MainFrm.Files.Filename
End If

End Function

Function GetPath(ByVal PathName As String) As String
If PathName Like "*\" Then
GetPath = PathName
Else
GetPath = PathName & "\"
End If
End Function


Function SpecielNumber1(ByVal Text As String) As Byte
Dim Value, Shift1, Shift2, ch

For i = 1 To Len(Text)
ch = Asc(Mid$(Text, i, 1))
Value = Value Xor Int(Shift1 * 10.4323)
Value = Value Xor Int(Shift2 * 4.23)

Shift1 = (Shift1 + 7) Mod 19
Shift2 = (Shift2 + 13) Mod 23
Next

SpecielNumber1 = Value
End Function

Function SpecielNumber2(ByVal Password As String) As Byte
Dim Value

Value = 194
For i = 1 To Len(Password)
ch = Asc(Mid$(Password, i, 1))
Value = Value Xor ch Xor i
If Value > 100 Then Value = (Value - 50) Xor 255
Next
SpecielNumber2 = Value
End Function

Function SpecielNumber3(ByVal Password As String) As Byte
Value = Len(Password) Mod 37

For i = 1 To Len(Password)
ch = Asc(Mid$(Password, i, 1))

If (Value Mod 2) And (ch > 10) Then ch = ch - 1

Value = (ch * Value * 17.3463) Mod 255

Next
SpecielNumber3 = Value
End Function

Function Fib(ByVal Num As Integer) As Long
Dim Temp As Integer, Temp2 As Integer, Temp3 As Integer

Temp = 1
Temp2 = 1
Temp3 = 1

For i = 3 To Num
Temp3 = Temp2
Temp2 = Temp
Temp = Temp + Temp3
Next

Fib = Temp
End Function

Function Pwd(ByVal Text As String, ByVal KeyTxt As String) As String

Dim KeyLen As Integer
Dim PassAsc As Byte
Dim SaveNum As Integer

Dim AfterETxt As String

Dim RandTxt1 As Integer, RandTxt2 As Integer, RandTxt3 As Integer
Dim Temp As Byte

RandTxt1 = SpecielNumber1(Text)
RandTxt2 = SpecielNumber2(KeyTxt)
RandTxt3 = SpecielNumber3(KeyTxt)


SaveNum = 1

KeyLen = Len(KeyTxt)

AfterETxt = ""

For i = 1 To Len(Text)
Temp = Asc(Mid(Text, i, 1))
PassAsc = Asc(Mid(KeyTxt, ((i - 1) Mod KeyLen) + 1, 1))

If RandTxt2 > RandTxt3 Then Temp = Temp Xor RandTxt1 Xor RandTxt3
If RandTxt1 > RandTxt3 Then Temp = Temp Xor RandTxt2


Temp = Temp Xor (Abs(RandTxt3 - i) Mod 256)

Temp = Temp Xor PassAsc
Temp = Temp Xor (Int(i * 2.423121) Mod 256)

Temp = Temp Xor (Int(Fib(i Mod 17) * 0.334534) Mod 256)

Temp = Temp Xor SaveNum
Temp = Temp Xor (KeyLen Mod SaveNum)
Temp = Temp Xor RandTxt3
Temp = Temp Xor (Len(Text) Mod 71)

Temp = Temp Xor Abs(RandTxt3 - RandTxt1)

Temp = Temp Xor Abs(((RandTxt1 Mod 23) * 10) Mod RandTxt2)

SaveNum = (Int(Fib(i Mod 7) * 0.334534) Mod 256)
SaveNum = SaveNum Xor (PassAsc * 45.92425) Mod 256

If (i >= 2) Then
    If PassAsc And 2 Then
    Temp = Temp Xor PassAsc
    Else
    Temp = Temp Xor (Int(PassAsc * 3.2145561) Mod 256)
    End If
Else
Temp = Temp Xor ((KeyLen * PassAsc + (i Mod 3)) Mod 256)
End If

AfterETxt = AfterETxt & Chr(Temp)
Next

Pwd = AfterETxt
End Function

Function GetTxtFile(ByVal Filename As String) As String
If Filename Like "*.txt" Then
GetTxtFile = Filename
Else
GetTxtFile = Filename & ".txt"
End If
End Function

Function ChangeEnable(ByVal Status As Boolean)
With MainFrm
.LoadBtn.Enabled = Status
.SaveBtn.Enabled = Status
.Mopen.Enabled = Status
.Msave.Enabled = Status
.Msaveas = Status
End With
End Function

Function SaveQuestion() As Byte
'if the user try to open/new/exit without saving his doc
'*******   1=YES, 2=NO, 3=CANCLE   ******************

Opt = MsgBox("You didnt save the last file." & vbCrLf & "Do you want to Save it?", vbQuestion Or vbYesNoCancel, "Save")
If Opt = vbYes Then

    If StartSave = True Then
        SaveQuestion = 1
    Else
        SaveQuestion = 3
    End If
ElseIf Opt = vbNo Then
    SaveQuestion = 2
Else
    SaveQuestion = 3
End If

End Function

Function StartSave() As Boolean
Dim Temp As String, Temp2 As String

StartSave = True
If OpenFilename = "" Then
    Temp = InputBox("Enter Filename", "Save file", MainFrm.Files.Filename)
    If Temp = "" Then StartSave = False: Exit Function 'only filename
    Temp = GetPath(MainFrm.Files.Path) & GetTxtFile(Temp) 'set temp to the full path
    If (Dir(Temp) <> "") Then 'if file exists
        If MsgBox("The file you entered is already exists." & vbCrLf & "Do you want to replace him?", vbQuestion Or vbYesNo, "File exists!") = vbNo Then StartSave = False: Exit Function
    End If

    Temp2 = VerifyPass
    If Temp2 = "" Then StartSave = False: Exit Function
    OpenFilename = Temp
    SaveFile OpenFilename, Temp2
    Saved = True

Else
    Temp = VerifyPass
    If Temp = "" Then StartSave = False: Exit Function

    SaveFile OpenFilename, Temp
    Saved = True
End If

End Function

Function SaveFile(ByVal Filename As String, ByVal Pass As String)
'free = FreeFile

Open Filename For Output As #1 'Binary save makes problems
Print #1, Pwd(MainFrm.Textbox, Pass)
Close #1
Saved = True
MainFrm.Files.Refresh
End Function

Function LoadFile(ByVal Filename As String, ByVal Pass As String)
Dim Dta As String
Dta = Space(FileLen(Filename))

free = FreeFile
Open Filename For Binary Access Read As #free
Get #free, , Dta
Close #free

Dta = Mid(Dta, 1, Len(Dta) - 2) 'because when u write the file, it ends with VbCrLf

'dta now contains the value of the file

MainFrm.Textbox = Pwd(Dta, Pass)
Saved = True
End Function

Function VerifyPass() As String
Dim Temp As String

Temp = InputBox("Reenter Password", "Save with password")
If Temp = "" Then Exit Function

If (Temp = MainFrm.PasswordTxt) Then
VerifyPass = Temp
Else
MsgBox "Password dont match!", vbCritical
VerifyPass = ""
End If

End Function

