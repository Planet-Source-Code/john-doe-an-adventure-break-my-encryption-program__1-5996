VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doc Saver 1.0 - Made by STeRoiD"
   ClientHeight    =   6930
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9780
   Icon            =   "Password6.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Control"
      Height          =   855
      Left            =   5160
      TabIndex        =   8
      Top             =   6000
      Width           =   4335
      Begin VB.CommandButton NewBtn 
         Caption         =   "New"
         Height          =   315
         Left            =   2880
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton SaveBtn 
         Caption         =   "Save"
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton LoadBtn 
         Caption         =   "Load"
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Choose0 
      Caption         =   "Choose 0"
      Height          =   255
      Left            =   8760
      TabIndex        =   6
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton DelFileBtn 
      Caption         =   "Del"
      Height          =   255
      Left            =   7920
      TabIndex        =   5
      Top             =   5640
      Width           =   855
   End
   Begin VB.DriveListBox Drives 
      Height          =   315
      Left            =   7920
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
   Begin VB.DirListBox FilePath 
      Height          =   1890
      Left            =   7920
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.FileListBox Files 
      Height          =   3210
      Left            =   7920
      Pattern         =   "*.txt"
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox PasswordTxt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   1
      Top             =   6120
      Width           =   3735
   End
   Begin VB.TextBox Textbox 
      Height          =   5895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   7695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   6240
      Width           =   945
   End
   Begin VB.Menu Mfile 
      Caption         =   "File"
      Begin VB.Menu Mnew 
         Caption         =   "New"
      End
      Begin VB.Menu Mopen 
         Caption         =   "Open"
      End
      Begin VB.Menu kav4 
         Caption         =   "-"
      End
      Begin VB.Menu Msave 
         Caption         =   "Save"
      End
      Begin VB.Menu Msaveas 
         Caption         =   "Save As..."
      End
      Begin VB.Menu k 
         Caption         =   "-"
      End
      Begin VB.Menu Mexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Mhelp 
      Caption         =   "Help"
      Begin VB.Menu Mhowto 
         Caption         =   "How to"
         Begin VB.Menu MLoadFiles 
            Caption         =   "Load Files"
         End
         Begin VB.Menu MSaveFiles 
            Caption         =   "Save Files"
         End
      End
      Begin VB.Menu kav84 
         Caption         =   "-"
      End
      Begin VB.Menu Mabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Choose0_Click()
Files.ListIndex = -1
End Sub

Private Sub DelFileBtn_Click()
If Files.Filename = "" Then Exit Sub

If MsgBox("Are you sure u want to DELETE this file?", vbExclamation Or vbYesNo, "Delete File") = vbYes Then
If MsgBox("All the information will lost! Are you sure u want to DELETE this file?", vbExclamation Or vbYesNo, "Delete File") = vbYes Then
Kill GetFileWithPath
End If
End If
Files.Refresh
End Sub

Private Sub Drives_Change()
FilePath.Path = Drives.Drive
Files.Path = Drives.Drive
End Sub

Private Sub FilePath_Change()
Files.Path = FilePath.Path
End Sub

Private Sub Files_DblClick()
MsgBox GetFileWithPath
End Sub

Private Sub Form_Load()
Drives.Drive = Left(App.Path, 2) 'set drive to the app drive
FilePath.Path = App.Path         'set path to the app path

OpenFilename = ""   'no filename
Saved = True        'no need to save
ChangeEnable False  'can`t save/open because text is empty
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Mexit_Click
End Sub

Private Sub LoadBtn_Click()
Mopen_Click
End Sub

Private Sub Mabout_Click()
AboutFrm.Show 1
End Sub

Private Sub Mexit_Click()
If Saved = False Then
    If SaveQuestion = 3 Then Exit Sub
End If
End
End Sub

Private Sub MLoadFiles_Click()
MsgBox "Navigate through the dir/path/file list from the right side to the file you want to load, write password in the password text and click load", vbInformation, "How to load files"
End Sub

Private Sub MSaveFiles_Click()
MsgBox "Navigate through the dir/path list from the right side to the path you want to save your file, click save button, write filename(without path!) and rewrite your password(for security)", vbInformation, "How to save files"
End Sub

Private Sub PasswordTxt_Change()
If Len(PasswordTxt) = 0 Then
ChangeEnable False  'can`t save/open
Else
ChangeEnable True   'can save/open
End If
End Sub

Private Sub SaveBtn_Click()
StartSave
End Sub

Private Sub Mnew_Click()
If Saved = False Then
    If SaveQuestion = 3 Then Exit Sub
End If

Textbox = ""
OpenFilename = ""
Saved = True
End Sub

Private Sub Mopen_Click()
If Saved = False Then
    If SaveQuestion = 3 Then Exit Sub
End If

If Files.Filename = "" Then MsgBox "Choose Filename", vbExclamation: Exit Sub

OpenFilename = GetFileWithPath
LoadFile OpenFilename, PasswordTxt

Saved = True
End Sub

Private Sub Msave_Click()
StartSave
End Sub

Private Sub Msaveas_Click()
Dim Temp As String, Temp2 As String
Temp = InputBox("Enter Filename", "Save file")
If Temp = "" Then Exit Sub
Temp = GetPath(Files.Path) & GetTxtFile(Temp)
If (Dir(Temp) <> "") Then
    If MsgBox("The file you entered is already exists." & vbCrLf & "Do you want to replace him?", vbQuestion Or vbYesNo, "File exists!") = vbNo Then Exit Sub
End If

Temp2 = VerifyPass
If Temp2 <> "" Then
OpenFilename = Temp
SaveFile OpenFilename, Temp2
End If
End Sub

Private Sub Textbox_Change()
Saved = False
End Sub
