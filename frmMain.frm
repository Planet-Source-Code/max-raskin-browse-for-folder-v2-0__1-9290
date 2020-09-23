VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Browse For Folder Example"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4530
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Display Browse For Folder opened with the start menu folder"
      Height          =   765
      Left            =   1110
      TabIndex        =   3
      Top             =   2190
      Width           =   2325
   End
   Begin VB.CheckBox chkFiles 
      Caption         =   "Browse With Files"
      Height          =   345
      Left            =   1440
      TabIndex        =   2
      Top             =   1710
      Width           =   1605
   End
   Begin VB.CommandButton cmdBrowseForFolder 
      Caption         =   "Display Browse For Folder"
      Height          =   705
      Left            =   1110
      TabIndex        =   0
      Top             =   240
      Width           =   2235
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   585
      Left            =   270
      TabIndex        =   1
      Top             =   1050
      Width           =   4005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BrowseForFolder *Updated* by Max Raskin, April 08 2000
'
'New feature - Special Folders on start up, with out crashes,
'originally if you put a PIDL without verfying its exists then
'the application crashes
'
'The module modCheck will take care of verifying the path, it also can be used for
'other needs when using special folders
'
'function from modCheck Syntax:
'
'CheckFolderID(Folders Enum Will Show Up Here)
'
'Return value is a string that contains the path of the folder

Private Sub cmdBrowseForFolder_Click()
    Dim ReturnValue As String 'Keeps up the return
    Dim WithFiles As Long 'Just for this project, to add browsing with files or not
   
    ReturnValue = BrowseForFolder(Me.hwnd, "Choose a folder:", WithFiles, RecycleBin)
    If ReturnValue <> "" Then
      lblInfo.Caption = "Path Selected: " & ReturnValue
    Else
      lblInfo.Caption = "Cancel selected or the folder/file type selected isn't from the file system"
    End If
End Sub

Private Sub Command1_Click()
    Dim ReturnValue As String 'Keeps up the return
    Dim WithFiles As Long 'Just for this project, to add browsing with files or not
    If chkFiles.Value = 1 Then WithFiles = BrowseForFolderFlags.BrowseIncludeFiles
    ReturnValue = BrowseForFolder(Me.hwnd, "Choose a folder:", BIF_DONTGOBELOWDOMAIN Or WithFiles, StartMenu)
    If ReturnValue <> "" Then
      lblInfo.Caption = "Path Selected: " & ReturnValue
    Else
      lblInfo.Caption = "Cancel selected or the folder/file type selected isn't from the file system"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
