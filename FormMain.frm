VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PSC zipfile renamer"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rename all files OR the selected ones"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   4200
      Width           =   6615
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin MSComctlLib.ListView ZipList 
      Height          =   3735
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "PSC zip-files"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PSC_Text As String = "PSC_ReadMe_"
Private CurrentFile As String
Private IsBusy As Boolean

Private Sub Command1_Click()
    Call RenameAll
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Dir1_Change()
    If IsBusy Then IsBusy = False: DoEvents
    Call GetFiles
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    With ZipList
        .ListItems.Clear
        .ColumnHeaders.Add , , "Now PSC name", .Width / 3
        .ColumnHeaders.Add , , "Rename to", .Width / 3 * 2
    End With
    Call GetFiles
End Sub

Private Sub GetFiles()
    Dim ArchName As String
    Dim FileName As String
    Dim CountFiles As Long
    On Error Resume Next
    ZipList.ListItems.Clear
    ArchName = Dir1.Path & "\*.zip"
    CountFiles = 0
    IsBusy = True
    ArchName = Dir(Dir1.Path & "\*.zip", vbNormal)
'search for ZIP files
    Do While ArchName <> ""
        If IsBusy = False Then Exit Sub
'Find out if it is as PSC- zip file and if so return the new name
        FileName = GetNewFileName(Dir1.Path & "\" & ArchName, PSC_Text)
        If FileName <> "" Then
            CountFiles = CountFiles + 1
            With ZipList
                .ListItems.Add CountFiles, , ArchName
                .ListItems(CountFiles).SubItems(1) = FileName
                .ListItems(CountFiles).Selected = False
            End With
        End If
        DoEvents
        ArchName = Dir
    Loop
    IsBusy = False
End Sub


'rename all (selected) files
Private Function RenameAll()
    Dim TheDir As String
    Dim SrcName As String
    Dim TargName As String
    Dim Sel As Boolean
    Dim I As Integer
    TheDir = Dir1.Path
    If Right(TheDir, 1) <> "\" And Right(TheDir, 1) <> "/" Then
        TheDir = TheDir & "\"
    End If
'find out if a selection is made
    For I = 1 To ZipList.ListItems.Count
        If ZipList.ListItems(I).Selected = True Then
            Sel = True
            Exit For
        End If
    Next
'find the files to rename
    For I = 1 To ZipList.ListItems.Count
        If ZipList.ListItems(I).Selected = Sel Then
            SrcName = TheDir & ZipList.ListItems(I)
            TargName = TheDir & ZipList.ListItems(I).SubItems(1)
            If SrcName <> TargName Then
                On Error Resume Next
                Name SrcName As TargName
            End If
        End If
    Next
    Call GetFiles
End Function
