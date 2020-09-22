VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSearch 
   Caption         =   "Find"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   3810
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ListView LstResults 
      Height          =   1590
      Left            =   0
      TabIndex        =   7
      Top             =   2205
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   2805
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   5080
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "In Folder"
         Object.Width           =   5080
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Modified"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4635
      TabIndex        =   6
      Top             =   1395
      Width           =   1185
   End
   Begin VB.CommandButton cmdStop 
      Cancel          =   -1  'True
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4635
      TabIndex        =   5
      Top             =   900
      Width           =   1185
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find Now"
      Default         =   -1  'True
      Height          =   375
      Left            =   4635
      TabIndex        =   4
      Top             =   405
      Width           =   1185
   End
   Begin VB.TextBox txtLookIn 
      Height          =   285
      Left            =   1485
      TabIndex        =   3
      Top             =   900
      Width           =   2760
   End
   Begin VB.TextBox txtNamed 
      Height          =   285
      Left            =   1485
      TabIndex        =   1
      Top             =   405
      Width           =   2760
   End
   Begin VB.Label lblLookIn 
      Caption         =   "&Look in:"
      Height          =   240
      Left            =   270
      TabIndex        =   2
      Top             =   900
      Width           =   1185
   End
   Begin VB.Label lblNamed 
      Caption         =   "&Named:"
      Height          =   240
      Left            =   270
      TabIndex        =   0
      Top             =   405
      Width           =   1185
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    ' Terminate program
    Unload Me
End Sub

Private Sub cmdFind_Click()
    ' Disable Find button
    cmdFind.Enabled = False
    ' Enable Stop button
    cmdStop.Enabled = True
    
    ' Clear list
    LstResults.ListItems.Clear
    
    ' Find file
    StatusBar.Panels(1).Text = _
        FindFile(txtNamed.Text, txtLookIn.Text, Me, LstResults, , StatusBar)
    
    ' Enable Find button
    cmdFind.Enabled = True
    ' Disable Stop button
    cmdStop.Enabled = False
End Sub

Private Sub cmdStop_Click()
    ' Enable Find button
    cmdFind.Enabled = True
    ' Disable Stop button
    cmdStop.Enabled = False
    
    ' Stop
    FindFile txtNamed.Text, txtLookIn.Text, Me, LstResults, , StatusBar
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    ' Relocate buttons
    cmdFind.Left = Width - cmdFind.Width * 1.5
    cmdStop.Left = Width - cmdStop.Width * 1.5
    cmdExit.Left = Width - cmdExit.Width * 1.5
    
    ' Resize text boxes
    txtNamed.Width = cmdFind.Left - cmdFind.Width * 0.5 - txtNamed.Left
    txtLookIn.Width = cmdFind.Left - cmdFind.Width * 0.5 - txtLookIn.Left

    ' Resize controls
    LstResults.Width = Width - 200
    LstResults.Height = Height - StatusBar.Height - LstResults.Top - 300
    
    ' Resize panels
    StatusBar.Panels(1).Width = Width
End Sub
