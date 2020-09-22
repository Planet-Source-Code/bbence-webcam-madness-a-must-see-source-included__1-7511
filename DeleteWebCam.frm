VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete WebCam"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "DeleteWebCam.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      DataField       =   "Name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      DataField       =   "MainCategory"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      DataField       =   "SubCategory"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000F&
      DataField       =   "Description"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H8000000F&
      DataField       =   "URL"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\webcams.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "WebCams"
      Top             =   2040
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   -600
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MainCategory:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   -600
      TabIndex        =   10
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SubCategory:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   -600
      TabIndex        =   9
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   -600
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "URL:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   -600
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   4800
      X2              =   4800
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   4815
      X2              =   4815
      Y1              =   120
      Y2              =   2400
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
retval = MsgBox("Do you really want to delete the entry?", vbQuestion + vbYesNo, "Delete WebCam")
If retval = vbYes Then
Data1.Recordset.Delete
Data1.Recordset.MoveLast
End If
End Sub

Private Sub Command3_Click()
MsgBox "Deleted entries will appear in the menu(but won't function anymore, though) as long as you restart the program.", vbInformation, "Info"
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Data1.DatabaseName = App.Path & "\webcams.mdb"
Data1.RecordSource = "WebCams"
Data1.Refresh
Data1.Recordset.MoveFirst
End Sub
