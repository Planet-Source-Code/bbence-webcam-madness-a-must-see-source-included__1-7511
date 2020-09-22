VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add WebCam"
   ClientHeight    =   3075
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "AddWebCam.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\webcams.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "WebCams"
      Top             =   2160
      Visible         =   0   'False
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
      TabIndex        =   12
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
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
      TabIndex        =   11
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
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
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text5 
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
      TabIndex        =   9
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox Text4 
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
      TabIndex        =   8
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox Text3 
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
      TabIndex        =   7
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox Text2 
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
      TabIndex        =   6
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"AddWebCam.frx":000C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   4815
      X2              =   4815
      Y1              =   120
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   4800
      X2              =   4800
      Y1              =   120
      Y2              =   3000
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
      TabIndex        =   5
      Top             =   1560
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
      TabIndex        =   4
      Top             =   1200
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
      TabIndex        =   3
      Top             =   840
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
      TabIndex        =   2
      Top             =   480
      Width           =   1815
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
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Data1.Recordset.AddNew
Command1.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
Data1.Recordset.Update
Data1.Refresh
Data1.Recordset.MoveLast
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
MsgBox "Any entries added to the database will only appear in the menu after restarting the program!", vbInformation, "Info"
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Data1.DatabaseName = App.Path & "\webcams.mdb"
Data1.RecordSource = "WebCams"
Data1.Refresh
Command2.Enabled = False
Data1.Recordset.MoveFirst
End Sub
