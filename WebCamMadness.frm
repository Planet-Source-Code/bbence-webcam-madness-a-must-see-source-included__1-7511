VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{A22D979F-2684-11D2-8E21-10B404C10000}#1.4#0"; "CPOPMENU.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WebCam Madness"
   ClientHeight    =   5895
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5175
   Icon            =   "WebCamMadness.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   3840
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   5640
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "18:46"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2000/04/19."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   4875
      TabIndex        =   5
      Top             =   240
      Width           =   4935
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   4150
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   4870
         ExtentX         =   8590
         ExtentY         =   7320
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   4935
      Begin VB.Label Label2 
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
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\webcams.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "WebCams"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin cPopMenu.PopMenu ctlPopMenu 
      Left            =   4680
      Top             =   4200
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
   End
   Begin MSComctlLib.ImageList ilsIcons 
      Left            =   4440
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":08CA
            Key             =   "SETTINGS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":09E6
            Key             =   "WINDOW"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":0E3A
            Key             =   "FAVO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":0F96
            Key             =   "WALLPAPER"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":1872
            Key             =   "SAVEAS"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":19CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":1B36
            Key             =   "REFRESH"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":1C92
            Key             =   "CONNECT"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":1DEE
            Key             =   "DELETE"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":1F0A
            Key             =   "SAVE"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":2026
            Key             =   "CAMERA"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":2142
            Key             =   "NEW"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":225E
            Key             =   "BYCAT"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":23BE
            Key             =   "BYPLACE"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":26DA
            Key             =   "FIND"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebCamMadness.frx":27EE
            Key             =   "HELP"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.Label dbSubCategory 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      DataField       =   "SubCategory"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label dbName 
      BackStyle       =   0  'Transparent
      DataField       =   "Name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu WebCams 
      Caption         =   "WebCams"
      Begin VB.Menu AddNew 
         Caption         =   "Add New"
      End
      Begin VB.Menu Delete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu Picture 
      Caption         =   "Picture"
      Begin VB.Menu ShowInWindow 
         Caption         =   "Show In Window"
      End
      Begin VB.Menu line6 
         Caption         =   "-"
      End
      Begin VB.Menu AddToFavourites 
         Caption         =   "Add To Favourites"
      End
      Begin VB.Menu line5 
         Caption         =   "-"
      End
      Begin VB.Menu SaveAs 
         Caption         =   "Save Picture As..."
      End
      Begin VB.Menu SetAsWallpaper 
         Caption         =   "Set as Wallpaper"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu Refresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu Select 
      Caption         =   "Select"
      Begin VB.Menu ByPlace 
         Caption         =   "By Place"
      End
      Begin VB.Menu ByCategory 
         Caption         =   "By Category"
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu Favourites 
         Caption         =   "Favourites"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu Hlp 
         Caption         =   "Help"
      End
      Begin VB.Menu line8 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SETDESKWALLPAPER = 20
Private Const SPIF_SENDWININICHANGE = &H2
Private Const SPIF_UPDATEINIFILE = &H1

Private Sub pSetIcon(ByVal sIconKey As String, ByVal sMenuKey As String)
Dim lIconIndex As Long
    lIconIndex = plGetIconIndex(sIconKey)
    ctlPopMenu.ItemIcon(sMenuKey) = lIconIndex
End Sub
Private Function plGetIconIndex(ByVal sKey As String) As Long
    plGetIconIndex = ilsIcons.ListImages.Item(sKey).Index - 1
End Function
Private Sub SetWallpaper(ByVal Filename As String)
Dim X As Long
X = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, Filename, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE)
End Sub

Private Sub ctlPopMenu_Click(ItemNumber As Long)
Select Case ctlPopMenu.MenuKey(ItemNumber)
Case "Exit"
    retval = MsgBox("Do you really want to exit?", vbQuestion + vbYesNo, "WebCam Madness")
    If retval = vbYes Then End
Case "AddNew"
    Form4.Show vbModal
Case "Delete"
    Form5.Show vbModal
Case "ByPlace", "ByCategory", "Favourites"
Case "About"
    Form2.Show vbModal
Case "Refresh"
    If Not WebBrowser1.LocationURL = "" Then WebBrowser1.Refresh
    StatusBar1.Panels(1).Text = "Loading..."
Case "SaveAs"
    If Not WebBrowser1.LocationURL = "" Then WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
Case "SetAsWallpaper"
MsgBox "Maybe in Next version!", , "Sorry"
Case "AddToFavourites"
    On Error Resume Next
    ctlPopMenu.AddItem dbName.Caption, "fav" & dbName.Caption, , , ctlPopMenu.MenuIndex("Favourites")
    List1.AddItem "fav" & dbName.Caption
Case "ShowInWindow"
    Form3.Show
Case "Hlp"
    Form6.Show vbModal
Case Else
    If Not ctlPopMenu.ItemIcon(ctlPopMenu.MenuKey(ItemNumber)) = 0 Then
    Data1.RecordSource = "SELECT * FROM [WebCams] WHERE [Name] = '" & ctlPopMenu.Caption(ctlPopMenu.MenuKey(ItemNumber)) & "'"
    Data1.Refresh
    WebBrowser1.Navigate Data1.Recordset(4).Value
    Form1.Caption = "WebCam Madness - " & dbName.Caption
    StatusBar1.Panels(1).Text = "Loading..."
    End If
End Select
End Sub

Private Sub Form_Load()
    With ctlPopMenu
        .ImageList = ilsIcons
        .SubClassMenu Me
    End With
pSetIcon "NEW", "AddNew"
pSetIcon "DELETE", "Delete"
pSetIcon "BYPLACE", "ByPlace"
pSetIcon "BYCAT", "ByCategory"
pSetIcon "REFRESH", "Refresh"
pSetIcon "SAVEAS", "SaveAs"
pSetIcon "WALLPAPER", "SetAsWallpaper"
pSetIcon "FAVO", "Favourites"
pSetIcon "FAVO", "AddToFavourites"
pSetIcon "WINDOW", "ShowInWindow"
pSetIcon "HELP", "Hlp"

Data1.DatabaseName = App.Path & "\webcams.mdb"
Data1.RecordSource = "WebCams"
Data1.Refresh

If Data1.Recordset.EOF Then
MsgBox "No WebCams in the database!", vbInformation, "Error"
Exit Sub
End If

Data1.Recordset.MoveFirst

Data1.RecordSource = "SELECT * FROM [WebCams] WHERE [MainCategory] = 'BP'"
Data1.Refresh
Do While Not Data1.Recordset.EOF
On Error Resume Next
ctlPopMenu.AddItem Data1.Recordset(2).Value, Data1.Recordset(2).Value, , , ctlPopMenu.MenuIndex("ByPlace")
Data1.Recordset.MoveNext
Loop

Data1.RecordSource = "WebCams"
Data1.Refresh
Data1.Recordset.MoveFirst

Data1.RecordSource = "SELECT * FROM [WebCams] WHERE [MainCategory] = 'BP'"
Data1.Refresh
Do While Not Data1.Recordset.EOF
ctlPopMenu.AddItem Data1.Recordset(0).Value, Data1.Recordset(0).Value, , , ctlPopMenu.MenuIndex(Data1.Recordset(2).Value)
pSetIcon "CAMERA", Data1.Recordset(0).Value
Data1.Recordset.MoveNext
Loop
'------------------------------------------------
Data1.RecordSource = "WebCams"
Data1.Refresh
Data1.Recordset.MoveFirst

Data1.RecordSource = "SELECT * FROM [WebCams] WHERE [MainCategory] = 'BC'"
Data1.Refresh
Do While Not Data1.Recordset.EOF
On Error Resume Next
ctlPopMenu.AddItem Data1.Recordset(2).Value, Data1.Recordset(2).Value, , , ctlPopMenu.MenuIndex("ByCategory")
Data1.Recordset.MoveNext
Loop

Data1.RecordSource = "WebCams"
Data1.Refresh
Data1.Recordset.MoveFirst

Data1.RecordSource = "SELECT * FROM [WebCams] WHERE [MainCategory] = 'BC'"
Data1.Refresh
Do While Not Data1.Recordset.EOF
ctlPopMenu.AddItem Data1.Recordset(0).Value, Data1.Recordset(0).Value, , , ctlPopMenu.MenuIndex(Data1.Recordset(2).Value)
pSetIcon "CAMERA", Data1.Recordset(0).Value
Data1.Recordset.MoveNext
Loop

Open App.Path & "\favourites.txt" For Input As #2

Do Until EOF(2)
Line Input #2, favvar
If favvar = "" Then Exit Sub
List1.AddItem favvar
ctlPopMenu.AddItem Mid(favvar, 4), favvar, , , ctlPopMenu.MenuIndex("Favourites")
pSetIcon "CAMERA", favvar
Loop
Close #2

End Sub



Private Sub Form_Unload(Cancel As Integer)
Open App.Path & "\favourites.txt" For Output As #1
For i = 0 To List1.ListCount - 1
Print #1, List1.List(i)
Next i
Close #1
End Sub

Private Sub Label1_Click()
ctlPopMenu.ShowPopupMenu Label1, "File", 0, Label1.Height

End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    StatusBar1.Panels(1).Text = "Done."
End Sub


