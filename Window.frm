VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form3 
   Caption         =   "WebCam Madness"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   Icon            =   "Window.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3765
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4048
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
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form3.Caption = "WebCam Madness - " & Form1.dbName.Caption
WebBrowser1.Navigate Form1.WebBrowser1.LocationURL
End Sub

Private Sub Form_Resize()
WebBrowser1.Move 0, 0, Form3.ScaleWidth, Form3.ScaleHeight
End Sub
