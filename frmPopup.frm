VERSION 5.00
Begin VB.Form frmPopup 
   BackColor       =   &H8000000C&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "www.fysoft.com"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   720
      Top             =   360
   End
   Begin VB.Image imgLogo 
      Height          =   900
      Left            =   0
      MouseIcon       =   "frmPopup.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmPopup.frx":0152
      Top             =   0
      Width           =   7020
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOp As String, ByVal lpFile As String, ByVal lpParam As String, ByVal lpDir As String, ByVal nShowCmd As Long) As Long

Dim IntAjout As Integer

Private Sub Form_Load()
    IntAjout = 0
    
End Sub

Private Sub imgLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveURL Me, "http://www.fysoft.com"
    Unload Me
End Sub

Private Sub Timer1_Timer()
    IntAjout = IntAjout + 1
    If IntAjout = 5 Then IntAjout = 1
    imgLogo.Picture = LoadResPicture(100 + IntAjout, vbResBitmap)
End Sub

Private Sub ActiveURL(frm As Form, strURL As String)
    Dim RetV As Long
    Dim fURL As String
    Const SHOWNORMAL = 1

    RetV = ShellExecute(frm.hwnd, vbNullString, strURL, vbNullString, "c:\", SHOWNORMAL)
    If RetV < 32 Then MsgBox "Impossible d'ouvrir le site Internet " + strURL + "." + vbLf + "Une erreur inattendue s'est produite. Vérifiez la connexion Internet.", vbCritical, "Popup"
               
End Sub
