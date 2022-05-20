VERSION 5.00
Begin VB.Form AyudaForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9570
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AyudaForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "AyudaForm.frx":058A
   ScaleHeight     =   3660
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lbl_Manual 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   4700
      MouseIcon       =   "AyudaForm.frx":D9E1
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2950
      Width           =   1275
   End
   Begin VB.Label lbl_Acerca 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   3660
      MouseIcon       =   "AyudaForm.frx":DCEB
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1180
      Width           =   1280
   End
End
Attribute VB_Name = "AyudaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbl_Acerca_Click()
Unload Me
AcercaForm.Show vbModal
End Sub

Private Sub lbl_Manual_Click()
Shell ("rundll32.exe url.dll,FileProtocolHandler " & App.Path & ("\Manual\ManualDeUsuario.pdf")), vbMaximizedFocus
Unload Me
End Sub
