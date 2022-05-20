VERSION 5.00
Begin VB.Form RestablecerPassForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restablecer contraseña 2/2"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RestablecerPassForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "RestablecerPassForm.frx":08CA
   ScaleHeight     =   4110
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Cancelar 
      Appearance      =   0  'Flat
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmd_Cambiar 
      Caption         =   "Cambiar "
      Height          =   615
      Left            =   4440
      TabIndex        =   2
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox txt_pass 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3240
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox txt_rpass 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3240
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label lbl_pass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label lbl_EJpass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTA: Debe ser mayor a 6 caracteres"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3240
      TabIndex        =   6
      Top             =   1680
      Width           =   2850
   End
   Begin VB.Label lbl_Rpass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Repita Contraseña:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   2400
      Width           =   2220
   End
   Begin VB.Label lbl_EJrpass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTA: Debe coincidir con la contraseña anterior"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3240
      TabIndex        =   4
      Top             =   2640
      Width           =   3630
   End
End
Attribute VB_Name = "RestablecerPassForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Cambiar_Click()
If txt_pass.Text = "" Then MsgBox ("Por favor escriba su nueva contraseña"), vbInformation, "Aviso": txt_pass.SetFocus: Exit Sub
          If Len(txt_pass.Text) < 6 Then MsgBox ("La contraseña debe ser mayor a 6 caracteres"), vbInformation, "Aviso": txt_pass.SetFocus: Exit Sub
If txt_rpass.Text = "" Then MsgBox ("Por favor, repita la contraseña"), vbInformation, "Aviso": txt_rpass.SetFocus: Exit Sub
          If Trim(txt_pass.Text) <> Trim(txt_rpass.Text) Then MsgBox ("Las contraseñas deben de coincidir"), vbInformation, "Aviso": txt_rpass.Text = "": txt_rpass.SetFocus: Exit Sub
          
With RsUsuarios
          .Requery
          .Find "Nomb_Usuario='" & Trim(ConfirmarDatosForm.txt_username.Text) & "'"
                    !Contraseña = Trim(txt_pass.Text)
          .UpdateBatch
          .Requery
          MsgBox ("Contraseña cambiada exitosamente"), vbInformation, "Aviso"
End With
Unload Me
Unload ConfirmarDatosForm
LoginForm.Show
End Sub

Private Sub cmd_Cancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Usuarios
End Sub
