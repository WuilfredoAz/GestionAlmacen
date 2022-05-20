VERSION 5.00
Begin VB.Form RegisterForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro - Shekina 4004"
   ClientHeight    =   10575
   ClientLeft      =   6375
   ClientTop       =   1155
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RegisterForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "RegisterForm.frx":08CA
   ScaleHeight     =   10575
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbo_Tarea 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3600
      TabIndex        =   26
      Text            =   "cbo_Tarea"
      Top             =   8760
      Width           =   4455
   End
   Begin VB.CommandButton cmd_atras 
      Appearance      =   0  'Flat
      Caption         =   "Atras"
      Height          =   615
      Left            =   1800
      TabIndex        =   9
      Top             =   9720
      Width           =   2295
   End
   Begin VB.CommandButton cmd_guardar 
      Caption         =   "Registrarse"
      Height          =   615
      Left            =   5760
      TabIndex        =   8
      Top             =   9720
      Width           =   2295
   End
   Begin VB.TextBox txt_correo 
      Height          =   495
      Left            =   3600
      MaxLength       =   30
      TabIndex        =   7
      Top             =   7680
      Width           =   4455
   End
   Begin VB.TextBox txt_rpass 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3600
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   6840
      Width           =   4455
   End
   Begin VB.TextBox txt_pass 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3600
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   5880
      Width           =   4455
   End
   Begin VB.TextBox txt_username 
      Height          =   495
      Left            =   3600
      MaxLength       =   15
      TabIndex        =   4
      Top             =   5040
      Width           =   4455
   End
   Begin VB.TextBox txt_cedula 
      Height          =   495
      Left            =   3600
      MaxLength       =   8
      TabIndex        =   3
      Top             =   4200
      Width           =   4455
   End
   Begin VB.TextBox txt_apellido 
      Height          =   495
      Left            =   3600
      MaxLength       =   30
      TabIndex        =   2
      Top             =   3360
      Width           =   4455
   End
   Begin VB.TextBox txt_nombre 
      Height          =   495
      Left            =   3600
      MaxLength       =   30
      TabIndex        =   1
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label lbl_descripcion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para ingresar al sistema, por favor rellene los siguientes campos:"
      Height          =   270
      Left            =   720
      TabIndex        =   25
      Top             =   1800
      Width           =   6780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione una opción (Rol dentro del almacén)"
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
      Left            =   3600
      TabIndex        =   24
      Top             =   9120
      Width           =   3555
   End
   Begin VB.Label lbl_tarea 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tarea:"
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
      Left            =   2280
      TabIndex        =   23
      Top             =   8880
      Width           =   735
   End
   Begin VB.Label lbl_EJcorreo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ejm: ejemplo@correo.com"
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
      Left            =   3600
      TabIndex        =   22
      Top             =   8160
      Width           =   1965
   End
   Begin VB.Label lbl_correo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Correo:"
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
      Left            =   2160
      TabIndex        =   21
      Top             =   7920
      Width           =   885
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
      Left            =   3600
      TabIndex        =   20
      Top             =   7320
      Width           =   3630
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
      Left            =   840
      TabIndex        =   19
      Top             =   7080
      Width           =   2220
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
      Left            =   3600
      TabIndex        =   18
      Top             =   6360
      Width           =   2850
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
      Left            =   1560
      TabIndex        =   17
      Top             =   6120
      Width           =   1425
   End
   Begin VB.Label lbl_EJusername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ejm:PedroP"
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
      Left            =   3600
      TabIndex        =   16
      Top             =   5520
      Width           =   870
   End
   Begin VB.Label lbl_username 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de Usuario:"
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
      Left            =   720
      TabIndex        =   15
      Top             =   5280
      Width           =   2310
   End
   Begin VB.Label lbl_EJcedula 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ejm:12080090"
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
      Left            =   3600
      TabIndex        =   14
      Top             =   4680
      Width           =   1050
   End
   Begin VB.Label lbl_cedula 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.I:"
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
      Left            =   2640
      TabIndex        =   13
      Top             =   4440
      Width           =   390
   End
   Begin VB.Label lbl_EJapellido 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ejm: Pérez"
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
      Left            =   3600
      TabIndex        =   12
      Top             =   3840
      Width           =   810
   End
   Begin VB.Label lbl_apellido 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido:"
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
      Left            =   2040
      TabIndex        =   11
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Label lbl_EJnombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ejm: Pedro"
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
      Left            =   3600
      TabIndex        =   10
      Top             =   3000
      Width           =   810
   End
   Begin VB.Label lbl_nombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
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
      Left            =   2040
      TabIndex        =   0
      Top             =   2760
      Width           =   1005
   End
End
Attribute VB_Name = "RegisterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Atras_Click()
txt_nombre.Text = ""
txt_apellido.Text = ""
txt_cedula.Text = ""
txt_username.Text = ""
txt_pass.Text = ""
txt_rpass.Text = ""
txt_correo.Text = ""
'txt_tarea.Text = ""
cbo_Tarea.ListIndex = 0
'opt_user = False
'opt_admi = False

Unload Me
LoginForm.Show

End Sub

Private Sub cmd_guardar_Click()
If txt_nombre.Text = "" Then MsgBox "Por favor, ingrese su nombre", vbInformation, "Aviso": txt_nombre.SetFocus: Exit Sub
          If IsNumeric(txt_nombre.Text) = True Then MsgBox ("Por favor, no escriba números en el campo nombre"), vbInformation, "Aviso": txt_nombre.SetFocus: Exit Sub
If txt_apellido.Text = "" Then MsgBox "Por favor, ingrese su apellido", vbInformation, "Aviso": txt_apellido.SetFocus: Exit Sub
          If IsNumeric(txt_apellido.Text) = True Then MsgBox ("Por favor, no ingrese números en el campo apellido"), vbInformation, "Aviso": txt_apellido.SetFocus: Exit Sub
If UCase(Trim(txt_apellido.Text)) = UCase(Trim(txt_nombre.Text)) Then MsgBox ("El nombre y apellido no pueden ser los mismos"), vbInformation, "Aviso": txt_apellido.SetFocus: Exit Sub
If txt_cedula.Text = "" Then MsgBox "Por favor, ingrese su cédula", vbInformation, "Aviso": txt_cedula.SetFocus: Exit Sub
          If Not IsNumeric(txt_cedula.Text) Then MsgBox "Los datos que debe introducir en la cédula deben ser sólo números", vbInformation, "Aviso": txt_cedula.SetFocus: Exit Sub
          If Val(txt_cedula.Text) < 1000000 Then MsgBox ("Cédula inválida, por favor verifique el número e intente de nuevo"), vbInformation, "Aviso": txt_cedula.SetFocus: Exit Sub
          
          With RsUsuarios
                    .Requery
                    .Find "CI='" & Trim(txt_cedula.Text) & "'"
                    If .EOF Then
                              'no hago nada
                    Else
                              MsgBox ("Ya existe un usuario registrado con esta cédula, por favor verifique"), vbInformation, "Aviso": Exit Sub
                    End If
          End With
          
If txt_username.Text = "" Then MsgBox "Por favor, ingrese su nombre de usuario", vbInformation, "Aviso": txt_username.SetFocus: Exit Sub
If txt_pass.Text = "" Then MsgBox "Por favor, ingrese una contraseña", vbInformation, "Aviso": txt_pass.SetFocus: Exit Sub
          If Len(txt_pass.Text) < 6 Then MsgBox ("La contraseña debe ser mayor a 6 caracteres"), vbInformation, "Aviso": txt_pass.SetFocus: Exit Sub
If txt_rpass.Text = "" Then MsgBox "Por favor, repita su contraseña", vbInformation, "Aviso": txt_rpass.SetFocus: Exit Sub
          If txt_pass.Text <> txt_rpass.Text Then MsgBox "Las contraseñas deben coincidir", vbInformation, "Aviso": txt_rpass.SetFocus: Exit Sub
If txt_correo.Text = "" Then MsgBox "Por favor, ingrese su correo", vbInformation, "Aviso": txt_correo.SetFocus: Exit Sub

     Dim ret As Boolean
     Dim Direccion As String
    Direccion = Trim(txt_correo.Text)
    If Direccion = vbNullString Then Exit Sub
    ' ejecuta la función
    ret = Comprobar_Mail(Direccion)
    ' Resultado
    If ret Then
                ' Ok
    Else
               MsgBox ("La dirección de correo ingresada es inválida, por favor cambiela"), vbInformation, "Aviso": Exit Sub ' El mail no es correcto
    End If

     With RsUsuarios
               .Requery
               .Find "Correo='" & Trim(txt_correo.Text) & "'"
               If .EOF Then
                         'no hago nada
               Else
                         MsgBox ("Ya existe un usuario registrado con este correo, por favor verifique"), vbInformation, "Aviso": Exit Sub
               End If
     End With

If cbo_Tarea.Text = "" Then MsgBox ("Por favor, indique o seleccione una tarea (rol) dentro del almacén"), vbInformation, "Aviso": cbo_Tarea.ListIndex = 0: cbo_Tarea.SetFocus: Exit Sub
If cbo_Tarea.Text = "Seleccione" Then MsgBox ("Por favor, indique cual es su tarea dentro del almacén (Rol)"), vbInformation, "Aviso": cbo_Tarea.SetFocus: Exit Sub
'If txt_tarea.Text = "" Then MsgBox "Por favor ingrese una tarea", vbInformation, "Aviso": txt_tarea.SetFocus: Exit Sub
'If opt_user = False And opt_admi = False Then MsgBox "Seleccione un rango", vbInformation, "Aviso": Exit Sub

With RsUsuarios
    .Requery
    .Find "nomb_usuario='" & Trim(txt_username.Text) & "'"
    
    If .EOF Then 'si no encontro nada
        .AddNew
            !Nombre = txt_nombre.Text
            !Apellido = txt_apellido.Text
            !CI = txt_cedula.Text
            !Nomb_Usuario = txt_username.Text
        
           ' If opt_user = True Then
                !rango = 0
            'Else
             '   !rango = 1
           ' End If
                
            !Contraseña = txt_pass.Text
            !Correo = txt_correo.Text
            !Tarea = cbo_Tarea.Text
    .Update
    .Requery
    MsgBox ("Usuario registrado con éxito"), vbInformation, "Aviso"
    Logs
    With RsLogs
               .Requery
               .AddNew
                         !User = "Sistema"
                         !Accion = "creó un usuario a " & txt_nombre & " " & txt_apellido & " (" & txt_username.Text & ") con rango 0 el"
                         !Fecha = Date + Time
               .Update
               .Requery
               .Close
    End With
    LogsUsuarios
    With RsLogsUsuarios
               .Requery
               .AddNew
                         !User = "Sistema"
                         !Accion = "creó un usuario a " & txt_nombre & " " & txt_apellido & " (" & txt_username.Text & ") con rango 0 el"
                         !Fecha = Date + Time
               .Update
               .Requery
               .Close
    End With
    Limpiar
    Unload Me
    LoginForm.Show
Else
    'si el usuario ya existe
    MsgBox "Ya existe una cuenta con ese nombre de usuario, por favor cambielo", vbInformation, "Aviso"
    txt_username.Text = ""
    txt_username.SetFocus
End If
End With
End Sub

Sub Limpiar()

txt_nombre.Text = ""
txt_apellido.Text = ""
txt_cedula.Text = ""
txt_username.Text = ""
txt_pass.Text = ""
txt_rpass.Text = ""
txt_correo.Text = ""
cbo_Tarea.ListIndex = 0
'txt_tarea.Text = ""
'opt_user = False
'opt_admi = False
End Sub

Private Sub Form_Load()
cbo_Tarea.AddItem "Seleccione"
cbo_Tarea.AddItem "Almacenista"
cbo_Tarea.AddItem "Supervisor de Almacén"
cbo_Tarea.AddItem "Vendedor"
cbo_Tarea.ListIndex = 0
End Sub


