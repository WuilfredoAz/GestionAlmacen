VERSION 5.00
Begin VB.Form LoginForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login - Shekina 4004"
   ClientHeight    =   5685
   ClientLeft      =   7860
   ClientTop       =   4035
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LoginForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "LoginForm.frx":08CA
   ScaleHeight     =   5685
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_registrarse 
      Caption         =   "Registrarse"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmd_iniciar 
      Caption         =   "Iniciar Sesion"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox txt_pass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3480
      Width           =   3255
   End
   Begin VB.TextBox txt_nombreUser 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   1
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label lbl_OlvidoContraseña 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¿Olvido su contraseña?"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3000
      MouseIcon       =   "LoginForm.frx":9F7C
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4200
      Width           =   1485
   End
   Begin VB.Label lbl_pass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1200
      TabIndex        =   5
      Top             =   3840
      Width           =   720
   End
   Begin VB.Label lbl_nombreUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de Usuario"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1200
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Intetos As Integer
Private Sub cmd_iniciar_Click()
If txt_nombreUser.Text = "" And txt_pass.Text = "" Then MsgBox ("Para entrar al sistema, primero ingrese sus datos"), vbInformation, "Aviso": txt_nombreUser.SetFocus: Exit Sub
If txt_nombreUser.Text = "" Then MsgBox "Ingrese nombre de usuario", vbInformation, "Aviso": txt_nombreUser.SetFocus: Exit Sub 'comprobar campo nombre este vacio
If txt_pass.Text = "" Then MsgBox "Ingrese la contraseña", vbInformation, "Aviso": txt_pass.SetFocus: Exit Sub 'comprobar campo pass este vacio
BuscamosImagenes

With RsUsuarios
    .Requery 'actualizar la db
    .Find "Nomb_Usuario='" & Trim(txt_nombreUser.Text) & "'" 'comparo nombre usuario de db con el introducido
    If .EOF Then
          LoginForm.Picture = LoadPicture(App.Path & Error) 'cambiar imagen a rojo por error
          MsgBox "Usuario Incorrecto", vbInformation, "Aviso"
                    Intetos = Intetos + 1
                    If Intetos = 3 Then
                              ActualizoImagenesForzosas
                              Unload Me
                    End If
        Exit Sub 'dejo de ejecutar este sub
    Else 'en caso contrario
        If !Contraseña = Trim(txt_pass.Text) Then 'pregunto si la clave es correcta
        
          ActualizoImagenes
          Intetos = 0
            vUsername = !Nomb_Usuario 'mando a una variable el nombre para mostrarla en el index
            vTarea = !Tarea 'mando a una variable la tarea para mostrarla en el index
            IndexForm.Show
            LoginForm.Hide
            txt_pass.Text = ""
            txt_nombreUser.Text = ""
                                    
            
                If !rango = 1 Then
                    IndexForm.Picture = LoadPicture(App.Path & "\Interfaz\IndexStacADM.jpg")
                    ConsultasForm.Picture = LoadPicture(App.Path & "\Interfaz\ConsultasADM.jpg")
                    DespachosForm.Picture = LoadPicture(App.Path & "\Interfaz\IndexADM.jpg")
                    ProductosForm.Picture = LoadPicture(App.Path & "\Interfaz\IndexADM.jpg")
                    UbicacionesForm.Picture = LoadPicture(App.Path & "\Interfaz\UbicacionesFormADM.jpg")
                    UsuariosForm.Picture = LoadPicture(App.Path & "\Interfaz\IndexADM.jpg")
                    ReportesForm.Picture = LoadPicture(App.Path & "\Interfaz\IndexADM.jpg")
                    MantenimientoForm.Picture = LoadPicture(App.Path & "\Interfaz\Mantenimiento.jpg")
                
                Else
                    IndexForm.cmd_productos.Visible = False
                    IndexForm.cmd_usuarios.Visible = False
                    IndexForm.cmd_Mantenimiento.Visible = False
                    IndexForm.cmd_Logs.Visible = False
                    
                    ConsultasForm.cmd_productos.Visible = False
                    ConsultasForm.cmd_usuarios.Visible = False
                    ConsultasForm.cmd_Mantenimiento.Visible = False
                    ConsultasForm.cmd_Logs.Visible = False
                    
                    UbicacionesForm.cmd_productos.Visible = False
                    UbicacionesForm.cmd_usuarios.Visible = False
                    UbicacionesForm.cmd_Mantenimiento.Visible = False
                    UbicacionesForm.cmd_Logs.Visible = False
                    
                    DespachosForm.cmd_productos.Visible = False
                    DespachosForm.cmd_usuarios.Visible = False
                    DespachosForm.cmd_Mantenimiento.Visible = False
                    DespachosForm.cmd_Logs.Visible = False
                    
                    ReportesForm.cmd_productos.Visible = False
                    ReportesForm.cmd_usuarios.Visible = False
                    ReportesForm.cmd_Mantenimiento.Visible = False
                    ReportesForm.cmd_Logs.Visible = False
                    
                    'Botones de reportes (opciones)
                    ReportesForm.cmd_ReporteUsuarios.Visible = False
                    ReportesForm.Label3.Visible = False
                    
                    ReportesForm.cmd_ReporteDespachos.Left = 6120
                    ReportesForm.Label4.Left = 6120
                    
                    ReportesForm.cmd_Devoluciones.Left = 9360
                    ReportesForm.Label7.Left = 9360
                    
                End If
            
        Else 'caso contrario
            LoginForm.Picture = LoadPicture(App.Path & Error) 'cambiar imagen a rojo por error
            MsgBox "Clave incorrecta", vbInformation, "Aviso" 'mensaje de error
                    Intetos = Intetos + 1
                    If Intetos = 3 Then
                              ActualizoImagenesForzosas
                              Unload Me
                    End If
            Exit Sub 'dejo de ejectutar este sub
        End If
    End If
End With


End Sub

Private Sub cmd_registrarse_Click()
Unload Me
RegisterForm.Show
End Sub

Private Sub Form_Load()
Usuarios
Config
BuscamosImagenes
LoginForm.Picture = LoadPicture(App.Path & Normal)
End Sub

Private Sub lbl_OlvidoContraseña_Click()
LoginForm.Hide
ConfirmarDatosForm.Show
End Sub

Private Sub txt_pass_Click()
txt_pass.Text = ""
End Sub

Sub BuscamosImagenes()
With RsConfig
          .Requery
          .Find "Id=1"
          Normal = !Ruta
End With
With RsConfig
          .Requery
          .Find "Id=2"
          Error = !Ruta
End With
End Sub

Sub ActualizoImagenesForzosas()
With RsConfig
          .Requery
          .Find "Id=1"
          !Ruta = "\Interfaz\LoginForzoso.jpg"
          .UpdateBatch
          .Requery
End With
With RsConfig
          .Requery
          .Find "Id=2"
          !Ruta = "\Interfaz\LoginForzoso_Error.jpg"
          .UpdateBatch
          .Requery
End With
End Sub

Sub ActualizoImagenes()
With RsConfig
          .Requery
          .Find "Id=1"
          !Ruta = "\Interfaz\Login.jpg"
          .UpdateBatch
          .Requery
End With
With RsConfig
          .Requery
          .Find "Id=2"
          !Ruta = "\Interfaz\Login_Error.jpg"
          .UpdateBatch
          .Requery
End With
End Sub
