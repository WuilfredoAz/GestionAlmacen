VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form UsuariosEditForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar datos de usuario"
   ClientHeight    =   11475
   ClientLeft      =   6120
   ClientTop       =   1155
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "UsuariosEditForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "UsuariosEditForm.frx":08CA
   ScaleHeight     =   11475
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid GrillaUsuarios 
      Height          =   2535
      Left            =   1050
      TabIndex        =   16
      Top             =   2640
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   -2147483641
      HeadLines       =   1
      RowHeight       =   23
      RowDividerStyle =   6
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Usuarios Registrados"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         SizeMode        =   1
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cbo_Tarea 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3960
      TabIndex        =   21
      Text            =   "cbo_Tarea"
      Top             =   9000
      Width           =   4455
   End
   Begin VB.CommandButton cmd_Eliminar 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   4200
      TabIndex        =   20
      Top             =   10560
      Width           =   1815
   End
   Begin VB.CommandButton cmd_atras 
      Appearance      =   0  'Flat
      Caption         =   "Atras"
      Height          =   615
      Left            =   1080
      TabIndex        =   6
      Top             =   10560
      Width           =   1695
   End
   Begin VB.CommandButton cmd_guardar 
      Caption         =   "Guardar Cambios"
      Height          =   615
      Left            =   6240
      TabIndex        =   5
      Top             =   10560
      Width           =   2055
   End
   Begin VB.OptionButton opt_admi 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Administrador"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   9840
      Width           =   2055
   End
   Begin VB.OptionButton opt_user 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Usuario"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   9840
      Width           =   1575
   End
   Begin VB.TextBox txt_rpass 
      Enabled         =   0   'False
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3960
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   8040
      Width           =   4455
   End
   Begin VB.TextBox txt_pass 
      Enabled         =   0   'False
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3960
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   7080
      Width           =   4455
   End
   Begin VB.TextBox txt_username 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   6240
      Width           =   4455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1200
      TabIndex        =   19
      Top             =   5640
      Width           =   225
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Realize las modificaciones correspondientes"
      Height          =   270
      Left            =   1560
      TabIndex        =   18
      Top             =   5760
      Width           =   4665
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1200
      TabIndex        =   17
      Top             =   1800
      Width           =   225
   End
   Begin VB.Label lbl_descripcion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el usuario que desea modificar"
      Height          =   270
      Left            =   1560
      TabIndex        =   15
      Top             =   1920
      Width           =   4440
   End
   Begin VB.Label lbl_rango 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rango:"
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
      Left            =   2520
      TabIndex        =   14
      Top             =   9840
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ejm: Supervisor de Almacén"
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
      Left            =   3960
      TabIndex        =   13
      Top             =   9360
      Width           =   2085
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
      Left            =   2640
      TabIndex        =   12
      Top             =   9120
      Width           =   735
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
      Left            =   3960
      TabIndex        =   11
      Top             =   8520
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
      Left            =   1200
      TabIndex        =   10
      Top             =   8280
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
      Left            =   3960
      TabIndex        =   9
      Top             =   7560
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
      Left            =   1920
      TabIndex        =   8
      Top             =   7320
      Width           =   1425
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
      Left            =   1080
      TabIndex        =   7
      Top             =   6480
      Width           =   2310
   End
End
Attribute VB_Name = "UsuariosEditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer 'valor que toma la id de los usuarios

Private Sub cmd_Atras_Click()
Limpiar
Unload Me
End Sub

Private Sub cmd_Eliminar_Click()
If Trim(GrillaUsuarios.Columns(1).Text) = "Administrador" Then
          MsgBox ("No se puede borrar este usuario"), vbCritical, "Error": Exit Sub
ElseIf MsgBox("Esta a punto de eliminar a " & GrillaUsuarios.Columns(2).Text & " " & GrillaUsuarios.Columns(3).Text & " del sistema, ¿Desea continuar con la operación?", vbinforamtion + vbYesNo, "Advertencia") = vbYes Then
          'si me elimino a mi mismo
          If UCase(Trim(GrillaUsuarios.Columns(1).Text)) = UCase(Trim(vUsername)) Then
                    'confirmar eliminarme a mi mismo
                   If MsgBox("Esta a punto de borrar sus datos del Sistema, al realizar esta operación su sesión será cerrada y todos sus datos serán eliminados. ¿Está seguro que desea continuar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                              With RsUsuarios
                                        .Find "Nomb_Usuario='" & Trim(GrillaUsuarios.Columns(1).Text) & "'"
                                        Logs
                                        With RsLogs
                                                  .Requery
                                                  .AddNew
                                                            !User = vUsername
                                                            !Accion = "eliminó al usuario  " & GrillaUsuarios.Columns(2).Text & " " & GrillaUsuarios.Columns(3).Text & " (" & GrillaUsuarios.Columns(1).Text & " se eliminó él mismo) del sistema"
                                                            !Fecha = Date + Time
                                                  .Update
                                                  .Requery
                                                  .Close
                                         End With
                                        LogsUsuarios
                                        With RsLogsUsuarios
                                                  .Requery
                                                  .AddNew
                                                            !User = vUsername
                                                            !Accion = "eliminó al usuario  " & GrillaUsuarios.Columns(2).Text & " " & GrillaUsuarios.Columns(3).Text & " (" & GrillaUsuarios.Columns(1).Text & " se eliminó él mismo) del sistema"
                                                            !Fecha = Date + Time
                                                  .Update
                                                  .Requery
                                                  .Close
                                         End With
                                        .Delete
                                        .Requery
                                        vTarea = ""
                                        vUsername = ""
                                        Unload IndexForm
                                        Unload ConsultasForm
                                        Unload UbicacionesForm
                                        Unload DespachosForm
                                        Unload ProductosForm
                                        Unload UsuariosForm
                                        Unload ReportesForm
                                        Unload MantenimientoForm
                                        Unload Me
                                        LoginForm.Picture = LoadPicture(App.Path & "\Interfaz\Login.jpg")
                                        LoginForm.Show
                              End With
                   Else
                              'si no me quiero eliminar
                              Exit Sub
                   End If
          'si no soy yo mismo
          Else
                    With RsUsuarios
                              .Find "Nomb_Usuario='" & Trim(GrillaUsuarios.Columns(1).Text) & "'"
                              Logs
                              With RsLogs
                                        .Requery
                                        .AddNew
                                                  !User = vUsername
                                                  !Accion = "eliminó al usuario  " & GrillaUsuarios.Columns(2).Text & " " & GrillaUsuarios.Columns(3).Text & " (" & GrillaUsuarios.Columns(1).Text & ") del sistema"
                                                  !Fecha = Date + Time
                                        .Update
                                        .Requery
                                        .Close
                              End With
                              LogsUsuarios
                              With RsLogsUsuarios
                                        .Requery
                                        .AddNew
                                                  !User = vUsername
                                                  !Accion = "eliminó al usuario  " & GrillaUsuarios.Columns(2).Text & " " & GrillaUsuarios.Columns(3).Text & " (" & GrillaUsuarios.Columns(1).Text & ") del sistema"
                                                  !Fecha = Date + Time
                                        .Update
                                        .Requery
                                        .Close
                              End With
                              .Delete
                              .Requery
                    End With
                    EstilosGrilla
          End If

Else
          Exit Sub
End If


End Sub

Private Sub cmd_guardar_Click()

txt_username.Enabled = False
txt_pass.Enabled = False
txt_rpass.Enabled = False
'txt_tarea.Enabled = False
cbo_Tarea.Enabled = False
opt_user.Enabled = False
opt_admi.Enabled = False

If txt_username.Text = "Administrador" And Not (vUsername = "Administrador") = True Then MsgBox ("Usted no tiene permiso para editar a este usuario"), vbInformation, "Aviso": Exit Sub
If txt_username.Text = "" Then MsgBox "Por favor ingrese su nombre de usuario", vbInformation, "Aviso":  Exit Sub 'OJO NO PUEDO MANDAR SETFOCUS EN NINGUNO PORQUE LE PUSE ENABLE=FALSE ACOMODAR DESPUES
If txt_pass.Text = "" Then MsgBox "Por favor ingrese una contraseña", vbInformation, "Aviso": Exit Sub
     If Len(txt_pass.Text) < 6 Then MsgBox ("La contraseña debe ser mayor a 6 caracteres"), vbInformation, "Aviso": Exit Sub
If txt_rpass.Text = "" Then MsgBox "Por favor repita su contraseña", vbInformation, "Aviso":  Exit Sub
If txt_pass.Text <> txt_rpass.Text Then MsgBox "Las contraseñas deben coincidir", vbInformation, "Aviso":  Exit Sub
If cbo_Tarea.Text = "" Then MsgBox "Por favor, indique o seleccione una tarea (rol) dentro del almacén", vbInformation, "Aviso":  Exit Sub
If cbo_Tarea.Text = "Seleccione" Then MsgBox ("Por favor, indique cual es su tarea dentro del almacén (Rol)"), vbInformation, "Aviso": Exit Sub
If opt_user = False And opt_admi = False Then MsgBox "Seleccione un rango", vbInformation, "Aviso": Exit Sub

With RsUsuarios
          .Requery
          .Find "id='" & Val(a) & "'"
                    !Nomb_Usuario = txt_username.Text
                    !Contraseña = txt_pass.Text
                    !Tarea = cbo_Tarea.Text
                    If opt_user = True Then
                              !rango = 0
                              Logs
                              With RsLogs
                                        .Requery
                                        .AddNew
                                                  !User = vUsername
                                                  !Accion = "editó  al usuario  " & GrillaUsuarios.Columns(2).Text & " " & GrillaUsuarios.Columns(3).Text & " (" & txt_username.Text & ") y le reasigno el rango: 0 el"
                                                  !Fecha = Date + Time
                                        .Update
                                        .Requery
                                        .Close
                              End With
                              LogsUsuarios
                              With RsLogsUsuarios
                                        .Requery
                                        .AddNew
                                                  !User = vUsername
                                                  !Accion = "editó  al usuario  " & GrillaUsuarios.Columns(2).Text & " " & GrillaUsuarios.Columns(3).Text & " (" & txt_username.Text & ") y le reasigno el rango: 0 el"
                                                  !Fecha = Date + Time
                                        .Update
                                        .Requery
                                        .Close
                              End With
                    Else
                              !rango = 1
                              Logs
                              With RsLogs
                                        .Requery
                                        .AddNew
                                                  !User = vUsername
                                                  !Accion = "editó  al usuario  " & GrillaUsuarios.Columns(2).Text & " " & GrillaUsuarios.Columns(3).Text & " (" & txt_username.Text & ") y le reasigno el rango: 1 el"
                                                  !Fecha = Date + Time
                                        .Update
                                        .Requery
                                        .Close
                              End With
                              LogsUsuarios
                              With RsLogsUsuarios
                                        .Requery
                                        .AddNew
                                                  !User = vUsername
                                                  !Accion = "editó  al usuario  " & GrillaUsuarios.Columns(2).Text & " " & GrillaUsuarios.Columns(3).Text & " (" & txt_username.Text & ") y le reasigno el rango: 1 el"
                                                  !Fecha = Date + Time
                                        .Update
                                        .Requery
                                        .Close
                              End With
                    End If
          .UpdateBatch
          .Requery
          EstilosGrilla
          Limpiar
          MsgBox ("Usuario editado con éxito"), vbInformation, "Aviso"
End With

End Sub
Sub Limpiar()
txt_username.Text = ""
txt_pass.Text = ""
txt_rpass.Text = ""
'txt_tarea.Text = ""
cbo_Tarea.ListIndex = 0
opt_user = False
opt_admi = False
End Sub

Sub EstilosGrilla()

'tamaños de la grilla
GrillaUsuarios.Columns(0).Width = 500
GrillaUsuarios.Columns(1).Width = 2400
GrillaUsuarios.Columns(2).Width = 2400
GrillaUsuarios.Columns(3).Width = 2000
GrillaUsuarios.Columns(4).Width = 2000
GrillaUsuarios.Columns(5).Width = 2000
GrillaUsuarios.Columns(6).Width = 500
GrillaUsuarios.Columns(7).Width = 2000
GrillaUsuarios.Columns(8).Width = 2400

'Caption de las grillas
GrillaUsuarios.Columns(0).Caption = "ID"
GrillaUsuarios.Columns(1).Caption = "Usuario"
GrillaUsuarios.Columns(2).Caption = "Nombre"
GrillaUsuarios.Columns(3).Caption = "Apellido"
GrillaUsuarios.Columns(4).Caption = "Contraseña"
GrillaUsuarios.Columns(5).Caption = "Email"
GrillaUsuarios.Columns(6).Caption = "Rank"
GrillaUsuarios.Columns(7).Caption = "Cédula"
GrillaUsuarios.Columns(8).Caption = "Tarea"

'alineacion
GrillaUsuarios.Columns(0).Alignment = dbgCenter
GrillaUsuarios.Columns(1).Alignment = dbgCenter
GrillaUsuarios.Columns(2).Alignment = dbgCenter
GrillaUsuarios.Columns(3).Alignment = dbgCenter
GrillaUsuarios.Columns(4).Alignment = dbgCenter
GrillaUsuarios.Columns(5).Alignment = dbgCenter
GrillaUsuarios.Columns(6).Alignment = dbgCenter
GrillaUsuarios.Columns(7).Alignment = dbgCenter
GrillaUsuarios.Columns(8).Alignment = dbgCenter

'cabeceras
GrillaUsuarios.HeadFont.Bold = True

'las que no quiero ver
GrillaUsuarios.Columns(0).Visible = False
GrillaUsuarios.Columns(4).Visible = False
GrillaUsuarios.Columns(5).Visible = False
GrillaUsuarios.Columns(6).Visible = False
GrillaUsuarios.Columns(7).Visible = False
GrillaUsuarios.Columns(8).Visible = False

End Sub

Private Sub Form_Load()
Usuarios
Set GrillaUsuarios.DataSource = RsUsuarios
EstilosGrilla
cmd_Eliminar.Enabled = False
cmd_guardar.Enabled = False

cbo_Tarea.AddItem "Seleccione"
cbo_Tarea.AddItem "Almacenista"
cbo_Tarea.AddItem "Supervisor de Almacén"
cbo_Tarea.AddItem "Vendedor"
cbo_Tarea.ListIndex = 0
cbo_Tarea.Enabled = False
End Sub
Private Sub GrillaUsuarios_Click()
With RsUsuarios
    If .BOF Or .EOF Then Exit Sub 'el valor de eof es el fin de la linea y el bof que es ningun es decir si existe ambos (o sea nada esta seleccionado) no haga nada
    .Find "id='" & Val(GrillaUsuarios.Columns(o).Text) & "'"
    a = !id ' valor declararo en "generales"
        
    txt_username.Text = !Nomb_Usuario
    txt_pass.Text = !Contraseña
    txt_rpass.Text = !Contraseña
    cbo_Tarea.Text = !Tarea
    
    If !rango = 1 Then
        opt_admi = True
    Else
        opt_user = True
    End If
    
    txt_username.Enabled = True
    txt_pass.Enabled = True
    txt_rpass.Enabled = True
    'txt_tarea.Enabled = True
    cbo_Tarea.Enabled = True
    opt_user.Enabled = True
    opt_admi.Enabled = True
    cmd_Eliminar.Enabled = True
    
    
End With
EstilosGrilla
cmd_Eliminar.Enabled = True
cmd_guardar.Enabled = False
End Sub

Private Sub opt_admi_GotFocus()
cmd_guardar.Enabled = True
End Sub

Private Sub opt_user_GotFocus()
cmd_guardar.Enabled = True
End Sub

Private Sub txt_pass_KeyPress(KeyAscii As Integer)
cmd_guardar.Enabled = True
End Sub

Private Sub txt_rpass_KeyPress(KeyAscii As Integer)
cmd_guardar.Enabled = True
End Sub

Private Sub cbo_Tarea_Click()
cmd_guardar.Enabled = True
End Sub

'Private Sub txt_tarea_KeyPress(KeyAscii As Integer)
'cmd_guardar.Enabled = True
'End Sub
