VERSION 5.00
Begin VB.Form UsuariosForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   10500
   ClientLeft      =   2460
   ClientTop       =   1515
   ClientWidth     =   16500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "UsuariosForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "UsuariosForm.frx":058A
   ScaleHeight     =   10500
   ScaleWidth      =   16500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Logs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Logs"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton cmd_Mantenimiento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Mantenimiento"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton cmd_reportes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Reportes"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000001&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   12480
      Top             =   840
   End
   Begin VB.CommandButton cmd_eliminaruser 
      Height          =   2655
      Left            =   11880
      Picture         =   "UsuariosForm.frx":188DE
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Width           =   3615
   End
   Begin VB.CommandButton cmd_editaruser 
      Height          =   2655
      Left            =   7320
      Picture         =   "UsuariosForm.frx":1D39C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4200
      Width           =   3615
   End
   Begin VB.CommandButton cmd_newuser 
      Height          =   2655
      Left            =   2880
      Picture         =   "UsuariosForm.frx":21A8C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   3615
   End
   Begin VB.CommandButton cmd_inicio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Inicio"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmd_Consultas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Consultas"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmd_despachos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Despachos"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmd_ubicaciones 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Ubicaciones"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmd_usuarios 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Usuarios"
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton cmd_productos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Productos"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Por favor, seleccione la opción que desea realizar:"
      ForeColor       =   &H80000011&
      Height          =   270
      Left            =   3480
      TabIndex        =   20
      Top             =   3480
      Width           =   5250
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GESTIÓN DE USUARIOS DEL SISTEMA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   510
      Left            =   2760
      TabIndex        =   19
      Top             =   2640
      Width           =   8055
   End
   Begin VB.Label lbl_username 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   510
      Left            =   1920
      TabIndex        =   18
      Top             =   600
      Width           =   2460
   End
   Begin VB.Label lbl_tarea 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAREA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label lbl_fecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HORA ACTUAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Index           =   0
      Left            =   13560
      TabIndex        =   16
      Top             =   600
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   570
      Left            =   12960
      TabIndex        =   15
      Top             =   720
      Width           =   2610
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Index           =   0
      Left            =   13800
      TabIndex        =   14
      Top             =   1560
      Width           =   45
   End
   Begin VB.Label lbl_cerrarsesion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CERRAR SESION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   13920
      MouseIcon       =   "UsuariosForm.frx":26A7D
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   1560
      Width           =   1605
   End
   Begin VB.Label lbl_faq 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AYUDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   12960
      MouseIcon       =   "UsuariosForm.frx":26D87
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   1560
      Width           =   675
   End
End
Attribute VB_Name = "UsuariosForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Consultas_Click()
UsuariosForm.Hide
ConsultasForm.Show
End Sub
Private Sub cmd_despachos_Click()
UsuariosForm.Hide
DespachosForm.Show
End Sub
Private Sub cmd_editaruser_Click()
UsuariosEditForm.Show vbModal
End Sub
Private Sub cmd_eliminaruser_Click()
Dim BorrarUsuarios As String
BorrarUsuarios = InputBox("Ingrese el nombre de usuario que desea Eliminar", "Eliminar Usuario")

If UCase(BorrarUsuarios) = UCase("Administrador") Then
          MsgBox ("Éste usuario no puede ser eliminado"), vbCritical, "Error": Exit Sub

ElseIf StrPtr(BorrarUsuarios) = vbEmpty Then
          'si el usuario lo cancela
          MsgBox ("Cancelado por el usuario")
          Exit Sub
Else
          'si esta vacio
          If BorrarUsuarios = "" Then
                    MsgBox ("Para proceder con la operación, introduzca primero un nombre de usuario")
                    Exit Sub
          Else
                    'si soy yo mismo
                    If UCase(Trim(BorrarUsuarios)) = UCase(Trim(vUsername)) Then
                              If MsgBox("Esta a punto de borrar sus datos del Sistema, al realizar esta operacion su sesión será cerrada y todos sus datos serán eliminados. ¿Esta seguro que desea continuar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                        With RsUsuarios
                                                  .Find "Nomb_Usuario='" & UCase(Trim(BorrarUsuarios)) & "'"
                                                            Logs
                                                            With RsLogs
                                                                      If .State = 1 Then .Close
                                                                      .Open "SELECT * FROM Logs"
                                                                      .Requery
                                                                      .AddNew
                                                                                !User = vUsername
                                                                                !Accion = "se eliminó él mismo del sistema"
                                                                                !Fecha = Date + Time
                                                                      .Update
                                                                      .Requery
                                                                      .Close
                                                            End With
                                                            LogsUsuarios
                                                            With RsLogsUsuarios
                                                                      If .State = 1 Then .Close
                                                                      .Open "SELECT * FROM LogsUsuarios"
                                                                      .Requery
                                                                      .AddNew
                                                                                !User = vUsername
                                                                                !Accion = "se eliminó él mismo del sistema"
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
                                        'MsgBox ("soy yo y me cancele")
                                        Exit Sub
                              End If
                                    
                    Else
                              With RsUsuarios
                                        .Requery
                                        .Find "Nomb_Usuario='" & UCase(Trim(BorrarUsuarios)) & "'"
                                        'si no lo encuentro
                                        If .EOF Then
                                                  MsgBox ("Nombre de usuario invalido"), vbInformation, "Aviso": Exit Sub
                                        Else
                                                  If MsgBox("Esta a punto de borrar al usuario: " & UCase(Trim(BorrarUsuarios)) & " del sistema, ¿Desea continuar con la operación?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                                            Logs
                                                            With RsLogs
                                                                      If .State = 1 Then .Close
                                                                      .Open "SELECT * FROM Logs"
                                                                      .Requery
                                                                      .AddNew
                                                                                !User = vUsername
                                                                                !Accion = "eliminó al usuario " & Trim(BorrarUsuarios) & " del sistema"
                                                                                !Fecha = Date + Time
                                                                      .Update
                                                                      .Requery
                                                                      .Close
                                                            End With
                                                            LogsUsuarios
                                                            With RsLogsUsuarios
                                                                      If .State = 1 Then .Close
                                                                      .Open "SELECT * FROM LogsUsuarios"
                                                                      .Requery
                                                                      .AddNew
                                                                                !User = vUsername
                                                                                !Accion = "eliminó al usuario " & Trim(BorrarUsuarios) & " del sistema"
                                                                                !Fecha = Date + Time
                                                                      .Update
                                                                      .Requery
                                                                      .Close
                                                            End With
                                                            .Delete
                                                            .Requery
                                                            UsuariosEditForm.EstilosGrilla
                                                            MsgBox ("Usuario eliminado con exito"), vbInformation, "Aviso"
                                                  Else
                                                            'MsgBox ("No soy yo y cancele"):
                                                            Exit Sub
                                                  End If
                                        End If
                              End With
                    End If
          End If
End If
End Sub
Private Sub cmd_inicio_Click()
UsuariosForm.Hide
IndexForm.Show
End Sub

Private Sub cmd_Logs_Click()
LogsForm.Show
End Sub

Private Sub cmd_Mantenimiento_Click()
UsuariosForm.Hide
MantenimientoForm.Show
End Sub

Private Sub cmd_newuser_Click()
UsuariosNewForm.Show
End Sub
Private Sub cmd_productos_Click()
UsuariosForm.Hide
ProductosForm.Show
End Sub

Private Sub cmd_reportes_Click()
UsuariosForm.Hide
ReportesForm.Show
End Sub

Private Sub cmd_ubicaciones_Click()
UsuariosForm.Hide
UbicacionesForm.Show
End Sub
Private Sub Form_Load()
lbl_tarea.Caption = vTarea
lbl_username.Caption = vUsername
End Sub

Private Sub lbl_cerrarsesion_Click()
If MsgBox("¿Desea salir del Sistema?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    vTarea = ""
    vUsername = ""
     Unload LoginForm
               Unload RegisterForm
               Unload ConfirmarDatosForm
                         Unload RestablecerPassForm
     Unload IndexForm
               Unload AcercaForm
               Unload AyudaForm
     Unload ConsultasForm
               Unload ConsultasVerForm
     Unload UbicacionesForm
               Unload UbicacionesSeccionesForm
     Unload DespachosForm
               Unload DespachosHistorialForm
                         Unload DespachosVerForm
               Unload DespachosNewForm
                         Unload DespachosAddProductosForm
               Unload DespachosShippingSelectForm
                         Unload DespachosShippingDetalles
               Unload DespachosDevolucionesHistorialForm
                         Unload DespachosDevolucionesVerForm
               Unload DespachosDevolucionesNewForm
                         Unload DespachosDevolucionesDetallesForm
               Unload DespachosConsultasForm
     Unload ReportesForm
     Unload ProductosForm
               Unload ProductoNewForm
                         Unload ProductosUbicacionHelpForm
                                   Unload ProductosNormasForm
               Unload ProductosEditForm
                         Unload ProductosEdit2Form
     Unload UsuariosForm
               Unload UsuariosNewForm
               Unload UsuariosEditForm
     Unload LogsForm
               Unload LogsProductosForm
               Unload LogsReportesForm
               Unload LogsMantenimientoForm
               Unload LogsDespachosForm
               Unload LogsUsuariosForm
               Unload LogsAllForm
     Unload MantenimientoForm
    LoginForm.Picture = LoadPicture(App.Path & "\Interfaz\Login.jpg")
    LoginForm.Show
End If
End Sub

Private Sub lbl_faq_Click()
AyudaForm.Show vbModal
End Sub

Private Sub Timer1_Timer()
'Label1.Caption = Format(Time, "hh:mm:ss")
Label1.Caption = Format(Now, "HH:MM AM/PM")
End Sub
