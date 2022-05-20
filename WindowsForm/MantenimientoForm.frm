VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MantenimientoForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento"
   ClientHeight    =   10500
   ClientLeft      =   2415
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
   Icon            =   "MantenimientoForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "MantenimientoForm.frx":058A
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
      TabIndex        =   8
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton cmd_Restablecer 
      Height          =   2295
      Left            =   12600
      Picture         =   "MantenimientoForm.frx":1AD28
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4440
      Width           =   2895
   End
   Begin VB.CommandButton cmd_CompactarReparar 
      Height          =   2295
      Left            =   9360
      Picture         =   "MantenimientoForm.frx":1F0B7
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   2895
   End
   Begin VB.CommandButton cmd_Restore 
      Height          =   2295
      Left            =   6120
      Picture         =   "MantenimientoForm.frx":2376C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   2895
   End
   Begin VB.CommandButton cmd_CrearBackup 
      Height          =   2295
      Left            =   2880
      Picture         =   "MantenimientoForm.frx":284CD
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   2895
   End
   Begin VB.CommandButton cmd_Mantenimiento 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Mantenimiento"
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton cmd_productos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Productos"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000005&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton cmd_usuarios 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Usuarios"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000005&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton cmd_ubicaciones 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Ubicaciones"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000005&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmd_despachos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Despachos"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000005&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmd_Consultas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Consultas"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000005&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmd_inicio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Inicio"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000005&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   12480
      Top             =   840
   End
   Begin VB.CommandButton cmd_reportes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Reportes"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000005&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   8520
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crea una copia de seguridad en la carpeta por defecto del Sistema con todos los datos actuales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   25
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Regresa el sistema a un punto anterior, a través de una copia realizada posteriormente por un administrador de la base de datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   24
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione ésta opción si está teniendo anomalías a la hora de visualizar productos, reportes o cualquier dato"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9360
      TabIndex        =   23
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Borra todos los datos del sistema para volver a su estado original"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      TabIndex        =   22
      Top             =   6960
      Width           =   2895
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
      MouseIcon       =   "MantenimientoForm.frx":2D4A2
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   1560
      Width           =   675
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
      MouseIcon       =   "MantenimientoForm.frx":2D7AC
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   1560
      Width           =   1605
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
      TabIndex        =   19
      Top             =   1560
      Width           =   45
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
      TabIndex        =   18
      Top             =   720
      Width           =   2610
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
      TabIndex        =   17
      Top             =   600
      Width           =   1365
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
      TabIndex        =   16
      Top             =   1080
      Width           =   825
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
      TabIndex        =   15
      Top             =   600
      Width           =   2460
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GESTIÓN DE DATOS DEL SISTEMA"
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
      TabIndex        =   14
      Top             =   2640
      Width           =   7290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Por favor, seleccione la opción que desea realizar:"
      ForeColor       =   &H80000011&
      Height          =   270
      Left            =   3480
      TabIndex        =   13
      Top             =   3480
      Width           =   5250
   End
End
Attribute VB_Name = "MantenimientoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_CompactarReparar_Click()
CompactarBD
End Sub

Private Sub cmd_CrearBackup_Click()

If MsgBox("Esta a punto de crear una copia de seguridad del sistema con todo el contenido actual (Usuarios, Productos, Reportes y Despachos), ¿Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
     directorio1 = App.Path & "\BDProyecto.mdb" 'mi archivo a respaldar
     directorio2 = App.Path & "\Backup\Copia de Seguridad [" & Day(Now) & "-" & Month(Now) & "-" & Year(Now) & "].mdb" 'la ubicacion elegida por el usuario para respaldar

     Set fs = CreateObject("scripting.filesystemobject")
     fs.copyfile directorio1, directorio2, (True) 'sobreescribir si es necesario
     
     Logs
     With RsLogs
     If .State = 1 Then .Close
               .Open "SELECT * FROM Logs"
               .Requery
               .AddNew
                         !User = vUsername
                         !Accion = "creó una copia de seguridad de la base de datos del sistema"
                         !Fecha = Date + Time
               .Update
               .Requery
               .Close
     End With
     
     LogsMantenimiento
     With RsLogsMantenimiento
     If .State = 1 Then .Close
               .Open "SELECT * FROM LogsMantenimiento"
               .Requery
               .AddNew
                         !User = vUsername
                         !Accion = "creó una copia de seguridad de la base de datos del sistema"
                         !Fecha = Date + Time
               .Update
               .Requery
               .Close
     End With
     
     MsgBox ("Copia de Seguridad creado con éxito en ->" & directorio2), vbInformation, "Aviso": Exit Sub

Else
     MsgBox ("Proceso cancelado"), vbInformation, "Aviso": Exit Sub
End If

End Sub

Private Sub cmd_Logs_Click()
LogsForm.Show
End Sub

Private Sub cmd_Restablecer_Click()
If MsgBox("ADVERTENCIA: Esta a punto de restablecer el sistema a su versión inicial (Sin datos). Todos los datos de los usuarios, reportes, productos,  despachos, devoluciones y Logs se borrarán. ¿Desea continuar?", vbCritical + vbYesNo, "Precaución") = vbYes Then
          Confirmacion = InputBox("Para realizar esta operación, ingrese su contaseña nuevamente", "Confirmación de seguridad", "******")
          With RsUsuarios
                    .Requery
                    .Find "Nomb_Usuario='" & Trim(vUsername) & "'"
                    If !Contraseña = Trim(Confirmacion) Then
                              CierrroTodo
                              With RsUsuarios
                                        If .State = 1 Then .Close
                                        adm = "Administrador"
                                        .Open "delete * from usuarios where Nomb_Usuario NOT LIKE '" & adm & "'"
                              End With
                              With RsProductos
                                        If .State = 1 Then .Close
                                        .Open "delete * from productos"
                              End With
                              With RsDevoluciones
                                        If .State = 1 Then .Close
                                        .Open "delete * from Devoluciones"
                              End With
                              With RsDespacho
                                        If .State = 1 Then .Close
                                        .Open "delete * from despacho"
                              End With
                              With RsDetallesDespacho
                                        If .State = 1 Then .Close
                                        .Open "delete * from DetallesDespacho"
                              End With
                              With RsDevolucionesDetalles
                                        If .State = 1 Then .Close
                                        .Open "delete * from DevolucionesDetalles"
                              End With
                              With RsTemporalDespacho
                                        If .State = 1 Then .Close
                                        .Open "delete * from TemporalDespacho"
                              End With
                              With RsTemporalDevoluciones
                                        If .State = 1 Then .Close
                                        .Open "delete * from TemporalDevoluciones"
                              End With
                              Logs
                              With RsLogs
                                        If .State = 1 Then .Close
                                        .Open "delete * from Logs"
                              End With
                              
                              LogsDespachos
                              With RsLogsDespachos
                                        If .State = 1 Then .Close
                                        .Open "delete * from LogsDespachos"
                              End With
                              
                              LogsMantenimiento
                              With RsLogsMantenimiento
                                        If .State = 1 Then .Close
                                        .Open "delete * from LogsMantenimiento"
                              End With
                              
                              LogsProductos
                              With RsLogsProductos
                                        If .State = 1 Then .Close
                                        .Open "delete * from LogsProductos"
                              End With
                              
                              LogsReportes
                              With RsLogsReportes
                                        If .State = 1 Then .Close
                                        .Open "delete * from LogsReportes"
                              End With

                              LogsUsuarios
                              With RsLogsUsuarios
                                        If .State = 1 Then .Close
                                        .Open "delete * from LogsUsuarios"
                              End With
                              
                              Guilty = Trim(vUsername)
                              vRestablecer = 1
                              vTarea = ""
                              vUsername = ""
                              Unload MantenimientoForm
                               With Base
                                        If .State = 1 Then .Close
                                        .Open
                              End With
                              
                              Load LoginForm
                              LoginForm.Show
                              


                              ElseIf Confirmacion = "" Then
                                        MsgBox ("Proceso cancelado"), vbInformation, "Aviso": Exit Sub
                              Else
                              
                                        MsgBox ("Contraseña invalida"), vbInformation, "Aviso": Exit Sub
                              End If
          End With
Else
          Exit Sub
End If
End Sub

Private Sub cmd_Restore_Click()
If MsgBox("Esta a punto de restaurar el sistema a un punto anterior, debe tener en cuenta que al realizar esta operación el sistema reiniciará para poder cargar los datos nuevamente, ¿Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
     Dialogo.DialogTitle = "Seleccione algun respaldo de la Base de Datos (Solo .MDB)"
     Dialogo.Filter = "Archivos MDB|*.mdb"
     Dialogo.ShowOpen 'para decirle al boton "abrir"
     RutaOrigenDB = Dialogo.FileName
     ArchivoNombreDB = Dialogo.FileTitle
     

 If Dialogo.FileName = "" Then MsgBox ("Canelado por el usuario"), vbInformation, "Aviso": Exit Sub 'SOLUCION A UN ERROR
 CierrroTodo
     directorio1 = RutaOrigenDB 'mi archivo a respaldar
     directorio2 = App.Path & "\BDProyecto.mdb" 'la ubicacion elegida por el usuario para respaldar

     Set fs = CreateObject("scripting.filesystemobject")
     fs.copyfile directorio1, directorio2, (True) 'sobreescribir si es necesario ME DIO ERROR AQUI
    Unload MantenimientoForm
    
        With Base
          If .State = 1 Then .Close
          .Open
     End With
     
         Logs
     With RsLogs
     If .State = 1 Then .Close
               .Open "SELECT * FROM Logs ORDER BY Fecha DESC"
               .Requery
               .AddNew
                         !User = vUsername
                         !Accion = "restauró el sistema a la base de datos: " & ArchivoNombreDB
                         !Fecha = Date + Time
               .Update
               .Requery
               .Close
     End With

     LogsMantenimiento
     With RsLogsMantenimiento
     If .State = 1 Then .Close
               .Open "SELECT * FROM LogsMantenimiento"
               .Requery
               .AddNew
                         !User = vUsername
                         !Accion = "restauró el sistema a la base de datos: " & ArchivoNombreDB
                         !Fecha = Date + Time
               .Update
               .Requery
               .Close
     End With


    vTarea = ""
    vUsername = ""


     Load LoginForm
     LoginForm.Show
     Else
          Exit Sub
End If

RutaOrigenDB = ""
RutaDestinoDB = ""
ArchivoNombreDB = ""

End Sub

Private Sub Form_Load()
lbl_tarea.Caption = vTarea
lbl_username.Caption = vUsername

DetallesDespacho
DevolucionesDetalles
TemporalDespacho
TemporalDevoluciones
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Val(vRestablecer) = 1 Then
          Logs
          With RsLogs
                    .Requery
                    .AddNew
                              !User = Guilty
                              !Accion = "restableció el sistema a su estado original (Sin datos)"
                              !Fecha = Date + Time
                    .Update
                    .Requery
                    .Close
          End With
          LogsMantenimiento
          With RsLogsMantenimiento
                    .Requery
                    .AddNew
                              !User = Guilty
                              !Accion = "restableció el sistema a su estado original (Sin datos)"
                              !Fecha = Date + Time
                    .Update
                    .Requery
                    .Close
          End With
          Guilty = ""
          vRestablecer = 0
Else
          'no hago nada
End If
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

Private Sub cmd_Consultas_Click()
MantenimientoForm.Hide
ConsultasForm.Show
End Sub

Private Sub cmd_despachos_Click()
MantenimientoForm.Hide
DespachosForm.Show
End Sub

Private Sub cmd_inicio_Click()
MantenimientoForm.Hide
IndexForm.Show
End Sub

Private Sub cmd_productos_Click()
MantenimientoForm.Hide
ProductosForm.Show
End Sub

Private Sub cmd_reportes_Click()
MantenimientoForm.Hide
ReportesForm.Show
End Sub

Private Sub cmd_ubicaciones_Click()
MantenimientoForm.Hide
UbicacionesForm.Show
End Sub

Private Sub cmd_usuarios_Click()
MantenimientoForm.Hide
UsuariosForm.Show
End Sub

Private Function CompactarBD()

'para indicar que esta cargando algo (Mouse en reloj de arena)
    Screen.MousePointer = 11
        
    On Error Resume Next
    
    Dim BaseDeDatos As String
    Dim BaseDeDatosCo As String
    
    
    BaseDeDatos = App.Path & "\BDProyecto.mdb" 'la direccion de la base de datos original
    BaseDeDatosCo = Mid$(BaseDeDatos, 1, Len(BaseDeDatos) - 4) & App.Path & "\Backup\Copia de Seguridad [" & Day(Now) & "-" & Month(Now) & "-" & Year(Now) & "].mdb" 'la direccion que tendra la copia
          
    'hago una copia de seguridad de la base de datos antes de compactar
    
    FileCopy BaseDeDatos, DameDirectorioAplicacion & "~bdatos.mdb"
    
    ' Me aseguro que no existe un archivo con el
    ' nombre de la base de datos compactada (de algún error anterior).
    If Dir(BaseDeDatosCo) <> "" Then _
        Kill BaseDeDatosCo
    
    ' Esta instrucción crea una versión compactada de la base de datos
    DoEvents
    DBEngine.CompactDatabase BaseDeDatos, _
        BaseDeDatosCo, dbLangGeneral
    
    'si nuestra bd tiene contraseña se haría con esta instrucción:
    
'    DBEngine.CompactDatabase BaseDeDatos, _
'        BaseDeDatosCo, dbLangSpanish & ";pwd =" & gClave, , ";pwd =" & gClave
    'si tiene contraseña, hay que añadir ,pwd="contraseña"

    'elimino la base de datos y copio la compactada con el nombre bueno
    
    If Dir(BaseDeDatosCo) <> "" Then
         Kill BaseDeDatos
    End If
    

    DoEvents
    FileCopy BaseDeDatosCo, BaseDeDatos
    
    'elimino las copias de seguridad
    Kill BaseDeDatosCo
    Kill DameDirectorioAplicacion & "~bdatos.mdb"

    Screen.MousePointer = 0
    
    MsgBox ("Base de datos compactada")
    DoEvents
    
     Logs
     With RsLogs
     If .State = 1 Then .Close
               .Open "SELECT * FROM Logs"
               .Requery
               .AddNew
                         !User = vUsername
                         !Accion = "Compactó y reparó la Base de datos del sistema"
                         !Fecha = Date + Time
               .Update
               .Requery
               .Close
     End With
     
     LogsMantenimiento
     With RsLogsMantenimiento
     If .State = 1 Then .Close
               .Open "SELECT * FROM LogsMantenimiento"
               .Requery
               .AddNew
                         !User = vUsername
                         !Accion = "Compactó y reparó la Base de datos del sistema"
                         !Fecha = Date + Time
               .Update
               .Requery
               .Close
     End With
End Function

Sub CierrroTodo()
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
End Sub

