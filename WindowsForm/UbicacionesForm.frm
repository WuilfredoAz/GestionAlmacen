VERSION 5.00
Begin VB.Form UbicacionesForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ubicaciones"
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
   Icon            =   "UbicacionesForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "UbicacionesForm.frx":058A
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
      TabIndex        =   4
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   12480
      Top             =   840
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
      TabIndex        =   3
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmd_ubicaciones 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Ubicaciones"
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmd_usuarios 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Usuarios"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "FILTROS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3000
      TabIndex        =   43
      Top             =   9700
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "EMPACADURAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3000
      TabIndex        =   42
      Top             =   9360
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "CABLES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3000
      TabIndex        =   41
      Top             =   9050
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "PARTES DE MOTOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3000
      TabIndex        =   40
      Top             =   8760
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ARTÍCULOS MÁS VENDIDOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3000
      TabIndex        =   39
      Top             =   8445
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "BOMBAS DE AGUA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3000
      TabIndex        =   38
      Top             =   8100
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ARTÍCULOS ELÉCTRICOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3000
      TabIndex        =   37
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ARTÍCULOS PESADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3000
      TabIndex        =   36
      Top             =   7480
      Width           =   2175
   End
   Begin VB.Label lbl_S19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   440
      Left            =   6990
      MouseIcon       =   "UbicacionesForm.frx":273D7
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   5490
      Width           =   540
   End
   Begin VB.Label lbl_S18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   450
      Left            =   6990
      MouseIcon       =   "UbicacionesForm.frx":276E1
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   4920
      Width           =   540
   End
   Begin VB.Label lbl_S17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   630
      Left            =   8000
      MouseIcon       =   "UbicacionesForm.frx":279EB
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   4150
      Width           =   945
   End
   Begin VB.Label lbl_S16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   630
      Left            =   6960
      MouseIcon       =   "UbicacionesForm.frx":27CF5
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   4150
      Width           =   945
   End
   Begin VB.Label lbl_S15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   6360
      MouseIcon       =   "UbicacionesForm.frx":27FFF
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   3795
      Width           =   3950
   End
   Begin VB.Label lbl_S14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5040
      MouseIcon       =   "UbicacionesForm.frx":28309
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   4250
      Width           =   780
   End
   Begin VB.Label lbl_S13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5040
      MouseIcon       =   "UbicacionesForm.frx":28613
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   3800
      Width           =   780
   End
   Begin VB.Label lbl_S12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   4560
      MouseIcon       =   "UbicacionesForm.frx":2891D
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   3840
      Width           =   300
   End
   Begin VB.Label lbl_S11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   2170
      Left            =   6200
      MouseIcon       =   "UbicacionesForm.frx":28C27
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   4800
      Width           =   300
   End
   Begin VB.Label lbl_S10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   1245
      Left            =   8130
      MouseIcon       =   "UbicacionesForm.frx":28F31
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   8220
      Width           =   300
   End
   Begin VB.Label lbl_S9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   1275
      Left            =   8130
      MouseIcon       =   "UbicacionesForm.frx":2923B
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   6840
      Width           =   300
   End
   Begin VB.Label lbl_S8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   1275
      Left            =   8130
      MouseIcon       =   "UbicacionesForm.frx":29545
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   5460
      Width           =   300
   End
   Begin VB.Label lbl_S7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   1250
      Left            =   9960
      MouseIcon       =   "UbicacionesForm.frx":2984F
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   8220
      Width           =   300
   End
   Begin VB.Label lbl_S6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   1250
      Left            =   9960
      MouseIcon       =   "UbicacionesForm.frx":29B59
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   6840
      Width           =   300
   End
   Begin VB.Label lbl_S5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   1275
      Left            =   9960
      MouseIcon       =   "UbicacionesForm.frx":29E63
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   5450
      Width           =   300
   End
   Begin VB.Label lbl_S4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   1250
      Left            =   11745
      MouseIcon       =   "UbicacionesForm.frx":2A16D
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   8000
      Width           =   300
   End
   Begin VB.Label lbl_S3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   1275
      Left            =   11745
      MouseIcon       =   "UbicacionesForm.frx":2A477
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   6600
      Width           =   300
   End
   Begin VB.Label lbl_S2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   1275
      Left            =   11760
      MouseIcon       =   "UbicacionesForm.frx":2A781
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   5205
      Width           =   300
   End
   Begin VB.Label lbl_S1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   10670
      MouseIcon       =   "UbicacionesForm.frx":2AA8B
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   4805
      Width           =   1345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UBICACIONES"
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
      TabIndex        =   16
      Top             =   2640
      Width           =   2985
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      MouseIcon       =   "UbicacionesForm.frx":2AD95
      MousePointer    =   99  'Custom
      TabIndex        =   10
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
      MouseIcon       =   "UbicacionesForm.frx":2B09F
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   1560
      Width           =   675
   End
End
Attribute VB_Name = "UbicacionesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Consultas_Click()
UbicacionesForm.Hide
ConsultasForm.Show
End Sub
Private Sub cmd_despachos_Click()
UbicacionesForm.Hide
DespachosForm.Show
End Sub
Private Sub cmd_inicio_Click()
UbicacionesForm.Hide
IndexForm.Show
End Sub

Private Sub cmd_Logs_Click()
LogsForm.Show
End Sub

Private Sub cmd_Mantenimiento_Click()
UbicacionesForm.Hide
MantenimientoForm.Show
End Sub

Private Sub cmd_productos_Click()
UbicacionesForm.Hide
ProductosForm.Show
End Sub

Private Sub cmd_reportes_Click()
UbicacionesForm.Hide
ReportesForm.Show
End Sub

Private Sub cmd_usuarios_Click()
UbicacionesForm.Hide
UsuariosForm.Show
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

Private Sub lbl_S1_Click()
vSeccion = "Segmento 1 (S1)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S1.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 1 (S1)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S10_Click()
vSeccion = "Segmento 10 (S10)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S10.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 10 (S10)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S11_Click()
vSeccion = "Segmento 11 (S11)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S11.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 11 (S11)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S12_Click()
vSeccion = "Segmento 12 (S12)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S12.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 12 (S12)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S13_Click()
vSeccion = "Segmento 13 (S13)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S13.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 13 (S13)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S14_Click()
vSeccion = "Segmento 14 (S14)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S14.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 14 (S14)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S15_Click()
vSeccion = "Segmento 15 (S15)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S15.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 15 (S15)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S16_Click()
vSeccion = "Segmento 16 (S16)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S16.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 16 (16)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S17_Click()
vSeccion = "Segmento 17 (S17)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S17.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 17 (S17)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S18_Click()
vSeccion = "Segmento 18 (S18)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S18.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 18 (S18)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S19_Click()
vSeccion = "Segmento 19 (S19)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S19.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 19 (S19)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S2_Click()
vSeccion = "Segmento 2 (S2)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S2.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 2 (S2)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S3_Click()
vSeccion = "Segmento 3 (S3)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S3.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 3 (S3)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S4_Click()
vSeccion = "Segmento 4 (S4)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S4.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 4 (S4)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S5_Click()
vSeccion = "Segmento 5 (S5)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S5.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 5 (S5)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S6_Click()
vSeccion = "Segmento 6 (S6)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S6.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 6 (S6)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S7_Click()
vSeccion = "Segmento 7 (S7)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S7.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 7 (S7)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S8_Click()
vSeccion = "Segmento 8 (S8)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S8.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 8 (S8)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub lbl_S9_Click()
vSeccion = "Segmento 9 (S9)"
MostrarSeccion
UbicacionesSeccionesForm.Picture = LoadPicture(App.Path & "\Interfaz\S9.jpg")
UbicacionesSeccionesForm.Caption = "Lista de Productos del Segmento 9 (S9)"
UbicacionesSeccionesForm.Show vbModal
End Sub

Private Sub Timer1_Timer()
'Label1.Caption = Format(Time, "hh:mm:ss")
Label1.Caption = Format(Now, "HH:MM AM/PM")
End Sub

