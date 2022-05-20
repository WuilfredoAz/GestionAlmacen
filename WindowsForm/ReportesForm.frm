VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ReportesForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes del Sistema"
   ClientHeight    =   10500
   ClientLeft      =   2505
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
   Icon            =   "ReportesForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "ReportesForm.frx":058A
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
      TabIndex        =   6
      Top             =   7320
      Width           =   1935
   End
   Begin VB.ComboBox cbo_DevolucionesAño 
      Height          =   390
      Left            =   12360
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   8160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox cbo_DevolucionesMes 
      Height          =   390
      Left            =   12360
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   7680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox cbo_DespachosAño 
      Height          =   390
      Left            =   9120
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   8160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox cbo_DespachoMes 
      Height          =   390
      Left            =   9120
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   7680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmd_DespachoFechaImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   9120
      TabIndex        =   27
      Top             =   8640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmd_DevolucionesFecha 
      Height          =   2200
      Left            =   9360
      Picture         =   "ReportesForm.frx":188DE
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7200
      Width           =   2895
   End
   Begin VB.CommandButton cmd_DespachosFecha 
      Height          =   2200
      Left            =   6120
      Picture         =   "ReportesForm.frx":1DFF8
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7200
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc AdoImpimirxSeccion 
      Height          =   375
      Left            =   5880
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_Impirmir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   5880
      TabIndex        =   26
      Top             =   8280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox cbo_Secciones 
      Height          =   390
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   7680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmd_InventarioxSecciones 
      Height          =   2200
      Left            =   2880
      Picture         =   "ReportesForm.frx":22DE5
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7200
      Width           =   2895
   End
   Begin VB.CommandButton cmd_ReporteInventario 
      Height          =   2200
      Left            =   2880
      Picture         =   "ReportesForm.frx":2798B
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton cmd_ReporteUsuarios 
      Height          =   2200
      Left            =   6120
      Picture         =   "ReportesForm.frx":2BCC5
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton cmd_ReporteDespachos 
      Height          =   2200
      Left            =   9360
      Picture         =   "ReportesForm.frx":30C2E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton cmd_Mantenimiento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Mantenimiento"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton cmd_reportes 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Reportes"
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000001&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmd_inicio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Inicio"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000001&
      Style           =   1  'Graphical
      TabIndex        =   0
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
      TabIndex        =   1
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
      BackColor       =   &H80000006&
      Caption         =   "Ubicaciones"
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   2
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   12480
      Top             =   840
   End
   Begin VB.CommandButton cmd_Devoluciones 
      Height          =   2200
      Left            =   12600
      Picture         =   "ReportesForm.frx":35861
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc AdoReporteInventario 
      Height          =   375
      Left            =   2880
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6960
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoDespachosXFecha 
      Height          =   375
      Left            =   9120
      Top             =   9240
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoDevolucionesXFecha 
      Height          =   375
      Left            =   12360
      Top             =   9240
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_DevolucionesFechaImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   12360
      TabIndex        =   30
      Top             =   8640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc AdoReporteUsuarios 
      Height          =   375
      Left            =   6120
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crea un documento listo para imprimir con todas las devoluciones realizadas en un determinado período."
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
      TabIndex        =   39
      Top             =   9600
      Width           =   2895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crea un documento listo para imprimir con todos los despachos realizados en un determinado período."
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
      TabIndex        =   38
      Top             =   9600
      Width           =   2895
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crea un documento listo para imprimir con los productos ubicados en un segmento determinado."
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
      TabIndex        =   37
      Top             =   9600
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crea un documento listo para imprimir donde se muestra la lista de todas las devoluciones realizadas en la empresa."
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
      Left            =   12600
      TabIndex        =   36
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crea un documento listo para imprimir donde se muestra una lista con todos los despachos que se han realizado en la empresa."
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
      TabIndex        =   35
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crea un documento listo para imprimir donde se muestran los datos de todos los usuarios registrados en el sistema."
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
      TabIndex        =   34
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"ReportesForm.frx":3A0FD
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
      TabIndex        =   33
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REPORTES DEL SISTEMA"
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
      TabIndex        =   24
      Top             =   2640
      Width           =   5385
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Por favor, seleccione el tipo de reporte que desea generar:"
      ForeColor       =   &H80000011&
      Height          =   270
      Left            =   3480
      TabIndex        =   23
      Top             =   3240
      Width           =   6120
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      MouseIcon       =   "ReportesForm.frx":3A18C
      MousePointer    =   99  'Custom
      TabIndex        =   17
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
      MouseIcon       =   "ReportesForm.frx":3A496
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   1560
      Width           =   675
   End
End
Attribute VB_Name = "ReportesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MesDespacho As String
Dim AñoDespacho As String
Dim MesDevolucion As String
Dim AñoDevolucion As String


Private Sub cmd_DespachoFechaImprimir_Click()
'Ambos campos en seleccione
If cbo_DespachoMes.Text = "Seleccione" And cbo_DespachosAño.Text = "Seleccione" Then MsgBox ("Por favor, especifique el período de tiempo que quiere consultar (Mes y Año)"), vbInformation, "Aviso": cbo_DespachoMes.SetFocus: Exit Sub

'Campo de mes en seleccione
If cbo_DespachoMes.Text = "Seleccione" Then MsgBox ("Por favor, especifique el mes que quiere consultar"), vbInformation, "Aviso": cbo_DespachoMes.SetFocus: Exit Sub

'Campo de año en seleccione
If cbo_DespachosAño.Text = "Seleccione" Then MsgBox ("Por favor, especifique el año que quiere consultar"), vbInformation, "Aviso": cbo_DespachosAño.SetFocus: Exit Sub

If cbo_DespachoMes.Text = "Enero" Then MesDespacho = "/01/"
If cbo_DespachoMes.Text = "Febrero" Then MesDespacho = "/02/"
If cbo_DespachoMes.Text = "Marzo" Then MesDespacho = "/03/"
If cbo_DespachoMes.Text = "Abril" Then MesDespacho = "/04/"
If cbo_DespachoMes.Text = "Mayo" Then MesDespacho = "/05/"
If cbo_DespachoMes.Text = "Junio" Then MesDespacho = "/06/"
If cbo_DespachoMes.Text = "Julio" Then MesDespacho = "/07/"
If cbo_DespachoMes.Text = "Agosto" Then MesDespacho = "/08/"
If cbo_DespachoMes.Text = "Septiembre" Then MesDespacho = "/09/"
If cbo_DespachoMes.Text = "Octubre" Then MesDespacho = "/10/"
If cbo_DespachoMes.Text = "Noviembre" Then MesDespacho = "/11/"
If cbo_DespachoMes.Text = "Diciembre" Then MesDespacho = "/12/"

If cbo_DespachosAño.Text = "2016" Then AñoDespacho = "2016"
If cbo_DespachosAño.Text = "2017" Then AñoDespacho = "2017"
If cbo_DespachosAño.Text = "2018" Then AñoDespacho = "2018"
If cbo_DespachosAño.Text = "2019" Then AñoDespacho = "2019"

With AdoDespachosXFecha
     .CursorLocation = adUseClient
     .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
     .RecordSource = "SELECT * FROM Despacho WHERE Fecha LIKE '%" & MesDespacho & AñoDespacho & "%'"
     .Refresh
     Set dr_RDespachosxFecha.DataSource = AdoDespachosXFecha
     
     'seccion4
     dr_RDespachosxFecha.Caption = "Reporte de Despacho de " & cbo_DespachoMes.Text & " del " & cbo_DespachosAño.Text
     dr_RDespachosxFecha.Sections("Sección4").Controls("Etiqueta1").Caption = "Reporte de Despachos realizados en el mes de " & cbo_DespachoMes.Text & " del " & cbo_DespachosAño.Text
     dr_RDespachosxFecha.WindowState = 2
     dr_RDespachosxFecha.Show
End With

cmd_DespachoFechaImprimir.Visible = False
cbo_DespachosAño.Visible = False
cbo_DespachoMes.Visible = False
cmd_DespachosFecha.Left = 6120
Label9.Left = 6120
cmd_DevolucionesFecha.Left = 9360
Label10.Left = 9360

Logs
With RsLogs
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir el reporte de despachos de " & cbo_DespachoMes.Text & " del " & cbo_DespachosAño.Text
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With

LogsReportes
With RsLogsReportes
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir el reporte de despachos de " & cbo_DespachoMes.Text & " del " & cbo_DespachosAño.Text
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With

End Sub

Private Sub cmd_DespachosFecha_Click()

     cbo_Secciones.Visible = False
          cbo_Secciones.ListIndex = 0
     cmd_Impirmir.Visible = False
     
cmd_DevolucionesFecha.Left = 12600
     cbo_DevolucionesMes.Visible = False
          cbo_DevolucionesMes.ListIndex = 0
     cbo_DevolucionesAño.Visible = False
          cbo_DevolucionesAño.ListIndex = 0
     cmd_DevolucionesFechaImprimir.Visible = False
     Label10.Left = 12600
     
cmd_DespachosFecha.Left = 6120
     cbo_DespachoMes.Visible = True
          cbo_DespachoMes.ListIndex = 0
     cbo_DespachosAño.Visible = True
          cbo_DespachosAño.ListIndex = 0
     cmd_DespachoFechaImprimir.Visible = True
     Label9.Left = 6120

End Sub

Private Sub cmd_Devoluciones_Click()
     cbo_Secciones.Visible = False
          cbo_Secciones.ListIndex = 0
     cmd_Impirmir.Visible = False

cmd_DespachosFecha.Left = 6120
     cbo_DespachoMes.Visible = False
          cbo_DespachoMes.ListIndex = 0
     cbo_DespachosAño.Visible = False
          cbo_DespachosAño.ListIndex = 0
     cmd_DespachoFechaImprimir.Visible = False
     Label9.Left = 6120
     
cmd_DevolucionesFecha.Left = 9360
     cbo_DevolucionesMes.Visible = False
          cbo_DevolucionesMes.ListIndex = 0
     cbo_DevolucionesAño.Visible = False
          cbo_DevolucionesAño.ListIndex = 0
     cmd_DevolucionesFechaImprimir.Visible = False
     Label10.Left = 9360
     
With RsDevoluciones
          If .State = 1 Then .Close
          .Open "select * from Devoluciones order by FechaDevo asc"
          Set dr_RDevoluciones.DataSource = RsDevoluciones
          dr_RDevoluciones.Orientation = rptOrientLandscape
          dr_RDevoluciones.WindowState = 2
          dr_RDevoluciones.Show
End With

Logs
With RsLogs
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir el reporte de todas las devoluciones"
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With

LogsReportes
With RsLogsReportes
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir el reporte de todas las devoluciones"
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With
End Sub

Private Sub cmd_DevolucionesFecha_Click()
     cbo_Secciones.Visible = False
          cbo_Secciones.ListIndex = 0
     cmd_Impirmir.Visible = False

cmd_DespachosFecha.Left = 6120
     cbo_DespachoMes.Visible = False
          cbo_DespachoMes.ListIndex = 0
     cbo_DespachosAño.Visible = False
          cbo_DespachosAño.ListIndex = 0
     cmd_DespachoFechaImprimir.Visible = False
     Label9.Left = 6120

cmd_DevolucionesFecha.Left = 9360
     cbo_DevolucionesMes.Visible = True
          cbo_DevolucionesMes.ListIndex = 0
     cbo_DevolucionesAño.Visible = True
          cbo_DevolucionesAño.ListIndex = 0
     cmd_DevolucionesFechaImprimir.Visible = True
     Label10.Left = 9360
End Sub

Private Sub cmd_DevolucionesFechaImprimir_Click()
'Ambos campos en seleccione
If cbo_DevolucionesMes.Text = "Seleccione" And cbo_DevolucionesAño.Text = "Seleccione" Then MsgBox ("Por favor, especifique el período de tiempo que quiere consultar (Mes y Año)"), vbInformation, "Aviso": cbo_DevolucionesMes.SetFocus: Exit Sub

'Campo de mes en seleccione
If cbo_DevolucionesMes.Text = "Seleccione" Then MsgBox ("Por favor, especifique el mes que quiere consultar"), vbInformation, "Aviso": cbo_DevolucionesMes.SetFocus: Exit Sub

'Campo de año en seleccione
If cbo_DevolucionesAño.Text = "Seleccione" Then MsgBox ("Por favor, especifique el año que quiere consultar"), vbInformation, "Aviso": cbo_DevolucionesAño.SetFocus: Exit Sub

If cbo_DevolucionesMes.Text = "Enero" Then MesDevolucion = "/01/"
If cbo_DevolucionesMes.Text = "Febrero" Then MesDevolucion = "/02/"
If cbo_DevolucionesMes.Text = "Marzo" Then MesDevolucion = "/03/"
If cbo_DevolucionesMes.Text = "Abril" Then MesDevolucion = "/04/"
If cbo_DevolucionesMes.Text = "Mayo" Then MesDevolucion = "/05/"
If cbo_DevolucionesMes.Text = "Junio" Then MesDevolucion = "/06/"
If cbo_DevolucionesMes.Text = "Julio" Then MesDevolucion = "/07/"
If cbo_DevolucionesMes.Text = "Agosto" Then MesDevolucion = "/08/"
If cbo_DevolucionesMes.Text = "Septiembre" Then MesDevolucion = "/09/"
If cbo_DevolucionesMes.Text = "Octubre" Then MesDevolucion = "/10/"
If cbo_DevolucionesMes.Text = "Noviembre" Then MesDevolucion = "/11/"
If cbo_DevolucionesMes.Text = "Diciembre" Then MesDevolucion = "/12/"

If cbo_DevolucionesAño.Text = "2016" Then AñoDevolucion = "2016"
If cbo_DevolucionesAño.Text = "2017" Then AñoDevolucion = "2017"
If cbo_DevolucionesAño.Text = "2018" Then AñoDevolucion = "2018"
If cbo_DevolucionesAño.Text = "2019" Then AñoDevolucion = "2019"

With AdoDevolucionesXFecha
     .CursorLocation = adUseClient
     .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
     .RecordSource = "SELECT * FROM Devoluciones WHERE FechaDevo LIKE '%" & MesDevolucion & AñoDevolucion & "%'"
     .Refresh
     Set dr_RDevolucionesxFecha.DataSource = AdoDevolucionesXFecha
     
     'seccion 4
     dr_RDevolucionesxFecha.Caption = "Reporte de Devoluciones realizadas en el mes de " & cbo_DevolucionesMes.Text & " del " & cbo_DevolucionesAño.Text
     dr_RDevolucionesxFecha.Sections("Sección4").Controls("Etiqueta1").Caption = "Reporte de Devoluciones realizadas en el mes de " & cbo_DevolucionesMes.Text & " del " & cbo_DevolucionesAño.Text
     dr_RDevolucionesxFecha.Orientation = rptOrientLandscape
     dr_RDevolucionesxFecha.WindowState = 2
     dr_RDevolucionesxFecha.Show
End With

cmd_DevolucionesFechaImprimir.Visible = False
cbo_DevolucionesAño.Visible = False
cbo_DevolucionesMes.Visible = False
cmd_DespachosFecha.Left = 6120
Label9.Left = 6120
cmd_DevolucionesFecha.Left = 9360
Label10.Left = 9360

Logs
With RsLogs
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir el reporte de devoluciones de " & cbo_DevolucionesMes.Text & " del " & cbo_DevolucionesAño.Text
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With

LogsReportes
With RsLogsReportes
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir el reporte de devoluciones de " & cbo_DevolucionesMes.Text & " del " & cbo_DevolucionesAño.Text
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With
End Sub

Private Sub cmd_Impirmir_Click()
With AdoImpimirxSeccion
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
          Busca = UCase(Trim(cbo_Secciones.Text))
          .RecordSource = "SELECT * FROM Productos WHERE Ubicacion LIKE '" & Busca & "'"
          .Refresh
          Set dr_RProductosSecciones.DataSource = AdoImpimirxSeccion
End With

cmd_Impirmir.Visible = False
cbo_Secciones.Visible = False
cmd_DespachosFecha.Left = 6120
Label9.Left = 6120
cmd_DevolucionesFecha.Left = 9360
Label10.Left = 9360

'Seccion 4
dr_RProductosSecciones.Sections("Sección4").Controls("Etiqueta1").Caption = "Reporte de productos " & cbo_Secciones.Text

dr_RProductosSecciones.WindowState = 2
dr_RProductosSecciones.Show

Logs
With RsLogs
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir el reporte de inventario del " & cbo_Secciones.Text
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With

LogsReportes
With RsLogsReportes
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir el reporte de inventario del " & cbo_Secciones.Text
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With

End Sub

Private Sub cmd_InventarioxSecciones_Click()
     cbo_Secciones.Visible = True
          cbo_Secciones.ListIndex = 0
     cmd_Impirmir.Visible = True

cmd_DespachosFecha.Left = 9480
     cbo_DespachoMes.Visible = False
          cbo_DespachoMes.ListIndex = 0
     cbo_DespachosAño.Visible = False
          cbo_DespachosAño.ListIndex = 0
     cmd_DespachoFechaImprimir.Visible = False
     Label9.Left = 9480
     
cmd_DevolucionesFecha.Left = 12600
     cbo_DevolucionesMes.Visible = False
          cbo_DevolucionesMes.ListIndex = 0
     cbo_DevolucionesAño.Visible = False
          cbo_DevolucionesAño.ListIndex = 0
     cmd_DevolucionesFechaImprimir.Visible = False
     Label10.Left = 12600
End Sub

Private Sub cmd_Logs_Click()
LogsForm.Show
End Sub

Private Sub cmd_Mantenimiento_Click()
     cbo_Secciones.Visible = False
          cbo_Secciones.ListIndex = 0
     cmd_Impirmir.Visible = False

cmd_DespachosFecha.Left = 6120
     cbo_DespachoMes.Visible = False
          cbo_DespachoMes.ListIndex = 0
     cbo_DespachosAño.Visible = False
          cbo_DespachosAño.ListIndex = 0
     cmd_DespachoFechaImprimir.Visible = False
     Label9.Left = 6120
     
cmd_DevolucionesFecha.Left = 9360
     cbo_DevolucionesMes.Visible = False
          cbo_DevolucionesMes.ListIndex = 0
     cbo_DevolucionesAño.Visible = False
          cbo_DevolucionesAño.ListIndex = 0
     cmd_DevolucionesFechaImprimir.Visible = False
     Label10.Left = 9360
     
ReportesForm.Hide
MantenimientoForm.Show
End Sub

Private Sub cmd_ReporteDespachos_Click()
     cbo_Secciones.Visible = False
          cbo_Secciones.ListIndex = 0
     cmd_Impirmir.Visible = False

cmd_DespachosFecha.Left = 6120
     cbo_DespachoMes.Visible = False
          cbo_DespachoMes.ListIndex = 0
     cbo_DespachosAño.Visible = False
          cbo_DespachosAño.ListIndex = 0
     cmd_DespachoFechaImprimir.Visible = False
     Label9.Left = 6120
     
cmd_DevolucionesFecha.Left = 9360
     cbo_DevolucionesMes.Visible = False
          cbo_DevolucionesMes.ListIndex = 0
     cbo_DevolucionesAño.Visible = False
          cbo_DevolucionesAño.ListIndex = 0
     cmd_DevolucionesFechaImprimir.Visible = False
     Label10.Left = 9360
     
With RsDespacho
        If .State = 1 Then .Close
        .Open "select * from Despacho order by Fecha asc"
        Set dr_RDespachos.DataSource = RsDespacho
        dr_RDespachos.WindowState = 2
        dr_RDespachos.Show
End With

Logs
With RsLogs
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir el reporte de todos los despachos"
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With
LogsReportes
With RsLogsReportes
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir el reporte de todos los despachos"
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With
End Sub

Private Sub cmd_ReporteInventario_Click()
     cbo_Secciones.Visible = False
          cbo_Secciones.ListIndex = 0
     cmd_Impirmir.Visible = False

cmd_DespachosFecha.Left = 6120
     cbo_DespachoMes.Visible = False
          cbo_DespachoMes.ListIndex = 0
     cbo_DespachosAño.Visible = False
          cbo_DespachosAño.ListIndex = 0
     cmd_DespachoFechaImprimir.Visible = False
     Label9.Left = 6120
     
     
cmd_DevolucionesFecha.Left = 9360
     cbo_DevolucionesMes.Visible = False
          cbo_DevolucionesMes.ListIndex = 0
     cbo_DevolucionesAño.Visible = False
          cbo_DevolucionesAño.ListIndex = 0
     cmd_DevolucionesFechaImprimir.Visible = False
     Label10.Left = 9360

With AdoReporteInventario
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
        '  Busca = ">0"
          .RecordSource = "SELECT * FROM Productos WHERE Cantidad > 0 ORDER BY Descripcion ASC"
          .Refresh
        Set dr_RProductos.DataSource = AdoReporteInventario
        dr_RProductos.WindowState = 2
        dr_RProductos.Show
End With

Logs
With RsLogs
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir el reporte de inventario correspondiente a la fecha de"
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With

LogsReportes
With RsLogsReportes
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir el reporte de inventario correspondiente a la fecha de"
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With

End Sub

Private Sub cmd_ReporteUsuarios_Click()
     cbo_Secciones.Visible = False
          cbo_Secciones.ListIndex = 0
     cmd_Impirmir.Visible = False

cmd_DespachosFecha.Left = 6120
     cbo_DespachoMes.Visible = False
          cbo_DespachoMes.ListIndex = 0
     cbo_DespachosAño.Visible = False
          cbo_DespachosAño.ListIndex = 0
     cmd_DespachoFechaImprimir.Visible = False
     Label9.Left = 6120
     
cmd_DevolucionesFecha.Left = 9360
     cbo_DevolucionesMes.Visible = False
          cbo_DevolucionesMes.ListIndex = 0
     cbo_DevolucionesAño.Visible = False
          cbo_DevolucionesAño.ListIndex = 0
     cmd_DevolucionesFechaImprimir.Visible = False
     Label10.Left = 9360
     
'codigo para mostrar todos los usuarios menos el admin
With AdoReporteUsuarios
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
          .RecordSource = "SELECT * FROM Usuarios WHERE Nomb_Usuario NOT IN ('Administrador') ORDER BY Nombre ASC"
          .Refresh
          Set dr_RUsuarios.DataSource = AdoReporteUsuarios
          dr_RUsuarios.WindowState = 2
          dr_RUsuarios.Show
End With

'Codigo para mostrar todos los usuarios
'Set dr_RUsuarios.DataSource = RsUsuarios
'dr_RUsuarios.WindowState = 2
'dr_RUsuarios.Show

Logs
With RsLogs
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir el reporte de usuarios correspondiente a la fecha de"
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With
LogsReportes
With RsLogsReportes
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir el reporte de usuarios correspondiente a la fecha de"
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With
End Sub

Private Sub Form_Load()
lbl_tarea.Caption = vTarea
lbl_username.Caption = vUsername
Despacho 'ME DECIA QUE ESTABA CERRADA O ALGO ASI QUE POR ESO CARGO ESTO ACA
Devoluciones


cbo_Secciones.AddItem "Segmento 1 (S1)"
cbo_Secciones.AddItem "Segmento 2 (S2)"
cbo_Secciones.AddItem "Segmento 3 (S3)"
cbo_Secciones.AddItem "Segmento 4 (S4)"
cbo_Secciones.AddItem "Segmento 5 (S5)"
cbo_Secciones.AddItem "Segmento 6 (S6)"
cbo_Secciones.AddItem "Segmento 7 (S7)"
cbo_Secciones.AddItem "Segmento 8 (S8)"
cbo_Secciones.AddItem "Segmento 9 (S9)"
cbo_Secciones.AddItem "Segmento 10 (S10)"
cbo_Secciones.AddItem "Segmento 11 (S11)"
cbo_Secciones.AddItem "Segmento 12 (S12)"
cbo_Secciones.AddItem "Segmento 13 (S13)"
cbo_Secciones.AddItem "Segmento 14 (S14)"
cbo_Secciones.AddItem "Segmento 15 (S15)"
cbo_Secciones.AddItem "Segmento 16 (S16)"
cbo_Secciones.AddItem "Segmento 17 (S17)"
cbo_Secciones.AddItem "Segmento 18 (S18)"
cbo_Secciones.AddItem "Segmento 19 (S19)"
cbo_Secciones.ListIndex = 0

cbo_DespachoMes.AddItem "Seleccione"
cbo_DespachoMes.AddItem "Enero"
cbo_DespachoMes.AddItem "Febrero"
cbo_DespachoMes.AddItem "Marzo"
cbo_DespachoMes.AddItem "Abril"
cbo_DespachoMes.AddItem "Mayo"
cbo_DespachoMes.AddItem "Junio"
cbo_DespachoMes.AddItem "Julio"
cbo_DespachoMes.AddItem "Agosto"
cbo_DespachoMes.AddItem "Septiembre"
cbo_DespachoMes.AddItem "Octubre"
cbo_DespachoMes.AddItem "Noviembre"
cbo_DespachoMes.AddItem "Diciembre"
cbo_DespachoMes.ListIndex = 0

cbo_DespachosAño.AddItem "Seleccione"
cbo_DespachosAño.AddItem "2016"
cbo_DespachosAño.AddItem "2017"
cbo_DespachosAño.AddItem "2018"
cbo_DespachosAño.AddItem "2019"
cbo_DespachosAño.ListIndex = 0

cbo_DevolucionesMes.AddItem "Seleccione"
cbo_DevolucionesMes.AddItem "Enero"
cbo_DevolucionesMes.AddItem "Febrero"
cbo_DevolucionesMes.AddItem "Marzo"
cbo_DevolucionesMes.AddItem "Abril"
cbo_DevolucionesMes.AddItem "Mayo"
cbo_DevolucionesMes.AddItem "Junio"
cbo_DevolucionesMes.AddItem "Julio"
cbo_DevolucionesMes.AddItem "Agosto"
cbo_DevolucionesMes.AddItem "Septiembre"
cbo_DevolucionesMes.AddItem "Octubre"
cbo_DevolucionesMes.AddItem "Noviembre"
cbo_DevolucionesMes.AddItem "Diciembre"
cbo_DevolucionesMes.ListIndex = 0

cbo_DevolucionesAño.AddItem "Seleccione"
cbo_DevolucionesAño.AddItem "2016"
cbo_DevolucionesAño.AddItem "2017"
cbo_DevolucionesAño.AddItem "2018"
cbo_DevolucionesAño.AddItem "2019"
cbo_DevolucionesAño.ListIndex = 0

cbo_Secciones.Visible = False
cmd_Impirmir.Visible = False
     
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
     cbo_Secciones.Visible = False
          cbo_Secciones.ListIndex = 0
     cmd_Impirmir.Visible = False

cmd_DespachosFecha.Left = 6120
     cbo_DespachoMes.Visible = False
          cbo_DespachoMes.ListIndex = 0
     cbo_DespachosAño.Visible = False
          cbo_DespachosAño.ListIndex = 0
     cmd_DespachoFechaImprimir.Visible = False
     Label9.Left = 6120
     
cmd_DevolucionesFecha.Left = 9360
     cbo_DevolucionesMes.Visible = False
          cbo_DevolucionesMes.ListIndex = 0
     cbo_DevolucionesAño.Visible = False
          cbo_DevolucionesAño.ListIndex = 0
     cmd_DevolucionesFechaImprimir.Visible = False
     Label10.Left = 9360
     
ReportesForm.Hide
ConsultasForm.Show
End Sub

Private Sub cmd_despachos_Click()
     cbo_Secciones.Visible = False
          cbo_Secciones.ListIndex = 0
     cmd_Impirmir.Visible = False

cmd_DespachosFecha.Left = 6120
     cbo_DespachoMes.Visible = False
          cbo_DespachoMes.ListIndex = 0
     cbo_DespachosAño.Visible = False
          cbo_DespachosAño.ListIndex = 0
     cmd_DespachoFechaImprimir.Visible = False
     Label9.Left = 6120
     
cmd_DevolucionesFecha.Left = 9360
     cbo_DevolucionesMes.Visible = False
          cbo_DevolucionesMes.ListIndex = 0
     cbo_DevolucionesAño.Visible = False
          cbo_DevolucionesAño.ListIndex = 0
     cmd_DevolucionesFechaImprimir.Visible = False
     Label10.Left = 9360
     
ReportesForm.Hide
DespachosForm.Show
End Sub

Private Sub cmd_inicio_Click()
     cbo_Secciones.Visible = False
          cbo_Secciones.ListIndex = 0
     cmd_Impirmir.Visible = False

cmd_DespachosFecha.Left = 6120
     cbo_DespachoMes.Visible = False
          cbo_DespachoMes.ListIndex = 0
     cbo_DespachosAño.Visible = False
          cbo_DespachosAño.ListIndex = 0
     cmd_DespachoFechaImprimir.Visible = False
     Label9.Left = 6120
     
cmd_DevolucionesFecha.Left = 9360
     cbo_DevolucionesMes.Visible = False
          cbo_DevolucionesMes.ListIndex = 0
     cbo_DevolucionesAño.Visible = False
          cbo_DevolucionesAño.ListIndex = 0
     cmd_DevolucionesFechaImprimir.Visible = False
     Label10.Left = 9360
     
ReportesForm.Hide
IndexForm.Show
End Sub

Private Sub cmd_productos_Click()
     cbo_Secciones.Visible = False
          cbo_Secciones.ListIndex = 0
     cmd_Impirmir.Visible = False

cmd_DespachosFecha.Left = 6120
     cbo_DespachoMes.Visible = False
          cbo_DespachoMes.ListIndex = 0
     cbo_DespachosAño.Visible = False
          cbo_DespachosAño.ListIndex = 0
     cmd_DespachoFechaImprimir.Visible = False
     Label9.Left = 6120
     
cmd_DevolucionesFecha.Left = 9360
     cbo_DevolucionesMes.Visible = False
          cbo_DevolucionesMes.ListIndex = 0
     cbo_DevolucionesAño.Visible = False
          cbo_DevolucionesAño.ListIndex = 0
     cmd_DevolucionesFechaImprimir.Visible = False
     Label10.Left = 9360
     
ReportesForm.Hide
ProductosForm.Show
End Sub

Private Sub cmd_ubicaciones_Click()
     cbo_Secciones.Visible = False
          cbo_Secciones.ListIndex = 0
     cmd_Impirmir.Visible = False

cmd_DespachosFecha.Left = 6120
     cbo_DespachoMes.Visible = False
          cbo_DespachoMes.ListIndex = 0
     cbo_DespachosAño.Visible = False
          cbo_DespachosAño.ListIndex = 0
     cmd_DespachoFechaImprimir.Visible = False
     Label9.Left = 6120
     
cmd_DevolucionesFecha.Left = 9360
     cbo_DevolucionesMes.Visible = False
          cbo_DevolucionesMes.ListIndex = 0
     cbo_DevolucionesAño.Visible = False
          cbo_DevolucionesAño.ListIndex = 0
     cmd_DevolucionesFechaImprimir.Visible = False
     Label10.Left = 9360
     
ReportesForm.Hide
UbicacionesForm.Show
End Sub

Private Sub cmd_usuarios_Click()
     cbo_Secciones.Visible = False
          cbo_Secciones.ListIndex = 0
     cmd_Impirmir.Visible = False

cmd_DespachosFecha.Left = 6120
     cbo_DespachoMes.Visible = False
          cbo_DespachoMes.ListIndex = 0
     cbo_DespachosAño.Visible = False
          cbo_DespachosAño.ListIndex = 0
     cmd_DespachoFechaImprimir.Visible = False
     Label9.Left = 6120
     
cmd_DevolucionesFecha.Left = 9360
     cbo_DevolucionesMes.Visible = False
          cbo_DevolucionesMes.ListIndex = 0
     cbo_DevolucionesAño.Visible = False
          cbo_DevolucionesAño.ListIndex = 0
     cmd_DevolucionesFechaImprimir.Visible = False
     Label10.Left = 9360
     
ReportesForm.Hide
UsuariosForm.Show
End Sub
