VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form LogsForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logs del Sistema"
   ClientHeight    =   10845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13545
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LogsForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "LogsForm.frx":08CA
   ScaleHeight     =   10845
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoLogsProductos 
      Height          =   375
      Left            =   3600
      Top             =   4800
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.ComboBox cbo_AllLogsAño 
      Height          =   390
      Left            =   10080
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   7320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox cbo_AllLogsMes 
      Height          =   390
      Left            =   10080
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   6840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmd_VerAllLogs 
      Caption         =   "Ver"
      Height          =   495
      Left            =   10080
      TabIndex        =   30
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox cbo_LogsUsuariosAño 
      Height          =   390
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   7320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox cbo_LogsUsuariosMes 
      Height          =   390
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   6840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmd_VerLogsUsuarios 
      Caption         =   "Ver"
      Height          =   495
      Left            =   6720
      TabIndex        =   27
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox cbo_LogsDespachosAño 
      Height          =   390
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   7320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox cbo_LogsDespachosMes 
      Height          =   390
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   6840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmd_VerLogsDespachos 
      Caption         =   "Ver"
      Height          =   495
      Left            =   3600
      TabIndex        =   24
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox cbo_LogsMantenimientosAño 
      Height          =   390
      Left            =   10080
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   3720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox cbo_LogsMantenimientoMes 
      Height          =   390
      Left            =   10080
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmd_VerLogsMantenimiento 
      Caption         =   "Ver"
      Height          =   495
      Left            =   10080
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox cbo_LogsReportesAño 
      Height          =   390
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox cbo_LogsReportesMes 
      Height          =   390
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmd_VerLogsReportes 
      Caption         =   "Ver"
      Height          =   495
      Left            =   6720
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox cbo_LogsProductosAño 
      Height          =   390
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox cbo_LogsProductoMes 
      Height          =   390
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmd_VerLogsProductos 
      Caption         =   "Ver"
      Height          =   495
      Left            =   3600
      TabIndex        =   15
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmd_AllLogs 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2200
      Left            =   10080
      Picture         =   "LogsForm.frx":87CE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "LogsForm.frx":C546
      Top             =   8880
      Width           =   2895
   End
   Begin VB.CommandButton cmd_Atras 
      Caption         =   "Atras"
      Height          =   615
      Left            =   5640
      TabIndex        =   13
      Top             =   10080
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
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
      Left            =   5280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "LogsForm.frx":C5B2
      Top             =   8880
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "LogsForm.frx":C625
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "LogsForm.frx":C695
      Top             =   5280
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "LogsForm.frx":C730
      Top             =   8880
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "LogsForm.frx":C7C8
      Top             =   5280
      Width           =   2895
   End
   Begin VB.CommandButton cmd_LogsUsuarios 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2200
      Left            =   5280
      Picture         =   "LogsForm.frx":C84E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   2895
   End
   Begin VB.CommandButton cmd_LogsReportes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2200
      Left            =   5280
      Picture         =   "LogsForm.frx":103B0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton cmd_LogsMantenimiento 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2200
      Left            =   10080
      Picture         =   "LogsForm.frx":13A62
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton cmd_LogsDespachos 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2200
      Left            =   480
      Picture         =   "LogsForm.frx":1937D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   2895
   End
   Begin VB.CommandButton cmd_LogsProductos 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2200
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "LogsForm.frx":1CF6F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc AdoLogsReportes 
      Height          =   375
      Left            =   6720
      Top             =   4800
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc AdoLogsMantenimiento 
      Height          =   375
      Left            =   10080
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc AdoLogsDespachos 
      Height          =   375
      Left            =   3600
      Top             =   8400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc AdoLogsUsuarios 
      Height          =   375
      Left            =   6720
      Top             =   8400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc AdoTodosLogs 
      Height          =   375
      Left            =   10080
      Top             =   8400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.Label lbl_1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el tipo de Logs que quiere consultar"
      Height          =   270
      Left            =   1800
      TabIndex        =   7
      Top             =   2160
      Width           =   4935
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
      Left            =   1440
      TabIndex        =   6
      Top             =   2040
      Width           =   225
   End
End
Attribute VB_Name = "LogsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_AllLogs_Click()
cbo_LogsProductoMes.ListIndex = 0
cbo_LogsProductosAño.ListIndex = 0
cbo_LogsReportesMes.ListIndex = 0
cbo_LogsReportesAño.ListIndex = 0
cbo_LogsMantenimientoMes.ListIndex = 0
cbo_LogsMantenimientosAño.ListIndex = 0
cbo_LogsDespachosMes.ListIndex = 0
cbo_LogsDespachosAño.ListIndex = 0
cbo_LogsUsuariosMes.ListIndex = 0
cbo_LogsUsuariosAño.ListIndex = 0
cbo_AllLogsMes.ListIndex = 0
cbo_AllLogsAño.ListIndex = 0

cmd_LogsProductos.Left = 480
          Text1.Left = 480
          cbo_LogsProductoMes.Visible = False
          cbo_LogsProductosAño.Visible = False
          cmd_VerLogsProductos.Visible = False
          
cmd_LogsReportes.Left = 5280
          Text4.Left = 5280
          cbo_LogsReportesMes.Visible = False
          cbo_LogsReportesAño.Visible = False
          cmd_VerLogsReportes.Visible = False
          
cmd_LogsMantenimiento.Left = 10080
          Text3.Left = 10080
          cbo_LogsMantenimientoMes.Visible = False
          cbo_LogsMantenimientosAño.Visible = False
          cmd_VerLogsMantenimiento.Visible = False

cmd_LogsDespachos.Left = 480
          Text2.Left = 480
          cbo_LogsDespachosMes.Visible = False
          cbo_LogsDespachosAño.Visible = False
          cmd_VerLogsDespachos.Visible = False
          
cmd_LogsUsuarios.Left = 3720
          Text5.Left = 3720
          cbo_LogsUsuariosMes.Visible = False
          cbo_LogsUsuariosAño.Visible = False
          cmd_VerLogsUsuarios.Visible = False
          
cmd_AllLogs.Left = 6960
          Text6.Left = 6960
          cbo_AllLogsMes.Visible = True
          cbo_AllLogsAño.Visible = True
          cmd_VerAllLogs.Visible = True
End Sub

Private Sub cmd_Atras_Click()
Unload Me
End Sub

Private Sub cmd_ReporteInventario_Click()

End Sub

Private Sub cmd_LogsDespachos_Click()
cbo_LogsProductoMes.ListIndex = 0
cbo_LogsProductosAño.ListIndex = 0
cbo_LogsReportesMes.ListIndex = 0
cbo_LogsReportesAño.ListIndex = 0
cbo_LogsMantenimientoMes.ListIndex = 0
cbo_LogsMantenimientosAño.ListIndex = 0
cbo_LogsDespachosMes.ListIndex = 0
cbo_LogsDespachosAño.ListIndex = 0
cbo_LogsUsuariosMes.ListIndex = 0
cbo_LogsUsuariosAño.ListIndex = 0
cbo_AllLogsMes.ListIndex = 0
cbo_AllLogsAño.ListIndex = 0

cmd_LogsProductos.Left = 480
          Text1.Left = 480
          cbo_LogsProductoMes.Visible = False
          cbo_LogsProductosAño.Visible = False
          cmd_VerLogsProductos.Visible = False
          
cmd_LogsReportes.Left = 5280
          Text4.Left = 5280
          cbo_LogsReportesMes.Visible = False
          cbo_LogsReportesAño.Visible = False
          cmd_VerLogsReportes.Visible = False
          
cmd_LogsMantenimiento.Left = 10080
          Text3.Left = 10080
          cbo_LogsMantenimientoMes.Visible = False
          cbo_LogsMantenimientosAño.Visible = False
          cmd_VerLogsMantenimiento.Visible = False

cmd_LogsDespachos.Left = 480
          Text2.Left = 480
          cbo_LogsDespachosMes.Visible = True
          cbo_LogsDespachosAño.Visible = True
          cmd_VerLogsDespachos.Visible = True
          
cmd_LogsUsuarios.Left = 6840
          Text5.Left = 6840
          cbo_LogsUsuariosMes.Visible = False
          cbo_LogsUsuariosAño.Visible = False
          cmd_VerLogsUsuarios.Visible = False
          
cmd_AllLogs.Left = 10080
          Text6.Left = 10080
          cbo_AllLogsMes.Visible = False
          cbo_AllLogsAño.Visible = False
          cmd_VerAllLogs.Visible = False
End Sub

Private Sub cmd_LogsMantenimiento_Click()
cbo_LogsProductoMes.ListIndex = 0
cbo_LogsProductosAño.ListIndex = 0
cbo_LogsReportesMes.ListIndex = 0
cbo_LogsReportesAño.ListIndex = 0
cbo_LogsMantenimientoMes.ListIndex = 0
cbo_LogsMantenimientosAño.ListIndex = 0
cbo_LogsDespachosMes.ListIndex = 0
cbo_LogsDespachosAño.ListIndex = 0
cbo_LogsUsuariosMes.ListIndex = 0
cbo_LogsUsuariosAño.ListIndex = 0
cbo_AllLogsMes.ListIndex = 0
cbo_AllLogsAño.ListIndex = 0

cmd_LogsProductos.Left = 480
          Text1.Left = 480
          cbo_LogsProductoMes.Visible = False
          cbo_LogsProductosAño.Visible = False
          cmd_VerLogsProductos.Visible = False
          
cmd_LogsReportes.Left = 3720
          Text4.Left = 3720
          cbo_LogsReportesMes.Visible = False
          cbo_LogsReportesAño.Visible = False
          cmd_VerLogsReportes.Visible = False
          
cmd_LogsMantenimiento.Left = 6960
          Text3.Left = 6960
          cbo_LogsMantenimientoMes.Visible = True
          cbo_LogsMantenimientosAño.Visible = True
          cmd_VerLogsMantenimiento.Visible = True

cmd_LogsDespachos.Left = 480
          Text2.Left = 480
          cbo_LogsDespachosMes.Visible = False
          cbo_LogsDespachosAño.Visible = False
          cmd_VerLogsDespachos.Visible = False
          
cmd_LogsUsuarios.Left = 5280
          Text5.Left = 5280
          cbo_LogsUsuariosMes.Visible = False
          cbo_LogsUsuariosAño.Visible = False
          cmd_VerLogsUsuarios.Visible = False
          
cmd_AllLogs.Left = 10080
          Text6.Left = 10080
          cbo_AllLogsMes.Visible = False
          cbo_AllLogsAño.Visible = False
          cmd_VerAllLogs.Visible = False
End Sub

Private Sub cmd_LogsProductos_Click()
cbo_LogsProductoMes.ListIndex = 0
cbo_LogsProductosAño.ListIndex = 0
cbo_LogsReportesMes.ListIndex = 0
cbo_LogsReportesAño.ListIndex = 0
cbo_LogsMantenimientoMes.ListIndex = 0
cbo_LogsMantenimientosAño.ListIndex = 0
cbo_LogsDespachosMes.ListIndex = 0
cbo_LogsDespachosAño.ListIndex = 0
cbo_LogsUsuariosMes.ListIndex = 0
cbo_LogsUsuariosAño.ListIndex = 0
cbo_AllLogsMes.ListIndex = 0
cbo_AllLogsAño.ListIndex = 0

cbo_LogsProductoMes.Visible = True
cbo_LogsProductosAño.Visible = True
cmd_VerLogsProductos.Visible = True

cmd_LogsReportes.Left = 6840
          Text4.Left = 6840
          cbo_LogsReportesMes.Visible = False
          cbo_LogsReportesAño.Visible = False
          cmd_VerLogsReportes.Visible = False
          
cmd_LogsMantenimiento.Left = 10080
          Text3.Left = 10080
          cbo_LogsMantenimientoMes.Visible = False
          cbo_LogsMantenimientosAño.Visible = False
          cmd_VerLogsMantenimiento.Visible = False
          
cmd_LogsDespachos.Left = 480
          Text2.Left = 480
          cbo_LogsDespachosMes.Visible = False
          cbo_LogsDespachosAño.Visible = False
          cmd_VerLogsDespachos.Visible = False
          
cmd_LogsUsuarios.Left = 5280
          Text5.Left = 5280
          cbo_LogsUsuariosMes.Visible = False
          cbo_LogsUsuariosAño.Visible = False
          cmd_VerLogsUsuarios.Visible = False
          
cmd_AllLogs.Left = 10080
          Text6.Left = 10080
          cbo_AllLogsMes.Visible = False
          cbo_AllLogsAño.Visible = False
          cmd_VerAllLogs.Visible = False
End Sub

Private Sub cmd_LogsReportes_Click()
cbo_LogsProductoMes.ListIndex = 0
cbo_LogsProductosAño.ListIndex = 0
cbo_LogsReportesMes.ListIndex = 0
cbo_LogsReportesAño.ListIndex = 0
cbo_LogsMantenimientoMes.ListIndex = 0
cbo_LogsMantenimientosAño.ListIndex = 0
cbo_LogsDespachosMes.ListIndex = 0
cbo_LogsDespachosAño.ListIndex = 0
cbo_LogsUsuariosMes.ListIndex = 0
cbo_LogsUsuariosAño.ListIndex = 0
cbo_AllLogsMes.ListIndex = 0
cbo_AllLogsAño.ListIndex = 0

cmd_LogsProductos.Left = 480
          Text1.Left = 480
          cbo_LogsProductoMes.Visible = False
          cbo_LogsProductosAño.Visible = False
          cmd_VerLogsProductos.Visible = False

cmd_LogsReportes.Left = 3720
          Text4.Left = 3720
          cbo_LogsReportesMes.Visible = True
          cbo_LogsReportesAño.Visible = True
          cmd_VerLogsReportes.Visible = True

cmd_LogsMantenimiento.Left = 10080
          Text3.Left = 10080
          cbo_LogsMantenimientoMes.Visible = False
          cbo_LogsMantenimientosAño.Visible = False
          cmd_VerLogsMantenimiento.Visible = False
          
cmd_LogsDespachos.Left = 480
          Text2.Left = 480
          cbo_LogsDespachosMes.Visible = False
          cbo_LogsDespachosAño.Visible = False
          cmd_VerLogsDespachos.Visible = False
          
cmd_LogsUsuarios.Left = 5280
          Text5.Left = 5280
          cbo_LogsUsuariosMes.Visible = False
          cbo_LogsUsuariosAño.Visible = False
          cmd_VerLogsUsuarios.Visible = False
          
cmd_AllLogs.Left = 10080
          Text6.Left = 10080
          cbo_AllLogsMes.Visible = False
          cbo_AllLogsAño.Visible = False
          cmd_VerAllLogs.Visible = False


End Sub

Private Sub cmd_LogsUsuarios_Click()
cbo_LogsProductoMes.ListIndex = 0
cbo_LogsProductosAño.ListIndex = 0
cbo_LogsReportesMes.ListIndex = 0
cbo_LogsReportesAño.ListIndex = 0
cbo_LogsMantenimientoMes.ListIndex = 0
cbo_LogsMantenimientosAño.ListIndex = 0
cbo_LogsDespachosMes.ListIndex = 0
cbo_LogsDespachosAño.ListIndex = 0
cbo_LogsUsuariosMes.ListIndex = 0
cbo_LogsUsuariosAño.ListIndex = 0
cbo_AllLogsMes.ListIndex = 0
cbo_AllLogsAño.ListIndex = 0

cmd_LogsProductos.Left = 480
          Text1.Left = 480
          cbo_LogsProductoMes.Visible = False
          cbo_LogsProductosAño.Visible = False
          cmd_VerLogsProductos.Visible = False
          
cmd_LogsReportes.Left = 5280
          Text4.Left = 5280
          cbo_LogsReportesMes.Visible = False
          cbo_LogsReportesAño.Visible = False
          cmd_VerLogsReportes.Visible = False
          
cmd_LogsMantenimiento.Left = 10080
          Text3.Left = 10080
          cbo_LogsMantenimientoMes.Visible = False
          cbo_LogsMantenimientosAño.Visible = False
          cmd_VerLogsMantenimiento.Visible = False
          
cmd_LogsDespachos.Left = 480
          Text2.Left = 480
          cbo_LogsDespachosMes.Visible = False
          cbo_LogsDespachosAño.Visible = False
          cmd_VerLogsDespachos.Visible = False
          
cmd_LogsUsuarios.Left = 3720
          Text5.Left = 3720
          cbo_LogsUsuariosMes.Visible = True
          cbo_LogsUsuariosAño.Visible = True
          cmd_VerLogsUsuarios.Visible = True
          
cmd_AllLogs.Left = 10080
          Text6.Left = 10080
          cbo_AllLogsMes.Visible = False
          cbo_AllLogsAño.Visible = False
          cmd_VerAllLogs.Visible = False
End Sub

Private Sub cmd_VerAllLogs_Click()
If cbo_AllLogsMes.ListIndex = 0 And cbo_AllLogsAño.ListIndex = 0 Then
     MsgBox ("Por favor, indique el mes y el año que desea consultar"), vbInformation, "Aviso": Exit Sub
ElseIf cbo_AllLogsMes.Text = "Seleccione" Then
     MsgBox ("Por favor, indique el mes que desea consultar"), vbInformation, "Aviso": cbo_AllLogsMes.SetFocus: Exit Sub
ElseIf cbo_AllLogsAño.Text = "Seleccione" Then
     MsgBox ("Por favor, indique el año que desea consultar"), vbInformation, "Aviso": cbo_AllLogsAño.SetFocus: Exit Sub
Else
     cmd_AllLogs.Left = 10080
          cbo_AllLogsMes.Visible = False
          cbo_AllLogsAño.Visible = False
          cmd_VerAllLogs.Visible = False
          Text6.Left = 10080
     cmd_LogsUsuarios.Left = 5280
          Text5.Left = 5280
          
     If cbo_AllLogsMes.Text = "Enero" Then vMesLogs = "/01/"
     If cbo_AllLogsMes.Text = "Febrero" Then vMesLogs = "/02/"
     If cbo_AllLogsMes.Text = "Marzo" Then vMesLogs = "/03/"
     If cbo_AllLogsMes.Text = "Abril" Then vMesLogs = "/04/"
     If cbo_AllLogsMes.Text = "Mayo" Then vMesLogs = "/05/"
     If cbo_AllLogsMes.Text = "Junio" Then vMesLogs = "/06/"
     If cbo_AllLogsMes.Text = "Julio" Then vMesLogs = "/07/"
     If cbo_AllLogsMes.Text = "Agosto" Then vMesLogs = "/08/"
     If cbo_AllLogsMes.Text = "Septiembre" Then vMesLogs = "/09/"
     If cbo_AllLogsMes.Text = "Octubre" Then vMesLogs = "/10/"
     If cbo_AllLogsMes.Text = "Noviembre" Then vMesLogs = "/11/"
     If cbo_AllLogsMes.Text = "Diciembre" Then vMesLogs = "/12/"
     
     If cbo_AllLogsAño.Text = "2016" Then vAñoLogs = "2016"
     If cbo_AllLogsAño.Text = "2017" Then vAñoLogs = "2017"
     If cbo_AllLogsAño.Text = "2018" Then vAñoLogs = "2018"
     If cbo_AllLogsAño.Text = "2019" Then vAñoLogs = "2019"
          
          
     With AdoTodosLogs
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
          .RecordSource = "SELECT * FROM Logs WHERE Fecha LIKE '%" & vMesLogs & vAñoLogs & "%'"
          .Refresh
          Set LogsAllForm.GrillaAllLogs.DataSource = LogsForm.AdoTodosLogs
          LogsAllForm.EstilosGrillaAllLogs
          LogsAllForm.lbl_Mes.Caption = UCase(Trim(cbo_AllLogsMes.Text))
          LogsAllForm.lbl_Año.Caption = Trim(cbo_AllLogsAño.Text)
          LogsAllForm.Show vbModal
     End With
End If
End Sub

Private Sub cmd_VerLogsDespachos_Click()
If cbo_LogsDespachosMes.ListIndex = 0 And cbo_LogsDespachosAño.ListIndex = 0 Then
     MsgBox ("Por favor, indique el mes y el año que desea consultar"), vbInformation, "Aviso": Exit Sub
ElseIf cbo_LogsDespachosMes.Text = "Seleccione" Then
     MsgBox ("Por favor, indique el mes que desea consultar"), vbInformation, "Aviso": cbo_LogsDespachosMes.SetFocus: Exit Sub
ElseIf cbo_LogsDespachosAño.Text = "Seleccione" Then
     MsgBox ("Por favor, indique el año que desea consultar"), vbInformation, "Aviso": cbo_LogsDespachosAño.SetFocus: Exit Sub
Else
          cbo_LogsDespachosMes.Visible = False
          cbo_LogsDespachosAño.Visible = False
          cmd_VerLogsDespachos.Visible = False
     cmd_LogsUsuarios.Left = 5280
     Text5.Left = 5280
     
     If cbo_LogsDespachosMes.Text = "Enero" Then vMesLogs = "/01/"
     If cbo_LogsDespachosMes.Text = "Febrero" Then vMesLogs = "/02/"
     If cbo_LogsDespachosMes.Text = "Marzo" Then vMesLogs = "/03/"
     If cbo_LogsDespachosMes.Text = "Abril" Then vMesLogs = "/04/"
     If cbo_LogsDespachosMes.Text = "Mayo" Then vMesLogs = "/05/"
     If cbo_LogsDespachosMes.Text = "Junio" Then vMesLogs = "/06/"
     If cbo_LogsDespachosMes.Text = "Julio" Then vMesLogs = "/07/"
     If cbo_LogsDespachosMes.Text = "Agosto" Then vMesLogs = "/08/"
     If cbo_LogsDespachosMes.Text = "Septiembre" Then vMesLogs = "/09/"
     If cbo_LogsDespachosMes.Text = "Octubre" Then vMesLogs = "/10/"
     If cbo_LogsDespachosMes.Text = "Noviembre" Then vMesLogs = "/11/"
     If cbo_LogsDespachosMes.Text = "Diciembre" Then vMesLogs = "/12/"
     
     If cbo_LogsDespachosAño.Text = "2016" Then vAñoLogs = "2016"
     If cbo_LogsDespachosAño.Text = "2017" Then vAñoLogs = "2017"
     If cbo_LogsDespachosAño.Text = "2018" Then vAñoLogs = "2018"
     If cbo_LogsDespachosAño.Text = "2019" Then vAñoLogs = "2019"
     
     With AdoLogsDespachos
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
          .RecordSource = "SELECT * FROM LogsDespachos WHERE Fecha LIKE '%" & vMesLogs & vAñoLogs & "%'"
          .Refresh
          Set LogsDespachosForm.GrillaLogsDespachos.DataSource = LogsForm.AdoLogsDespachos
          LogsDespachosForm.EstilosGrillaLogsDespachos
          LogsDespachosForm.lbl_Mes.Caption = UCase(Trim(cbo_LogsDespachosMes.Text))
          LogsDespachosForm.lbl_Año.Caption = Trim(cbo_LogsDespachosAño.Text)
          LogsDespachosForm.Show vbModal
     End With
End If
End Sub

Private Sub cmd_VerLogsMantenimiento_Click()
If cbo_LogsMantenimientoMes.ListIndex = 0 And cbo_LogsMantenimientosAño.ListIndex = 0 Then
     MsgBox ("Por favor, indique el mes y el año que desea consultar"), vbInformation, "Aviso": Exit Sub
ElseIf cbo_LogsMantenimientoMes.Text = "Seleccione" Then
     MsgBox ("Por favor, indique el mes que desea consultar"), vbInformation, "Aviso": cbo_LogsMantenimientoMes.SetFocus: Exit Sub
ElseIf cbo_LogsMantenimientosAño.Text = "Seleccione" Then
     MsgBox ("Por favor, indique el año que desea consultar"), vbInformation, "Aviso": cbo_LogsMantenimientosAño.SetFocus: Exit Sub
Else
     cmd_LogsMantenimiento.Left = 10080
          Text3.Left = 10080
          cbo_LogsMantenimientoMes.Visible = False
          cbo_LogsMantenimientosAño.Visible = False
          cmd_VerLogsMantenimiento.Visible = False
     
     cmd_LogsReportes.Left = 5280
          Text4.Left = 5280
          
          
'     If cbo_LogsMantenimientoMes.Text = "Enero" Then vMesLogs = "/01/"
'     If cbo_LogsMantenimientoMes.Text = "Febreo" Then vMesLogs = "/02/"
'     If cbo_LogsMantenimientoMes.Text = "Marzo" Then vMesLogs = "/03/"
'     If cbo_LogsMantenimientoMes.Text = "Abril" Then vMesLogs = "/04/"
'     If cbo_LogsMantenimientoMes.Text = "Mayo" Then vMesLogs = "/05/"
'     If cbo_LogsMantenimientoMes.Text = "Junio" Then vMesLogs = "/06/"
'     If cbo_LogsMantenimientoMes.Text = "Julio" Then vMesLogs = "/07/"
'     If cbo_LogsMantenimientoMes.Text = "Agosto" Then vMesLogs = "/08/"
'     If cbo_LogsMantenimientoMes.Text = "Septiembre" Then vMesLogs = "/09/"
'     If cbo_LogsMantenimientoMes.Text = "Octubre" Then vMesLogs = "/10/"
'     If cbo_LogsMantenimientoMes.Text = "Noviembre" Then vMesLogs = "/11/"
'     If cbo_LogsMantenimientoMes.Text = "Diciembre" Then vMesLogs = "/12/"

     If cbo_LogsMantenimientoMes.ListIndex = 1 Then vMesLogs = "/01/"
     If cbo_LogsMantenimientoMes.ListIndex = 2 Then vMesLogs = "/02/"
     If cbo_LogsMantenimientoMes.ListIndex = 3 Then vMesLogs = "/03/"
     If cbo_LogsMantenimientoMes.ListIndex = 4 Then vMesLogs = "/04/"
     If cbo_LogsMantenimientoMes.ListIndex = 5 Then vMesLogs = "/05/"
     If cbo_LogsMantenimientoMes.ListIndex = 6 Then vMesLogs = "/06/"
     If cbo_LogsMantenimientoMes.ListIndex = 7 Then vMesLogs = "/07/"
     If cbo_LogsMantenimientoMes.ListIndex = 8 Then vMesLogs = "/08/"
     If cbo_LogsMantenimientoMes.ListIndex = 9 Then vMesLogs = "/09/"
     If cbo_LogsMantenimientoMes.ListIndex = 10 Then vMesLogs = "/10/"
     If cbo_LogsMantenimientoMes.ListIndex = 11 Then vMesLogs = "/11/"
     If cbo_LogsMantenimientoMes.ListIndex = 12 Then vMesLogs = "/12/"
     
     If cbo_LogsMantenimientosAño.Text = "2016" Then vAñoLogs = "2016"
     If cbo_LogsMantenimientosAño.Text = "2017" Then vAñoLogs = "2017"
     If cbo_LogsMantenimientosAño.Text = "2018" Then vAñoLogs = "2018"
     If cbo_LogsMantenimientosAño.Text = "2019" Then vAñoLogs = "2019"
          
     With AdoLogsMantenimiento
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
          .RecordSource = "SELECT * FROM LogsMantenimiento WHERE Fecha LIKE '%" & vMesLogs & vAñoLogs & "%' "
          .Refresh
          Set LogsMantenimientoForm.GrillaLogsMantenimiento.DataSource = LogsForm.AdoLogsMantenimiento
          LogsMantenimientoForm.EstilosGrillaLogsMantenimiento
          LogsMantenimientoForm.lbl_Mes.Caption = UCase(Trim(cbo_LogsMantenimientoMes.Text))
          LogsMantenimientoForm.lbl_Año.Caption = Trim(cbo_LogsMantenimientosAño.Text)
          LogsMantenimientoForm.Show vbModal
     End With
     
End If
End Sub

Private Sub cmd_VerLogsProductos_Click()
If cbo_LogsProductoMes.ListIndex = 0 And cbo_LogsProductosAño.ListIndex = 0 Then
     MsgBox ("Por favor, indique el mes y el año que quiere consultar"), vbInformation, "Aviso": Exit Sub

ElseIf cbo_LogsProductoMes.Text = "Seleccione" Then
     MsgBox ("Por favor, indique el mes que desea consultar"), vbInformation, "Aviso": cbo_LogsProductoMes.SetFocus: Exit Sub
     
ElseIf cbo_LogsProductosAño.Text = "Seleccione" Then
     MsgBox ("Por favor, indique el año que desea consultar"), vbInformation, "Aviso": cbo_LogsProductosAño.SetFocus: Exit Sub

Else
     cmd_LogsProductos.Left = 480
               cbo_LogsProductoMes.Visible = False
               cbo_LogsProductosAño.Visible = False
               cmd_VerLogsProductos.Visible = False
     cmd_LogsReportes.Left = 5280
               Text4.Left = 5280
     cmd_LogsMantenimiento.Left = 10080
     cmd_LogsDespachos.Left = 480
     cmd_LogsUsuarios.Left = 5280
     cmd_LogsMantenimiento.Left = 10080
     
     'AQUI DETERMINO QUE MES Y AÑO ESCOGIO ANTES DE MOSTRAR
     If cbo_LogsProductoMes.Text = "Enero" Then vMesLogs = "/01/"
     If cbo_LogsProductoMes.Text = "Febrero" Then vMesLogs = "/02/"
     If cbo_LogsProductoMes.Text = "Marzo" Then vMesLogs = "/03/"
     If cbo_LogsProductoMes.Text = "Abril" Then vMesLogs = "/04/"
     If cbo_LogsProductoMes.Text = "Mayo" Then vMesLogs = "/05/"
     If cbo_LogsProductoMes.Text = "Junio" Then vMesLogs = "/06/"
     If cbo_LogsProductoMes.Text = "Julio" Then vMesLogs = "/07/"
     If cbo_LogsProductoMes.Text = "Agosto" Then vMesLogs = "/08/"
     If cbo_LogsProductoMes.Text = "Septiembre" Then vMesLogs = "/09/"
     If cbo_LogsProductoMes.Text = "Octubre" Then vMesLogs = "/10/"
     If cbo_LogsProductoMes.Text = "Noviembre" Then vMesLogs = "/11/"
     If cbo_LogsProductoMes.Text = "Diciembre" Then vMesLogs = "/12/"
     
     If cbo_LogsProductosAño.Text = "2016" Then vAñoLogs = "2016"
     If cbo_LogsProductosAño.Text = "2017" Then vAñoLogs = "2017"
     If cbo_LogsProductosAño.Text = "2018" Then vAñoLogs = "2018"
     If cbo_LogsProductosAño.Text = "2019" Then vAñoLogs = "2019"
     
     With AdoLogsProductos
               .CursorLocation = adUseClient
               .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
               .RecordSource = "SELECT * FROM LogsProductos WHERE Fecha LIKE '%" & vMesLogs & vAñoLogs & "%'"
               .Refresh
               Set LogsProductosForm.GrillaLogsProductos.DataSource = LogsForm.AdoLogsProductos
               LogsProductosForm.EstilosGrillaLogsProductos
               LogsProductosForm.lbl_Mes.Caption = UCase(Trim(cbo_LogsProductoMes.Text))
               LogsProductosForm.lbl_Año.Caption = UCase(Trim(cbo_LogsProductosAño.Text))
               LogsProductosForm.Show vbModal
     End With
End If

End Sub

Private Sub cmd_VerLogsReportes_Click()
If cbo_LogsReportesMes.ListIndex = 0 And cbo_LogsReportesAño.ListIndex = 0 Then
     MsgBox ("Por favor, indique el mes y el año que desea consultar"), vbInformation, "Aviso": Exit Sub
ElseIf cbo_LogsReportesMes.Text = "Seleccione" Then
     MsgBox ("Por favor, indique el mes que desea consultar"), vbInformation, "Aviso": cbo_LogsReportesMes.SetFocus: Exit Sub
ElseIf cbo_LogsReportesAño.Text = "Seleccione" Then
     MsgBox ("Por favor, indique el año que desea consultar"), vbInformation, "Aviso": cbo_LogsReportesAño.SetFocus: Exit Sub

Else
     cmd_LogsReportes.Left = 5280
          Text4.Left = 5280
          cbo_LogsReportesMes.Visible = False
          cbo_LogsReportesAño.Visible = False
          cmd_VerLogsReportes.Visible = False
          
          If cbo_LogsReportesMes.Text = "Enero" Then vMesLogs = "/01/"
          If cbo_LogsReportesMes.Text = "Febrero" Then vMesLogs = "/02/"
          If cbo_LogsReportesMes.Text = "Marzo" Then vMesLogs = "/03/"
          If cbo_LogsReportesMes.Text = "Abril" Then vMesLogs = "/04/"
          If cbo_LogsReportesMes.Text = "Mayo" Then vMesLogs = "/05/"
          If cbo_LogsReportesMes.Text = "Junio" Then vMesLogs = "/06/"
          If cbo_LogsReportesMes.Text = "Julio" Then vMesLogs = "/07/"
          If cbo_LogsReportesMes.Text = "Agosto" Then vMesLogs = "/08/"
          If cbo_LogsReportesMes.Text = "Septiembre" Then vMesLogs = "/09/"
          If cbo_LogsReportesMes.Text = "Octubre" Then vMesLogs = "/10/"
          If cbo_LogsReportesMes.Text = "Noviembre" Then vMesLogs = "/11/"
          If cbo_LogsReportesMes.Text = "Diciembre" Then vMesLogs = "/12/"
          
          If cbo_LogsReportesAño.Text = "2016" Then vAñoLogs = "2016"
          If cbo_LogsReportesAño.Text = "2017" Then vAñoLogs = "2017"
          If cbo_LogsReportesAño.Text = "2018" Then vAñoLogs = "2018"
          If cbo_LogsReportesAño.Text = "2019" Then vAñoLogs = "2019"
          
          With AdoLogsReportes
               .CursorLocation = adUseClient
               .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
               .RecordSource = "SELECT * FROM LogsReportes WHERE Fecha LIKE '%" & vMesLogs & vAñoLogs & "%'"
               .Refresh
               Set LogsReportesForm.GrillaLogsReportes.DataSource = LogsForm.AdoLogsReportes
               LogsReportesForm.EstilosGrillaLogsReportes
               LogsReportesForm.lbl_Mes.Caption = UCase(Trim(cbo_LogsReportesMes.Text))
               LogsReportesForm.lbl_Año.Caption = Trim(cbo_LogsReportesAño.Text)
               LogsReportesForm.Show vbModal
          End With
End If
    
End Sub

Private Sub cmd_VerLogsUsuarios_Click()
If cbo_LogsUsuariosMes.ListIndex = 0 And cbo_LogsUsuariosAño.ListIndex = 0 Then
     MsgBox ("Por favor, indique el mes y el año que desea consultar"), vbInformation, "Aviso": Exit Sub
ElseIf cbo_LogsUsuariosMes.Text = "Seleccione" Then
     MsgBox ("Por favor, indique el mes que desea consultar"), vbInformation, "Aviso": cbo_LogsUsuariosMes.SetFocus: Exit Sub
ElseIf cbo_LogsUsuariosAño.Text = "Seleccione" Then
     MsgBox ("Por favor, indique el año que desea consultar"), vbInformation, "Aviso": cbo_LogsUsuariosAño.SetFocus: Exit Sub
Else
     cmd_LogsUsuarios.Left = 5280
          cbo_LogsUsuariosMes.Visible = False
          cbo_LogsUsuariosAño.Visible = False
          cmd_VerLogsUsuarios.Visible = False
          Text5.Left = 5280
          
     If cbo_LogsUsuariosMes.Text = "Enero" Then vMesLogs = "/01/"
     If cbo_LogsUsuariosMes.Text = "Febrero" Then vMesLogs = "/02/"
     If cbo_LogsUsuariosMes.Text = "Marzo" Then vMesLogs = "/03/"
     If cbo_LogsUsuariosMes.Text = "Abril" Then vMesLogs = "/04/"
     If cbo_LogsUsuariosMes.Text = "Mayo" Then vMesLogs = "/05/"
     If cbo_LogsUsuariosMes.Text = "Junio" Then vMesLogs = "/06/"
     If cbo_LogsUsuariosMes.Text = "Julio" Then vMesLogs = "/07/"
     If cbo_LogsUsuariosMes.Text = "Agosto" Then vMesLogs = "/08/"
     If cbo_LogsUsuariosMes.Text = "Septiembre" Then vMesLogs = "/09/"
     If cbo_LogsUsuariosMes.Text = "Octubre" Then vMesLogs = "/10/"
     If cbo_LogsUsuariosMes.Text = "Noviembre" Then vMesLogs = "/11/"
     If cbo_LogsUsuariosMes.Text = "Diciembre" Then vMesLogs = "/12/"
     
     If cbo_LogsUsuariosAño.Text = "2016" Then vAñoLogs = "2016"
     If cbo_LogsUsuariosAño.Text = "2017" Then vAñoLogs = "2017"
     If cbo_LogsUsuariosAño.Text = "2018" Then vAñoLogs = "2018"
     If cbo_LogsUsuariosAño.Text = "2019" Then vAñoLogs = "2019"
          
          
     With AdoLogsUsuarios
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
          .RecordSource = "SELECT * FROM LogsUsuarios WHERE Fecha LIKE '%" & vMesLogs & vAñoLogs & "%'"
          .Refresh
          Set LogsUsuariosForm.GrillaLogsUsuarios.DataSource = LogsForm.AdoLogsUsuarios
          LogsUsuariosForm.EstilosGrillaLogsUsuarios
          LogsUsuariosForm.lbl_Mes.Caption = UCase(Trim(cbo_LogsUsuariosMes.Text))
          LogsUsuariosForm.lbl_Año.Caption = Trim(cbo_LogsUsuariosAño.Text)
          LogsUsuariosForm.Show vbModal
     End With
End If

End Sub

Private Sub Form_Load()
cbo_LogsProductoMes.AddItem "Seleccione"
cbo_LogsProductoMes.AddItem "Enero"
cbo_LogsProductoMes.AddItem "Febrero"
cbo_LogsProductoMes.AddItem "Marzo"
cbo_LogsProductoMes.AddItem "Abril"
cbo_LogsProductoMes.AddItem "Mayo"
cbo_LogsProductoMes.AddItem "Junio"
cbo_LogsProductoMes.AddItem "Julio"
cbo_LogsProductoMes.AddItem "Agosto"
cbo_LogsProductoMes.AddItem "Septiembre"
cbo_LogsProductoMes.AddItem "Octubre"
cbo_LogsProductoMes.AddItem "Noviembre"
cbo_LogsProductoMes.AddItem "Diciembre"
cbo_LogsProductoMes.ListIndex = 0

cbo_LogsProductosAño.AddItem "Seleccione"
cbo_LogsProductosAño.AddItem "2016"
cbo_LogsProductosAño.AddItem "2017"
cbo_LogsProductosAño.AddItem "2018"
cbo_LogsProductosAño.AddItem "2019"
cbo_LogsProductosAño.ListIndex = 0

cbo_LogsReportesMes.AddItem "Seleccione"
cbo_LogsReportesMes.AddItem "Enero"
cbo_LogsReportesMes.AddItem "Febrero"
cbo_LogsReportesMes.AddItem "Marzo"
cbo_LogsReportesMes.AddItem "Abril"
cbo_LogsReportesMes.AddItem "Mayo"
cbo_LogsReportesMes.AddItem "Junio"
cbo_LogsReportesMes.AddItem "Julio"
cbo_LogsReportesMes.AddItem "Agosto"
cbo_LogsReportesMes.AddItem "Septiembre"
cbo_LogsReportesMes.AddItem "Octubre"
cbo_LogsReportesMes.AddItem "Noviembre"
cbo_LogsReportesMes.AddItem "Diciembre"
cbo_LogsReportesMes.ListIndex = 0

cbo_LogsReportesAño.AddItem "Seleccione"
cbo_LogsReportesAño.AddItem "2016"
cbo_LogsReportesAño.AddItem "2017"
cbo_LogsReportesAño.AddItem "2018"
cbo_LogsReportesAño.AddItem "2019"
cbo_LogsReportesAño.ListIndex = 0

cbo_LogsMantenimientoMes.AddItem "Seleccione"
cbo_LogsMantenimientoMes.AddItem "Enero"
cbo_LogsMantenimientoMes.AddItem "Febrero"
cbo_LogsMantenimientoMes.AddItem "Marzo"
cbo_LogsMantenimientoMes.AddItem "Abril"
cbo_LogsMantenimientoMes.AddItem "Mayo"
cbo_LogsMantenimientoMes.AddItem "Junio"
cbo_LogsMantenimientoMes.AddItem "Julio"
cbo_LogsMantenimientoMes.AddItem "Agosto"
cbo_LogsMantenimientoMes.AddItem "Septiembre"
cbo_LogsMantenimientoMes.AddItem "Octubre"
cbo_LogsMantenimientoMes.AddItem "Noviembre"
cbo_LogsMantenimientoMes.AddItem "Diciembre"
cbo_LogsMantenimientoMes.ListIndex = 0

cbo_LogsMantenimientosAño.AddItem "Seleccione"
cbo_LogsMantenimientosAño.AddItem "2016"
cbo_LogsMantenimientosAño.AddItem "2017"
cbo_LogsMantenimientosAño.AddItem "2018"
cbo_LogsMantenimientosAño.AddItem "2019"
cbo_LogsMantenimientosAño.ListIndex = 0

cbo_LogsDespachosMes.AddItem "Seleccione"
cbo_LogsDespachosMes.AddItem "Enero"
cbo_LogsDespachosMes.AddItem "Febrero"
cbo_LogsDespachosMes.AddItem "Marzo"
cbo_LogsDespachosMes.AddItem "Abril"
cbo_LogsDespachosMes.AddItem "Mayo"
cbo_LogsDespachosMes.AddItem "Junio"
cbo_LogsDespachosMes.AddItem "Julio"
cbo_LogsDespachosMes.AddItem "Agosto"
cbo_LogsDespachosMes.AddItem "Septiembre"
cbo_LogsDespachosMes.AddItem "Octubre"
cbo_LogsDespachosMes.AddItem "Noviembre"
cbo_LogsDespachosMes.AddItem "Diciembre"
cbo_LogsDespachosMes.ListIndex = 0

cbo_LogsDespachosAño.AddItem "Seleccione"
cbo_LogsDespachosAño.AddItem "2016"
cbo_LogsDespachosAño.AddItem "2017"
cbo_LogsDespachosAño.AddItem "2018"
cbo_LogsDespachosAño.AddItem "2019"
cbo_LogsDespachosAño.ListIndex = 0

cbo_LogsUsuariosMes.AddItem "Seleccione"
cbo_LogsUsuariosMes.AddItem "Enero"
cbo_LogsUsuariosMes.AddItem "Febrero"
cbo_LogsUsuariosMes.AddItem "Marzo"
cbo_LogsUsuariosMes.AddItem "Abril"
cbo_LogsUsuariosMes.AddItem "Mayo"
cbo_LogsUsuariosMes.AddItem "Junio"
cbo_LogsUsuariosMes.AddItem "Julio"
cbo_LogsUsuariosMes.AddItem "Agosto"
cbo_LogsUsuariosMes.AddItem "Septiembre"
cbo_LogsUsuariosMes.AddItem "Octubre"
cbo_LogsUsuariosMes.AddItem "Noviembre"
cbo_LogsUsuariosMes.AddItem "Diciembre"
cbo_LogsUsuariosMes.ListIndex = 0

cbo_LogsUsuariosAño.AddItem "Seleccione"
cbo_LogsUsuariosAño.AddItem "2016"
cbo_LogsUsuariosAño.AddItem "2017"
cbo_LogsUsuariosAño.AddItem "2018"
cbo_LogsUsuariosAño.AddItem "2019"
cbo_LogsUsuariosAño.ListIndex = 0

cbo_AllLogsMes.AddItem "Seleccione"
cbo_AllLogsMes.AddItem "Enero"
cbo_AllLogsMes.AddItem "Febrero"
cbo_AllLogsMes.AddItem "Marzo"
cbo_AllLogsMes.AddItem "Abril"
cbo_AllLogsMes.AddItem "Mayo"
cbo_AllLogsMes.AddItem "Junio"
cbo_AllLogsMes.AddItem "Julio"
cbo_AllLogsMes.AddItem "Agosto"
cbo_AllLogsMes.AddItem "Septiembre"
cbo_AllLogsMes.AddItem "Octubre"
cbo_AllLogsMes.AddItem "Noviembre"
cbo_AllLogsMes.AddItem "Diciembre"
cbo_AllLogsMes.ListIndex = 0

cbo_AllLogsAño.AddItem "Seleccione"
cbo_AllLogsAño.AddItem "2016"
cbo_AllLogsAño.AddItem "2017"
cbo_AllLogsAño.AddItem "2018"
cbo_AllLogsAño.AddItem "2019"
cbo_AllLogsAño.ListIndex = 0

End Sub
