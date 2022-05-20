VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form IndexForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Index - Shekina 4004"
   ClientHeight    =   10500
   ClientLeft      =   2460
   ClientTop       =   1515
   ClientWidth     =   16485
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "IndexForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "IndexForm.frx":058A
   ScaleHeight     =   10500
   ScaleWidth      =   16485
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
   Begin MSDataGridLib.DataGrid GrillaTotalProductos 
      Height          =   255
      Left            =   16800
      TabIndex        =   28
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoTotalProductos 
      Height          =   330
      Left            =   16800
      Top             =   4920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   12480
      Top             =   840
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
   Begin VB.CommandButton cmd_inicio 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Inicio"
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000001&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid GrillaTotalRegistros 
      Height          =   255
      Left            =   16800
      TabIndex        =   29
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoTotalRegistros 
      Height          =   330
      Left            =   16800
      Top             =   3480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSDataGridLib.DataGrid GrillaProductosMas50 
      Height          =   255
      Left            =   16800
      TabIndex        =   30
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoProductosMas50 
      Height          =   330
      Left            =   16800
      Top             =   4200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSDataGridLib.DataGrid GrillaProductosMenos50 
      Height          =   255
      Left            =   18240
      TabIndex        =   31
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoProductosMenos50 
      Height          =   330
      Left            =   18240
      Top             =   4200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSDataGridLib.DataGrid GrillaProductosOUT 
      Height          =   255
      Left            =   18240
      TabIndex        =   32
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoProductosOUT 
      Height          =   330
      Left            =   18240
      Top             =   3480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSDataGridLib.DataGrid GrillaTotalDespachos 
      Height          =   255
      Left            =   16800
      TabIndex        =   33
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoTotalDespachos 
      Height          =   330
      Left            =   16800
      Top             =   2760
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSDataGridLib.DataGrid GrillaTotalDevoluciones 
      Height          =   255
      Left            =   18240
      TabIndex        =   38
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoTotalDevoluciones 
      Height          =   330
      Left            =   18240
      Top             =   2760
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin VB.Label lbl_TotalDevoluciones 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   840
      Left            =   7560
      MouseIcon       =   "IndexForm.frx":19F6C
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   3000
      Width           =   3045
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DE DEVOLUCIONES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   7500
      MouseIcon       =   "IndexForm.frx":1A276
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   4080
      Width           =   3075
   End
   Begin VB.Label lbl_TotalOperaciones 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   840
      Left            =   10890
      TabIndex        =   35
      Top             =   3000
      Width           =   3045
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DE OPERACIONES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   11190
      MousePointer    =   1  'Arrow
      TabIndex        =   34
      Top             =   4080
      Width           =   2445
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTOS SIN STOCK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   10900
      MouseIcon       =   "IndexForm.frx":1A580
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   5985
      Width           =   2970
   End
   Begin VB.Label lbl_fuerastock 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   840
      Left            =   10860
      MouseIcon       =   "IndexForm.frx":1A88A
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   4875
      Width           =   3045
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTOS CON < 50 STOCK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   7580
      MousePointer    =   1  'Arrow
      TabIndex        =   25
      Top             =   5990
      Width           =   2910
   End
   Begin VB.Label lbl_menos50 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   840
      Left            =   7560
      TabIndex        =   24
      Top             =   4875
      Width           =   2925
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTOS CON > 50 STOCKS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   4100
      MousePointer    =   1  'Arrow
      TabIndex        =   23
      Top             =   5990
      Width           =   3030
   End
   Begin VB.Label lbl_mas50 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   840
      Left            =   4200
      TabIndex        =   22
      Top             =   4875
      Width           =   2925
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DE DESPACHOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   4080
      MouseIcon       =   "IndexForm.frx":1AB94
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   4080
      Width           =   3075
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPOS DE PRODUCTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   4155
      MouseIcon       =   "IndexForm.frx":1AE9E
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   7980
      Width           =   9780
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DE PRODUCTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   4080
      MousePointer    =   1  'Arrow
      TabIndex        =   19
      Top             =   9910
      Width           =   9885
   End
   Begin VB.Label lbl_totaldespachos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   840
      Left            =   4080
      MouseIcon       =   "IndexForm.frx":1B1A8
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   3000
      Width           =   3045
   End
   Begin VB.Label lbl_totalregistros 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   840
      Left            =   4200
      MouseIcon       =   "IndexForm.frx":1B4B2
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   6900
      Width           =   9645
   End
   Begin VB.Label lbl_totalproductos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   840
      Left            =   4080
      TabIndex        =   16
      Top             =   8760
      Width           =   9885
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
      MouseIcon       =   "IndexForm.frx":1B7BC
      MousePointer    =   99  'Custom
      TabIndex        =   15
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
      MouseIcon       =   "IndexForm.frx":1BAC6
      MousePointer    =   99  'Custom
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   600
      Width           =   2460
   End
End
Attribute VB_Name = "IndexForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Consultas_Click()
IndexForm.Hide
ConsultasForm.Show
End Sub

Private Sub cmd_despachos_Click()
IndexForm.Hide
DespachosForm.Show
End Sub

Private Sub cmd_inicio_Click()
IndexForm.Enabled = False
End Sub

Private Sub cmd_Logs_Click()
LogsForm.Show
End Sub

Private Sub cmd_Mantenimiento_Click()
IndexForm.Hide
MantenimientoForm.Show
End Sub

Private Sub cmd_productos_Click()
IndexForm.Hide
ProductosForm.Show
End Sub

Private Sub cmd_reportes_Click()
IndexForm.Hide
ReportesForm.Show
End Sub

Private Sub cmd_ubicaciones_Click()
IndexForm.Hide
UbicacionesForm.Show
End Sub

Private Sub cmd_usuarios_Click()
IndexForm.Hide
UsuariosForm.Show
End Sub

Private Sub Command1_Click()
dr_prueba.Show
dr_prueba.WindowState = 2
End Sub

Private Sub Form_Load()
lbl_tarea.Caption = vTarea
lbl_username.Caption = vUsername
Productos

'todos los indicadores de las estadisticas
TotalProductos
TotalRegistros
ProductosMas50
ProductosMenos50
ProductosOUT
TotalDespachos
TotalDevoluciones
TotalOperaciones
End Sub

Private Sub Label10_Click()
With RsUsuarios
          .Requery
          .Find "Nomb_Usuario='" & Trim(lbl_username.Caption) & "'"
          If !rango = 0 Then
                    Exit Sub
          Else
                    ProductosEditForm.CboFiltrar.Text = "Cantidad"
                    ProductosEditForm.TxtFiltrar.Text = "0"
                    ProductosEditForm.Show
          End If
End With
End Sub

Private Sub Label11_Click()
DespachosDevolucionesHistorialForm.Show
End Sub

Private Sub Label3_Click()
IndexForm.Hide
ConsultasForm.Show
End Sub

Private Sub Label4_Click()
DespachosHistorialForm.Show
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

Private Sub lbl_fuerastock_Click()
With RsUsuarios
          .Requery
          .Find "Nomb_Usuario='" & Trim(lbl_username.Caption) & "'"
          If !rango = 0 Then
                    Exit Sub
          Else
                    ProductosEditForm.CboFiltrar.Text = "Cantidad"
                    ProductosEditForm.TxtFiltrar.Text = "0"
                    ProductosEditForm.Show
          End If
End With
End Sub

Private Sub lbl_totaldespachos_Click()
DespachosHistorialForm.Show
End Sub

Private Sub lbl_TotalDevoluciones_Click()
DespachosDevolucionesHistorialForm.Show
End Sub

Private Sub lbl_totalregistros_Click()
IndexForm.Hide
ConsultasForm.Show
End Sub

Private Sub Timer1_Timer()
'Label1.Caption = Format(Time, "hh:mm:ss")
Label1.Caption = Format(Now, "HH:MM AM/PM")
End Sub

Sub TotalProductos()
'programar consulta
With AdoTotalProductos
    .CursorLocation = adUseClient
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
    .RecordSource = "select sum(cantidad) from Productos" 'PARA SUMAR TODA LA CANTIDAD DE PRODUCTOS
    .Refresh
End With
Set GrillaTotalProductos.DataSource = AdoTotalProductos
lbl_totalproductos.Caption = GrillaTotalProductos.Columns(0).Text

If lbl_totalproductos.Caption = "" Then
     lbl_totalproductos.Caption = "0"
Else
     Exit Sub
End If
End Sub

Sub TotalRegistros()
'programar consulta
With AdoTotalRegistros
    .CursorLocation = adUseClient
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
    .RecordSource = "select count(*)  from Productos" 'PARA MOSTRAR LA CANTIDAD DE REGISTROS POR TIPO DE PRODUCTO
    .Refresh
End With
Set GrillaTotalRegistros.DataSource = AdoTotalRegistros
lbl_totalregistros.Caption = GrillaTotalRegistros.Columns(0).Text
End Sub

Sub ProductosMas50()
'programar consulta
With AdoProductosMas50
    .CursorLocation = adUseClient
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
    .RecordSource = "select count(*)  from Productos where cantidad >=50"  'PARA MOSTRAR LOS PRODUCTOS CON STOCK MAYOR A 50 PIEZAS
    .Refresh
End With
Set GrillaProductosMas50.DataSource = AdoProductosMas50
lbl_mas50.Caption = GrillaProductosMas50.Columns(0).Text
End Sub

Sub ProductosMenos50()
'programar consulta
With AdoProductosMenos50
    .CursorLocation = adUseClient
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
    .RecordSource = "select count(*)  from Productos where cantidad  between 49 and 1"  'PARA MOSTRAR LOS PRODUCTOS CON STOCK MAYOR A 50 PIEZAS
    .Refresh
End With
Set GrillaProductosMenos50.DataSource = AdoProductosMenos50
lbl_menos50.Caption = GrillaProductosMenos50.Columns(0).Text
End Sub

Sub ProductosOUT()
'programar consulta
With AdoProductosOUT
    .CursorLocation = adUseClient
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
    .RecordSource = "select count(*)  from Productos where cantidad =0"  'PARA MOSTRAR LOS PRODUCTOS CON STOCK MAYOR A 50 PIEZAS
    .Refresh
End With
Set GrillaProductosOUT.DataSource = AdoProductosOUT
lbl_fuerastock.Caption = GrillaProductosOUT.Columns(0).Text
End Sub

Sub TotalDespachos()
'programar consulta
With AdoTotalDespachos
    .CursorLocation = adUseClient
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
    .RecordSource = "select count(*)  from Despacho"  'PARA LA CANTIDAD DE DESPACHOS QUE SE HAN HECHO
    .Refresh
End With
Set GrillaTotalDespachos.DataSource = AdoTotalDespachos
lbl_totaldespachos.Caption = GrillaTotalDespachos.Columns(0).Text

End Sub

Sub TotalDevoluciones()
With AdoTotalDevoluciones
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
          .RecordSource = "SELECT COUNT (*) FROM Devoluciones"
          .Refresh
End With
Set GrillaTotalDevoluciones.DataSource = AdoTotalDevoluciones
lbl_TotalDevoluciones.Caption = GrillaTotalDevoluciones.Columns(0).Text
End Sub

Sub TotalOperaciones()
Dim Total As Integer
Total = Val(lbl_totaldespachos.Caption) + Val(lbl_TotalDevoluciones.Caption)
lbl_TotalOperaciones.Caption = Total
End Sub
