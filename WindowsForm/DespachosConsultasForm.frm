VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form DespachosConsultasForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta detallada de despacho"
   ClientHeight    =   11535
   ClientLeft      =   3855
   ClientTop       =   1335
   ClientWidth     =   13515
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DespachosConsultasForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DespachosConsultasForm.frx":058A
   ScaleHeight     =   11535
   ScaleWidth      =   13515
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid GrillaDespachosContenido 
      Height          =   1995
      Left            =   840
      TabIndex        =   24
      Top             =   8430
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3519
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
   Begin MSDataGridLib.DataGrid GrillaDespachosDespachos 
      Height          =   1995
      Left            =   840
      TabIndex        =   13
      Top             =   5790
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3519
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
   Begin MSDataGridLib.DataGrid GrillaDespachosUsuarios 
      Height          =   2000
      Left            =   840
      TabIndex        =   4
      Top             =   3120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3519
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
   Begin MSDataGridLib.DataGrid GrillaDevuelto 
      Height          =   735
      Left            =   10080
      TabIndex        =   40
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
   Begin MSAdodcLib.Adodc AdoTotalDespachosUsuarios 
      Height          =   330
      Left            =   10920
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSAdodcLib.Adodc AdoUsuarios 
      Height          =   330
      Left            =   10920
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.OptionButton opt_Busqueda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cédula"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   36
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmd_Atras 
      Caption         =   "Atras"
      Height          =   615
      Left            =   5760
      TabIndex        =   35
      Top             =   10700
      Width           =   2295
   End
   Begin VB.TextBox txt_parametro 
      Height          =   390
      Left            =   5160
      MaxLength       =   15
      TabIndex        =   3
      Top             =   2400
      Width           =   3615
   End
   Begin VB.OptionButton opt_Busqueda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nombre de Usuario"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc AdoDespachos 
      Height          =   330
      Left            =   10920
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSAdodcLib.Adodc AdoDetalles 
      Height          =   330
      Left            =   10920
      Top             =   8040
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   735
      Left            =   10080
      TabIndex        =   37
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
   Begin MSAdodcLib.Adodc AdoDevuelto 
      Height          =   330
      Left            =   10920
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¿Presenta Devolución?"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   9960
      TabIndex        =   39
      Top             =   7080
      Width           =   1665
   End
   Begin VB.Label lbl_Devuelto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   9960
      TabIndex        =   38
      Top             =   7320
      Width           =   60
   End
   Begin VB.Label lbl_PiezasxKit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   9960
      TabIndex        =   34
      Top             =   10080
      Width           =   60
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Piezas del kit"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   9960
      TabIndex        =   33
      Top             =   9840
      Width           =   915
   End
   Begin VB.Label lbl_Kit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   9960
      TabIndex        =   32
      Top             =   9480
      Width           =   60
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¿El producto es un kit?"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   9960
      TabIndex        =   31
      Top             =   9240
      Width           =   1635
   End
   Begin VB.Label lbl_CantidadVendida 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   6120
      TabIndex        =   30
      Top             =   9960
      Width           =   60
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad vendida:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   6120
      TabIndex        =   29
      Top             =   9720
      Width           =   1290
   End
   Begin VB.Label lbl_MarcaProducto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   6120
      TabIndex        =   28
      Top             =   9360
      Width           =   60
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   6120
      TabIndex        =   27
      Top             =   9120
      Width           =   495
   End
   Begin VB.Label lbl_DescripcionProducto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   6120
      TabIndex        =   26
      Top             =   8760
      Width           =   60
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción del producto:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   6120
      TabIndex        =   25
      Top             =   8520
      Width           =   1800
   End
   Begin VB.Label lbl_Despachador 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   9960
      TabIndex        =   23
      Top             =   6720
      Width           =   60
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Despachado por:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   9960
      TabIndex        =   22
      Top             =   6480
      Width           =   1230
   End
   Begin VB.Label lbl_Fecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   9960
      TabIndex        =   21
      Top             =   6120
      Width           =   60
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Realizado el:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   9960
      TabIndex        =   20
      Top             =   5880
      Width           =   915
   End
   Begin VB.Label lbl_Vendedor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   6120
      TabIndex        =   19
      Top             =   7320
      Width           =   60
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   6120
      TabIndex        =   18
      Top             =   7080
      Width           =   750
   End
   Begin VB.Label lbl_Zona 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   6120
      TabIndex        =   17
      Top             =   6720
      Width           =   60
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zona del pedido:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   6120
      TabIndex        =   16
      Top             =   6480
      Width           =   1200
   End
   Begin VB.Label lbl_NombreCliente 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   6120
      TabIndex        =   15
      Top             =   6120
      Width           =   60
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Cliente:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   6120
      TabIndex        =   14
      Top             =   5880
      Width           =   1425
   End
   Begin VB.Label lbl_TotalDespachos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   9960
      TabIndex        =   12
      Top             =   3480
      Width           =   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total de despachos realizados:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   9960
      TabIndex        =   11
      Top             =   3240
      Width           =   2205
   End
   Begin VB.Label lbl_Tarea 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   6120
      TabIndex        =   10
      Top             =   4680
      Width           =   60
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tarea:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   6120
      TabIndex        =   9
      Top             =   4440
      Width           =   480
   End
   Begin VB.Label lbl_Cedula 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   6120
      TabIndex        =   8
      Top             =   4080
      Width           =   60
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cédula de identidad:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   6120
      TabIndex        =   7
      Top             =   3840
      Width           =   1470
   End
   Begin VB.Label lbl_NombreUsuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   6120
      TabIndex        =   6
      Top             =   3480
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del usuario:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   6120
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
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
      Left            =   960
      TabIndex        =   1
      Top             =   1710
      Width           =   225
   End
   Begin VB.Label lbl_1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione los parámetros de búsqueda:"
      Height          =   270
      Left            =   1320
      TabIndex        =   0
      Top             =   1875
      Width           =   4275
   End
End
Attribute VB_Name = "DespachosConsultasForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Atras_Click()
Unload Me
Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
ConsultasForm.EstilosGrillaConsultas
End Sub

Private Sub Form_Load()
'abrimos las tablas necesarias
Usuarios
Despacho
DetallesDespacho
'Productos
Devoluciones

'conectamos el ado de los usuarios
AdoUsuarios.CursorLocation = adUseClient
AdoUsuarios.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
AdoUsuarios.RecordSource = "select * from Usuarios where [Nomb_Usuario] like '" & Busca & "'"
AdoUsuarios.Refresh
Set GrillaDespachosUsuarios.DataSource = AdoUsuarios
EstiloGrillaDespachosUsuarios

'conectamos el ado de los despachos de ese usuario
AdoDespachos.CursorLocation = adUseClient
AdoDespachos.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
AdoDespachos.RecordSource = "select * from Despacho where [Despachador] like '" & BuscarDespacho & "'"
AdoDespachos.Refresh
Set GrillaDespachosDespachos.DataSource = AdoDespachos
EstiloGrillaDespachosDespachos

'conectamos el ado de los detalles de ese despacho
AdoDetalles.CursorLocation = adUseClient
AdoDetalles.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
AdoDetalles.RecordSource = "select * from DetallesDespacho where [CodigoDespacho] like '" & BuscarDetalle & "'"
AdoDetalles.Refresh
Set GrillaDespachosContenido.DataSource = AdoDetalles
EstiloGrillaDespachosContenido

'para que aparezca como primera opcion de busqueda el nombre de usuario
opt_Busqueda(0).Value = True

GrillaDespachosContenido.Enabled = False
GrillaDespachosUsuarios.Enabled = False
GrillaDespachosDespachos.Enabled = False


End Sub

Sub EstiloGrillaDespachosUsuarios()

With RsUsuarios
        'If .BOF And .EOF = True Then Exit Sub para que cargue con el estilo exista o no datos
            'tamaños de la grilla
            GrillaDespachosUsuarios.Columns(0).Width = 500
            GrillaDespachosUsuarios.Columns(1).Width = 4450
            GrillaDespachosUsuarios.Columns(2).Width = 2400
            GrillaDespachosUsuarios.Columns(3).Width = 2000
            GrillaDespachosUsuarios.Columns(4).Width = 2000
            GrillaDespachosUsuarios.Columns(5).Width = 2000
            GrillaDespachosUsuarios.Columns(6).Width = 500
            GrillaDespachosUsuarios.Columns(7).Width = 2000
            GrillaDespachosUsuarios.Columns(8).Width = 2400

            'Caption de las grillas
            GrillaDespachosUsuarios.Columns(0).Caption = "ID"
            GrillaDespachosUsuarios.Columns(1).Caption = "Usuario"
            GrillaDespachosUsuarios.Columns(2).Caption = "Nombre"
            GrillaDespachosUsuarios.Columns(3).Caption = "Apellido"
            GrillaDespachosUsuarios.Columns(4).Caption = "Contraseña"
            GrillaDespachosUsuarios.Columns(5).Caption = "Email"
            GrillaDespachosUsuarios.Columns(6).Caption = "Rank"
            GrillaDespachosUsuarios.Columns(7).Caption = "Cédula"
            GrillaDespachosUsuarios.Columns(8).Caption = "Tarea"

            'alineacion
            GrillaDespachosUsuarios.Columns(0).Alignment = dbgCenter
            GrillaDespachosUsuarios.Columns(1).Alignment = dbgCenter
            GrillaDespachosUsuarios.Columns(2).Alignment = dbgCenter
            GrillaDespachosUsuarios.Columns(3).Alignment = dbgCenter
            GrillaDespachosUsuarios.Columns(4).Alignment = dbgCenter
            GrillaDespachosUsuarios.Columns(5).Alignment = dbgCenter
            GrillaDespachosUsuarios.Columns(6).Alignment = dbgCenter
            GrillaDespachosUsuarios.Columns(7).Alignment = dbgCenter
            GrillaDespachosUsuarios.Columns(8).Alignment = dbgCenter

            'cabeceras
            GrillaDespachosUsuarios.HeadFont.Bold = True

            'las que no quiero ver
            GrillaDespachosUsuarios.Columns(0).Visible = False
            GrillaDespachosUsuarios.Columns(2).Visible = False
            GrillaDespachosUsuarios.Columns(3).Visible = False
            GrillaDespachosUsuarios.Columns(4).Visible = False
            GrillaDespachosUsuarios.Columns(5).Visible = False
            GrillaDespachosUsuarios.Columns(6).Visible = False
            GrillaDespachosUsuarios.Columns(7).Visible = False
            GrillaDespachosUsuarios.Columns(8).Visible = False
End With
End Sub

Sub EstiloGrillaDespachosDespachos()

With RsDespacho
        'If .BOF Or .EOF Then Exit Sub para que cargue el estilo exista o no el estilo
    
            'tamaños de la grilla
            GrillaDespachosDespachos.Columns(0).Width = 500
            GrillaDespachosDespachos.Columns(1).Width = 4400
            GrillaDespachosDespachos.Columns(2).Width = 1600
            GrillaDespachosDespachos.Columns(3).Width = 1800
            GrillaDespachosDespachos.Columns(4).Width = 1700
            GrillaDespachosDespachos.Columns(5).Width = 1700
            GrillaDespachosDespachos.Columns(6).Width = 1700

            'Caption de las grillas
            GrillaDespachosDespachos.Columns(0).Caption = "ID"
            GrillaDespachosDespachos.Columns(1).Caption = "Codigo de los Despachos"
            GrillaDespachosDespachos.Columns(2).Caption = "Cliente"
            GrillaDespachosDespachos.Columns(3).Caption = "Fecha"
            GrillaDespachosDespachos.Columns(4).Caption = "Despachador"
            GrillaDespachosDespachos.Columns(5).Caption = "Zona"
            GrillaDespachosDespachos.Columns(6).Caption = "Vendedor"
            

            'alineacion
            GrillaDespachosDespachos.Columns(0).Alignment = dbgCenter
            GrillaDespachosDespachos.Columns(2).Alignment = dbgCenter
            GrillaDespachosDespachos.Columns(3).Alignment = dbgCenter
            GrillaDespachosDespachos.Columns(4).Alignment = dbgCenter

            'cabeceras
            GrillaDespachosDespachos.HeadFont.Bold = True

            'las que no quiero ver
            GrillaDespachosDespachos.Columns(0).Visible = False
            GrillaDespachosDespachos.Columns(2).Visible = False
            GrillaDespachosDespachos.Columns(3).Visible = False
            GrillaDespachosDespachos.Columns(4).Visible = False
            GrillaDespachosDespachos.Columns(5).Visible = False
            GrillaDespachosDespachos.Columns(6).Visible = False

End With
End Sub

Sub EstiloGrillaDespachosContenido()

With RsDetallesDespacho
    'If .BOF Or .EOF Then Exit Sub para que carga el estilo exista o no productos
                
                GrillaDespachosContenido.Columns(0).Width = 1500
                GrillaDespachosContenido.Columns(1).Width = 1500
                GrillaDespachosContenido.Columns(2).Width = 4400
                GrillaDespachosContenido.Columns(3).Width = 1500
                GrillaDespachosContenido.Columns(4).Width = 1500
                GrillaDespachosContenido.Columns(5).Width = 1500
                GrillaDespachosContenido.Columns(6).Width = 1500
                GrillaDespachosContenido.Columns(7).Width = 1500


    
    'caption de las grillas
                GrillaDespachosContenido.Columns(0).Caption = "ID"
                GrillaDespachosContenido.Columns(1).Caption = "Codigo de Despacho"
                GrillaDespachosContenido.Columns(2).Caption = "Codigo de los Producto"
                GrillaDespachosContenido.Columns(3).Caption = "Descripción"
                GrillaDespachosContenido.Columns(4).Caption = "Cantidad"
                GrillaDespachosContenido.Columns(5).Caption = "Marca"
                GrillaDespachosContenido.Columns(6).Caption = "Kit"
                GrillaDespachosContenido.Columns(7).Caption = "Piezas"

    'alineacion
                GrillaDespachosContenido.Columns(0).Alignment = dbgCenter
                GrillaDespachosContenido.Columns(1).Alignment = dbgCenter
                GrillaDespachosContenido.Columns(2).Alignment = dbgCenter
                GrillaDespachosContenido.Columns(3).Alignment = dbgCenter
                GrillaDespachosContenido.Columns(4).Alignment = dbgCenter
                GrillaDespachosContenido.Columns(5).Alignment = dbgCenter
                GrillaDespachosContenido.Columns(6).Alignment = dbgCenter
                GrillaDespachosContenido.Columns(7).Alignment = dbgCenter

'cabeceras
                GrillaDespachosContenido.HeadFont.Bold = True

'las que no quiero ver
                GrillaDespachosContenido.Columns(0).Visible = False
                GrillaDespachosContenido.Columns(1).Visible = False
                GrillaDespachosContenido.Columns(3).Visible = False
                GrillaDespachosContenido.Columns(4).Visible = False
                GrillaDespachosContenido.Columns(5).Visible = False
                GrillaDespachosContenido.Columns(6).Visible = False
                GrillaDespachosContenido.Columns(7).Visible = False
                
End With
End Sub

Private Sub GrillaDespachosContenido_Click()
With RsDetallesDespacho
        .Requery
        If .BOF Or .EOF Or GrillaDespachosContenido.ApproxCount = 0 Then
               Exit Sub
          Else
               .Find "CodigoProducto='" & UCase(Trim(GrillaDespachosContenido.Columns(2).Text)) & "'"
               lbl_DescripcionProducto.Caption = GrillaDespachosContenido.Columns(3).Text
               lbl_MarcaProducto.Caption = GrillaDespachosContenido.Columns(5).Text
               lbl_CantidadVendida.Caption = GrillaDespachosContenido.Columns(4).Text
               
               If !Kit = 0 Then
                   lbl_Kit.Caption = "No"
               Else
                   lbl_Kit.Caption = "Sí"
               End If
               lbl_PiezasxKit.Caption = GrillaDespachosContenido.Columns(7).Text
        End If
End With

End Sub

Private Sub GrillaDespachosDespachos_Click()
With RsDespacho
        If .BOF Or .EOF Or GrillaDespachosDespachos.ApproxCount = 0 Then
               Exit Sub
          Else
               lbl_NombreCliente.Caption = GrillaDespachosDespachos.Columns(2).Text
               lbl_Fecha.Caption = GrillaDespachosDespachos.Columns(3).Text
               lbl_Despachador.Caption = GrillaDespachosDespachos.Columns(4).Text
               lbl_Zona.Caption = GrillaDespachosDespachos.Columns(5).Text
               lbl_Vendedor.Caption = GrillaDespachosDespachos.Columns(6).Text
        End If
End With
BuscarLosDetalles
Devolucion
GrillaDespachosContenido.Enabled = True
End Sub

Private Sub GrillaDespachosUsuarios_Click()
          With RsUsuarios
                    If .EOF And .BOF = True Then MsgBox ("No hay registro disponibles"), vbInformation, "Aviso"
                         Dim x 'varible para guardar el error
                         On Error GoTo x 'guardor el error en la variable
                    lbl_NombreUsuario.Caption = GrillaDespachosUsuarios.Columns(2).Text & " " & GrillaDespachosUsuarios.Columns(3).Text
                    lbl_cedula.Caption = GrillaDespachosUsuarios.Columns(7).Text
                    lbl_tarea.Caption = GrillaDespachosUsuarios.Columns(8).Text
         
                    With AdoTotalDespachosUsuarios
                    .CursorLocation = adUseClient
                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
                    .RecordSource = "select count(*)  from Despacho WHERE[Despachador] like '" & GrillaDespachosUsuarios.Columns(1).Text & "'"
                    .Refresh
                    
                     Set DataGrid1.DataSource = AdoTotalDespachosUsuarios
                    lbl_TotalDespachos.Caption = DataGrid1.Columns(0).Text
                    BuscarLosDespachos
                    GrillaDespachosDespachos.Enabled = True
                    
                    End With
                    Exit Sub
x:                    MsgBox ("No existe un registro disponible para selección"), vbInformation, "Aviso" 'muestro el error
          End With

End Sub

Private Sub opt_Busqueda_Click(Index As Integer)
If opt_Busqueda(0).Value = True Then
          txt_parametro.Text = ""
Else
          txt_parametro.Text = ""
End If
          
End Sub

Private Sub txt_parametro_Change()
If opt_Busqueda(0).Value = True Then BuscaNombre
If opt_Busqueda(1).Value = True Then BuscaCI

GrillaDespachosUsuarios.Enabled = True

End Sub

Sub BuscaNombre()

Dim Busca As String
Busca = UCase(Trim(txt_parametro.Text)) & "%"
AdoUsuarios.RecordSource = "select * from Usuarios where [Nomb_Usuario] like '" & Busca & "'"
AdoUsuarios.Refresh
EstiloGrillaDespachosUsuarios

End Sub

Sub BuscaCI()

Dim Busca As String
Busca = UCase(Trim(txt_parametro.Text)) & "%"
AdoUsuarios.RecordSource = "select * from Usuarios where [CI] like '" & Busca & "'"
AdoUsuarios.Refresh
EstiloGrillaDespachosUsuarios

End Sub

Sub BuscarLosDespachos()

Dim BuscarDespacho As String
BuscarDespacho = UCase(Trim(GrillaDespachosUsuarios.Columns(1).Text)) 'borre porcentaje porque quiero solo los que son iguales no parecidos
AdoDespachos.RecordSource = "select * from Despacho where [Despachador] like '" & BuscarDespacho & "'"
Set GrillaDespachosDespachos.DataSource = AdoDespachos
AdoDespachos.Refresh
EstiloGrillaDespachosDespachos

End Sub

Sub BuscarLosDetalles()

Dim BuscarDetalle As String
BuscarDetalle = UCase(Trim(GrillaDespachosDespachos.Columns(1).Text)) 'borre porcentaje porque quiero solo los que son iguales no parecidos
AdoDetalles.RecordSource = "select * from DetallesDespacho where [CodigoDespacho] like '" & BuscarDetalle & "'"
Set GrillaDespachosContenido.DataSource = AdoDetalles
AdoDetalles.Refresh
EstiloGrillaDespachosContenido

End Sub

Sub Devolucion()
With AdoDevuelto
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
          Busca = UCase(Trim(GrillaDespachosDespachos.Columns(1).Text))
          .RecordSource = "SELECT COUNT (*) FROM Devoluciones WHERE CodigoDespacho LIKE '" & Busca & "'"
          .Refresh
          Set GrillaDevuelto.DataSource = AdoDevuelto
End With
If Val(GrillaDevuelto.Columns(0).Text) = 0 Then
          lbl_Devuelto.Caption = "No"
Else
          lbl_Devuelto.Caption = "Sí"
End If
End Sub
