VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form UbicacionesSeccionesForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Secciones"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "UbicacionesSeccionesForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "UbicacionesSeccionesForm.frx":058A
   ScaleHeight     =   9330
   ScaleWidth      =   13500
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid GrillaSecciones 
      Height          =   4695
      Left            =   1200
      TabIndex        =   2
      Top             =   3405
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8281
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
      Caption         =   "Productos existentes"
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
   Begin VB.CommandButton cmd_Atras 
      Caption         =   "Atras"
      Height          =   615
      Left            =   5640
      TabIndex        =   6
      Top             =   8520
      Width           =   2175
   End
   Begin VB.CommandButton cmd_QuitarFiltro 
      Caption         =   "Quitar Filtro"
      Height          =   580
      Left            =   7080
      TabIndex        =   1
      Top             =   2445
      Width           =   2175
   End
   Begin VB.TextBox TxtFiltrar 
      Height          =   390
      Left            =   3960
      TabIndex        =   0
      Top             =   2565
      Width           =   2950
   End
   Begin MSAdodcLib.Adodc AdoSecciones 
      Height          =   330
      Left            =   4560
      Top             =   1200
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
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
      TabIndex        =   5
      Top             =   1800
      Width           =   225
   End
   Begin VB.Label lbl_1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de productos ubicados en el segmento:"
      Height          =   270
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   4665
   End
   Begin MSForms.ComboBox CboFiltrar 
      Height          =   390
      Left            =   1080
      TabIndex        =   3
      Top             =   2565
      Width           =   2595
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4568;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "UbicacionesSeccionesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CampoSeccion As String
Sub EstilosSecciones()
GrillaSecciones.Columns(0).Width = 500
GrillaSecciones.Columns(1).Width = 2500
GrillaSecciones.Columns(2).Width = 4800
GrillaSecciones.Columns(3).Width = 2000
GrillaSecciones.Columns(4).Width = 1200
GrillaSecciones.Columns(5).Width = 2300
GrillaSecciones.Columns(6).Width = 500
GrillaSecciones.Columns(7).Width = 1800
GrillaSecciones.Columns(8).Width = 2400
GrillaSecciones.Columns(9).Width = 2400
GrillaSecciones.Columns(10).Width = 2400
GrillaSecciones.Columns(11).Width = 2300
GrillaSecciones.Columns(12).Width = 2400


'Caption de las grillas
GrillaSecciones.Columns(0).Caption = "ID"
GrillaSecciones.Columns(1).Caption = "Código del producto"
GrillaSecciones.Columns(2).Caption = "Descripción del Producto"
GrillaSecciones.Columns(3).Caption = "Aplicable para"
GrillaSecciones.Columns(4).Caption = "Cantidad"
GrillaSecciones.Columns(5).Caption = "Fecha de Recibido"
GrillaSecciones.Columns(6).Caption = "Kit"
GrillaSecciones.Columns(7).Caption = "Piezas por kit"
GrillaSecciones.Columns(8).Caption = "Cant. Max. para la Venta"
GrillaSecciones.Columns(9).Caption = "Cant. Min. para la Venta"
GrillaSecciones.Columns(10).Caption = "Ubicación"
GrillaSecciones.Columns(11).Caption = "Marca"
GrillaSecciones.Columns(12).Caption = "REF"

'alineacion
GrillaSecciones.Columns(0).Alignment = dbgCenter
GrillaSecciones.Columns(2).Alignment = dbgLeft
GrillaSecciones.Columns(3).Alignment = dbgCenter
GrillaSecciones.Columns(4).Alignment = dbgCenter
GrillaSecciones.Columns(5).Alignment = dbgCenter
GrillaSecciones.Columns(6).Alignment = dbgCenter
GrillaSecciones.Columns(7).Alignment = dbgCenter
GrillaSecciones.Columns(8).Alignment = dbgCenter
GrillaSecciones.Columns(9).Alignment = dbgCenter
GrillaSecciones.Columns(10).Alignment = dbgCenter
GrillaSecciones.Columns(11).Alignment = dbgCenter

'cabeceras
GrillaSecciones.HeadFont.Bold = True

'las que no quiero ver
GrillaSecciones.Columns(0).Visible = False
GrillaSecciones.Columns(3).Visible = False
GrillaSecciones.Columns(5).Visible = False
GrillaSecciones.Columns(6).Visible = False
GrillaSecciones.Columns(7).Visible = False
GrillaSecciones.Columns(8).Visible = False
GrillaSecciones.Columns(9).Visible = False
GrillaSecciones.Columns(10).Visible = False
GrillaSecciones.Columns(12).Visible = False
End Sub

Private Sub cmd_Atras_Click()
vSeccion = ""

With RsProductos
          If .State = 1 Then .Close
          .Open "SELECT * FROM Productos ORDER BY Descripcion ASC"
          .Requery
End With

Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
ConsultasForm.EstilosGrillaConsultas
Unload Me
End Sub

Private Sub cmd_QuitarFiltro_Click()
TxtFiltrar.Text = ""
End Sub

Private Sub Form_Load()
CboFiltrar.AddItem "Codigo de producto"
CboFiltrar.AddItem "Descripcion producto"

CboFiltrar.ListIndex = 0

End Sub

Sub BuscarProductos()
If CboFiltrar.Text = "Codigo de producto" Then CampoSeccion = "CodigoProducto"
If CboFiltrar.Text = "Descripcion producto" Then CampoSeccion = "Descripcion"

With AdoSecciones
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
          Busca = UCase(Trim(TxtFiltrar.Text)) & "%"
          .RecordSource = "SELECT * FROM Productos WHERE [" & CampoSeccion & "] LIKE '" & Busca & "' and Ubicacion LIKE '" & vSeccion & "' ORDER BY Descripcion ASC"
          .Refresh
          Set GrillaSecciones.DataSource = AdoSecciones
          EstilosSecciones
End With
End Sub

Private Sub Form_Unload(Cancel As Integer) 'evento cuando cierra con la x
vSeccion = ""
With RsProductos
          If .State = 1 Then .Close
          .Open "SELECT * FROM Productos ORDER BY Descripcion ASC"
          .Requery
End With
Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
ConsultasForm.EstilosGrillaConsultas
End Sub

Private Sub TxtFiltrar_Change()
If CboFiltrar.Text = "" Then Exit Sub
BuscarProductos
EstilosSecciones
End Sub
