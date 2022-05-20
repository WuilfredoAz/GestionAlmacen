VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form ConsultasForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultas"
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
   Icon            =   "ConsultasForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "ConsultasForm.frx":058A
   ScaleHeight     =   10500
   ScaleWidth      =   16500
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid GrillaConsultas 
      Height          =   4695
      Left            =   3120
      TabIndex        =   21
      Top             =   4560
      Width           =   11535
      _ExtentX        =   20346
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
   Begin MSAdodcLib.Adodc AdoFiltrarUsuario 
      Height          =   330
      Left            =   6600
      Top             =   10560
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
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
   Begin VB.CommandButton cmd_QuitarFiltro 
      Caption         =   "Quitar Filtro"
      Height          =   580
      Left            =   9120
      TabIndex        =   11
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox TxtFiltrar 
      Height          =   390
      Left            =   6000
      TabIndex        =   10
      Top             =   3720
      Width           =   2890
   End
   Begin VB.CommandButton cmd_verproducto 
      Caption         =   "Ver producto"
      Height          =   495
      Left            =   12720
      TabIndex        =   12
      Top             =   9600
      Width           =   2175
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
      BackColor       =   &H8000000A&
      Caption         =   "Consultas"
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   0
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
   Begin MSForms.ComboBox CboFiltrar 
      Height          =   390
      Left            =   3120
      TabIndex        =   9
      Top             =   3720
      Width           =   2535
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "4471;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTOS REGISTRADOS"
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
      TabIndex        =   20
      Top             =   2640
      Width           =   6030
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      MouseIcon       =   "ConsultasForm.frx":1A972
      MousePointer    =   99  'Custom
      TabIndex        =   14
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
      MouseIcon       =   "ConsultasForm.frx":1AC7C
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   1560
      Width           =   675
   End
End
Attribute VB_Name = "ConsultasForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NombreCampo As String 'Portador del nombre campo que contenga la tabla (o sea campo por el que quiera filtrar)

Private Sub cmd_despachos_Click()
ConsultasForm.Hide
DespachosForm.Show
End Sub

Private Sub cmd_inicio_Click()
ConsultasForm.Hide
IndexForm.Show
End Sub

Private Sub cmd_Logs_Click()
LogsForm.Show
End Sub

Private Sub cmd_Mantenimiento_Click()
ConsultasForm.Hide
MantenimientoForm.Show
End Sub

Private Sub cmd_productos_Click()
ConsultasForm.Hide
ProductosForm.Show
End Sub

Private Sub cmd_QuitarFiltro_Click()
TxtFiltrar.Text = ""
End Sub

Private Sub cmd_reportes_Click()
ConsultasForm.Hide
ReportesForm.Show
End Sub

Private Sub cmd_ubicaciones_Click()
ConsultasForm.Hide
UbicacionesForm.Show
End Sub

Private Sub cmd_usuarios_Click()
ConsultasForm.Hide
UsuariosForm.Show
End Sub

Private Sub cmd_verproducto_Click()
With RsProductos 'esto para que cuando haga clic y este vacio no haga nada
     If .BOF And .EOF = True Then
          Exit Sub
     Else
          ConsultasForm.EstilosGrillaConsultas
          ' verificamos si la tabla esta vacia
          If TxtFiltrar.Text = "" Then
               With RsProductos
                    If RecordCount = "0" Then Exit Sub
               End With
          Else
          With AdoFiltrarUsuario
               If .Recordset.RecordCount = 0 Then Exit Sub
          End With
          End If
'obtener el codigo de usuario
vMostrarProducto = GrillaConsultas.Columns(0).Text
'llamamos al fonrmulario de vista de producto
ConsultasVerForm.Show 'QUITE EL MODAL DE AQUI
     End If
End With
End Sub

Private Sub Form_Load()
lbl_tarea.Caption = vTarea
lbl_username.Caption = vUsername

'llenamo la grilla con los productos
Productos
Set GrillaConsultas.DataSource = RsProductos
EstilosGrillaConsultas 'me decia error 9 porque primero debo llenar la grilla antes de meter el estilo yo metia estilo y despues llenaba

'Colocamos las opciones del filtro
CboFiltrar.AddItem "Código Producto"
CboFiltrar.AddItem "Descripción Producto"

'campo del filtro predetermiado
CboFiltrar.ListIndex = 0

End Sub

Private Sub GrillaConsultas_Click()
GrillaConsultas.MarqueeStyle = dbgHighlightRowRaiseCell
End Sub

Private Sub GrillaConsultas_DblClick()
cmd_verproducto_Click
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

Sub EstilosGrillaConsultas()

'tamaños de la grilla
GrillaConsultas.Columns(0).Width = 500
GrillaConsultas.Columns(1).Width = 2500
GrillaConsultas.Columns(2).Width = 4800
GrillaConsultas.Columns(3).Width = 2000
GrillaConsultas.Columns(4).Width = 1200
GrillaConsultas.Columns(5).Width = 2300
GrillaConsultas.Columns(6).Width = 500
GrillaConsultas.Columns(7).Width = 1800
GrillaConsultas.Columns(8).Width = 2400
GrillaConsultas.Columns(9).Width = 2400
GrillaConsultas.Columns(10).Width = 2400
GrillaConsultas.Columns(11).Width = 2400
GrillaConsultas.Columns(12).Width = 2400


'Caption de las grillas
GrillaConsultas.Columns(0).Caption = "ID"
GrillaConsultas.Columns(1).Caption = "Código del producto"
GrillaConsultas.Columns(2).Caption = "Descripción del Producto"
GrillaConsultas.Columns(3).Caption = "Aplicable para"
GrillaConsultas.Columns(4).Caption = "Cantidad"
GrillaConsultas.Columns(5).Caption = "Fecha de Recibido"
GrillaConsultas.Columns(6).Caption = "Kit"
GrillaConsultas.Columns(7).Caption = "Piezas por kit"
GrillaConsultas.Columns(8).Caption = "Cant. Max. para la Venta"
GrillaConsultas.Columns(9).Caption = "Cant. Min. para la Venta"
GrillaConsultas.Columns(10).Caption = "Ubicación"
GrillaConsultas.Columns(11).Caption = "Marca"
GrillaConsultas.Columns(12).Caption = "REF"

'alineacion
GrillaConsultas.Columns(0).Alignment = dbgCenter
GrillaConsultas.Columns(2).Alignment = dbgLeft
GrillaConsultas.Columns(3).Alignment = dbgCenter
GrillaConsultas.Columns(4).Alignment = dbgCenter
GrillaConsultas.Columns(5).Alignment = dbgCenter
GrillaConsultas.Columns(6).Alignment = dbgCenter
GrillaConsultas.Columns(7).Alignment = dbgCenter
GrillaConsultas.Columns(8).Alignment = dbgCenter
GrillaConsultas.Columns(9).Alignment = dbgCenter
GrillaConsultas.Columns(10).Alignment = dbgCenter
GrillaConsultas.Columns(11).Alignment = dbgCenter

'cabeceras
GrillaConsultas.HeadFont.Bold = True

'las que no quiero ver
GrillaConsultas.Columns(0).Visible = False
GrillaConsultas.Columns(3).Visible = False
GrillaConsultas.Columns(5).Visible = False
GrillaConsultas.Columns(6).Visible = False
GrillaConsultas.Columns(7).Visible = False
GrillaConsultas.Columns(8).Visible = False
GrillaConsultas.Columns(9).Visible = False
GrillaConsultas.Columns(10).Visible = False
GrillaConsultas.Columns(12).Visible = False

End Sub

Sub FiltrarUsuarios()

If CboFiltrar.Text = "Código Producto" Then NombreCampo = "CodigoProducto"
If CboFiltrar.Text = "Descripción Producto" Then NombreCampo = "Descripcion"

'programar filtro
With AdoFiltrarUsuario
    .CursorLocation = adUseClient
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
    Busca = UCase(Trim(TxtFiltrar.Text)) & "%"
    
     .RecordSource = "select * from Productos WHERE[" & NombreCampo & "] like '" & Busca & "' order by Descripcion asc"
     .Refresh
End With

Set GrillaConsultas.DataSource = AdoFiltrarUsuario
EstilosGrillaConsultas
End Sub

Private Sub TxtFiltrar_Change()
If CboFiltrar.Text = "" Then Exit Sub
FiltrarUsuarios
EstilosGrillaConsultas
End Sub
