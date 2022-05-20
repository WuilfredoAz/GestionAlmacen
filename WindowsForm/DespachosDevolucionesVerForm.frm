VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form DespachosDevolucionesVerForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalles de la Devolucion"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DespachosDevolucionesVerForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DespachosDevolucionesVerForm.frx":058A
   ScaleHeight     =   10500
   ScaleWidth      =   18000
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid GrillaContenidoDevolucion 
      Height          =   2895
      Left            =   1755
      TabIndex        =   10
      Top             =   6075
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   5106
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
      Caption         =   "Lista de productos del Despacho"
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoVerProductosDevueltos 
      Height          =   330
      Left            =   1080
      Top             =   10560
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
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
   Begin VB.TextBox txt_FechaDevo 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   14880
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2610
      Width           =   1815
   End
   Begin VB.TextBox txt_Motivo 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   4305
      Width           =   3855
   End
   Begin VB.CommandButton cmd_Atras 
      Appearance      =   0  'Flat
      Caption         =   "Atras"
      Height          =   615
      Left            =   5640
      TabIndex        =   9
      Top             =   9480
      Width           =   2295
   End
   Begin VB.CommandButton cmd_Imprimir 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   9600
      TabIndex        =   8
      Top             =   9480
      Width           =   2295
   End
   Begin VB.TextBox txt_NombreDespachador 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3450
      Width           =   3855
   End
   Begin VB.TextBox txt_CodFactura 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2610
      Width           =   3855
   End
   Begin VB.TextBox txt_NombreCliente 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3450
      Width           =   3855
   End
   Begin VB.TextBox txt_ZonaDespacho 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2610
      Width           =   3855
   End
   Begin VB.TextBox txt_NombreVendedor 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3450
      Width           =   3855
   End
   Begin VB.TextBox txt_Fecha 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2610
      Width           =   1815
   End
   Begin VB.TextBox txt_Observaciones 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   4305
      Width           =   3855
   End
   Begin VB.TextBox txt_DevueltoPor 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   4305
      Width           =   3855
   End
   Begin VB.Label lbl_DevoTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   615
      Left            =   14280
      TabIndex        =   27
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label8 
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
      Left            =   1320
      TabIndex        =   24
      Top             =   5190
      Width           =   225
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vista previa de los productos que fueron devueltos"
      Height          =   270
      Left            =   1680
      TabIndex        =   23
      Top             =   5340
      Width           =   5235
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Despachado por"
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
      Left            =   12840
      TabIndex        =   22
      Top             =   3945
      Width           =   1215
   End
   Begin VB.Label lbl_EJusername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código de la FACTURA"
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
      Left            =   1920
      TabIndex        =   21
      Top             =   3105
      Width           =   1740
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
      Left            =   1320
      TabIndex        =   20
      Top             =   1880
      Width           =   225
   End
   Begin VB.Label lbl_descripcion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vista previa de la información general"
      Height          =   270
      Left            =   1680
      TabIndex        =   19
      Top             =   2040
      Width           =   3930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Cliente"
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
      Left            =   1920
      TabIndex        =   18
      Top             =   3945
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zona a donde fue el despacho"
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
      Left            =   7200
      TabIndex        =   17
      Top             =   3105
      Width           =   2220
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del vendedor"
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
      Left            =   7200
      TabIndex        =   16
      Top             =   3945
      Width           =   1590
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de despacho"
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
      Left            =   12840
      TabIndex        =   15
      Top             =   3105
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de devolución"
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
      Left            =   15000
      TabIndex        =   14
      Top             =   3105
      Width           =   1530
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Motivo de la devolución"
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
      Left            =   1920
      TabIndex        =   13
      Top             =   4800
      Width           =   1725
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones pertinentes a la devolución"
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
      Left            =   7200
      TabIndex        =   12
      Top             =   4800
      Width           =   3165
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Devuelto por"
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
      Left            =   12840
      TabIndex        =   11
      Top             =   4800
      Width           =   930
   End
End
Attribute VB_Name = "DespachosDevolucionesVerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Atras_Click()
DespachosDevolucionesHistorialForm.EstilosGrillaDevolucion1
vMostrarDetallesDevoluciones = 0
Unload Me
End Sub

Private Sub cmd_Imprimir_Click() 'devoluciones
'igualamos las variables
vNFacturaDevo = txt_CodFactura.Text
vClienteDevo = txt_NombreCliente.Text
vMotivo = txt_Motivo.Text
vZonaDevo = txt_ZonaDespacho.Text
vVendedorDevo = txt_NombreVendedor.Text
vObservaciones = txt_Observaciones.Text
vEntregado = txt_Fecha.Text
vDevuelto = txt_FechaDevo.Text
vDespachadorDevo = txt_NombreDespachador.Text
vDevueltoPor = txt_DevueltoPor.Text

'buscamos los productos devueltos
With RsDevolucionesDetalles
          If .State = 1 Then .Close
          Busca = UCase(Trim(txt_CodFactura.Text))
          .Open "SELECT * FROM DevolucionesDetalles WHERE CodigoDespacho LIKE '" & Busca & "'"
          Set dr_RDevolucionIndividual.DataSource = RsDevolucionesDetalles
End With

'salimos
Unload Me

'Igualamos las secciones -->
'Seccion 4
dr_RDevolucionIndividual.Sections("Sección4").Controls("Etiqueta3").Caption = vNFacturaDevo

'Seccion 2
dr_RDevolucionIndividual.Sections("Sección2").Controls("Etiqueta12").Caption = vNFacturaDevo
dr_RDevolucionIndividual.Sections("Sección2").Controls("Etiqueta17").Caption = vClienteDevo
dr_RDevolucionIndividual.Sections("Sección2").Controls("Etiqueta18").Caption = vDespachadorDevo
dr_RDevolucionIndividual.Sections("Sección2").Controls("Etiqueta22").Caption = vDevueltoPor
dr_RDevolucionIndividual.Sections("Sección2").Controls("Etiqueta24").Caption = vMotivo
dr_RDevolucionIndividual.Sections("Sección2").Controls("Etiqueta20").Caption = vZonaDevo
dr_RDevolucionIndividual.Sections("Sección2").Controls("Etiqueta19").Caption = vEntregado
dr_RDevolucionIndividual.Sections("Sección2").Controls("Etiqueta26").Caption = vDevuelto
dr_RDevolucionIndividual.Sections("Sección2").Controls("Etiqueta28").Caption = vObservaciones

'Mostramos
dr_RDevolucionIndividual.Show
dr_RDevolucionIndividual.WindowState = 2
End Sub

Private Sub Form_Load()
Devoluciones
DevolucionesDetalles
Despacho
VerInformacion
VerProductosDevueltos
DespachosDevolucionesHistorialForm.EstilosGrillaDevolucion1
ComprobarDevolucion
End Sub

Sub VerInformacion()
With RsDevoluciones
          .Requery
          .Find "CodigoDespacho='" & Trim(vMostrarDetallesDevoluciones) & "'"
                    txt_CodFactura.Text = !CodigoDespacho
                    txt_NombreCliente.Text = !Cliente
                    txt_Motivo.Text = !Motivo
                    txt_ZonaDespacho.Text = !Zona
                    txt_NombreVendedor.Text = !Vendedor
                    txt_Observaciones.Text = !Observaciones
                    txt_Fecha.Text = !Fecha
                    txt_FechaDevo.Text = !FechaDevo
                    txt_NombreDespachador.Text = !Despachador
                    txt_DevueltoPor.Text = !By
End With
End Sub

Sub VerProductosDevueltos()
With AdoVerProductosDevueltos
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
          Busca = UCase(Trim(txt_CodFactura.Text))
          .RecordSource = "Select * from DevolucionesDetalles WHERE CodigoDespacho like '" & Busca & "'"
          .Refresh
End With

Set GrillaContenidoDevolucion.DataSource = AdoVerProductosDevueltos
EstilosGrillaDevolucion2
End Sub
Sub ComprobarDevolucion()
With RsDespacho
          .Requery
          .Find "CodigoDespacho='" & Trim(vMostrarDetallesDevoluciones) & "'"
          If .EOF = True Then
                    lbl_DevoTotal.Visible = True
                    lbl_DevoTotal.ForeColor = &H8000000D
          Else
                    lbl_DevoTotal.Visible = True
                    lbl_DevoTotal.ForeColor = &H80000000
          End If
End With
End Sub
Sub EstilosGrillaDevolucion2()
     'tamaños
                GrillaContenidoDevolucion.Columns(0).Width = 0
                GrillaContenidoDevolucion.Columns(1).Width = 0
                GrillaContenidoDevolucion.Columns(2).Width = 4150
                GrillaContenidoDevolucion.Columns(3).Width = 7000
                GrillaContenidoDevolucion.Columns(4).Width = 1500
                GrillaContenidoDevolucion.Columns(5).Width = 1500
                GrillaContenidoDevolucion.Columns(6).Width = 0
                GrillaContenidoDevolucion.Columns(7).Width = 0
    
    'caption de las grillas
                GrillaContenidoDevolucion.Columns(0).Caption = "ID"
                GrillaContenidoDevolucion.Columns(1).Caption = "Código Despacho"
                GrillaContenidoDevolucion.Columns(2).Caption = "Código Producto"
                GrillaContenidoDevolucion.Columns(3).Caption = "Descripción"
                GrillaContenidoDevolucion.Columns(4).Caption = "Cantidad"
                GrillaContenidoDevolucion.Columns(5).Caption = "Marca"
                GrillaContenidoDevolucion.Columns(6).Caption = "Kit"
                GrillaContenidoDevolucion.Columns(7).Caption = "Piezas por Kit"

    'alineacion
                GrillaContenidoDevolucion.Columns(0).Alignment = dbgCenter
                GrillaContenidoDevolucion.Columns(1).Alignment = dbgCenter
                GrillaContenidoDevolucion.Columns(2).Alignment = dbgCenter
                GrillaContenidoDevolucion.Columns(3).Alignment = dbgCenter
                GrillaContenidoDevolucion.Columns(4).Alignment = dbgCenter
                GrillaContenidoDevolucion.Columns(5).Alignment = dbgCenter
                GrillaContenidoDevolucion.Columns(6).Alignment = dbgCenter
                GrillaContenidoDevolucion.Columns(7).Alignment = dbgCenter


'cabeceras
                GrillaContenidoDevolucion.HeadFont.Bold = True

'las que no quiero ver
                GrillaContenidoDevolucion.Columns(0).Visible = False
                GrillaContenidoDevolucion.Columns(1).Visible = False
                GrillaContenidoDevolucion.Columns(6).Visible = False
                GrillaContenidoDevolucion.Columns(7).Visible = False

End Sub
