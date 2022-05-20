VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form DespachosVerForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalles del despacho"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17985
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DespachosVerForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DespachosVerForm.frx":058A
   ScaleHeight     =   9930
   ScaleWidth      =   17985
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoVerProductos 
      Height          =   375
      Left            =   840
      Top             =   10560
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
   Begin VB.TextBox txt_Fecha 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2680
      Width           =   3855
   End
   Begin VB.TextBox txt_NombreVendedor 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3520
      Width           =   3855
   End
   Begin VB.TextBox txt_ZonaDespacho 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2680
      Width           =   3855
   End
   Begin VB.TextBox txt_NombreCliente 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3520
      Width           =   3855
   End
   Begin VB.TextBox txt_CodFactura 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2680
      Width           =   3855
   End
   Begin VB.TextBox txt_NombreDespachador 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3520
      Width           =   3855
   End
   Begin VB.CommandButton cmd_Atras 
      Appearance      =   0  'Flat
      Caption         =   "Atras"
      Height          =   615
      Left            =   6360
      TabIndex        =   0
      Top             =   9000
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid GrillaVerDespacho 
      Height          =   2895
      Left            =   1725
      TabIndex        =   7
      Top             =   5445
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   5106
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd_Imprimir 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   9360
      TabIndex        =   1
      Top             =   9000
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc AdoImprimirProductos 
      Height          =   375
      Left            =   11040
      Top             =   10560
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
   Begin VB.Label lbl_DevoParcial 
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
      Left            =   12720
      TabIndex        =   19
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de realización"
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
      TabIndex        =   17
      Top             =   3180
      Width           =   1560
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
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
      Top             =   4005
      Width           =   705
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
      TabIndex        =   15
      Top             =   3180
      Width           =   2220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
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
      TabIndex        =   14
      Top             =   4005
      Width           =   525
   End
   Begin VB.Label lbl_descripcion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Información básica:"
      Height          =   270
      Left            =   1680
      TabIndex        =   13
      Top             =   2100
      Width           =   2025
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
      TabIndex        =   12
      Top             =   1950
      Width           =   225
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
      TabIndex        =   11
      Top             =   3180
      Width           =   1740
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
      TabIndex        =   10
      Top             =   4000
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Productos despachados:"
      Height          =   270
      Left            =   1680
      TabIndex        =   9
      Top             =   4750
      Width           =   2610
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
      TabIndex        =   8
      Top             =   4600
      Width           =   225
   End
End
Attribute VB_Name = "DespachosVerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub VerDespacho()
With RsDespacho
    .Requery
    .Find "Id='" & Val(vMostrarDetallesDespacho) & "'"

'igualamos los campos
txt_CodFactura.Text = !CodigoDespacho
txt_NombreCliente.Text = !Cliente
txt_Fecha.Text = !Fecha
txt_NombreDespachador.Text = !Despachador
txt_ZonaDespacho.Text = !Zona
txt_NombreVendedor.Text = !Vendedor

End With
End Sub

Private Sub cmd_Atras_Click()
Unload Me
vMostrarDetallesDespacho = 0
DespachosHistorialForm.EstilosGrillaDespacho
End Sub

Private Sub cmd_Imprimir_Click()
'igualamos las variables
 vNFactura = txt_CodFactura.Text
vCliente = txt_NombreCliente.Text
vVendedor = txt_NombreVendedor.Text
vDespachador = txt_NombreDespachador.Text
vFecha = txt_Fecha.Text
vZona = txt_ZonaDespacho.Text

'coloque esto aqui porque no me dejaba imprimir correctamente ya que no actualizaba la grila del historial (Formulario anterior)
With RsDetallesDespacho
        If .State = 1 Then .Close
        Busca = UCase(Trim(txt_CodFactura.Text))
        .Open "select* from DetallesDespacho WHERE CodigoDespacho like '" & Busca & "'"
        Set dr_RDespachosIndividual.DataSource = RsDetallesDespacho
End With

'salimos
Unload Me

' s e c c i o n 2
dr_RDespachosIndividual.Sections("Sección2").Controls("Etiqueta12").Caption = vCliente
dr_RDespachosIndividual.Sections("Sección2").Controls("Etiqueta17").Caption = vVendedor
dr_RDespachosIndividual.Sections("Sección2").Controls("Etiqueta18").Caption = vDespachador
dr_RDespachosIndividual.Sections("Sección2").Controls("Etiqueta19").Caption = vFecha
dr_RDespachosIndividual.Sections("Sección2").Controls("Etiqueta20").Caption = vZona

' s e c c i o n 4
dr_RDespachosIndividual.Sections("Sección4").Controls("Etiqueta3").Caption = vNFactura



'mostramos
dr_RDespachosIndividual.Show
dr_RDespachosIndividual.WindowState = 2


End Sub

Private Sub Form_Load()
Despacho
DetallesDespacho
Devoluciones
VerDespacho
VerProductosDelDespacho
ComprobarDevolucion


DespachosHistorialForm.EstilosGrillaDespacho


End Sub
Sub ComprobarDevolucion()
With RsDevoluciones
          .Requery
          .Find "CodigoDespacho='" & Trim(txt_CodFactura.Text) & "'"
          If .EOF = True Then
                    lbl_DevoParcial.Visible = False
          Else
                    lbl_DevoParcial.Visible = True
          End If
End With
End Sub
Sub VerProductosDelDespacho()
With AdoVerProductos
    .CursorLocation = adUseClient
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
    Busca = UCase(Trim(txt_CodFactura.Text))
    .RecordSource = "select * from DetallesDespacho WHERE CodigoDespacho  like '" & Busca & "'"
    .Refresh
End With
    
Set GrillaVerDespacho.DataSource = AdoVerProductos
EstiloGrillaVerDetalles

End Sub


Sub EstiloGrillaVerDetalles()

                GrillaVerDespacho.Columns(0).Width = 0
                GrillaVerDespacho.Columns(1).Width = 0
                GrillaVerDespacho.Columns(2).Width = 2750
                GrillaVerDespacho.Columns(3).Width = 8450
                GrillaVerDespacho.Columns(4).Width = 1500
                GrillaVerDespacho.Columns(5).Width = 1500
                GrillaVerDespacho.Columns(6).Width = 0
                GrillaVerDespacho.Columns(7).Width = 0
    
    'caption de las grillas
                GrillaVerDespacho.Columns(0).Caption = "ID"
                GrillaVerDespacho.Columns(1).Caption = "Código Despacho"
                GrillaVerDespacho.Columns(2).Caption = "Código Producto"
                GrillaVerDespacho.Columns(3).Caption = "Descripción"
                GrillaVerDespacho.Columns(4).Caption = "Cantidad"
                GrillaVerDespacho.Columns(5).Caption = "Marca"
                GrillaVerDespacho.Columns(6).Caption = "Kit"
                GrillaVerDespacho.Columns(7).Caption = "Piezas por Kit"

    'alineacion
                GrillaVerDespacho.Columns(0).Alignment = dbgCenter
                GrillaVerDespacho.Columns(1).Alignment = dbgCenter
                GrillaVerDespacho.Columns(2).Alignment = dbgCenter
                GrillaVerDespacho.Columns(3).Alignment = dbgLeft
                GrillaVerDespacho.Columns(4).Alignment = dbgCenter
                GrillaVerDespacho.Columns(5).Alignment = dbgCenter
                GrillaVerDespacho.Columns(6).Alignment = dbgCenter
                GrillaVerDespacho.Columns(7).Alignment = dbgCenter


'cabeceras
                GrillaVerDespacho.HeadFont.Bold = True

'las que no quiero ver
                GrillaVerDespacho.Columns(0).Visible = False
                GrillaVerDespacho.Columns(1).Visible = False
                GrillaVerDespacho.Columns(6).Visible = False
                GrillaVerDespacho.Columns(7).Visible = False

End Sub

