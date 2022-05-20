VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form DespachosShippingSelectForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Despacho"
   ClientHeight    =   11715
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
   Icon            =   "DespachosShippingSelectForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DespachosShippingSelectForm.frx":058A
   ScaleHeight     =   11715
   ScaleWidth      =   13500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_salir 
      Caption         =   "Salir"
      Height          =   615
      Left            =   4095
      TabIndex        =   4
      Top             =   10800
      Width           =   2175
   End
   Begin VB.CommandButton cmd_seleccionar 
      Caption         =   "Seleccionar"
      Height          =   615
      Left            =   7335
      TabIndex        =   3
      Top             =   10800
      Width           =   2175
   End
   Begin VB.CommandButton cmd_QuitarFiltro 
      Caption         =   "Quitar Filtro"
      Height          =   580
      Left            =   7215
      TabIndex        =   2
      Top             =   2805
      Width           =   2175
   End
   Begin VB.TextBox TxtFiltrar 
      Height          =   390
      Left            =   4080
      TabIndex        =   1
      Top             =   2895
      Width           =   2890
   End
   Begin MSDataGridLib.DataGrid GrillaHistorialDespachoShipping 
      Height          =   6300
      Left            =   1290
      TabIndex        =   0
      Top             =   3960
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   11113
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
      Caption         =   "Historial de despachos realizados"
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
   Begin MSAdodcLib.Adodc AdoFiltrarDespachosShipping 
      Height          =   330
      Left            =   4335
      Top             =   1200
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
   Begin VB.Label lbl_1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el Despacho al que desea crearle etiquetas de envío:"
      Height          =   270
      Left            =   1695
      TabIndex        =   7
      Top             =   2160
      Width           =   6750
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
      Left            =   1335
      TabIndex        =   6
      Top             =   2040
      Width           =   225
   End
   Begin MSForms.ComboBox CboFiltrar 
      Height          =   390
      Left            =   1200
      TabIndex        =   5
      Top             =   2895
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
End
Attribute VB_Name = "DespachosShippingSelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NombreCampo As String

Private Sub cmd_QuitarFiltro_Click()
TxtFiltrar.Text = ""
End Sub

Private Sub cmd_Salir_Click()
Unload Me
End Sub

Private Sub cmd_seleccionar_Click()
' verificamos si la tabla esta vacia
With RsDespacho
          If .BOF And .EOF = True Then
                    Exit Sub
          Else
                    DespachosShippingSelectForm.EstiloGrillaDespachoShipping
                    If TxtFiltrar.Text = "" Then
                              With RsDespacho
                                        If .RecordCount = "0" Then Exit Sub
                              End With
                    Else
                              With AdoFiltrarDespachosShipping
                                        If .Recordset.RecordCount = 0 Then Exit Sub
                              End With
                    End If
          'obtener el codigo de usuario
          vCodigoShipping = GrillaHistorialDespachoShipping.Columns(1).Text
          'llamamos al fonrmulario de vista de producto
          Unload Me
          DespachosShippingDetalles.Show vbModal
          End If
End With
End Sub

Private Sub Form_Load()
CboFiltrar.AddItem "Código despacho"
CboFiltrar.AddItem "Cliente"
CboFiltrar.AddItem "Despachador"
CboFiltrar.AddItem "Fecha"
CboFiltrar.AddItem "Zona"
CboFiltrar.ListIndex = 0

Set GrillaHistorialDespachoShipping.DataSource = RsDespacho
EstiloGrillaDespachoShipping
End Sub

Sub EstiloGrillaDespachoShipping()
'tamaños
GrillaHistorialDespachoShipping.Columns(0).Width = 0
GrillaHistorialDespachoShipping.Columns(1).Width = 2500
GrillaHistorialDespachoShipping.Columns(2).Width = 2800
GrillaHistorialDespachoShipping.Columns(3).Width = 1600
GrillaHistorialDespachoShipping.Columns(4).Width = 1700
GrillaHistorialDespachoShipping.Columns(5).Width = 1700
GrillaHistorialDespachoShipping.Columns(6).Width = 1700

'captions
GrillaHistorialDespachoShipping.Columns(0).Caption = "ID"
GrillaHistorialDespachoShipping.Columns(1).Caption = "Código de Despacho"
GrillaHistorialDespachoShipping.Columns(2).Caption = "Cliente"
GrillaHistorialDespachoShipping.Columns(3).Caption = "Fecha"
GrillaHistorialDespachoShipping.Columns(4).Caption = "Despachador"
GrillaHistorialDespachoShipping.Columns(5).Caption = "Zona"
GrillaHistorialDespachoShipping.Columns(6).Caption = "Fecha"

'alianeacion
GrillaHistorialDespachoShipping.Columns(0).Alignment = dbgCenter
GrillaHistorialDespachoShipping.Columns(3).Alignment = dbgCenter
GrillaHistorialDespachoShipping.Columns(4).Alignment = dbgCenter
GrillaHistorialDespachoShipping.Columns(5).Alignment = dbgCenter

'cabeceras
GrillaHistorialDespachoShipping.HeadFont.Bold = True

'las que no quiero ver
GrillaHistorialDespachoShipping.Columns(0).Visible = False
GrillaHistorialDespachoShipping.Columns(6).Visible = False

End Sub

Sub FiltrarDespachosShipping()
'obtenemos el nombre del campo por el que vamos a filtrar
          If CboFiltrar.Text = "Código despacho" Then NombreCampo = "CodigoDespacho"
          If CboFiltrar.Text = "Cliente" Then NombreCampo = "Cliente"
          If CboFiltrar.Text = "Despachador" Then NombreCampo = "Despachador"
          If CboFiltrar.Text = "Fecha" Then NombreCampo = "Fecha"
          If CboFiltrar.Text = "Zona" Then NombreCampo = "Zona"
          
          'programamos el filtro
          With AdoFiltrarDespachosShipping
                    .CursorLocation = adUseClient
                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
                    Busca = UCase(Trim(TxtFiltrar.Text)) & "%"
                    .RecordSource = "SELECT * FROM Despacho WHERE [" & NombreCampo & "] LIKE '" & Busca & "'"
                    .Refresh
          End With
          Set GrillaHistorialDespachoShipping.DataSource = AdoFiltrarDespachosShipping
          EstiloGrillaDespachoShipping
End Sub
Private Sub GrillaHistorialDespachoShipping_DblClick()
cmd_seleccionar_Click
End Sub

Private Sub TxtFiltrar_Change()
If CboFiltrar.Text = "" Then Exit Sub
FiltrarDespachosShipping
End Sub
