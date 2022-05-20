VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form DespachosDevolucionesHistorialForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial de Devoluciones"
   ClientHeight    =   10350
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
   Icon            =   "DespachosDevolucionesHistorialForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DespachosDevolucionesHistorialForm.frx":058A
   ScaleHeight     =   10350
   ScaleWidth      =   18000
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid GrillaHistorialDevoluciones 
      Height          =   5475
      Left            =   1150
      TabIndex        =   2
      Top             =   3600
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   9657
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
      Caption         =   "Historial de Devoluciones"
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
   Begin VB.TextBox TxtFiltrar 
      Height          =   390
      Left            =   3930
      TabIndex        =   6
      Top             =   2760
      Width           =   2890
   End
   Begin VB.CommandButton cmd_QuitarFiltro 
      Caption         =   "Quitar Filtro"
      Height          =   580
      Left            =   7100
      TabIndex        =   5
      Top             =   2625
      Width           =   2175
   End
   Begin VB.CommandButton cmd_atras 
      Caption         =   "Atras"
      Height          =   615
      Left            =   6360
      TabIndex        =   1
      Top             =   9480
      Width           =   2175
   End
   Begin VB.CommandButton cmd_ver 
      Caption         =   "Ver"
      Height          =   615
      Left            =   9480
      TabIndex        =   0
      Top             =   9480
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc AdoFiltrarDevoluciones 
      Height          =   330
      Left            =   4200
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
   Begin MSForms.ComboBox CboFiltrar 
      Height          =   390
      Left            =   1050
      TabIndex        =   7
      Top             =   2760
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
   Begin VB.Label lbl_1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione la devolución que desea ver:"
      Height          =   270
      Left            =   1680
      TabIndex        =   4
      Top             =   1875
      Width           =   4185
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
      TabIndex        =   3
      Top             =   1725
      Width           =   225
   End
End
Attribute VB_Name = "DespachosDevolucionesHistorialForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NombreCampo As String
Private Sub cmd_Atras_Click()
Unload Me
End Sub

Private Sub cmd_QuitarFiltro_Click()
TxtFiltrar.Text = ""
End Sub

Private Sub cmd_ver_Click()
With RsDevoluciones
          If .BOF And .EOF = True Then
                    Exit Sub
          Else
                    DespachosDevolucionesHistorialForm.EstilosGrillaDevolucion1
                    If TxtFiltrar.Text = "" Then
                              With RsDevoluciones
                                        If .RecordCount = "0" Then Exit Sub
                              End With
                    Else
                              With AdoFiltrarDevoluciones
                                        If .Recordset.RecordCount = 0 Then Exit Sub
                              End With
                    End If
                    
          vMostrarDetallesDevoluciones = GrillaHistorialDevoluciones.Columns(1).Text
          DespachosDevolucionesVerForm.Show vbModal
          End If
End With
End Sub

Private Sub Form_Load()
Devoluciones
DevolucionesDetalles

Set GrillaHistorialDevoluciones.DataSource = RsDevoluciones
EstilosGrillaDevolucion1

CboFiltrar.AddItem "Código Despacho"
CboFiltrar.AddItem "Cliente"
CboFiltrar.AddItem "Fecha Devolución"
CboFiltrar.AddItem "Motivo"
CboFiltrar.AddItem "Usuario"
CboFiltrar.AddItem "Vendedor"
CboFiltrar.ListIndex = 0
End Sub

Sub EstilosGrillaDevolucion1()
'Tamaño
GrillaHistorialDevoluciones.Columns(0).Width = 0 'ID
GrillaHistorialDevoluciones.Columns(1).Width = 2600 'Codigo de Despacho
GrillaHistorialDevoluciones.Columns(2).Width = 2800 'cliente
GrillaHistorialDevoluciones.Columns(3).Width = 0 'Despachador
GrillaHistorialDevoluciones.Columns(4).Width = 0 'Zona
GrillaHistorialDevoluciones.Columns(5).Width = 2800 'Vendedor
GrillaHistorialDevoluciones.Columns(6).Width = 0 'Fecha
GrillaHistorialDevoluciones.Columns(7).Width = 1300 'Fecha de Devolucion
GrillaHistorialDevoluciones.Columns(8).Width = 4000 'Motivo
GrillaHistorialDevoluciones.Columns(9).Width = 0 'Observacion
GrillaHistorialDevoluciones.Columns(10).Width = 1720 'by
GrillaHistorialDevoluciones.Columns(11).Width = 0 'tipo

'Captions
GrillaHistorialDevoluciones.Columns(0).Caption = "ID" '
GrillaHistorialDevoluciones.Columns(1).Caption = "Código de Despacho"
GrillaHistorialDevoluciones.Columns(2).Caption = "Cliente"
GrillaHistorialDevoluciones.Columns(3).Caption = "Despachador" '
GrillaHistorialDevoluciones.Columns(4).Caption = "Zona" '
GrillaHistorialDevoluciones.Columns(5).Caption = "Vendedor"
GrillaHistorialDevoluciones.Columns(6).Caption = "Fecha V" '
GrillaHistorialDevoluciones.Columns(7).Caption = "Devuelto"
GrillaHistorialDevoluciones.Columns(8).Caption = "Motivo"
GrillaHistorialDevoluciones.Columns(9).Caption = "Observaciones" '
GrillaHistorialDevoluciones.Columns(10).Caption = "Devuelto por"
GrillaHistorialDevoluciones.Columns(11).Caption = "Tipo" '

'alineaciones
GrillaHistorialDevoluciones.Columns(0).Alignment = dbgCenter
GrillaHistorialDevoluciones.Columns(1).Alignment = dbgLeft
GrillaHistorialDevoluciones.Columns(2).Alignment = dbgLeft
GrillaHistorialDevoluciones.Columns(3).Alignment = dbgCenter
GrillaHistorialDevoluciones.Columns(4).Alignment = dbgCenter
GrillaHistorialDevoluciones.Columns(5).Alignment = dbgLeft
GrillaHistorialDevoluciones.Columns(6).Alignment = dbgCenter
GrillaHistorialDevoluciones.Columns(7).Alignment = dbgLeft
GrillaHistorialDevoluciones.Columns(8).Alignment = dbgCenter
GrillaHistorialDevoluciones.Columns(9).Alignment = dbgCenter
GrillaHistorialDevoluciones.Columns(10).Alignment = dbgCenter
GrillaHistorialDevoluciones.Columns(11).Alignment = dbgCenter


'negritas
GrillaHistorialDevoluciones.HeadFont.Bold = True

'las que no quiero ver
GrillaHistorialDevoluciones.Columns(0).Visible = False
GrillaHistorialDevoluciones.Columns(3).Visible = False
GrillaHistorialDevoluciones.Columns(4).Visible = False
GrillaHistorialDevoluciones.Columns(6).Visible = False
GrillaHistorialDevoluciones.Columns(9).Visible = False
GrillaHistorialDevoluciones.Columns(11).Visible = False

End Sub

Sub FiltrarDevoluciones()
          If CboFiltrar.Text = "Código Despacho" Then NombreCampo = "CodigoDespacho"
          If CboFiltrar.Text = "Cliente" Then NombreCampo = "Cliente"
          If CboFiltrar.Text = "Fecha Devolución" Then NombreCampo = "FechaDevo"
          If CboFiltrar.Text = "Motivo" Then NombreCampo = "Motivo"
          If CboFiltrar.Text = "Usuario" Then NombreCampo = "By"
          If CboFiltrar.Text = "Vendedor" Then NombreCampo = "Vendedor"

          With AdoFiltrarDevoluciones
                    .CursorLocation = adUseClient
                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
                    Busca = UCase(Trim(TxtFiltrar.Text)) & "%"
                    .RecordSource = "SELECT* FROM Devoluciones WHERE [" & NombreCampo & "] LIKE '" & Busca & "'"
                    .Refresh
                    Set GrillaHistorialDevoluciones.DataSource = AdoFiltrarDevoluciones
                    EstilosGrillaDevolucion1
          End With
End Sub

Private Sub GrillaHistorialDevoluciones_DblClick()
cmd_ver_Click
End Sub

Private Sub TxtFiltrar_Change()
If CboFiltrar.Text = "" Then Exit Sub
FiltrarDevoluciones
End Sub
