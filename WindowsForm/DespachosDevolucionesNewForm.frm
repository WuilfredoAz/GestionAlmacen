VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form DespachosDevolucionesNewForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar devolución"
   ClientHeight    =   10155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DespachosDevolucionesNewForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DespachosDevolucionesNewForm.frx":058A
   ScaleHeight     =   10155
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid GrillaHistorialDespachoDevoluciones 
      Height          =   5535
      Left            =   1320
      TabIndex        =   2
      Top             =   2880
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9763
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
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   3
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
   Begin VB.CommandButton cmd_atras 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   9360
      Width           =   2175
   End
   Begin VB.CommandButton cmd_Ver 
      Caption         =   "Ver"
      Height          =   615
      Left            =   5400
      TabIndex        =   0
      Top             =   9360
      Width           =   2175
   End
   Begin VB.Label lbl_1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el Despacho que desea devolver:"
      Height          =   270
      Left            =   1320
      TabIndex        =   4
      Top             =   2160
      Width           =   4710
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
      TabIndex        =   3
      Top             =   2040
      Width           =   225
   End
End
Attribute VB_Name = "DespachosDevolucionesNewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Atras_Click()
Unload Me
End Sub

Private Sub cmd_ver_Click()
' verificamos si la tabla esta vacia escojo la tabla de despacho porque en base a los despachos realizados es que se realizan las devoluciones
'sin despachos realizados no se puede realizar una devolucion **
DespachosDevolucionesNewForm.BorrarTemporalDevolucion
With RsDespacho
        If .BOF And .EOF = True Then
               Exit Sub
        Else
               vMostrarDetallesDevoluciones = GrillaHistorialDespachoDevoluciones.Columns(0).Text
               DespachosDevolucionesDetallesForm.Show vbModal
          End If
End With
'cargamos el estilo
DespachosDevolucionesNewForm.EstilosGrillaDespachoDevoluciones

End Sub

Private Sub Form_Load()
'tabla necesaria
Despacho
TemporalDevoluciones
'valores a la grilla
Set GrillaHistorialDespachoDevoluciones.DataSource = RsDespacho
EstilosGrillaDespachoDevoluciones
'borramos el temporal de las devoluciones
'BorrarTemporalDevolucion
End Sub

Sub EstilosGrillaDespachoDevoluciones()
'tamaños de la grilla
GrillaHistorialDespachoDevoluciones.Columns(0).Width = 500 'id
GrillaHistorialDespachoDevoluciones.Columns(1).Width = 1700 'cod despacho
GrillaHistorialDespachoDevoluciones.Columns(2).Width = 2600 'cliente
GrillaHistorialDespachoDevoluciones.Columns(3).Width = 1600 'fecha
GrillaHistorialDespachoDevoluciones.Columns(4).Width = 1700 'despachador
GrillaHistorialDespachoDevoluciones.Columns(5).Width = 1700
GrillaHistorialDespachoDevoluciones.Columns(6).Width = 1700

'Caption de las grillas
GrillaHistorialDespachoDevoluciones.Columns(0).Caption = "ID"
GrillaHistorialDespachoDevoluciones.Columns(1).Caption = "Código"
GrillaHistorialDespachoDevoluciones.Columns(2).Caption = "Cliente"
GrillaHistorialDespachoDevoluciones.Columns(3).Caption = "Fecha"
GrillaHistorialDespachoDevoluciones.Columns(4).Caption = "Despachador"
GrillaHistorialDespachoDevoluciones.Columns(5).Caption = "Zona"
GrillaHistorialDespachoDevoluciones.Columns(6).Caption = "Vendedor"


'alineacion
GrillaHistorialDespachoDevoluciones.Columns(0).Alignment = dbgCenter
GrillaHistorialDespachoDevoluciones.Columns(2).Alignment = dbgLeft
GrillaHistorialDespachoDevoluciones.Columns(3).Alignment = dbgCenter
GrillaHistorialDespachoDevoluciones.Columns(4).Alignment = dbgCenter


'cabeceras
GrillaHistorialDespachoDevoluciones.HeadFont.Bold = True

'las que no quiero ver
GrillaHistorialDespachoDevoluciones.Columns(0).Visible = False
GrillaHistorialDespachoDevoluciones.Columns(5).Visible = False
GrillaHistorialDespachoDevoluciones.Columns(6).Visible = False


End Sub
Sub BorrarTemporalDevolucion()
     With RsTemporalDevoluciones
               .Requery
               If .BOF And .EOF = True Then Exit Sub
               For x = 1 To .RecordCount
                         .Delete
                         If .BOF And .EOF = True Then Exit Sub
                         .MoveNext
               Next
     End With
End Sub

Private Sub GrillaHistorialDespachoDevoluciones_DblClick()
cmd_ver_Click
End Sub
