VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form LogsUsuariosForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logs de los Usuarios"
   ClientHeight    =   10800
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
   Icon            =   "LogsUsuariosForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "LogsUsuariosForm.frx":058A
   ScaleHeight     =   10800
   ScaleWidth      =   18000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_QuitarFiltro 
      Caption         =   "Quitar Filtro"
      Height          =   465
      Left            =   7185
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmd_Atras 
      Caption         =   "Atras"
      Height          =   615
      Left            =   7905
      TabIndex        =   1
      Top             =   9960
      Width           =   2175
   End
   Begin VB.TextBox TxtFiltrar 
      Height          =   390
      Left            =   3945
      TabIndex        =   0
      Top             =   3030
      Width           =   2950
   End
   Begin MSDataGridLib.DataGrid GrillaLogsUsuarios 
      Height          =   5535
      Left            =   1185
      TabIndex        =   2
      Top             =   3960
      Width           =   15645
      _ExtentX        =   27596
      _ExtentY        =   9763
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
      Caption         =   "Logs de los productos"
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
   Begin MSAdodcLib.Adodc AdoFiltrarLogsUsuarios 
      Height          =   330
      Left            =   5640
      Top             =   1200
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
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
      Left            =   1080
      TabIndex        =   7
      Top             =   3030
      Width           =   2550
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "4498;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lbl_Mes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEPTIEMBRE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   675
      Left            =   1305
      TabIndex        =   6
      Top             =   1440
      Width           =   3450
   End
   Begin VB.Label lbl_Año 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2016"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   360
      Left            =   1305
      TabIndex        =   5
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Logs de éste período"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   1305
      TabIndex        =   4
      Top             =   2505
      Width           =   1680
   End
End
Attribute VB_Name = "LogsUsuariosForm"
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

Private Sub Form_Load()

LogsUsuarios
Set GrillaLogsUsuarios.DataSource = RsLogsUsuarios
EstilosGrillaLogsUsuarios

CboFiltrar.AddItem "Usuario"
CboFiltrar.AddItem "Acción"
CboFiltrar.AddItem "Día/hora"
CboFiltrar.ListIndex = 0
End Sub

Sub EstilosGrillaLogsUsuarios()
'tamaños
GrillaLogsUsuarios.Columns(0).Width = 0
GrillaLogsUsuarios.Columns(1).Width = 2000
GrillaLogsUsuarios.Columns(2).Width = 10000
GrillaLogsUsuarios.Columns(3).Width = 3000


'captios
GrillaLogsUsuarios.Columns(0).Caption = "ID"
GrillaLogsUsuarios.Columns(1).Caption = "Usuario"
GrillaLogsUsuarios.Columns(2).Caption = "Acción Realizada"
GrillaLogsUsuarios.Columns(3).Caption = "Fecha/Hora"

'alineacion
GrillaLogsUsuarios.Columns(0).Alignment = dbgCenter
GrillaLogsUsuarios.Columns(1).Alignment = dbgLeft
GrillaLogsUsuarios.Columns(2).Alignment = dbgLeft
GrillaLogsUsuarios.Columns(3).Alignment = dbgCenter

'negrita
GrillaLogsUsuarios.HeadFont.Bold = True

'los que no quiero ver
GrillaLogsUsuarios.Columns(0).Visible = False

End Sub

Sub FiltrarLogsUsuarios()
          If CboFiltrar.Text = "Usuario" Then NombreCampo = "User"
          If CboFiltrar.Text = "Acción" Then NombreCampo = "Accion"
          If CboFiltrar.Text = "Día/hora" Then NombreCampo = "Fecha"
          
          With AdoFiltrarLogsUsuarios
                    .CursorLocation = adUseClient
                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
                    Busca = UCase(Trim(TxtFiltrar.Text)) & "%"
                    .RecordSource = "SELECT * FROM LogsUsuarios WHERE [" & NombreCampo & "] LIKE '" & Busca & "' AND Fecha LIKE '%" & vMesLogs & vAñoLogs & "%'"
                    .Refresh
                    Set GrillaLogsUsuarios.DataSource = AdoFiltrarLogsUsuarios
                    EstilosGrillaLogsUsuarios
          End With
End Sub

Private Sub TxtFiltrar_Change()
If CboFiltrar.Text = "" Then Exit Sub
FiltrarLogsUsuarios
End Sub
