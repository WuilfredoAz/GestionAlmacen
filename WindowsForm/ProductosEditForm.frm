VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form ProductosEditForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar Productos"
   ClientHeight    =   10305
   ClientLeft      =   1860
   ClientTop       =   1335
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
   Icon            =   "ProductosEditForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ProductosEditForm.frx":058A
   ScaleHeight     =   10305
   ScaleWidth      =   17985
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid GrillaProductos 
      Height          =   5380
      Left            =   1395
      TabIndex        =   2
      Top             =   3480
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   9499
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
   Begin VB.TextBox TxtFiltrar 
      Height          =   390
      Left            =   4180
      TabIndex        =   7
      Top             =   2570
      Width           =   2890
   End
   Begin VB.CommandButton cmd_QuitarFiltro 
      Caption         =   "Quitar Filtro"
      Height          =   580
      Left            =   7320
      TabIndex        =   6
      Top             =   2500
      Width           =   2175
   End
   Begin VB.CommandButton cmd_eliminarproducto 
      Caption         =   "Eliminar producto"
      Height          =   615
      Left            =   13800
      TabIndex        =   5
      Top             =   9360
      Width           =   2295
   End
   Begin VB.CommandButton cmd_atras 
      Appearance      =   0  'Flat
      Caption         =   "Atras"
      Height          =   615
      Left            =   1440
      TabIndex        =   4
      Top             =   9360
      Width           =   2295
   End
   Begin VB.CommandButton cmd_editar 
      Caption         =   "Editar"
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   9360
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc AdoFiltrarProductosEdit 
      Height          =   330
      Left            =   5160
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
   Begin MSForms.ComboBox CboFiltrar 
      Height          =   390
      Left            =   1320
      TabIndex        =   8
      Top             =   2570
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
      TabIndex        =   1
      Top             =   1920
      Width           =   225
   End
   Begin VB.Label lbl_descripcion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el producto que desea editar"
      Height          =   270
      Left            =   1680
      TabIndex        =   0
      Top             =   2040
      Width           =   4215
   End
End
Attribute VB_Name = "ProductosEditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NombreCampo As String
Private Sub cmd_Atras_Click()
Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
ConsultasForm.EstilosGrillaConsultas
Unload Me
End Sub

Private Sub cmd_editar_Click()
'verificar si la tabla esta vacia
With RsProductos
          If .RecordCount = 0 Then
                    Exit Sub
          Else
                    If TxtFiltrar.Text = "" Then
                              With RsProductos
                                        If RecordCount = "0" Then Exit Sub
                              End With
                    Else
                              With AdoFiltrarProductosEdit
                                        If .Recordset.RecordCount = 0 Then Exit Sub
                              End With
                    End If
          'obetener el codigo del producto
          vIDProductos = Val(GrillaProductos.Columns(0).Text)
          'llamar al formulario de editar productos
          ProductosEdit2Form.Show vbModal
          ProductosEditForm.EstilosGrillaProducto
          End If
End With
End Sub

Private Sub cmd_EliminarProducto_Click()

With RsProductos 'esto para que cuando haga clic y este vacio no haga nada
     If .BOF And .EOF = True Then
          Exit Sub
     Else
          ProductosEditForm.EstilosGrillaProducto
          ' verificamos si la tabla esta vacia
          If TxtFiltrar.Text = "" Then
               With RsProductos
                    If RecordCount = "0" Then Exit Sub
               End With
          Else
          With AdoFiltrarProductosEdit
               If .Recordset.RecordCount = 0 Then Exit Sub
          End With
          End If
          'preguntar si se va a eliminar el usuario seleccionado
          If MsgBox("¿Desea eliminar el siguiente producto COD: " & GrillaProductos.Columns(1).Text & " |  Descripción: " & GrillaProductos.Columns(2).Text, vbInformation + vbYesNo, "Advertencia") = vbYes Then
          'elimino el registro
                    vIDProductos = Val(GrillaProductos.Columns(0).Text)
                    With RsProductos
                              .Requery
                              .Find "Id='" & Val(vIDProductos) & "'" 'ya tengo ubicado el registro de productos
                              Logs
                              With RsLogs
                                        .Requery
                                        .AddNew
                                                  !User = vUsername
                                                  !Accion = "eliminó el producto asociado al código [" & GrillaProductos.Columns(1).Text & "] teniendo la cantidad de: " & GrillaProductos.Columns(4).Text
                                                  !Fecha = Date + Time
                                        .Update
                                        .Requery
                                        .Close
                              End With
                              LogsProductos
                              With RsLogsProductos
                                        .Requery
                                        .AddNew
                                                  !User = vUsername
                                                  !Accion = "eliminó el producto asociado al código [" & GrillaProductos.Columns(1).Text & "] teniendo la cantidad de: " & GrillaProductos.Columns(4).Text
                                                  !Fecha = Date + Time
                                        .Update
                                        .Requery
                                        .Close
                              End With
                              .Delete ' LO ELIMINO
                              .Requery
                              EstilosGrillaProducto
                              ConsultasForm.EstilosGrillaConsultas
                              vIDProductos = 0
                              TxtFiltrar.Text = ""
                    End With
          Else
                    Exit Sub
          End If
     End If
End With

Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
ConsultasForm.EstilosGrillaConsultas

'actualizamos la consulta de total de productos del index
IndexForm.AdoTotalProductos.Refresh
IndexForm.lbl_totalproductos.Caption = IndexForm.GrillaTotalProductos.Columns(0).Text

'actualizamos la consulta de total de registtos del index
IndexForm.AdoTotalRegistros.Refresh
IndexForm.lbl_totalregistros.Caption = IndexForm.GrillaTotalRegistros.Columns(0).Text

'actualizamos la consulta de productos con mas de 50 de stocks
IndexForm.AdoProductosMas50.Refresh
IndexForm.lbl_mas50.Caption = IndexForm.GrillaProductosMas50.Columns(0).Text

'actualizamos la consulta de productos con menos de 50 de stocks
IndexForm.AdoProductosMenos50.Refresh
IndexForm.lbl_menos50.Caption = IndexForm.GrillaProductosMenos50.Columns(0).Text

'actualizamos la consulta de productos fuera de stocks
IndexForm.AdoProductosOUT.Refresh
IndexForm.lbl_fuerastock.Caption = IndexForm.GrillaProductosOUT.Columns(0).Text
End Sub

Private Sub cmd_QuitarFiltro_Click()
TxtFiltrar.Text = ""
End Sub

Private Sub Form_Load()
Productos
Set GrillaProductos.DataSource = RsProductos
EstilosGrillaProducto


CboFiltrar.AddItem "Código Producto"
CboFiltrar.AddItem "Descripción Producto"
CboFiltrar.AddItem "Ubicacion"
CboFiltrar.AddItem "Cantidad"

CboFiltrar.ListIndex = 0

End Sub

Sub EstilosGrillaProducto()

'tamaños de la grilla
GrillaProductos.Columns(0).Width = 500
GrillaProductos.Columns(1).Width = 2500
GrillaProductos.Columns(2).Width = 4900
GrillaProductos.Columns(3).Width = 2000
GrillaProductos.Columns(4).Width = 1200
GrillaProductos.Columns(5).Width = 2300
GrillaProductos.Columns(6).Width = 500
GrillaProductos.Columns(7).Width = 800
GrillaProductos.Columns(8).Width = 2400
GrillaProductos.Columns(9).Width = 2400
GrillaProductos.Columns(10).Width = 2400
GrillaProductos.Columns(11).Width = 2400
GrillaProductos.Columns(12).Width = 2400

'Caption de las grillas
GrillaProductos.Columns(0).Caption = "ID"
GrillaProductos.Columns(1).Caption = "Código del producto"
GrillaProductos.Columns(2).Caption = "Descripción del Producto"
GrillaProductos.Columns(3).Caption = "Aplicable para"
GrillaProductos.Columns(4).Caption = "Cantidad"
GrillaProductos.Columns(5).Caption = "Fecha de Recibido"
GrillaProductos.Columns(6).Caption = "Kit"
GrillaProductos.Columns(7).Caption = "Pzas"
GrillaProductos.Columns(8).Caption = "Cant. Max. para la Venta"
GrillaProductos.Columns(9).Caption = "Cant. Min. para la Venta"
GrillaProductos.Columns(10).Caption = "Ubicación"
GrillaProductos.Columns(11).Caption = "Marca"
GrillaProductos.Columns(12).Caption = "REF"

'alineacion
GrillaProductos.Columns(0).Alignment = dbgCenter
GrillaProductos.Columns(1).Alignment = dbgCenter
GrillaProductos.Columns(2).Alignment = dbgLeft
GrillaProductos.Columns(3).Alignment = dbgCenter
GrillaProductos.Columns(4).Alignment = dbgCenter
GrillaProductos.Columns(5).Alignment = dbgCenter
GrillaProductos.Columns(6).Alignment = dbgCenter
GrillaProductos.Columns(7).Alignment = dbgCenter
GrillaProductos.Columns(8).Alignment = dbgCenter

'cabeceras
GrillaProductos.HeadFont.Bold = True

'las que no quiero ver
GrillaProductos.Columns(0).Visible = False
GrillaProductos.Columns(3).Visible = False
GrillaProductos.Columns(6).Visible = False
GrillaProductos.Columns(8).Visible = False
GrillaProductos.Columns(9).Visible = False
GrillaProductos.Columns(10).Visible = False
GrillaProductos.Columns(12).Visible = False

End Sub

Sub FiltrarProductosEdit()

If CboFiltrar = "Código Producto" Then NombreCampo = "CodigoProducto"
If CboFiltrar = "Descripción Producto" Then NombreCampo = "Descripcion"
If CboFiltrar = "Ubicacion" Then NombreCampo = "Ubicacion"
If CboFiltrar = "Cantidad" Then NombreCampo = "Cantidad"

'programacion del filtro
With AdoFiltrarProductosEdit
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
          Busca = UCase(Trim(TxtFiltrar.Text)) & "%"
          .RecordSource = "SELECT * FROM Productos WHERE [" & NombreCampo & "] LIKE '" & Busca & "' ORDER BY Descripcion ASC"
          .Refresh
End With
Set GrillaProductos.DataSource = AdoFiltrarProductosEdit
EstilosGrillaProducto
End Sub

Private Sub GrillaProductos_DblClick()
cmd_editar_Click
End Sub

Private Sub TxtFiltrar_Change()
If CboFiltrar.Text = "" Then Exit Sub
FiltrarProductosEdit
EstilosGrillaProducto
End Sub
