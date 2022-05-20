VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form DespachosAddProductosForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Productos"
   ClientHeight    =   11175
   ClientLeft      =   4395
   ClientTop       =   1605
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
   Icon            =   "DespachosAddProductosForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DespachosAddProductosForm.frx":08CA
   ScaleHeight     =   11175
   ScaleWidth      =   13500
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid GrillaADDProductos 
      Height          =   4695
      Left            =   1200
      TabIndex        =   7
      Top             =   3400
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
   Begin MSAdodcLib.Adodc AdoFiltrarProductosDespacho 
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
   Begin VB.TextBox txt_cantidad 
      Height          =   390
      Left            =   9240
      MaxLength       =   4
      TabIndex        =   8
      Top             =   9520
      Width           =   795
   End
   Begin VB.TextBox TxtFiltrar 
      Height          =   390
      Left            =   3960
      TabIndex        =   5
      Top             =   2570
      Width           =   2950
   End
   Begin VB.CommandButton cmd_QuitarFiltro 
      Caption         =   "Quitar Filtro"
      Height          =   580
      Left            =   7080
      TabIndex        =   4
      Top             =   2445
      Width           =   2175
   End
   Begin VB.CommandButton cmd_Añadir 
      Caption         =   "Añadir"
      Height          =   615
      Left            =   7680
      TabIndex        =   3
      Top             =   10320
      Width           =   2295
   End
   Begin VB.CommandButton cmd_Atras 
      Appearance      =   0  'Flat
      Caption         =   "Atras"
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   10320
      Width           =   2295
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   11280
      TabIndex        =   17
      Top             =   9570
      Width           =   60
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MAX:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   11520
      TabIndex        =   16
      Top             =   9585
      Width           =   615
   End
   Begin VB.Label lbl_Max 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   12240
      TabIndex        =   15
      Top             =   9585
      Width           =   60
   End
   Begin VB.Label lbl_Min 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   10800
      TabIndex        =   14
      Top             =   9585
      Width           =   60
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MIN:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   10200
      TabIndex        =   13
      Top             =   9585
      Width           =   525
   End
   Begin VB.Label Label4 
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
      Left            =   960
      TabIndex        =   12
      Top             =   8730
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione la cantidad que desea añadir"
      Height          =   270
      Left            =   1320
      TabIndex        =   11
      Top             =   8880
      Width           =   4260
   End
   Begin VB.Label lbl_ProductoDescripcion 
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
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Top             =   9570
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CANT:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   8400
      TabIndex        =   9
      Top             =   9585
      Width           =   765
   End
   Begin MSForms.ComboBox CboFiltrar 
      Height          =   390
      Left            =   1080
      TabIndex        =   6
      Top             =   2570
      Width           =   2595
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4568;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lbl_1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione los productos que desea añadir"
      Height          =   270
      Left            =   1320
      TabIndex        =   1
      Top             =   1880
      Width           =   4515
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
      TabIndex        =   0
      Top             =   1720
      Width           =   225
   End
End
Attribute VB_Name = "DespachosAddProductosForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NombreCampoDespacho As String 'variable que va a tener el nombre del campo que contenga mi tabla y sera el campo por el que filtrare

Private Sub cmd_Añadir_Click()

'validamos que al darle clic a añadir si no hay nada no haga nada
With RsProductos
     If .BOF And .EOF = True Then MsgBox ("No existen productos registrados que pueda seleccionar"), vbInformation, "Aviso": Exit Sub
End With

'validamos que se alla seleccionado un producto e ingresado una cantidad
If lbl_ProductoDescripcion.Caption = "" Then MsgBox "Debe elegir al menos un producto de la lista", vbInformation, "Aviso": Exit Sub
If txt_cantidad.Text = "" Then MsgBox ("Por favor, ingrese la cantidad del producto"), vbInformation, "Aviso": txt_cantidad.SetFocus: Exit Sub

'validamos que lo que introduce en el txt_cantidad sea numerico
If Not IsNumeric(txt_cantidad.Text) Then MsgBox "Por favor, ingrese sólo números en el campo de cantidad", vbInformation, "Aviso": txt_cantidad.SetFocus: Exit Sub
     If Val(txt_cantidad.Text) <= "0" Then MsgBox "Cantidad inválida", vbInformation, "Aviso": txt_cantidad.SetFocus: Exit Sub

'validamos la cantidad que ingreso es mayor a la cantidad maxima/minima de venta por cliente
If Val(txt_cantidad.Text) > GrillaADDProductos.Columns(4).Text Then MsgBox "No contamos con la cantidad solicitada para dicho despacho", vbInformation, "aviso": txt_cantidad.SetFocus: Exit Sub
    'no sobrepase la cantidad maxima
    If Val(txt_cantidad.Text) > GrillaADDProductos.Columns(8).Text Then MsgBox "La cantidad ingresada no concuerda con la cantidad máxima para la venta", vbInformation, "Aviso": txt_cantidad.SetFocus: Exit Sub
    'no se salga de la cantidad minima
    If Val(txt_cantidad.Text) < GrillaADDProductos.Columns(9).Text Then MsgBox "La cantidad ingresada no concuerda con la cantidad mínima para la venta", vbInformation, "Aviso": txt_cantidad.SetFocus: Exit Sub

'validar que el producto no se encuentra ya seleccionado
With RsTemporalDespacho
    .Requery
    .Find "DescripcionProducto='" & Trim(lbl_ProductoDescripcion.Caption) & "'" ' si encuentran la descripcion en la db temporal. (CAMBIAR A VAL SI VALIDARA NUMERO)
    DespachosNewForm.EstiloGrillaDespacho
    If .EOF Then Else MsgBox "Éste producto ya ha sido agragado a éste despacho", vbInformation, "Aviso": Exit Sub
End With

'damos advertencias de cantidad minima
If Val(GrillaADDProductos.Columns(4).Text) - Val(txt_cantidad.Text) <= 3 Then
          If MsgBox("Realizar ésta operación, implicaría dejar a este producto con un stock bajo. ¿Desea continuar de todos modos?", vbExclamation + vbYesNo, "Advertencia") = vbYes Then
                    'no hago nada y sigo con mi proceso
          Else
                    txt_cantidad.Text = ""
                    Exit Sub
          End If
Else
          'sigo con mi proceso porque la cantidad no queda menor o igual a 3
End If
                                                                                                ' F I N     D E    V A L I D A C I O NE S
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

'grabar en el temporal
     With RsTemporalDespacho
               .Requery
               .AddNew
                         !CodigoDespacho = DespachosNewForm.txt_CodFactura.Text
                         !CodigoProducto = GrillaADDProductos.Columns(1).Text
                         !DescripcionProducto = GrillaADDProductos.Columns(2).Text
                         !CANTProductos = txt_cantidad.Text
                         !Marca = GrillaADDProductos.Columns(11).Text
                         If GrillaADDProductos.Columns(6).Value = 0 Then
                                   !Kit = 0
                         Else
                                   !Kit = 1
                         End If
                         !PiezasxKit = GrillaADDProductos.Columns(7).Text
                         .Update
     End With
     DespachosNewForm.EstiloGrillaDespacho
End Sub

Private Sub cmd_Atras_Click()
'para que consultas no se vea afectada
Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
ConsultasForm.EstilosGrillaConsultas

Set DespachosNewForm.GrillaDespachos.DataSource = RsTemporalDespacho
DespachosNewForm.EstiloGrillaDespacho
Unload Me
End Sub

Private Sub cmd_QuitarFiltro_Click()
TxtFiltrar.Text = ""

End Sub

Private Sub Form_Load()
'abrimos las tablas que estaran involucradas
Productos
Despacho
DetallesDespacho
TemporalDespacho


Set GrillaADDProductos.DataSource = RsProductos
EstilosGrillaADDProductos

'Colocamos las opciones del filtro_
CboFiltrar.AddItem "Código Producto"
CboFiltrar.AddItem "Descripción Producto"

'campo del filtro predetermiado
CboFiltrar.ListIndex = 0 'elige cod de producto


End Sub

Sub EstilosGrillaADDProductos()

With RsProductos
'tamaños de la grilla
GrillaADDProductos.Columns(0).Width = 500
GrillaADDProductos.Columns(1).Width = 2500
GrillaADDProductos.Columns(2).Width = 4800
GrillaADDProductos.Columns(3).Width = 2000
GrillaADDProductos.Columns(4).Width = 1200
GrillaADDProductos.Columns(5).Width = 2300
GrillaADDProductos.Columns(6).Width = 500
GrillaADDProductos.Columns(7).Width = 1800
GrillaADDProductos.Columns(8).Width = 2400
GrillaADDProductos.Columns(9).Width = 2400
GrillaADDProductos.Columns(10).Width = 2400
GrillaADDProductos.Columns(11).Width = 2300
GrillaADDProductos.Columns(12).Width = 2400


'Caption de las grillas
GrillaADDProductos.Columns(0).Caption = "ID"
GrillaADDProductos.Columns(1).Caption = "Código del producto"
GrillaADDProductos.Columns(2).Caption = "Descripción del Producto"
GrillaADDProductos.Columns(3).Caption = "Aplicable para"
GrillaADDProductos.Columns(4).Caption = "Cantidad"
GrillaADDProductos.Columns(5).Caption = "Fecha de Recibido"
GrillaADDProductos.Columns(6).Caption = "Kit"
GrillaADDProductos.Columns(7).Caption = "Piezas por kit"
GrillaADDProductos.Columns(8).Caption = "Cant. Max. para la Venta"
GrillaADDProductos.Columns(9).Caption = "Cant. Min. para la Venta"
GrillaADDProductos.Columns(10).Caption = "Ubicación"
GrillaADDProductos.Columns(11).Caption = "Marca"
GrillaADDProductos.Columns(12).Caption = "REF"

'alineacion
GrillaADDProductos.Columns(0).Alignment = dbgCenter
GrillaADDProductos.Columns(2).Alignment = dbgLeft
GrillaADDProductos.Columns(3).Alignment = dbgCenter
GrillaADDProductos.Columns(4).Alignment = dbgCenter
GrillaADDProductos.Columns(5).Alignment = dbgCenter
GrillaADDProductos.Columns(6).Alignment = dbgCenter
GrillaADDProductos.Columns(7).Alignment = dbgCenter
GrillaADDProductos.Columns(8).Alignment = dbgCenter
GrillaADDProductos.Columns(9).Alignment = dbgCenter
GrillaADDProductos.Columns(10).Alignment = dbgCenter
GrillaADDProductos.Columns(11).Alignment = dbgCenter

'cabeceras
GrillaADDProductos.HeadFont.Bold = True

'las que no quiero ver
GrillaADDProductos.Columns(0).Visible = False
GrillaADDProductos.Columns(3).Visible = False
GrillaADDProductos.Columns(5).Visible = False
GrillaADDProductos.Columns(6).Visible = False
GrillaADDProductos.Columns(7).Visible = False
GrillaADDProductos.Columns(8).Visible = False
GrillaADDProductos.Columns(9).Visible = False
GrillaADDProductos.Columns(10).Visible = False
GrillaADDProductos.Columns(12).Visible = False
End With
End Sub

Private Sub GrillaADDProductos_Click()
With RsProductos
    If .BOF And .EOF = True Then Exit Sub 'si la tabla en la bd esta vacia
    If GrillaADDProductos.ApproxCount = 0 Then Exit Sub 'si la grilla queda vacia despues de la busqueda
    lbl_ProductoDescripcion.Caption = Trim(GrillaADDProductos.Columns(2).Text)
    lbl_Min.Caption = GrillaADDProductos.Columns(9).Text
    lbl_Max.Caption = GrillaADDProductos.Columns(8).Text
    txt_cantidad.Text = ""
    lbl_Min.Refresh
    lbl_Max.Refresh
    lbl_ProductoDescripcion.Refresh
End With
End Sub

Sub FiltrarProductosDespacho()

'obtener el nombre del campo por el que se va a filtrar
If CboFiltrar.Text = "Código Producto" Then NombreCampoDespacho = "CodigoProducto"
If CboFiltrar.Text = "Descripción Producto" Then NombreCampoDespacho = "Descripcion"

'programar filtro * M O D I F I Q U E     A  Q U I
With AdoFiltrarProductosDespacho
    .CursorLocation = adUseClient
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
    Busca = UCase(Trim(TxtFiltrar.Text)) & "%"
    
    .RecordSource = "select * from Productos WHERE[" & NombreCampoDespacho & "] like '" & Busca & "' order by Descripcion asc"
    .Refresh
End With

Set GrillaADDProductos.DataSource = AdoFiltrarProductosDespacho
EstilosGrillaADDProductos
End Sub

Private Sub TxtFiltrar_Change()
If CboFiltrar.Text = "" Then Exit Sub
        FiltrarProductosDespacho
        EstilosGrillaADDProductos
        FiltrarProductosDespacho
        txt_cantidad.Text = ""
        lbl_ProductoDescripcion.Caption = ""
        lbl_Max.Caption = ""
        lbl_Min.Caption = ""
End Sub
