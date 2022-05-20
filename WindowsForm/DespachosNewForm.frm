VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form DespachosNewForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear nuevo despacho"
   ClientHeight    =   9795
   ClientLeft      =   1845
   ClientTop       =   1500
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
   Icon            =   "DespachosNewForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DespachosNewForm.frx":058A
   ScaleHeight     =   9795
   ScaleWidth      =   18000
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid GrillaDespachos 
      Height          =   2895
      Left            =   1680
      TabIndex        =   20
      Top             =   5450
      Width           =   13455
      _ExtentX        =   23733
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
         SizeMode        =   1
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd_EliminarProducto 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   15480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmd_AñadirProducto 
      Caption         =   "Añadir"
      Height          =   495
      Left            =   15480
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmd_Cancelar 
      Appearance      =   0  'Flat
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   5640
      TabIndex        =   9
      Top             =   8880
      Width           =   2295
   End
   Begin VB.CommandButton cmd_Registrar 
      Caption         =   "Registrar"
      Height          =   615
      Left            =   9600
      TabIndex        =   8
      Top             =   8880
      Width           =   2295
   End
   Begin VB.TextBox txt_NombreDespachador 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3520
      Width           =   3855
   End
   Begin VB.TextBox txt_CodFactura 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1920
      MaxLength       =   15
      TabIndex        =   1
      Top             =   2680
      Width           =   3855
   End
   Begin VB.TextBox txt_NombreCliente 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   2
      Top             =   3520
      Width           =   3855
   End
   Begin VB.TextBox txt_ZonaDespacho 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      MaxLength       =   15
      TabIndex        =   3
      Top             =   2680
      Width           =   3855
   End
   Begin VB.TextBox txt_NombreVendedor 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      MaxLength       =   30
      TabIndex        =   4
      Top             =   3520
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTP_fecha 
      Height          =   495
      Left            =   12840
      TabIndex        =   5
      Top             =   2680
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   -2147483635
      Format          =   117440513
      CurrentDate     =   42645
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
      TabIndex        =   19
      Top             =   4600
      Width           =   225
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Añada los productos correspondientes al despacho"
      Height          =   270
      Left            =   1680
      TabIndex        =   18
      Top             =   4750
      Width           =   5355
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
      TabIndex        =   17
      Top             =   4000
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
      TabIndex        =   16
      Top             =   3180
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
      TabIndex        =   15
      Top             =   1950
      Width           =   225
   End
   Begin VB.Label lbl_descripcion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rellene la información solicitada"
      Height          =   270
      Left            =   1680
      TabIndex        =   14
      Top             =   2100
      Width           =   3360
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
      TabIndex        =   13
      Top             =   4005
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zona a donde va el despacho"
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
      Top             =   3180
      Width           =   2160
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
      TabIndex        =   11
      Top             =   4005
      Width           =   705
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
      TabIndex        =   10
      Top             =   3180
      Width           =   1455
   End
End
Attribute VB_Name = "DespachosNewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_AñadirProducto_Click()

If txt_CodFactura.Text = "" Then MsgBox "Antes de proceder, por favor ingrese un código de factura", vbInformation, "Aviso": txt_CodFactura.SetFocus: Exit Sub
If txt_NombreCliente.Text = "" Then MsgBox "Antes de proceder, por favor ingrese el nombre del cliente", vbInformation, "Aviso": txt_NombreCliente.SetFocus: Exit Sub
     If IsNumeric(txt_NombreCliente.Text) Then MsgBox ("No escriba números en el nombre del cliente"), vbInformation, "Aviso": txt_NombreCliente.SetFocus: Exit Sub
If txt_ZonaDespacho.Text = "" Then MsgBox "Antes de proceder, por favor indique la zona a donde irá el despacho", vbInformation, "Aviso": txt_ZonaDespacho.SetFocus: Exit Sub
If txt_NombreVendedor.Text = "" Then MsgBox "Antes de proceder, ingrese el nombre del vendedor", vbInformation, "Aviso": txt_NombreVendedor.SetFocus: Exit Sub
     If IsNumeric(txt_NombreVendedor.Text) Then MsgBox ("No escriba números en el nombre del vendedor"), vbInformation, "Aviso": txt_NombreVendedor.SetFocus: Exit Sub
     If UCase(Trim(txt_NombreCliente.Text)) = UCase(Trim(txt_NombreVendedor.Text)) Then MsgBox ("El cliente y el vendedor no pueden ser los mismos"), vbInformation, "Aviso": txt_NombreVendedor.Text = "": txt_NombreVendedor.SetFocus: Exit Sub
'validamos que el codigo no se repita en despachos
With RsDespacho
          .Requery
          .Find "CodigoDespacho='" & Trim(txt_CodFactura.Text) & "'"
          If .EOF = True Then
                    EstiloGrillaDespacho
                    'no hago nada
          Else
                    EstiloGrillaDespacho
                    MsgBox ("Ya existe un despacho con ese código, por favor cambielo"), vbInformation, "Aviso": txt_CodFactura.SetFocus: Exit Sub
          End If
End With

'validamos que no se repita en devoluciones
With RsDevoluciones
          .Requery
          .Find "CodigoDespacho='" & Trim(txt_CodFactura.Text) & "'"
          If .EOF = True Then
                    EstiloGrillaDespacho
                    'no hago nada
          Else
                    EstiloGrillaDespacho
                    MsgBox ("Ya existe un despacho devuelto asociado a ese código, por favor cambielo"), vbInformation, "Aviso": txt_CodFactura.SetFocus: Exit Sub
          End If
End With

DespachosAddProductosForm.Show vbModal
End Sub

Private Sub cmd_cancelar_Click()
Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
ConsultasForm.GrillaConsultas.Refresh
ConsultasForm.EstilosGrillaConsultas
Unload Me
End Sub

Private Sub cmd_EliminarProducto_Click()
'verificar si la tabla esta vacia
With RsTemporalDespacho
    If .RecordCount = 0 Then Set ConsultasForm.GrillaConsultas.DataSource = RsProductos: ConsultasForm.EstilosGrillaConsultas: Exit Sub
End With

'obetener el codigo del producto
vDProductos = GrillaDespachos.Columns(0).Text

'preguntar si se va a eliminar el usuario seleccionado
If MsgBox("Eliminará éste producto del despacho. ¿Desea continuar?", vbInformation + vbYesNo, "Advertencia") = vbYes Then
'elimino el registro
        With RsTemporalDespacho
                .Requery
                        .Find "Id='" & Val(vDProductos) & "'" 'ya tengo ubicado el registro de productos
                        .Delete ' LO ELIMINO
                        .Requery
                                EstiloGrillaDespacho
                                DespachosAddProductosForm.EstilosGrillaADDProductos
                                vDProductos = 0
        End With
End If
Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
ConsultasForm.EstilosGrillaConsultas

Set GrillaDespachos.DataSource = RsTemporalDespacho
EstiloGrillaDespacho
End Sub

Private Sub cmd_Registrar_Click()
'validaciones generales
If txt_CodFactura.Text = "" Then MsgBox "Antes de proceder, por favor ingrese un código de factura", vbInformation, "Aviso": txt_CodFactura.SetFocus: Exit Sub
If txt_NombreCliente.Text = "" Then MsgBox "Antes de proceder, por favor ingrese el nombre del cliente", vbInformation, "Aviso": txt_NombreCliente.SetFocus: Exit Sub
     If IsNumeric(txt_NombreCliente.Text) Then MsgBox ("No escriba números en el nombre del cliente"), vbInformation, "Aviso": txt_NombreCliente.SetFocus: Exit Sub
If txt_ZonaDespacho.Text = "" Then MsgBox "Antes de proceder, por favor indique la zona a donde irá el despacho", vbInformation, "Aviso": txt_ZonaDespacho.SetFocus: Exit Sub
If txt_NombreVendedor.Text = "" Then MsgBox "Antes de proceder, Ingrese el nombre del vendedor", vbInformation, "Aviso": txt_NombreVendedor.SetFocus: Exit Sub
     If IsNumeric(txt_NombreVendedor.Text) Then MsgBox ("No escriba números en el nombre del vendedor"), vbInformation, "Aviso": txt_NombreVendedor.SetFocus: Exit Sub
     If UCase(Trim(txt_NombreCliente.Text)) = UCase(Trim(txt_NombreVendedor.Text)) Then MsgBox ("El cliente y el vendedor no pueden ser los mismos"), vbInformation, "Aviso": txt_NombreVendedor.Text = "": txt_NombreVendedor.SetFocus: Exit Sub
'validamos que en el Despacho existan productos
With RsTemporalDespacho
        .Requery
                If .BOF Or .EOF Then
                         EstiloGrillaDespacho
                        MsgBox "Para crear el reporte del despacho debe añadirle productos al mismo", vbInformation, "Aviso"
                        cmd_AñadirProducto.SetFocus
                        Exit Sub
                End If
End With

'validamos que el codigo del despacho no se este repitiendo
With RsDespacho
        .Requery
        .Find "CodigoDespacho='" & Trim(txt_CodFactura.Text) & "'"
                 If .EOF Then 'si no lo encuentra carga el estilo y no hace nada y sigue la secuencia del programa :D
                        EstiloGrillaDespacho
                 Else
                        EstiloGrillaDespacho
                        MsgBox "El código de despacho introducido ya fue ingresado con anterioridad, por favor cambielo e intente de nuevo", vbInformation, "Aviso": txt_CodFactura.SetFocus: Exit Sub
                End If
End With

'validamos que no se repita en devoluciones
With RsDevoluciones
          .Requery
          .Find "CodigoDespacho='" & Trim(txt_CodFactura.Text) & "'"
          If .EOF Then
                    EstiloGrillaDespacho
                    'no hago nada
          Else
                    EstiloGrillaDespacho
                    MsgBox ("Ya existe un despacho devuelto asociado a ese código, por favor cambielo"), vbInformation, "Aviso": txt_CodFactura.SetFocus: Exit Sub
          End If
End With
                                                                                                ' F I N     D E    V A L I D A C I O N E S
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

'Creamos el Despacho
With RsDespacho
        .Requery
        .AddNew
                !CodigoDespacho = txt_CodFactura.Text
                !Cliente = txt_NombreCliente.Text
                !Fecha = DTP_fecha
                !Despachador = txt_NombreDespachador.Text
                !Zona = txt_ZonaDespacho.Text
                !Vendedor = txt_NombreVendedor
        .Update
End With

'añadimos los detalles del despacho y restamos
Dim Registros As Integer
Registros = RsTemporalDespacho.RecordCount
RsTemporalDespacho.Requery
RsTemporalDespacho.MoveFirst
For x = 1 To Registros

    'RESTAMOS LOS PRODUCTOS
        With RsProductos
                .Requery
                        .Find "CodigoProducto='" & Trim(GrillaDespachos.Columns(2).Text) & "'"
                        !Cantidad = !Cantidad - Val(GrillaDespachos.Columns(4).Text)
                .UpdateBatch
        End With

        With RsDetallesDespacho
                .Requery
                .AddNew
                        !CodigoDespacho = GrillaDespachos.Columns(1).Text
                        !CodigoProducto = GrillaDespachos.Columns(2).Text
                        !DescripcionProducto = GrillaDespachos.Columns(3).Text
                        !CANTProductos = GrillaDespachos.Columns(4).Text
                        !Marca = GrillaDespachos.Columns(5).Text
                                If GrillaDespachos.Columns(6).Value = 0 Then
                                        !Kit = 0
                                Else
                                        !Kit = 1
                                End If
                        !PiezasxKit = GrillaDespachos.Columns(7).Text
                .Update
        End With
        If x = Registros Then Else RsTemporalDespacho.MoveNext

Next
EstiloGrillaDespacho
MsgBox "El despacho fue creado exitosamente", vbInformation, "Aviso"



                                                                                                ' F I N     D E    C R E A C I O N  D E   D E S P A C H O
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

'igualamos las variables
 vNFactura = txt_CodFactura.Text
vCliente = txt_NombreCliente.Text
vVendedor = txt_NombreVendedor.Text
vDespachador = txt_NombreDespachador.Text
vFecha = DTP_fecha
vZona = txt_ZonaDespacho.Text
vNProductos = Registros


Logs
With RsLogs
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "creó un despacho asociado al código de factura N° [" & vNFactura & "] con un total de " & Registros & " producto(s)"
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With

LogsDespachos
With RsLogsDespachos
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "creó un despacho asociado al código de factura N° [" & vNFactura & "] con un total de " & Registros & " producto(s)"
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
          vNProductos = ""
End With


'salimos
Unload Me

Set dr_RDespachosNew.DataSource = RsTemporalDespacho
dr_RDespachosNew.WindowState = 2

' s e c c i o n 2
dr_RDespachosNew.Sections("Sección2").Controls("Etiqueta12").Caption = vCliente
dr_RDespachosNew.Sections("Sección2").Controls("Etiqueta17").Caption = vVendedor
dr_RDespachosNew.Sections("Sección2").Controls("Etiqueta18").Caption = vDespachador
dr_RDespachosNew.Sections("Sección2").Controls("Etiqueta19").Caption = vFecha
dr_RDespachosNew.Sections("Sección2").Controls("Etiqueta20").Caption = vZona

' s e c c i o n 4
dr_RDespachosNew.Sections("Sección4").Controls("Etiqueta3").Caption = vNFactura


'mostramos
dr_RDespachosNew.Show


                                                                                                         ' F I N     D E    I M P R E S I O N
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/


'borramos en la tabla temporal
'BorrarTemporal

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

'actualizamos la consulta de total de despachos
IndexForm.AdoTotalDespachos.Refresh
IndexForm.lbl_TotalDespachos.Caption = IndexForm.GrillaTotalDespachos.Columns(0).Text

'actualizamos la consulta de total de devoluciones
IndexForm.AdoTotalDevoluciones.Refresh
IndexForm.lbl_TotalDevoluciones.Caption = IndexForm.GrillaTotalDevoluciones.Columns(0).Text

'actualizamos el total de operaciones
IndexForm.TotalOperaciones

EstiloGrillaDespacho

'para que consultas no se vea afectada
Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
ConsultasForm.EstilosGrillaConsultas


End Sub

Sub BorrarTemporal()
TemporalDespacho
With RsTemporalDespacho
        .Requery
        If .BOF Or .EOF Then Exit Sub
        For x = 1 To .RecordCount
                .Delete
                If .BOF Or .EOF Then Exit Sub
                .MoveNext
         Next
End With
End Sub

Private Sub Form_Load()
txt_NombreDespachador.Text = vUsername
txt_NombreDespachador.Refresh
DTP_fecha = Date

'abrimos las tablas que estaran involucradas
Productos
Despacho
DetallesDespacho
TemporalDespacho
Devoluciones

'cargar detalle de factura (YA QUE ES LA TABLA QUE CONTENDRA LOS PRODUCTOS QUE SE ESTAN VENDIENDO EN ESTE PROCESO DE DESPACHO
Set GrillaDespachos.DataSource = RsTemporalDespacho
EstiloGrillaDespacho

End Sub

Sub EstiloGrillaDespacho()

                GrillaDespachos.Columns(0).Width = 0
                GrillaDespachos.Columns(1).Width = 0
                GrillaDespachos.Columns(2).Width = 2750
                GrillaDespachos.Columns(3).Width = 7000
                GrillaDespachos.Columns(4).Width = 1500
                GrillaDespachos.Columns(5).Width = 1500
                GrillaDespachos.Columns(6).Width = 0
                GrillaDespachos.Columns(7).Width = 0
    
    'caption de las grillas
                GrillaDespachos.Columns(0).Caption = "ID"
                GrillaDespachos.Columns(1).Caption = "Código Despacho"
                GrillaDespachos.Columns(2).Caption = "Código Producto"
                GrillaDespachos.Columns(3).Caption = "Descripción"
                GrillaDespachos.Columns(4).Caption = "Cantidad"
                GrillaDespachos.Columns(5).Caption = "Marca"
                GrillaDespachos.Columns(6).Caption = "Kit"
                GrillaDespachos.Columns(7).Caption = "Piezas por Kit"

    'alineacion
                GrillaDespachos.Columns(0).Alignment = dbgCenter
                GrillaDespachos.Columns(1).Alignment = dbgCenter
                GrillaDespachos.Columns(2).Alignment = dbgCenter
                GrillaDespachos.Columns(3).Alignment = dbgLeft
                GrillaDespachos.Columns(4).Alignment = dbgCenter
                GrillaDespachos.Columns(5).Alignment = dbgCenter
                GrillaDespachos.Columns(6).Alignment = dbgCenter
                GrillaDespachos.Columns(7).Alignment = dbgCenter


'cabeceras
                GrillaDespachos.HeadFont.Bold = True

'las que no quiero ver
                GrillaDespachos.Columns(0).Visible = False
                GrillaDespachos.Columns(1).Visible = False
                GrillaDespachos.Columns(6).Visible = False
                GrillaDespachos.Columns(7).Visible = False
                
End Sub
