VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form DespachosDevolucionesDetallesForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear devolucion"
   ClientHeight    =   10470
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
   Icon            =   "DespachosDetallesDevolucionForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DespachosDetallesDevolucionForm.frx":058A
   ScaleHeight     =   10470
   ScaleWidth      =   17985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Descontar 
      Caption         =   "Descontar"
      Height          =   495
      Left            =   15480
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid GrillaContenidoDevolucion 
      Height          =   2895
      Left            =   1680
      TabIndex        =   7
      Top             =   6120
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
   Begin MSAdodcLib.Adodc AdoImprimirModificado 
      Height          =   375
      Left            =   12600
      Top             =   9360
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VB.TextBox txt_DevueltoPor 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   4300
      Width           =   3855
   End
   Begin VB.CommandButton cmd_Restablecer 
      Caption         =   "Restablecer"
      Height          =   495
      Left            =   15480
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmd_Sacar 
      Caption         =   "Sacar"
      Height          =   495
      Left            =   15480
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox txt_Observaciones 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      TabIndex        =   23
      Top             =   4300
      Width           =   3855
   End
   Begin VB.ComboBox cbo_Motivo 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   4320
      Width           =   3855
   End
   Begin MSAdodcLib.Adodc AdoVerProductos 
      Height          =   375
      Left            =   7680
      Top             =   10560
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Top             =   2610
      Width           =   1815
   End
   Begin VB.TextBox txt_NombreVendedor 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3450
      Width           =   3855
   End
   Begin VB.TextBox txt_ZonaDespacho 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2610
      Width           =   3855
   End
   Begin VB.TextBox txt_NombreCliente 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3450
      Width           =   3855
   End
   Begin VB.TextBox txt_CodFactura 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2610
      Width           =   3855
   End
   Begin VB.TextBox txt_NombreDespachador 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3450
      Width           =   3855
   End
   Begin VB.CommandButton cmd_Devolver 
      Caption         =   "Devolver"
      Height          =   615
      Left            =   9600
      TabIndex        =   1
      Top             =   9480
      Width           =   2295
   End
   Begin VB.CommandButton cmd_Cancelar 
      Appearance      =   0  'Flat
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   5640
      TabIndex        =   0
      Top             =   9480
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTP_fecha 
      Height          =   495
      Left            =   14880
      TabIndex        =   21
      Top             =   2610
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   -2147483635
      Format          =   117440513
      CurrentDate     =   42645
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   375
      Left            =   9840
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
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
      TabIndex        =   29
      Top             =   4800
      Width           =   930
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
      TabIndex        =   25
      Top             =   4800
      Width           =   3165
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el motivo de la devolución"
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
      TabIndex        =   24
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label lbl_DiferenciaTiempo 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   7800
      TabIndex        =   20
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
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
      TabIndex        =   19
      Top             =   3105
      Width           =   1530
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
      TabIndex        =   17
      Top             =   3100
      Width           =   1455
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
      Top             =   3940
      Width           =   1590
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
      Top             =   3105
      Width           =   2220
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
      TabIndex        =   14
      Top             =   3940
      Width           =   1440
   End
   Begin VB.Label lbl_descripcion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rellene la información solicitada (Campos en amarillo)"
      Height          =   270
      Left            =   1680
      TabIndex        =   13
      Top             =   2040
      Width           =   5670
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
      Top             =   1920
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
      Top             =   3105
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
      Top             =   3940
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gestione los productos que desea devolver con los botones"
      Height          =   270
      Left            =   1680
      TabIndex        =   9
      Top             =   5370
      Width           =   6195
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
      Top             =   5200
      Width           =   225
   End
End
Attribute VB_Name = "DespachosDevolucionesDetallesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancelar_Click()
'vMostrarDetallesDevoluciones = 0 CAMBIE ESTO AQUI A COMENTARIO Y DE INTEGER A STRING
DespachosDevolucionesNewForm.EstilosGrillaDespachoDevoluciones
DespachosDevolucionesNewForm.BorrarTemporalDevolucion
Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
ConsultasForm.GrillaConsultas.Refresh
ConsultasForm.EstilosGrillaConsultas
Unload Me
End Sub

Private Sub cmd_Descontar_Click()
'Verificamos si esta vacio
With RsTemporalDevoluciones
          If .BOF And .EOF Then
               MsgBox ("No existe un producto disponible al cual descantarle una cantidad, restablezca el despacho e intente de nuevo"), vbInformation, "Aviso": cmd_Restablecer.SetFocus
               Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
               ConsultasForm.EstilosGrillaConsultas
               Exit Sub
          End If
End With

'le pedimos que ingrese la cantidad a descontar

vDescuento = InputBox("Por favor ingrese la cantidad que desea descontar del producto", "Ingrese cantidad")
If vDescuento = "" Then
          MsgBox ("Proceso cancelado"), vbInformation, "Aviso": Exit Sub
ElseIf IsNumeric(vDescuento) = False Then
          MsgBox ("Por favor, no ingrese letras"), vbInformation, "Aviso": Exit Sub
ElseIf Val(vDescuento) <= 0 Then
          MsgBox ("La cantidad introducida es inválida, por favor rectifique e intente de nuevo"), vbInformation, "Aviso": Exit Sub
ElseIf Val(vDescuento) > Val(GrillaContenidoDevolucion.Columns(4).Text) Then
          MsgBox ("No puede ingresar una cantidad mayor de la que dispone el producto en el despacho, por favor rectifique e intente de nuevo"), vbInformation, "Aviso": Exit Sub
ElseIf vDescuento <> CInt(vDescuento) Then
          MsgBox ("Se ha detectado el uso de un '.' por favor solo ingrese números (enteros)"), vbInformation, "Aviso": Exit Sub
ElseIf Val(vDescuento) - Int(Val(vDescuento)) <> 0 Then
           MsgBox ("Se ha detectado el uso de una ',' por favor solo ingrese números (enteros)"), vbInformation, "Aviso": Exit Sub
Else
          'obtenemo el codigo del producto a eliminar del temporal
          vDProductosDevolucion = Trim(GrillaContenidoDevolucion.Columns(2).Text)
          
          With RsTemporalDevoluciones
                    .Requery
                    .Find "CodigoProducto='" & Trim(vDProductosDevolucion) & "'"
                              !CANTProductos = Val(!CANTProductos) - Val(vDescuento)
                              If Val(!CANTProductos) = 0 Then .Delete
                    .UpdateBatch
                    .Requery
          End With
          
          Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
          ConsultasForm.EstilosGrillaConsultas
          
          Set GrillaContenidoDevolucion.DataSource = RsTemporalDevoluciones
          EstiloGrillaVerDevolucionesDetalles

          vDProductosDevolucion = ""
          vDespachador = 1
          Label13.Caption = vDespachador
End If
End Sub

Private Sub cmd_Devolver_Click()
'//Querido colega programador:
'//
'//Cuando escribí este código, sólo Dios y yo
'//sabíamos como funcionaba
'//Ahora, ¡Sólo Dios lo sabe!
'//
'//Así que si está tratando de 'optimizarlo'
'// y fracasa (seguramente), por favor,
'// incremente el contador a continuación
'//como advertencia para su siguiente colega:
'//
'//total_horas_perdidas_aquí = 180

'Validamos que este dispononible la devolucion por limite de tiempo (fecha)
Dim FechaDespacho As Date
Dim FechaDevolucion As Date
Dim Dias As Integer

FechaDespacho = Format(txt_Fecha, "dd/mm/yyyy") 'guardamos la fecha del txt_fecha (fecha originaria del despacho)
FechaDevolucion = Format(DTP_fecha, "dd/mm/yyyy") 'guardamos la fecha del dtp_fecha (fecha actual o sea cuando se hace el despacho)
Dias = DateDiff("d", FechaDespacho, FechaDevolucion) 'si fuera dias utilzo "d" y si fuera años utilizo "yyyy" (hacemos la diferencia

lbl_DiferenciaTiempo.Caption = Format(Dias, "###0")
If lbl_DiferenciaTiempo.Caption > Val(3) Then MsgBox ("No puedes devolver una mercancía enviada hace 3 días"), vbInformation, "Aviso": Exit Sub


'observacion obligatoria si es otros motivos
If cbo_Motivo.Text = "Otros motivos" And txt_Observaciones.Text = "" Then MsgBox ("Por haber seleccionado 'Otros motivos' debe de ingresar una observación"), vbInformation, "Aviso": txt_Observaciones.SetFocus: Exit Sub


'Confirmamos si no ingresara observaciones
If txt_Observaciones.Text = "" Then
          If MsgBox("Esta a punto de continuar con la devolución sin espeficicar alguna observación, ¿Desea continuar de todos modos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
               'no hago nada continuo con mi proceso
          Else
               txt_Observaciones.SetFocus
               Exit Sub
          End If
Else
End If

'validamos que exista al menos un producto para devolver
With RsTemporalDevoluciones
          .Requery
          If .BOF And .EOF = True Then
               EstiloGrillaVerDevolucionesDetalles
               MsgBox ("Debe dejar al menos un (1) producto en la devolución para seguir con la operación"), vbInformation, "Aviso"
               cmd_Restablecer.SetFocus
               Exit Sub
          End If
End With

'validamos que no se haya devuelto por segunda vez dicho pedido
With RsDevoluciones
          .Requery
          .Find "CodigoDespacho='" & Trim(txt_CodFactura.Text) & "'"
          If .EOF Then
                    'no hago nada
          Else
                    MsgBox ("Este despacho ya fue devuelto con anterioridad, por lo tanto no se puede devolver nuevamente. Rectifique"), vbInformation, "Aviso": cmd_cancelar.SetFocus: Exit Sub
          End If
End With
                                                                                                ' F I N     D E    V A L I D A C I O N E S
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

'CREAMOS LA DEVOLUCION
'comparamos si ha sufrido alguna modificacion para determinar con que funcion vamos a grabar
If Val(vDespachador) = 1 Then
     GrabarConModificaciones 'grabar¿?
     EliminarDescontando
     ImprimirModificado
ElseIf Val(RsTemporalDevoluciones.RecordCount) < Val(RsDetallesDespacho.RecordCount) Then 'hice modificaciones?
      GrabarConModificaciones 'grabo con moficaciones
     EliminarParcial
      ImprimirModificado
Else 'deje todo igual?
    GrabarNormal 'grabo normal entonces
    EliminarTodo
    ImprimirTodo
End If
                                                                           ' F I N     D E    C R E A R   D E V O L U C I O N
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/


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

'EstiloGrillaVerDevolucionesDetalles ES ESTEEEEEEEEEE MADAFACAAAAA

'para que consultas no se vea afectada
Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
ConsultasForm.EstilosGrillaConsultas


End Sub

Private Sub cmd_Restablecer_Click()
If MsgBox("¿Esta a punto de reestablecer el despacho a su estado original, ¿Desea continuar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
          DespachosDevolucionesNewForm.BorrarTemporalDevolucion
          VerProductosDelDespacho
Else
          Exit Sub
End If
vDespachador = 0
Label13.Caption = vDespachador
Label13.Refresh
End Sub

Private Sub cmd_Sacar_Click()
'verificamos si esta vacio el despacho
With RsTemporalDevoluciones
          If .RecordCount = 0 Then
               MsgBox ("No hay producto disponible para sacar, restablezca el despacho e intente de nuevo"), vbInformation, "Aviso": cmd_Restablecer.SetFocus
               Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
               ConsultasForm.EstilosGrillaConsultas
               Exit Sub
          End If
End With

'obtenemo el codigo del producto a eliminar del temporal
vDProductosDevolucion = GrillaContenidoDevolucion.Columns(2).Text

'preguntamos primero si lo quiere eliminar
If MsgBox("¿Está seguro que quiere sacar éste producto de la devolución?", vbInformation + vbYesNo, "Advertencia") = vbYes Then
          'si dice si lo elimino
          With RsTemporalDevoluciones
                    .Requery
                              .Find "CodigoProducto='" & Trim(vDProductosDevolucion) & "'" 'lo ubico
                              .Delete 'LO BORRO
                              .Requery
                              EstiloGrillaVerDevolucionesDetalles
                              vDProductosDevolucion = ""
          End With
Else
          Exit Sub
End If
Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
ConsultasForm.EstilosGrillaConsultas

Set GrillaContenidoDevolucion.DataSource = RsTemporalDevoluciones
EstiloGrillaVerDevolucionesDetalles

End Sub

Private Sub Form_Load()

'tablas
Productos
Despacho
DetallesDespacho
Devoluciones
DevolucionesDetalles
TemporalDevoluciones

'funciones
VerDespacho
VerProductosDelDespacho

'fecha de hoy
DTP_fecha = Now
'desabilito que puedan editar fecha manualmente OPCIONAL**
'DTP_fecha.Enabled = False

'estilo
DespachosDevolucionesNewForm.EstilosGrillaDespachoDevoluciones

cbo_Motivo.AddItem "Mercancía Dañada/Defectuosa"
cbo_Motivo.AddItem "Falta de presupuesto"
cbo_Motivo.AddItem "Inconformidad con el artículo"
cbo_Motivo.AddItem "Mercancía añadida por accidente"
cbo_Motivo.AddItem "Otros motivos"
cbo_Motivo.ListIndex = 0

txt_DevueltoPor.Text = vUsername
txt_DevueltoPor.Refresh

End Sub

Sub VerDespacho()
With RsDespacho
          .Requery
          .Find "Id='" & Val(vMostrarDetallesDevoluciones) & "'"
          If .EOF Then
                    Exit Sub
          Else
                   'igualamos los campos
                    txt_CodFactura.Text = !CodigoDespacho 'error 3021 aqui
                    txt_NombreCliente.Text = !Cliente
                    txt_Fecha.Text = !Fecha
                    txt_NombreDespachador.Text = !Despachador
                    txt_ZonaDespacho.Text = !Zona
                    txt_NombreVendedor.Text = !Vendedor
          End If
End With
End Sub

Sub VerProductosDelDespacho()

With RsDetallesDespacho
          If .State = 1 Then .Close
          Busca = UCase(Trim(txt_CodFactura.Text))
          .Open "select * from DetallesDespacho WHERE CodigoDespacho  like '" & Busca & "'"
          .Requery
          
          If .EOF Then
                    Exit Sub
          
          Else
                    Set GrillaContenidoDevolucion.DataSource = RsDetallesDespacho
          
                    Dim Devolucion As Integer
                    Devolucion = RsDetallesDespacho.RecordCount
                    RsDetallesDespacho.Requery
                    RsDetallesDespacho.MoveFirst 'error 3021
          
          'el problema  empezaba aqui V
                    For x = 1 To Devolucion
                              With RsTemporalDevoluciones
                                             .Requery
                                             .AddNew
                                                       !CodigoDespacho = txt_CodFactura.Text
                                                       !CodigoProducto = GrillaContenidoDevolucion.Columns(2).Text
                                                       !DescripcionProducto = GrillaContenidoDevolucion.Columns(3).Text
                                                       !CANTProductos = GrillaContenidoDevolucion.Columns(4).Text
                                                       !Marca = GrillaContenidoDevolucion.Columns(5).Text
                                                       !Kit = GrillaContenidoDevolucion.Columns(6).Text
                                                       !PiezasxKit = GrillaContenidoDevolucion.Columns(7).Text
                                             .Update
                              End With
                              If x = Devolucion Then Else RsDetallesDespacho.MoveNext
                    Next
          End If
'terminaba aqui ^
End With
'Set GrillaContenidoDevolucion.DataSource = RsDetallesDespacho
EstiloGrillaVerDevolucionesDetalles
End Sub

Sub GrabarNormal()
'Creamos la devolucion
With RsDevoluciones
        .Requery
        .AddNew
                !CodigoDespacho = txt_CodFactura.Text
                !Cliente = txt_NombreCliente.Text
                !Fecha = txt_Fecha.Text
                !FechaDevo = DTP_fecha
                !Despachador = txt_NombreDespachador.Text
                !Zona = txt_ZonaDespacho.Text
                !Vendedor = txt_NombreVendedor
                !By = txt_DevueltoPor.Text
                !Observaciones = txt_Observaciones.Text
                !Motivo = cbo_Motivo.Text
                !Tipo = "Devolución Total"
        .Update
 End With
'añadimos los detalles de la devolucion  y sumamos

Dim Devo As Integer
Devo = RsDetallesDespacho.RecordCount 'tambien habia problema aqui puesto que tenia puesto para el temporal
RsDetallesDespacho.Requery 'tambien habia problema aqui puesto que tenia puesto para el temporal
RsDetallesDespacho.MoveFirst 'tambien habia problema aqui puesto que tenia puesto para el temporal
For x = 1 To Devo

If cbo_Motivo.Text = "Mercancía Dañada/Defectuosa" Then

Else
    'SUMAMOS LOS PRODUCTOS
        With RsProductos
               .Requery
                        .Find "CodigoProducto='" & Trim(GrillaContenidoDevolucion.Columns(2).Text) & "'"
                        !Cantidad = !Cantidad + Val(GrillaContenidoDevolucion.Columns(4).Text)
                .UpdateBatch
        End With
End If

        With RsDevolucionesDetalles
                .Requery
                .AddNew
                        !CodigoDespacho = GrillaContenidoDevolucion.Columns(1).Text
                        !CodigoProducto = GrillaContenidoDevolucion.Columns(2).Text
                        !DescripcionProducto = GrillaContenidoDevolucion.Columns(3).Text
                        !CANTProductos = GrillaContenidoDevolucion.Columns(4).Text
                        !Marca = GrillaContenidoDevolucion.Columns(5).Text
                                If GrillaContenidoDevolucion.Columns(6).Value = 0 Then
                                        !Kit = 0
                                Else
                                        !Kit = 1
                                End If
                        !PiezasxKit = GrillaContenidoDevolucion.Columns(7).Text
                .Update
        End With
        If x = Devo Then Else RsDetallesDespacho.MoveNext
  Next
EstiloGrillaVerDevolucionesDetalles
Logs
With RsLogs
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "Realizó una  devolución total del despacho asociado al código de factura N° [" & txt_CodFactura & "]"
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
                    !Accion = "Realizó una  devolución total del despacho asociado al código de factura N° [" & txt_CodFactura & "]"
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With
MsgBox "La devolución fue creada exitosamente", vbInformation, "Aviso"
End Sub

Sub GrabarConModificaciones()
'Creamos la devolucion
With RsDevoluciones
        .Requery
        .AddNew
                !CodigoDespacho = txt_CodFactura.Text
                !Cliente = txt_NombreCliente.Text
                !Fecha = txt_Fecha.Text
                !FechaDevo = DTP_fecha
                !Despachador = txt_NombreDespachador.Text
                !Zona = txt_ZonaDespacho.Text
                !Vendedor = txt_NombreVendedor
                !By = txt_DevueltoPor.Text
                !Observaciones = txt_Observaciones.Text
                !Motivo = cbo_Motivo.Text
                !Tipo = "Devolución Parcial"
        .Update
        .Requery
 End With
 
'añadimos los detalles de la devolucion con modificaciones y sumamos
Dim DevoM As Integer
DevoM = RsTemporalDevoluciones.RecordCount 'tambien habia problema aqui puesto que tenia puesto para el temporal
RsTemporalDevoluciones.Requery 'tambien habia problema aqui puesto que tenia puesto para el temporal
RsTemporalDevoluciones.MoveFirst 'tambien habia problema aqui puesto que tenia puesto para el temporal
For x = 1 To DevoM

If cbo_Motivo.Text = "Mercancía Dañada/Defectuosa" Then
          'no hago nada
Else
    'SUMAMOS LOS PRODUCTOS
        With RsProductos
               .Requery
                        .Find "CodigoProducto='" & Trim(GrillaContenidoDevolucion.Columns(2).Text) & "'"
                        !Cantidad = !Cantidad + Val(GrillaContenidoDevolucion.Columns(4).Text)
                .UpdateBatch
        End With
End If
        With RsDevolucionesDetalles
                .Requery
                .AddNew
                        !CodigoDespacho = GrillaContenidoDevolucion.Columns(1).Text
                        !CodigoProducto = GrillaContenidoDevolucion.Columns(2).Text
                        !DescripcionProducto = GrillaContenidoDevolucion.Columns(3).Text
                        !CANTProductos = GrillaContenidoDevolucion.Columns(4).Text
                        !Marca = GrillaContenidoDevolucion.Columns(5).Text
                                If GrillaContenidoDevolucion.Columns(6).Value = 0 Then
                                        !Kit = 0
                                Else
                                        !Kit = 1
                                End If
                        !PiezasxKit = GrillaContenidoDevolucion.Columns(7).Text
                .Update
        End With
        If x = DevoM Then Else RsTemporalDevoluciones.MoveNext
  Next
EstiloGrillaVerDevolucionesDetalles
Logs
With RsLogs
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "Realizó una  devolución parcial del despacho asociado al código de factura N° [" & txt_CodFactura & "]"
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
                    !Accion = "Realizó una  devolución parcial del despacho asociado al código de factura N° [" & txt_CodFactura & "]"
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With
MsgBox "La devolución fue creada exitosamente", vbInformation, "Aviso"
End Sub
Sub EliminarTodo()
     With RsDespacho
               .Requery
               .Find "CodigoDespacho='" & Trim(txt_CodFactura.Text) & "'"
               If .EOF Then
                            'no hago nada
               Else 'si lo encuentro
                         .Delete
                         .Requery
               End If
     End With
End Sub
Sub EliminarParcial()
     With RsDetallesDespacho
               If .State = 1 Then .Close
               Busca = UCase(Trim(txt_CodFactura.Text))
               .Open "select * from DetallesDespacho WHERE CodigoDespacho like '" & Busca & "'"
               .Requery

               Dim Cierra As Integer
               Cierra = 1
               RsDetallesDespacho.MoveFirst
               
               With RsTemporalDevoluciones
                         .Requery
                         .MoveFirst
               End With
               
               While Cierra = 1
                          If !CodigoDespacho = Trim(txt_CodFactura.Text) And !CodigoProducto = Trim(GrillaContenidoDevolucion.Columns(2).Text) Then
                                   .Delete
                                   .Requery 'cambie por update
                                   RsDetallesDespacho.MoveFirst
                                   With RsTemporalDevoluciones
                                             .Requery
                                             .Find "CodigoProducto='" & GrillaContenidoDevolucion.Columns(2).Text & "'"
                                                       .Delete
                                                       .Requery
                                                       If (.BOF And .EOF) = True Then
                                                                 Cierra = 0
                                                       Else
                                                                 .MoveFirst
                                                       End If
                                   End With
                         Else
                                   RsDetallesDespacho.MoveNext
                         End If
               Wend
               
     End With
     EstiloGrillaVerDevolucionesDetalles
End Sub
Sub EliminarDescontando()
     With RsDetallesDespacho
               If .State = 1 Then .Close
               Busca = UCase(Trim(txt_CodFactura.Text))
               .Open "select * from DetallesDespacho WHERE CodigoDespacho like '" & Busca & "'"
               .Requery

               Dim CierraDescontado As Integer
               CierraDescontado = 1
               RsDetallesDespacho.MoveFirst
               
               With RsTemporalDevoluciones
                         If .State = 1 Then .Close
                         .Open
                         .Requery
                         .MoveFirst
               End With
               
               While CierraDescontado = 1
                          If !CodigoDespacho = Trim(txt_CodFactura.Text) And !CodigoProducto = Trim(GrillaContenidoDevolucion.Columns(2).Text) Then
                                   If Val(!CANTProductos) > Val(GrillaContenidoDevolucion.Columns(4).Text) Then
                                             If (Val(!CANTProductos) - Val(GrillaContenidoDevolucion.Columns(4).Text)) <> 0 Then
                                                       !CANTProductos = Val(!CANTProductos) - Val(GrillaContenidoDevolucion.Columns(4).Text)
                                                       .UpdateBatch
                                             Else
                                                       .Delete
                                             End If
                                   Else
                                             If (Val(GrillaContenidoDevolucion.Columns(4).Text) - Val(!CANTProductos)) <> 0 Then
                                                       !CANTProductos = Val(GrillaContenidoDevolucion.Columns(4).Text) - Val(!CANTProductos)
                                                       .UpdateBatch
                                             Else
                                                       .Delete
                                             End If
                                   End If
                                   .Requery 'cambie por update
                                   RsDetallesDespacho.MoveFirst
                                   With RsTemporalDevoluciones
                                             .Requery
                                             .Find "CodigoProducto='" & GrillaContenidoDevolucion.Columns(2).Text & "'"
                                                       .Delete
                                                       .Requery
                                                       If (.BOF And .EOF) = True Then
                                                                 CierraDescontado = 0
                                                       Else
                                                                 .MoveFirst
                                                       End If
                                   End With
                         Else
                                   RsDetallesDespacho.MoveNext
                         End If
               Wend
               
     End With
     EstiloGrillaVerDevolucionesDetalles
End Sub


Sub ImprimirModificado()
'igualamos variables

vNFacturaDevo = txt_CodFactura.Text '
vClienteDevo = txt_NombreCliente.Text '
vMotivo = cbo_Motivo.Text '

vZonaDevo = txt_ZonaDespacho.Text '
vVendedorDevo = txt_NombreVendedor.Text '
vObservaciones = txt_Observaciones.Text '

vEntregado = txt_Fecha.Text '
vDevuelto = DTP_fecha '
vDespachadorDevo = txt_NombreDespachador.Text '
vDevueltoPor = txt_DevueltoPor.Text '

Unload Me


With AdoImprimirModificado
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
          Busca = Trim(txt_CodFactura.Text)
          .RecordSource = "SELECT * FROM DevolucionesDetalles WHERE CodigoDespacho LIKE '" & Busca & "'"
          .Refresh
End With

Set dr_RDevolucionNew.DataSource = AdoImprimirModificado
dr_RDevolucionNew.WindowState = 2

'Seccion 4
dr_RDevolucionNew.Sections("Sección4").Controls("Etiqueta3").Caption = vNFacturaDevo

'Seccion 2
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta12").Caption = vNFacturaDevo
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta17").Caption = vClienteDevo
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta18").Caption = vDespachadorDevo
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta22").Caption = vDevueltoPor
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta24").Caption = vMotivo
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta20").Caption = vZonaDevo
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta19").Caption = vEntregado
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta26").Caption = vDevuelto
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta28").Caption = vObservaciones

'lo mostramos
dr_RDevolucionNew.Show
End Sub

Sub ImprimirTodo()

'igualamos variables

vNFacturaDevo = txt_CodFactura.Text '
vClienteDevo = txt_NombreCliente.Text '
vMotivo = cbo_Motivo.Text '

vZonaDevo = txt_ZonaDespacho.Text '
vVendedorDevo = txt_NombreVendedor.Text '
vObservaciones = txt_Observaciones.Text '

vEntregado = txt_Fecha.Text '
vDevuelto = DTP_fecha '
vDespachadorDevo = txt_NombreDespachador.Text '
vDevueltoPor = txt_DevueltoPor.Text '

Unload Me

Set dr_RDevolucionNew.DataSource = RsTemporalDevoluciones
dr_RDevolucionNew.WindowState = 2

'Seccion 4
dr_RDevolucionNew.Sections("Sección4").Controls("Etiqueta3").Caption = vNFacturaDevo

'Seccion 2
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta12").Caption = vNFacturaDevo
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta17").Caption = vClienteDevo
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta18").Caption = vDespachadorDevo
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta22").Caption = vDevueltoPor
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta24").Caption = vMotivo
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta20").Caption = vZonaDevo
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta19").Caption = vEntregado
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta26").Caption = vDevuelto
dr_RDevolucionNew.Sections("Sección2").Controls("Etiqueta28").Caption = vObservaciones

'lo mostramos
dr_RDevolucionNew.Show

End Sub
Sub EstiloGrillaVerDevolucionesDetalles()
                
     'TAMAÑOS
                GrillaContenidoDevolucion.Columns(0).Width = 0
                GrillaContenidoDevolucion.Columns(1).Width = 0
                GrillaContenidoDevolucion.Columns(2).Width = 2750
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



