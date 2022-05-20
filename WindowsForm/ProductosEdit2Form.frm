VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ProductosEdit2Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar el Producto Seleccionado"
   ClientHeight    =   8280
   ClientLeft      =   1860
   ClientTop       =   1515
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
   Icon            =   "ProductosEdit2Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ProductosEdit2Form.frx":08CA
   ScaleHeight     =   8280
   ScaleWidth      =   17985
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Cbo_Ubicacion 
      Height          =   390
      Left            =   12840
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   2700
      Width           =   3855
   End
   Begin VB.CommandButton cmd_AgregarImagen 
      Caption         =   "Cambiar Imagen"
      Height          =   495
      Left            =   12840
      TabIndex        =   28
      Top             =   3520
      Width           =   1935
   End
   Begin VB.CommandButton cmd_ElimarImagen 
      Caption         =   "Eliminar Imagen"
      Height          =   495
      Left            =   14880
      TabIndex        =   27
      Top             =   3520
      Width           =   1835
   End
   Begin VB.CommandButton cmd_atras 
      Appearance      =   0  'Flat
      Caption         =   "Atras"
      Height          =   615
      Left            =   6360
      TabIndex        =   12
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton cmd_guardar 
      Caption         =   "Terminar editado"
      Height          =   615
      Left            =   9360
      TabIndex        =   11
      Top             =   7320
      Width           =   2295
   End
   Begin VB.TextBox txt_CodProducto 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   18
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2700
      Width           =   3855
   End
   Begin VB.TextBox txt_DescripcionProducto 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1920
      MaxLength       =   79
      TabIndex        =   1
      Top             =   3520
      Width           =   3855
   End
   Begin VB.TextBox txt_cantproducto 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      MaxLength       =   5
      TabIndex        =   2
      Top             =   2700
      Width           =   3855
   End
   Begin VB.TextBox txt_marcaproducto 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      MaxLength       =   18
      TabIndex        =   3
      Top             =   3520
      Width           =   3855
   End
   Begin VB.TextBox txt_cantmaxventa 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   12840
      MaxLength       =   5
      TabIndex        =   13
      Top             =   5200
      Width           =   3855
   End
   Begin VB.TextBox txt_by 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      MaxLength       =   30
      TabIndex        =   8
      Top             =   6050
      Width           =   3855
   End
   Begin VB.TextBox txt_piezasxkit 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   6
      Top             =   6050
      Width           =   3855
   End
   Begin VB.TextBox txt_cantminventa 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   12840
      MaxLength       =   5
      TabIndex        =   10
      Top             =   6050
      Width           =   3855
   End
   Begin VB.OptionButton opt_kityes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Si"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   5200
      Width           =   615
   End
   Begin VB.OptionButton opt_kitno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "No"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   5200
      Width           =   615
   End
   Begin MSComCtl2.DTPicker DTP_fecha 
      Height          =   495
      Left            =   7200
      TabIndex        =   7
      Top             =   5200
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   393216
      Format          =   115539969
      CurrentDate     =   42645
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   16800
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación del producto"
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
      TabIndex        =   31
      Top             =   3180
      Width           =   1710
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imagen referencial para reconocer el producto"
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
      Top             =   4040
      Width           =   3420
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
      TabIndex        =   26
      Top             =   1920
      Width           =   225
   End
   Begin VB.Label Label7 
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
      TabIndex        =   25
      Top             =   4440
      Width           =   225
   End
   Begin VB.Label lbl_EJusername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código del producto"
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
      Top             =   3180
      Width           =   1485
   End
   Begin VB.Label lbl_descripcion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rellene los datos Principales"
      Height          =   270
      Left            =   1680
      TabIndex        =   23
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción del producto"
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
      TabIndex        =   22
      Top             =   4035
      Width           =   1860
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad del producto"
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
      TabIndex        =   21
      Top             =   3180
      Width           =   1620
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca del producto"
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
      TabIndex        =   20
      Top             =   4040
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rellene los datos Secundarios"
      Height          =   270
      Left            =   1680
      TabIndex        =   19
      Top             =   4590
      Width           =   3165
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad máxima para la venta (Por cliente)"
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
      TabIndex        =   9
      Top             =   5700
      Width           =   3240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aplicable para (Información extra)"
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
      TabIndex        =   18
      Top             =   6550
      Width           =   2535
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha que llegó el producto"
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
      Top             =   5700
      Width           =   2055
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Piezas por kit"
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
      Top             =   6550
      Width           =   1020
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad mínima para la venta (Por cliente)"
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
      Top             =   6555
      Width           =   3195
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "indique si el producto es kit"
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
      Top             =   5700
      Width           =   2040
   End
End
Attribute VB_Name = "ProductosEdit2Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ComparoImagen As String
Private Sub cmd_AgregarImagen_Click()
Dialogo.DialogTitle = "Seleccione una imagen (Solo JPG o GIF)"
Dialogo.Filter = "Archivos Jpg|*.jpg|Archivos Gif|*.gif|"
Dialogo.ShowSave


RutaOrigen = Dialogo.FileName
' ImageProducto.Picture = LoadPicture(RutaOrigen) AQUI NO MUETRO IMAGEN
ArchivoNombre = Dialogo.FileTitle

End Sub

Private Sub cmd_Atras_Click()
Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
ConsultasForm.EstilosGrillaConsultas
vIDProductos = "1"
ArchivoNombre = ""
Unload Me 'PONER UNLOADME EN TODOS LOS CERRAR OJO
End Sub

Private Sub cmd_ElimarImagen_Click()
ArchivoNombre = "" 'el app.path esta configurado (mas bien ya usado en la ruta donde esta la db
'ImageProducto.Picture = LoadPicture(RutaOrigen) AQUI NO MUESTRO IMAGEN
End Sub

Private Sub cmd_guardar_Click()

'validamos todos los campos
If txt_CodProducto.Text = "" Then MsgBox "El campo donde va el CÓDIGO DEL PRODUCTO no puede estar vacío", vbInformation, "Aviso": txt_CodProducto.SetFocus: Exit Sub
If txt_DescripcionProducto = "" Then MsgBox "El campo donde va la DESCRIPCIÓN DEL PRODUCTO no puede estar vacío", vbInformation, "Aviso": txt_DescripcionProducto.SetFocus: Exit Sub
If txt_cantproducto.Text = "" Then MsgBox "El campo donde va la CANTIDAD DEL PRODUCTO no puede estar vacío", vbInformation, "Aviso": txt_cantproducto.SetFocus: Exit Sub
          If Not IsNumeric(txt_cantproducto) Then MsgBox "Por favor, en el campo CANTIDAD debe ir sólo números", vbInformation, "Aviso": txt_cantproducto.SetFocus: Exit Sub
          If Val(txt_cantproducto.Text) <= 0 Then MsgBox ("La cantidad ingresada del producto, es inválida"), vbInformation, "Aviso": txt_cantproducto.SetFocus: Exit Sub
If txt_marcaproducto = "" Then MsgBox "El campo donde va la MARCA DEL PRODUCTO no puede estar vacío", vbInformation, "Aviso": txt_marcaproducto.SetFocus: Exit Sub
If Cbo_Ubicacion.Text = "" Then MsgBox ("Por favor, seleccione la UBICACIÓN DEL PRODUCTO"), vbInformation, "Aviso": Cbo_Ubicacion.SetFocus: Exit Sub
'If txt_proveedorproducto = "" Then MsgBox "El campo donde va el PROVEEDOR DEL PRODUCTO no puede estar vacio", vbInformation, "Aviso": txt_proveedorproducto.SetFocus: Exit Sub
If opt_kityes.Value = False And opt_kitno.Value = False Then MsgBox "Indique si el producto ES KIT O NO", vbInformation, "Aviso": Exit Sub
          If txt_piezasxkit.Text = "" Then MsgBox ("Por favor especifique cuantos productos contiene el kit"), vbInformation, "Aviso": txt_piezasxkit.SetFocus: Exit Sub
          If opt_kityes.Value = True And Val(txt_piezasxkit.Text) < 2 Then MsgBox ("Error. Un producto que es kit no puede contener  menos de dos (2) piezas"), vbInformation, "Aviso": txt_piezasxkit.SetFocus: Exit Sub
If txt_by.Text = "" Then MsgBox "Por favor, especifique para QUE SE USA el producto", vbInformation, "Aviso": txt_by.SetFocus: Exit Sub
If txt_cantmaxventa.Text = "" Then MsgBox "Por favor, especifique cual es la CANTIDAD MÁXIMA PARA LA VENTA", vbInformation, "Aviso": txt_cantmaxventa.SetFocus: Exit Sub
          If Not IsNumeric(txt_cantmaxventa.Text) Then MsgBox "Por favor, en el campo CANTIDAD MÁXIMA PARA LA VENTA debe ir sólo numeros", vbInformation, "Aviso": txt_cantmaxventa.SetFocus: Exit Sub
          If Val(txt_cantmaxventa.Text) <= 0 Then MsgBox ("Error. Cantidad máxima para la venta inválida."), vbInformation, "Aviso": txt_cantmaxventa.SetFocus: Exit Sub
If txt_cantminventa.Text = "" Then MsgBox "Por favor, especifique cual es la CANTIDAD MÍNIMA PARA LA VENTA", vbInformation, "Aviso": txt_cantminventa.SetFocus: Exit Sub
          If Not IsNumeric(txt_cantminventa.Text) Then MsgBox "Por favor, en el campo CANTIDAD MÍNIMA PARA LA VENTA debe ir sólo números", vbInformation, "Aviso": txt_cantminventa.SetFocus: Exit Sub
          If Val(txt_cantminventa.Text) <= 0 Then MsgBox ("Error. Cantidad mínima para la venta inválida"), vbInformation, "Aviso": txt_cantminventa.SetFocus: Exit Sub
If Val(txt_cantminventa.Text) > Val(txt_cantmaxventa.Text) Then MsgBox ("La cantidad mínima no puede ser mayor a la cantidad máxima para la venta"), vbInformation, "Aviso": txt_cantminventa.SetFocus: Exit Sub

'editamos el producto sin validar
With RsProductos
    .Requery 'actualizamos la tabla
    .Find "CodigoProducto='" & Trim(txt_CodProducto.Text) & "'"
        
        !descripcion = txt_DescripcionProducto.Text
        !By = txt_by.Text
        !Cantidad = txt_cantproducto.Text
        !Fecha = DTP_fecha.Value
        If opt_kityes = True Then
            !Kit = 1
        Else
            !Kit = 0
        End If
        !PiezaxKit = txt_piezasxkit.Text
        !CantMAXVenta = txt_cantmaxventa.Text
        !CantMINVenta = txt_cantminventa.Text
        '!Proveedor = txt_proveedorproducto.Text
        !Ubicacion = Cbo_Ubicacion.Text
        !Marca = txt_marcaproducto.Text
        
        If ArchivoNombre = "" Then
            !Imagen = "\ImagenesProductos\predeterminada.jpg"
        Else
               If UCase(Trim(ArchivoNombre)) <> UCase(Trim(ComparoImagen)) Then
                         !Imagen = "\ImagenesProductos\" & Trim(ArchivoNombre)
               Else
                         !Imagen = Trim(ArchivoNombre)
               End If
        End If
        
    .UpdateBatch 'actualizamos datos
    
     Logs
     With RsLogs
               .Requery
               .AddNew
                         !User = vUsername
                         !Accion = "editó el producto asociado al codigo [" & txt_CodProducto.Text & "]"
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
                         !Accion = "editó el producto asociado al codigo [" & txt_CodProducto.Text & "]"
                         !Fecha = Date + Time
               .Update
               .Requery
               .Close
     End With
    
    Set ProductosEditForm.GrillaProductos.DataSource = RsProductos
    Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
    
    'NombredelFormulario+Estilo
    ProductosEditForm.EstilosGrillaProducto
    ConsultasForm.EstilosGrillaConsultas
    
    
    MsgBox "Producto  editado con exito", vbInformation, "Aviso"
    vIDProductos = 0 ' no sirve no hace nada hasta que edite algo
    Unload Me 'segun tutorial
End With

If ArchivoNombre = "" Then
Else
    RutaDestino = App.Path & "\ImagenesProductos\"
    Set GDB = Nothing
    Set fs = CreateObject("scripting.filesystemobject")
    fs.copyfile RutaOrigen, RutaDestino
    End If
    
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

ArchivoNombre = ""
End Sub

Private Sub Form_Load()

Cbo_Ubicacion.AddItem "Segmento 1 (S1)"
Cbo_Ubicacion.AddItem "Segmento 2 (S2)"
Cbo_Ubicacion.AddItem "Segmento 3 (S3)"
Cbo_Ubicacion.AddItem "Segmento 4 (S4)"
Cbo_Ubicacion.AddItem "Segmento 5 (S5)"
Cbo_Ubicacion.AddItem "Segmento 6 (S6)"
Cbo_Ubicacion.AddItem "Segmento 7 (S7)"
Cbo_Ubicacion.AddItem "Segmento 8 (S8)"
Cbo_Ubicacion.AddItem "Segmento 9 (S9)"
Cbo_Ubicacion.AddItem "Segmento 10 (S10)"
Cbo_Ubicacion.AddItem "Segmento 11 (S11)"
Cbo_Ubicacion.AddItem "Segmento 12 (S12)"
Cbo_Ubicacion.AddItem "Segmento 13 (S13)"
Cbo_Ubicacion.AddItem "Segmento 14 (S14)"
Cbo_Ubicacion.AddItem "Segmento 15 (S15)"
Cbo_Ubicacion.AddItem "Segmento 16 (S16)"
Cbo_Ubicacion.AddItem "Segmento 17 (S17)"
Cbo_Ubicacion.AddItem "Segmento 18 (S18)"
Cbo_Ubicacion.AddItem "Segmento 19 (S19)"

Cbo_Ubicacion.ListIndex = 0

CargarProductos
RutaOrigen = App.Path & ProductosEditForm.GrillaProductos.Columns(12).Text
End Sub

Sub CargarProductos()
With RsProductos
     .Requery
    .Find "Id='" & Val(vIDProductos) & "'"
        
    'igualamos los campos para mostrarlos
       txt_CodProducto.Text = !CodigoProducto
       txt_DescripcionProducto.Text = !descripcion
       txt_by.Text = !By
       txt_cantproducto.Text = !Cantidad
       DTP_fecha.Value = !Fecha
       txt_piezasxkit.Text = !PiezaxKit
       txt_cantmaxventa.Text = !CantMAXVenta
       txt_cantminventa.Text = !CantMINVenta
      ' txt_proveedorproducto.Text = !Proveedor
      Cbo_Ubicacion.Text = !Ubicacion
       txt_marcaproducto.Text = !Marca
       ArchivoNombre = !Imagen
       
       ComparoImagen = ArchivoNombre
    
            If !Kit = 1 Then
                opt_kityes = True
                txt_piezasxkit.Enabled = True
            Else
                opt_kitno = True
                txt_piezasxkit.Enabled = False
                txt_piezasxkit.Text = "1"
            End If
    
End With
End Sub

Private Sub opt_kitno_Click()
 If opt_kityes.Value = True Then
       txt_piezasxkit.Enabled = True
       txt_piezasxkit.Text = ""
    Else
        txt_piezasxkit.Enabled = False
        txt_piezasxkit.Text = "1"
    End If
End Sub

Private Sub opt_kityes_Click()
 If opt_kityes.Value = True Then
       txt_piezasxkit.Enabled = True
    Else
        txt_piezasxkit.Enabled = False
        txt_piezasxkit.Text = ""
    End If
End Sub
