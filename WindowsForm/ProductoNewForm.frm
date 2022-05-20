VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ProductoNewForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresar nuevo producto"
   ClientHeight    =   8310
   ClientLeft      =   1860
   ClientTop       =   2190
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
   Icon            =   "ProductoNewForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ProductoNewForm.frx":08CA
   ScaleHeight     =   8310
   ScaleWidth      =   18000
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Cbo_Ubicacion 
      Height          =   390
      Left            =   12840
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2700
      Width           =   3855
   End
   Begin VB.CommandButton cmd_ElimarImagen 
      Caption         =   "Eliminar Imagen"
      Height          =   495
      Left            =   14880
      TabIndex        =   30
      Top             =   3520
      Width           =   1835
   End
   Begin VB.CommandButton cmd_AgregarImagen 
      Caption         =   "Añadir Imagen"
      Height          =   495
      Left            =   12840
      TabIndex        =   28
      Top             =   3520
      Width           =   1835
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   16800
      Top             =   3520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Format          =   117112833
      CurrentDate     =   42645
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
   Begin VB.TextBox txt_cantminventa 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   12840
      MaxLength       =   5
      TabIndex        =   10
      Top             =   6050
      Width           =   3855
   End
   Begin VB.TextBox txt_piezasxkit 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   6
      Top             =   6050
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
   Begin VB.TextBox txt_cantmaxventa 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   12840
      MaxLength       =   5
      TabIndex        =   9
      Top             =   5200
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
   Begin VB.TextBox txt_cantproducto 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      MaxLength       =   5
      TabIndex        =   2
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
   Begin VB.TextBox txt_CodProducto 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1920
      MaxLength       =   18
      TabIndex        =   0
      Top             =   2700
      Width           =   3855
   End
   Begin VB.CommandButton cmd_atras 
      Appearance      =   0  'Flat
      Caption         =   "Atras"
      Height          =   615
      Left            =   6360
      TabIndex        =   11
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton cmd_guardar 
      Caption         =   "Registrar"
      Height          =   615
      Left            =   9360
      TabIndex        =   12
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ayúdame a elegir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   15360
      MouseIcon       =   "ProductoNewForm.frx":13B45
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   3180
      Width           =   1275
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
      TabIndex        =   27
      Top             =   5700
      Width           =   2040
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
      TabIndex        =   26
      Top             =   6555
      Width           =   3195
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
      TabIndex        =   25
      Top             =   6550
      Width           =   1020
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de llegada del producto"
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
      TabIndex        =   24
      Top             =   5700
      Width           =   2235
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
      TabIndex        =   23
      Top             =   6550
      Width           =   2535
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
      TabIndex        =   22
      Top             =   5700
      Width           =   3240
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
      TabIndex        =   21
      Top             =   4440
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rellene los datos Secundarios"
      Height          =   270
      Left            =   1680
      TabIndex        =   20
      Top             =   4590
      Width           =   3165
   End
   Begin VB.Label Label6 
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
      TabIndex        =   19
      Top             =   3180
      Width           =   1710
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
      TabIndex        =   18
      Top             =   4040
      Width           =   1425
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
      TabIndex        =   17
      Top             =   3180
      Width           =   1620
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
      TabIndex        =   16
      Top             =   4035
      Width           =   1860
   End
   Begin VB.Label lbl_descripcion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rellene los datos Principales"
      Height          =   270
      Left            =   1680
      TabIndex        =   15
      Top             =   2040
      Width           =   3015
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
      TabIndex        =   14
      Top             =   1920
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
      TabIndex        =   13
      Top             =   3180
      Width           =   1485
   End
End
Attribute VB_Name = "ProductoNewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_AgregarImagen_Click()
Dialogo.DialogTitle = "Seleccione una imagen (Solo JPG o GIF)"
Dialogo.Filter = "Archivos Jpg|*.jpg|Archivos Gif|*.gif|"
Dialogo.ShowSave
RutaOrigen = Dialogo.FileName
ArchivoNombre = Dialogo.FileTitle

End Sub

Private Sub cmd_Atras_Click()
LimpiarProductos
ArchivoNombre = ""
Unload Me
Set ConsultasForm.GrillaConsultas.DataSource = RsProductos
ConsultasForm.EstilosGrillaConsultas
End Sub

Private Sub lbl_EJpass_Click()

End Sub

Private Sub cmd_ElimarImagen_Click()
RutaOrigen = (App.Path & "\ImagenesProductos\predeterminada.jpg") 'el app.path esta configurado (mas bien ya usado en la ruta donde esta la db
'ImageProducto.Picture = LoadPicture(RutaOrigen) aqui no muestro imagen
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
          If Not IsNumeric(txt_cantmaxventa.Text) Then MsgBox "Por favor, en el campo CANTIDAD MÁXIMA PARA LA VENTA debe ir sólo números", vbInformation, "Aviso": txt_cantmaxventa.SetFocus: Exit Sub
          If Val(txt_cantmaxventa.Text) <= 0 Then MsgBox ("Error. Cantidad máxima para la venta inválida."), vbInformation, "Aviso": txt_cantmaxventa.SetFocus: Exit Sub
If txt_cantminventa.Text = "" Then MsgBox "Por favor, especifique cual es la CANTIDAD MÍNIMA PARA LA VENTA", vbInformation, "Aviso": txt_cantminventa.SetFocus: Exit Sub
          If Not IsNumeric(txt_cantminventa.Text) Then MsgBox "Por favor, en el campo CANTIDAD MÍNIMA PARA LA VENTA debe ir sólo números", vbInformation, "Aviso": txt_cantminventa.SetFocus: Exit Sub
          If Val(txt_cantminventa.Text) <= 0 Then MsgBox ("Error. Cantidad mínima para la venta inválida"), vbInformation, "Aviso": txt_cantminventa.SetFocus: Exit Sub
If Val(txt_cantminventa.Text) > Val(txt_cantmaxventa.Text) Then MsgBox ("La cantidad mínima no puede ser mayor a la cantidad máxima para la venta"), vbInformation, "Aviso": txt_cantminventa.SetFocus: Exit Sub


'validaciones para la ubicacion

'#PESADOS
If Cbo_Ubicacion.Text = "Segmento 1 (S1)" Or Cbo_Ubicacion.Text = "Segmento 2 (S2)" Or Cbo_Ubicacion.Text = "Segmento 3 (S3)" Or Cbo_Ubicacion.Text = "Segmento 4 (S4)" Then
          If UCase(Trim(txt_DescripcionProducto)) Like "*BUJE*" = True Then
                    'no hago nada paso al siguiente
          ElseIf UCase(Trim(txt_DescripcionProducto)) Like "*BUJES*" = True Then
                    'no hago nada paso al siguiente
          ElseIf UCase(Trim(txt_DescripcionProducto)) Like "*BASE*" = True Then
                    'no hago nada paso al siguiente
          ElseIf UCase(Trim(txt_DescripcionProducto)) Like "*BASES*" = True Then
                    'no hago nada paso al siguiente
          Else
                    MsgBox ("No pueder ubicar este producto en dicho segmento 1"), vbInformation, "Aviso": Exit Sub
          End If
'#ELECTRICOS
ElseIf Cbo_Ubicacion.Text = "Segmento 5 (S5)" Or Cbo_Ubicacion.Text = "Segmento 6 (S6)" Or Cbo_Ubicacion.Text = "Segmento 8 (S8)" Or Cbo_Ubicacion.Text = "Segmento 9 (S9)" Then
          If UCase(Trim(txt_DescripcionProducto)) Like "*LUCES*" = True Then
                    'no hago nada paso al siguiente
          ElseIf UCase(Trim(txt_DescripcionProducto)) Like "*LED*" = True Then
                    'no hago nada paso al siguiente
          ElseIf UCase(Trim(txt_DescripcionProducto)) Like "*FUSIBLE*" = True Then
                    'no hago nada paso al siguiente
          ElseIf UCase(Trim(txt_DescripcionProducto)) Like "*BUJIAS*" = True Then
                    'no hago nada paso al siguiente
          ElseIf UCase(Trim(txt_DescripcionProducto)) Like "*BOBINA*" = True Then
                    'no hago nada paso al siguiente
          Else
                    MsgBox ("No pueder ubicar este producto en dicho segmento 2"), vbInformation, "Aviso": Exit Sub
          End If
'#PARTES DE MOTOR
ElseIf Cbo_Ubicacion.Text = "Segmento 11 (S11)" Then
          If UCase(Trim(txt_DescripcionProducto)) Like "*PISTON*" = True Then
                    'no hago nada paso al siguiente
          ElseIf UCase(Trim(txt_DescripcionProducto)) Like "*MASTER KITS*" = True Then
                    'no hago nada paso al siguiente
          ElseIf UCase(Trim(txt_DescripcionProducto)) Like "*MASTER KIT*" = True Then
                    'no hago nada paso al siguiente
          Else
                    MsgBox ("No pueder ubicar este producto en dicho segmento 3"), vbInformation, "Aviso": Exit Sub
          End If
'#BOMBAS
ElseIf Cbo_Ubicacion.Text = "Segmento 7 (S7)" Then
          If UCase(Trim(txt_DescripcionProducto)) Like "*BOMBA*" = False Then
                    MsgBox ("No pueder ubicar este producto en dicho segmento 11"), vbInformation, "Aviso": Exit Sub
          End If
'#EMPACADURAS
ElseIf Cbo_Ubicacion.Text = "Segmento 16 (S16)" Or Cbo_Ubicacion.Text = "Segmento 17 (S17)" Or Cbo_Ubicacion.Text = "Segmento 18 (S18)" Or Cbo_Ubicacion.Text = "Segmento 19 (S19)" Then
          If UCase(Trim(txt_DescripcionProducto)) Like "*EMPACADURA*" = False Then
                    MsgBox ("No pueder ubicar este producto en dicho segmento 12"), vbInformation, "Aviso": Exit Sub
          End If
'#FILTROS
ElseIf Cbo_Ubicacion.Text = "Segmento 12 (S12)" Or Cbo_Ubicacion.Text = "Segmento 13 (S13)" Or Cbo_Ubicacion.Text = "Segmento 14 (S14)" Then
          If UCase(Trim(txt_DescripcionProducto)) Like "*FILTRO*" = False Then
                    MsgBox ("No pueder ubicar este producto en dicho segmento 13"), vbInformation, "Aviso": Exit Sub
          End If
'#CABLES
ElseIf Cbo_Ubicacion.Text = "Segmento 15 (S15)" Then
          If UCase(Trim(txt_DescripcionProducto)) Like "*CABLE*" = False Then
                    MsgBox ("No pueder ubicar este producto en dicho segmento 14"), vbInformation, "Aviso": Exit Sub
          End If
End If

With RsProductos
        .Requery 'actualizamos la tabla
        .Find "CodigoProducto='" & Trim(txt_CodProducto.Text) & "'"
    
        If .EOF Then
            .AddNew 'creamos un nuevo registro
                !CodigoProducto = txt_CodProducto.Text
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
              '  !Proveedor = txt_proveedorproducto.Text
                !Ubicacion = Cbo_Ubicacion.Text
                !Marca = txt_marcaproducto.Text
        
                        If ArchivoNombre = "" Then
                                !Imagen = "\ImagenesProductos\predeterminada.jpg"
                        Else
                                !Imagen = "\ImagenesProductos\" & ArchivoNombre
                        End If
            .Update
            .Requery
            Logs
            With RsLogs
                    .Requery
                    .AddNew
                              !User = vUsername
                              !Accion = "añadió un producto asociado al código [" & txt_CodProducto.Text & "] con la cantidad de: " & txt_cantproducto.Text
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
                              !Accion = "añadió un producto asociado al código [" & txt_CodProducto.Text & "] con la cantidad de: " & txt_cantproducto.Text
                              !Fecha = Date + Time
                    .Update
                    .Requery
                    .Close
            End With
            LimpiarProductos
            MsgBox "Producto creado con éxito", vbInformation, "Aviso"
        Else
            'si encuentra el mismo codigo de producto
            MsgBox "Ya existe un producto con el mismo código, por favor cambielo", vbInformation, "Aviso"
            txt_CodProducto.Text = ""
            txt_CodProducto.SetFocus
            Exit Sub
    End If
    
End With
If ArchivoNombre = "" Then
Else
    RutaDestino = App.Path & "\ImagenesProductos\"
    Set GDB = Nothing
    Set fs = CreateObject("scripting.filesystemobject")
    fs.copyfile RutaOrigen, RutaDestino
    End If
    
ConsultasForm.GrillaConsultas.Refresh 'actualizar el apartado de consultas cada vez que introduzco datos
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

'actualizamos la consulta de total de despachos
IndexForm.AdoTotalDespachos.Refresh
IndexForm.lbl_TotalDespachos.Caption = IndexForm.GrillaTotalDespachos.Columns(0).Text

'actualizamos la consulta de total de devoluciones
IndexForm.AdoTotalDevoluciones.Refresh
IndexForm.lbl_TotalDevoluciones.Caption = IndexForm.GrillaTotalDevoluciones.Columns(0).Text

'actualizamos el total de operaciones
IndexForm.TotalOperaciones


ArchivoNombre = ""
Unload Me
End Sub

Sub LimpiarProductos()

txt_CodProducto.Text = ""
txt_DescripcionProducto = ""
txt_cantproducto.Text = ""
txt_marcaproducto = ""
txt_proveedorproducto = ""
opt_kityes.Value = False
opt_kitno.Value = False
txt_by.Text = ""
txt_cantmaxventa.Text = ""
txt_cantminventa.Text = ""
txt_piezasxkit.Text = ""




End Sub

Private Sub Form_Load()
Productos
'cargamos imagenes
RutaOrigen = App.Path & "\ImagenesProductos\predeterminada.jpg" 'el app.path esta configurado (mas bien ya usado en la ruta donde esta la db

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

End Sub

Private Sub Label15_Click()
ProductosUbicacionHelpForm.Show
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
       txt_piezasxkit.Text = ""
    Else
        txt_piezasxkit.Enabled = False
        txt_piezasxkit.Text = ""
    End If
End Sub

Private Sub txt_cantproducto_Change()
If Trim(txt_cantproducto.Text) = "" Then
     Exit Sub
Else
     txt_cantmaxventa.Text = Trim(txt_cantproducto.Text)
     txt_cantminventa.Text = "1"
End If
End Sub
