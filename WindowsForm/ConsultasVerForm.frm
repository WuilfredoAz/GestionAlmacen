VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ConsultasVerForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalles del Producto"
   ClientHeight    =   8925
   ClientLeft      =   6330
   ClientTop       =   2070
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ConsultasVerForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ConsultasVerForm.frx":058A
   ScaleHeight     =   8925
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoEtiquetas 
      Height          =   375
      Left            =   5760
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.TextBox txt_Descripcion 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   21
      Text            =   "ConsultasVerForm.frx":67A9
      Top             =   2280
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imagen Referencial"
      Height          =   3255
      Left            =   480
      TabIndex        =   20
      Top             =   1920
      Width           =   2775
      Begin VB.Image ImageProducto 
         Height          =   3255
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmd_atras 
      Caption         =   "Atras"
      Height          =   615
      Left            =   2400
      TabIndex        =   19
      Top             =   7920
      Width           =   2175
   End
   Begin VB.CommandButton cmd_CrearEtiqueta 
      Caption         =   "Imprimir etiqueta"
      Height          =   615
      Left            =   4800
      TabIndex        =   22
      Top             =   7920
      Width           =   2175
   End
   Begin VB.Label lbl_CantMinCliente 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XXXXXXXXXXX"
      Height          =   270
      Left            =   4920
      TabIndex        =   18
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label lbl_CantMaxCliente 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XXXXXXXXXXX"
      Height          =   270
      Left            =   4920
      TabIndex        =   17
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label lbl_By 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Height          =   270
      Left            =   3840
      TabIndex        =   16
      Top             =   4440
      Width           =   5280
   End
   Begin VB.Label lbl_Fecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XXXXXXXXXXXX"
      Height          =   270
      Left            =   2640
      TabIndex        =   15
      Top             =   5760
      Width           =   1980
   End
   Begin VB.Label lbl_PiezasXKit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XX"
      Height          =   270
      Left            =   6600
      TabIndex        =   14
      Top             =   5280
      Width           =   330
   End
   Begin VB.Label lbl_Kit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XX"
      Height          =   270
      Left            =   6960
      TabIndex        =   13
      Top             =   4800
      Width           =   330
   End
   Begin VB.Label lbl_Ubicacion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XXXXXXXXXXXXXXXX"
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   6240
      Width           =   2640
   End
   Begin VB.Label lbl_Marca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XXXXXXXXXXXXXXXXXX"
      Height          =   270
      Left            =   6240
      TabIndex        =   11
      Top             =   3600
      Width           =   2970
   End
   Begin VB.Label lbl_Stock 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XXXX"
      Height          =   270
      Left            =   6000
      TabIndex        =   10
      Top             =   3120
      Width           =   660
   End
   Begin VB.Label lbl_Codigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COD-XXXXXXXXXXXXXX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   675
      Left            =   3600
      TabIndex        =   9
      Top             =   1680
      Width           =   5625
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad Mínima de venta por cliente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   7200
      Width           =   4290
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad Máxima de venta por cliente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   6720
      Width           =   4350
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aplicable para:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   6
      Top             =   4080
      Width           =   1710
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Registro:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Piezas que trae el KIT:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Top             =   5280
      Width           =   2550
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto vendido por KIT:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Top             =   4800
      Width           =   3045
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   6240
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca del Producto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Top             =   3600
      Width           =   2310
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Disponible:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   0
      Top             =   3120
      Width           =   2040
   End
End
Attribute VB_Name = "ConsultasVerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Atras_Click()
Unload Me
vMostrarProducto = 0
ConsultasForm.EstilosGrillaConsultas
End Sub

Private Sub cmd_CrearEtiqueta_Click()
Logs
With RsLogs
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir etiqueta del producto asociado al codigo [" & lbl_codigo.Caption & "]"
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With

LogsReportes
With RsLogsReportes
          .Requery
          .AddNew
                    !User = vUsername
                    !Accion = "hizo click en el botón de imprimir etiqueta del producto asociado al codigo [" & lbl_codigo.Caption & "]"
                    !Fecha = Date + Time
          .Update
          .Requery
          .Close
End With

If lbl_Ubicacion.Caption = "Segmento 1 (S1)" Then TagSeccion = "S1": TagSeccion1 = "S1": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Pesados1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Pesados].jpg")
If lbl_Ubicacion.Caption = "Segmento 2 (S2)" Then TagSeccion = "S2": TagSeccion1 = "S2": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Pesados1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Pesados].jpg")
If lbl_Ubicacion.Caption = "Segmento 3 (S3)" Then TagSeccion = "S3": TagSeccion1 = "S3": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Pesados1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Pesados].jpg")
If lbl_Ubicacion.Caption = "Segmento 4 (S4)" Then TagSeccion = "S4": TagSeccion1 = "S4": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Pesados1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Pesados].jpg")
If lbl_Ubicacion.Caption = "Segmento 5 (S5)" Then TagSeccion = "S5": TagSeccion1 = "S5": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Electricos1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Electricos].jpg")
If lbl_Ubicacion.Caption = "Segmento 6 (S6)" Then TagSeccion = "S6": TagSeccion1 = "S6": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Electricos1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Electricos].jpg")
If lbl_Ubicacion.Caption = "Segmento 7 (S7)" Then TagSeccion = "S7": TagSeccion1 = "S7": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[B.Agua1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[B.Agua].jpg")
If lbl_Ubicacion.Caption = "Segmento 8 (S8)" Then TagSeccion = "S8": TagSeccion1 = "S8": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Electricos1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Electricos].jpg")
If lbl_Ubicacion.Caption = "Segmento 9 (S9)" Then TagSeccion = "S9": TagSeccion1 = "S9": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Electricos1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Electricos].jpg")
If lbl_Ubicacion.Caption = "Segmento 10 (S10)" Then TagSeccion = "S10": TagSeccion1 = "S10": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[+Vendidos1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[+Vendidos].jpg")
If lbl_Ubicacion.Caption = "Segmento 11 (S11)" Then TagSeccion = "S11": TagSeccion1 = "S11": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[P.Motor1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[P.Motor].jpg")
If lbl_Ubicacion.Caption = "Segmento 12 (S12)" Then TagSeccion = "S12": TagSeccion1 = "S12": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Filtros1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Filtros].jpg")
If lbl_Ubicacion.Caption = "Segmento 13 (S13)" Then TagSeccion = "S13": TagSeccion1 = "S13": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Filtros1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Filtros].jpg")
If lbl_Ubicacion.Caption = "Segmento 14 (S14)" Then TagSeccion = "S14": TagSeccion1 = "S14": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Filtros1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Filtros].jpg")
If lbl_Ubicacion.Caption = "Segmento 15 (S15)" Then TagSeccion = "S15": TagSeccion1 = "S15": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Cables1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Cables].jpg")
If lbl_Ubicacion.Caption = "Segmento 16 (S16)" Then TagSeccion = "S16": TagSeccion1 = "S16": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Empacaduras1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Empacaduras].jpg")
If lbl_Ubicacion.Caption = "Segmento 17 (S17)" Then TagSeccion = "S17": TagSeccion1 = "S17": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Empacaduras1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Empacaduras].jpg")
If lbl_Ubicacion.Caption = "Segmento 18 (S18)" Then TagSeccion = "S18": TagSeccion1 = "S18": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Empacaduras1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Empacaduras].jpg")
If lbl_Ubicacion.Caption = "Segmento 19 (S19)" Then TagSeccion = "S19": TagSeccion1 = "S19": Set dr_Etiquetas.Sections("Sección2").Controls("Image1").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Empacaduras1].jpg"): Set dr_Etiquetas.Sections("Sección2").Controls("Image2").Picture = LoadPicture(App.Path & "\Interfaz\MaterialPrinter[Empacaduras].jpg")

TagCodigo = lbl_codigo.Caption
     TagCodigo1 = lbl_codigo.Caption
TagDescripcion = txt_Descripcion.Text
     TagDescripcion1 = txt_Descripcion.Text
TagAplicable = lbl_By.Caption
TagMin = lbl_CantMinCliente.Caption
TagMax = lbl_CantMaxCliente.Caption
TagKit = lbl_Kit.Caption
TagPzas = lbl_PiezasxKit.Caption

dr_Etiquetas.Sections("Sección2").Controls("Etiqueta1").Caption = TagCodigo1
dr_Etiquetas.Sections("Sección2").Controls("Etiqueta2").Caption = TagDescripcion1
dr_Etiquetas.Sections("Sección2").Controls("Etiqueta3").Caption = TagSeccion1

dr_Etiquetas.Sections("Sección2").Controls("Etiqueta6").Caption = TagSeccion
dr_Etiquetas.Sections("Sección2").Controls("Etiqueta10").Caption = TagMin
dr_Etiquetas.Sections("Sección2").Controls("Etiqueta11").Caption = TagMax
dr_Etiquetas.Sections("Sección2").Controls("Etiqueta7").Caption = TagCodigo
dr_Etiquetas.Sections("Sección2").Controls("Etiqueta8").Caption = TagDescripcion
dr_Etiquetas.Sections("Sección2").Controls("Etiqueta9").Caption = TagAplicable
dr_Etiquetas.Sections("Sección2").Controls("Etiqueta17").Caption = TagKit
dr_Etiquetas.Sections("Sección2").Controls("Etiqueta14").Caption = TagPzas

With AdoEtiquetas
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
          Busca = Trim(lbl_codigo.Caption)
          .RecordSource = "SELECT * FROM Productos WHERE CodigoProducto LIKE '" & Busca & "'"
          .Refresh
        Set dr_Etiquetas.DataSource = AdoEtiquetas
        dr_Etiquetas.Width = 10000
        dr_Etiquetas.Height = 8000
        dr_Etiquetas.Show vbModal
End With


End Sub

Private Sub Form_Load()
VerProducto
ConsultasForm.EstilosGrillaConsultas
End Sub

Sub VerProducto()
With RsProductos
    .Requery
    .Find "Id='" & Val(vMostrarProducto) & "'"

'igualamos los campos
    lbl_codigo.Caption = !CodigoProducto
 '   lbl_Descripcion.Caption = !descripcion
    txt_Descripcion.Text = !descripcion
    lbl_Stock.Caption = !Cantidad
    lbl_Marca.Caption = !Marca
    lbl_Ubicacion.Caption = !Ubicacion
    lbl_PiezasxKit.Caption = !PiezaxKit
    lbl_Fecha.Caption = !Fecha
    lbl_By.Caption = !By
    lbl_CantMaxCliente.Caption = !CantMAXVenta
    lbl_CantMinCliente.Caption = !CantMINVenta
  

        If !Kit = 0 Then
            lbl_Kit.Caption = "No"
        Else
            lbl_Kit.Caption = "Sí"
        End If
        
RutaOrigen = App.Path & ConsultasForm.GrillaConsultas.Columns(12).Text
ImageProducto.Picture = LoadPicture(RutaOrigen)

End With
End Sub

