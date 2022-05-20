VERSION 5.00
Begin VB.Form DespachosShippingDetalles 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalles de la etiqueta"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   375
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
   Icon            =   "DespachosShippingDetalles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DespachosShippingDetalles.frx":058A
   ScaleHeight     =   10095
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   2640
      TabIndex        =   17
      Top             =   9360
      Width           =   1695
   End
   Begin VB.CommandButton cmd_generar 
      Caption         =   "Generar"
      Height          =   615
      Left            =   5040
      TabIndex        =   16
      Top             =   9360
      Width           =   1695
   End
   Begin VB.TextBox txt_Bultos 
      Height          =   495
      Left            =   3960
      MaxLength       =   5
      TabIndex        =   14
      Top             =   8160
      Width           =   4455
   End
   Begin VB.TextBox txt_Direccion 
      Height          =   1335
      Left            =   3960
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   6480
      Width           =   4455
   End
   Begin VB.TextBox txt_Fecha 
      Height          =   495
      Left            =   3960
      MaxLength       =   30
      TabIndex        =   10
      Top             =   4920
      Width           =   4455
   End
   Begin VB.TextBox txt_Zona 
      Height          =   495
      Left            =   3960
      MaxLength       =   30
      TabIndex        =   8
      Top             =   4080
      Width           =   4455
   End
   Begin VB.TextBox txt_Codigo 
      Height          =   495
      Left            =   3960
      MaxLength       =   30
      TabIndex        =   6
      Top             =   3240
      Width           =   4455
   End
   Begin VB.TextBox txt_Cliente 
      Height          =   495
      Left            =   3960
      MaxLength       =   30
      TabIndex        =   4
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad de cajas contenedoras de productos"
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
      Left            =   3960
      TabIndex        =   19
      Top             =   8640
      Width           =   3390
   End
   Begin VB.Label lbl_EJnombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Min 66 caracteres"
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
      Left            =   3960
      TabIndex        =   18
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bultos:"
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
      Left            =   2760
      TabIndex        =   15
      Top             =   8400
      Width           =   840
   End
   Begin VB.Label lbl_direccion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección de envío:"
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
      Left            =   1440
      TabIndex        =   13
      Top             =   6720
      Width           =   2220
   End
   Begin VB.Label lbl_Fecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   5160
      Width           =   795
   End
   Begin VB.Label lbl_Zona 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zona:"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   4320
      Width           =   660
   End
   Begin VB.Label lbl_codigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código de despacho:"
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
      Left            =   1200
      TabIndex        =   7
      Top             =   3480
      Width           =   2475
   End
   Begin VB.Label lbl_Cliente 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
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
      Left            =   2760
      TabIndex        =   5
      Top             =   2640
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección del envío"
      Height          =   270
      Left            =   1440
      TabIndex        =   3
      Top             =   5970
      Width           =   1980
   End
   Begin VB.Label Label1 
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
      Left            =   1080
      TabIndex        =   2
      Top             =   5800
      Width           =   225
   End
   Begin VB.Label lbl_1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos del despacho"
      Height          =   270
      Left            =   1440
      TabIndex        =   1
      Top             =   1845
      Width           =   2100
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
      Left            =   1080
      TabIndex        =   0
      Top             =   1680
      Width           =   225
   End
End
Attribute VB_Name = "DespachosShippingDetalles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancelar_Click()
Unload Me
End Sub

Private Sub cmd_generar_Click()
If txt_Direccion.Text = "" Then MsgBox ("Por favor, ingrese una dirección para el envío"), vbInformation, "Aviso": Exit Sub
          If IsNumeric(txt_Direccion.Text) Then MsgBox ("Dirección inválida"), vbInformation, "Aviso": Exit Sub
          If Len(txt_Direccion.Text) < 66 Then MsgBox ("Dirección inválida"), vbInformation, "Aviso": Exit Sub
If Not IsNumeric(txt_Bultos.Text) Then MsgBox ("Sólo ingrese números en la cantidad de bultos"), vbInformation, "Aviso": Exit Sub

vClienteShipping = Trim(txt_Cliente.Text)
vZonaShipping = Trim(txt_Zona.Text)
vFechaShipping = Trim(txt_Fecha.Text)
vDireccionShipping = Trim(txt_Direccion.Text)
vBultoShipping = Trim(txt_Bultos.Text)
          
For x = 1 To vBultoShipping
          With RsShipping
                    .Requery
                    .AddNew
                              Dim Cantidad As Integer
                              Cantidad = Val(Cantidad) + 1
                              !Bulto = Cantidad
                    .Update
                    .Requery
          End With
Next
          
Unload Me
dr_Shipping.WindowState = 2

' s e c c i o n 1
dr_Shipping.Sections("Sección1").Controls("Etiqueta7").Caption = vZonaShipping
dr_Shipping.Sections("Sección1").Controls("Etiqueta3").Caption = vDireccionShipping
dr_Shipping.Sections("Sección1").Controls("Etiqueta6").Caption = vCodigoShipping
dr_Shipping.Sections("Sección1").Controls("Etiqueta9").Caption = vClienteShipping
dr_Shipping.Sections("Sección1").Controls("Etiqueta11").Caption = vFechaShipping
dr_Shipping.Sections("Sección1").Controls("Etiqueta13").Caption = vBultoShipping

Set dr_Shipping.DataSource = RsShipping

'mostramos
dr_Shipping.Show


End Sub

Private Sub Form_Load()
CargarDatos
txt_Cliente.Enabled = False
txt_Codigo.Enabled = False
txt_Zona.Enabled = False
txt_Fecha.Enabled = False
Shipping
End Sub

Sub CargarDatos()
With RsDespacho
          .Requery
          .Find "CodigoDespacho='" & Trim(vCodigoShipping) & "'"
                    txt_Cliente.Text = !Cliente
                    txt_Codigo.Text = !CodigoDespacho
                    txt_Zona.Text = !Zona
                    txt_Fecha.Text = !Fecha
End With
End Sub

