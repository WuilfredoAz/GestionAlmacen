VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ConfirmarDatosForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restablecer contraseña 1/2"
   ClientHeight    =   9465
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
   Icon            =   "ConfirmarDatosForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ConfirmarDatosForm.frx":08CA
   ScaleHeight     =   9465
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid GrillaVerificar 
      Height          =   975
      Left            =   960
      TabIndex        =   18
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
   Begin MSAdodcLib.Adodc AdoBuscarUsuario 
      Height          =   375
      Left            =   960
      Top             =   6360
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
   Begin VB.CommandButton cmd_Verificar 
      Caption         =   "Verficar"
      Height          =   615
      Left            =   5400
      TabIndex        =   6
      Top             =   8160
      Width           =   2295
   End
   Begin VB.CommandButton cmd_Cancelar 
      Appearance      =   0  'Flat
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   1680
      TabIndex        =   7
      Top             =   8160
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "ConfirmarDatosForm.frx":6F42
      Top             =   2040
      Width           =   7455
   End
   Begin VB.TextBox txt_nombre 
      Height          =   495
      Left            =   3840
      MaxLength       =   30
      TabIndex        =   1
      Top             =   3240
      Width           =   4455
   End
   Begin VB.TextBox txt_apellido 
      Height          =   495
      Left            =   3840
      MaxLength       =   30
      TabIndex        =   2
      Top             =   4080
      Width           =   4455
   End
   Begin VB.TextBox txt_cedula 
      Height          =   495
      Left            =   3840
      MaxLength       =   8
      TabIndex        =   3
      Top             =   4920
      Width           =   4455
   End
   Begin VB.TextBox txt_username 
      Height          =   495
      Left            =   3840
      MaxLength       =   15
      TabIndex        =   4
      Top             =   5760
      Width           =   4455
   End
   Begin VB.TextBox txt_correo 
      Height          =   495
      Left            =   3840
      MaxLength       =   30
      TabIndex        =   5
      Top             =   6600
      Width           =   4455
   End
   Begin VB.Label lbl_nombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
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
      Left            =   2280
      TabIndex        =   17
      Top             =   3480
      Width           =   1005
   End
   Begin VB.Label lbl_EJnombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ejm: Pedro"
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
      Left            =   3840
      TabIndex        =   16
      Top             =   3720
      Width           =   810
   End
   Begin VB.Label lbl_apellido 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido:"
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
      Left            =   2280
      TabIndex        =   15
      Top             =   4320
      Width           =   1020
   End
   Begin VB.Label lbl_EJapellido 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ejm: Pérez"
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
      Left            =   3840
      TabIndex        =   14
      Top             =   4560
      Width           =   810
   End
   Begin VB.Label lbl_cedula 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.I:"
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
      TabIndex        =   13
      Top             =   5160
      Width           =   390
   End
   Begin VB.Label lbl_EJcedula 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ejm:12080090"
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
      Left            =   3840
      TabIndex        =   12
      Top             =   5400
      Width           =   1050
   End
   Begin VB.Label lbl_username 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de Usuario:"
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
      Left            =   960
      TabIndex        =   11
      Top             =   6000
      Width           =   2310
   End
   Begin VB.Label lbl_EJusername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ejm:PedroP"
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
      Left            =   3840
      TabIndex        =   10
      Top             =   6240
      Width           =   870
   End
   Begin VB.Label lbl_correo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Correo:"
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
      Left            =   2400
      TabIndex        =   9
      Top             =   6840
      Width           =   885
   End
   Begin VB.Label lbl_EJcorreo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ejm: ejemplo@correo.com"
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
      Left            =   3840
      TabIndex        =   8
      Top             =   7080
      Width           =   1965
   End
End
Attribute VB_Name = "ConfirmarDatosForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Cancelar_Click()
Unload Me
LoginForm.Show
End Sub

Private Sub cmd_Verificar_Click()

If txt_nombre.Text = "" Then MsgBox ("Por favor, ingrese el nombre que utilizó para registrarse"), vbInformation, "Aviso": txt_nombre.SetFocus: Exit Sub
          If IsNumeric(txt_nombre.Text) = True Then MsgBox ("Por favor, no escriba números en el campo nombre"), vbInformation, "Aviso": txt_nombre.SetFocus: Exit Sub
If txt_apellido.Text = "" Then MsgBox ("Por favor, ingrese el apellido que utilizó para registrarse"), vbInformation, "Aviso": txt_apellido.SetFocus: Exit Sub
          If IsNumeric(txt_apellido.Text) = True Then MsgBox ("Por favor, no escriba números en el campo apellido"), vbInformation, "Aviso": txt_apellido.SetFocus: Exit Sub
          If UCase(Trim(txt_nombre.Text)) = UCase(Trim(txt_apellido.Text)) Then MsgBox ("Su nombre no puede ser igual al apellido, por favor verifique"), vbInformation, "Aviso": txt_apellido.Text = "": Exit Sub
If txt_cedula.Text = "" Then MsgBox ("Por favor, ingrese la cédula que utilizó para registrarse"), vbInformation, "Aviso": txt_cedula.SetFocus: Exit Sub
          If Not IsNumeric(txt_cedula.Text) Then MsgBox ("Los datos que debe introducir en la cédula deben ser sólo números"), vbInformation, "Aviso": txt_cedula.SetFocus: Exit Sub
          If Val(txt_cedula.Text) < 1000000 Then MsgBox ("Cédula inválida, por favor verifique el número e intente de nuevo"), vbInformation, "Aviso": txt_cedula.SetFocus: Exit Sub
If txt_username.Text = "" Then MsgBox ("Por favor, ingrese el nombre de usuario con el que ingresaba al sistema"), vbInformation, "Aviso": txt_username.SetFocus: Exit Sub
If txt_correo.Text = "" Then MsgBox ("Por favor, ingrese el correo que utilizó para registrarse"), vbInformation, "Aviso": txt_correo.SetFocus: Exit Sub
          Dim ret As Boolean
          Dim Direccion As String
          Direccion = Trim(txt_correo.Text)
          If Direccion = vbNullString Then Exit Sub
          ' ejecuta la función
          ret = Comprobar_Mail(Direccion)
          ' Resultado
          If ret Then
                    ' Ok
         Else
                    MsgBox ("La dirección de correo ingresada es inválida, por favor cambiela"), vbInformation, "Aviso": Exit Sub ' El mail no es correcto
          End If

'verificamos si la informacion coincide

With AdoBuscarUsuario
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
          .RecordSource = "SELECT * FROM Usuarios WHERE Nombre LIKE '" & txt_nombre.Text & "' AND Apellido LIKE '" & txt_apellido.Text & "' AND CI LIKE '" & txt_cedula.Text & "' AND Nomb_Usuario LIKE '" & txt_username.Text & "' AND Correo LIKE '" & txt_correo.Text & "'"
          .Refresh
          Set GrillaVerificar.DataSource = AdoBuscarUsuario
End With

If GrillaVerificar.ApproxCount = 0 Then MsgBox ("No existe ningún usuario asociado a los datos que introdujo. Por favor tenga en cuenta que debe escribir los datos de la misma manera que cuando se registro. Verifique las mayúsculas y acentos."), vbInformation, "Aviso": Exit Sub

RestablecerPassForm.Show vbModal

End Sub

Private Sub Form_Load()
Usuarios
End Sub
