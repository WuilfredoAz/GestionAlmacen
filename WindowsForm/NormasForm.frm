VERSION 5.00
Begin VB.Form ProductosNormasForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Normas de la clasificación"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NormasForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "NormasForm.frx":058A
   ScaleHeight     =   8970
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_cerrar 
      Caption         =   "Cerrar"
      Height          =   615
      Left            =   5280
      TabIndex        =   8
      Top             =   8160
      Width           =   1815
   End
   Begin VB.TextBox txt_Norma5 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "NormasForm.frx":8AC9
      Top             =   6840
      Width           =   10215
   End
   Begin VB.TextBox txt_Norma4 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "NormasForm.frx":8B9C
      Top             =   6000
      Width           =   10215
   End
   Begin VB.TextBox txt_Norma3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "NormasForm.frx":8C6F
      Top             =   5160
      Width           =   10215
   End
   Begin VB.TextBox txt_Norma2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "NormasForm.frx":8D42
      Top             =   4320
      Width           =   10215
   End
   Begin VB.TextBox txt_Norma1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "NormasForm.frx":8E15
      Top             =   3840
      Width           =   10215
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "NormasForm.frx":8E85
      Top             =   3240
      Width           =   10695
   End
   Begin VB.TextBox txt_Descripcion 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "NormasForm.frx":8EFE
      Top             =   2280
      Width           =   10215
   End
   Begin VB.Label lbl_Clasificacion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ARTÍCULOS PESADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1080
      TabIndex        =   0
      Top             =   1800
      Width           =   3930
   End
End
Attribute VB_Name = "ProductosNormasForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cerrar_Click()
Unload Me
End Sub
