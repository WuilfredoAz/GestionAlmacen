VERSION 5.00
Begin VB.Form ProductosUbicacionHelpForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Distribución del almacén"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ProductosUbicacionHelpForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ProductosUbicacionHelpForm.frx":058A
   ScaleHeight     =   9465
   ScaleWidth      =   15000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "ProductosUbicacionHelpForm.frx":15C1E
      Top             =   8175
      Width           =   5655
   End
   Begin VB.TextBox Text7 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   27
      Text            =   "ProductosUbicacionHelpForm.frx":15C69
      Top             =   7455
      Width           =   5295
   End
   Begin VB.TextBox Text6 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "ProductosUbicacionHelpForm.frx":15CAF
      Top             =   6680
      Width           =   5295
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   21
      Text            =   "ProductosUbicacionHelpForm.frx":15CEA
      Top             =   5865
      Width           =   5295
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "ProductosUbicacionHelpForm.frx":15D25
      Top             =   5100
      Width           =   5295
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "ProductosUbicacionHelpForm.frx":15D60
      Top             =   4320
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "ProductosUbicacionHelpForm.frx":15D9A
      Top             =   3520
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "ProductosUbicacionHelpForm.frx":15DE1
      Top             =   2750
      Width           =   5295
   End
   Begin VB.CommandButton cmd_atras 
      Caption         =   "Atras"
      Height          =   615
      Left            =   6600
      TabIndex        =   0
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label lbl_S19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9370
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":15E28
      MousePointer    =   99  'Custom
      TabIndex        =   51
      Top             =   4200
      Width           =   540
   End
   Begin VB.Label lbl_S18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9370
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":16132
      MousePointer    =   99  'Custom
      TabIndex        =   50
      Top             =   3670
      Width           =   540
   End
   Begin VB.Label lbl_S17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   10300
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":1643C
      MousePointer    =   99  'Custom
      TabIndex        =   49
      Top             =   2920
      Width           =   900
   End
   Begin VB.Label lbl_S16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   9350
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":16746
      MousePointer    =   99  'Custom
      TabIndex        =   48
      Top             =   2920
      Width           =   900
   End
   Begin VB.Label lbl_S11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   8660
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":16A50
      MousePointer    =   99  'Custom
      TabIndex        =   47
      Top             =   3600
      Width           =   300
   End
   Begin VB.Label lbl_S12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7140
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":16D5A
      MousePointer    =   99  'Custom
      TabIndex        =   46
      Top             =   2660
      Width           =   300
   End
   Begin VB.Label lbl_S10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   10440
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":17064
      MousePointer    =   99  'Custom
      TabIndex        =   45
      Top             =   6770
      Width           =   300
   End
   Begin VB.Label lbl_S9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   10440
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":1736E
      MousePointer    =   99  'Custom
      TabIndex        =   44
      Top             =   5500
      Width           =   300
   End
   Begin VB.Label lbl_S8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   10440
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":17678
      MousePointer    =   99  'Custom
      TabIndex        =   43
      Top             =   4200
      Width           =   300
   End
   Begin VB.Label lbl_S7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   12120
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":17982
      MousePointer    =   99  'Custom
      TabIndex        =   42
      Top             =   6770
      Width           =   300
   End
   Begin VB.Label lbl_S6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   12120
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":17C8C
      MousePointer    =   99  'Custom
      TabIndex        =   41
      Top             =   5480
      Width           =   300
   End
   Begin VB.Label lbl_S5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   12120
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":17F96
      MousePointer    =   99  'Custom
      TabIndex        =   40
      Top             =   4200
      Width           =   300
   End
   Begin VB.Label lbl_S4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   13760
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":182A0
      MousePointer    =   99  'Custom
      TabIndex        =   39
      Top             =   6540
      Width           =   300
   End
   Begin VB.Label lbl_S3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   13760
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":185AA
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   5250
      Width           =   300
   End
   Begin VB.Label lbl_S2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   13740
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":188B4
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   3960
      Width           =   300
   End
   Begin VB.Label lbl_S14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7560
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":18BBE
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   3045
      Width           =   735
   End
   Begin VB.Label lbl_S13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7560
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":18EC8
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   2630
      Width           =   735
   End
   Begin VB.Label lbl_S15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8760
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":191D2
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   2620
      Width           =   3735
   End
   Begin VB.Label lbl_S1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   12720
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":194DC
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   3560
      Width           =   1340
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2730
      TabIndex        =   32
      Top             =   7875
      Width           =   45
   End
   Begin VB.Label lbl_NormasEmpacaduras 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NORMAS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2880
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":197E6
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   7875
      Width           =   720
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2010
      TabIndex        =   29
      Top             =   7080
      Width           =   45
   End
   Begin VB.Label lbl_NormasFiltros 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NORMAS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2160
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":19AF0
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   7080
      Width           =   720
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2010
      TabIndex        =   26
      Top             =   6300
      Width           =   45
   End
   Begin VB.Label lbl_NormasCables 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NORMAS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2160
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":19DFA
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   6300
      Width           =   720
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3930
      TabIndex        =   23
      Top             =   5520
      Width           =   45
   End
   Begin VB.Label lbl_NormasHot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NORMAS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4080
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":1A104
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   5520
      Width           =   720
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3090
      TabIndex        =   20
      Top             =   4755
      Width           =   45
   End
   Begin VB.Label lbl_NormasPMotor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NORMAS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3240
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":1A40E
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   4755
      Width           =   720
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2970
      TabIndex        =   17
      Top             =   3960
      Width           =   45
   End
   Begin VB.Label lbl_NormasBombas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NORMAS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":1A718
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   3960
      Width           =   720
   End
   Begin VB.Label lbl_NormasElectricos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NORMAS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":1AA22
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   3165
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3690
      TabIndex        =   13
      Top             =   3165
      Width           =   45
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3330
      TabIndex        =   11
      Top             =   2355
      Width           =   45
   End
   Begin VB.Label lbl_NormasPesados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NORMAS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3480
      MouseIcon       =   "ProductosUbicacionHelpForm.frx":1AD2C
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   2355
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FILTROS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   8
      Top             =   7080
      Width           =   840
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPACADURAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   7
      Top             =   7875
      Width           =   1500
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CABLES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   6
      Top             =   6300
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ARTÍCULOS MÁS VENDIDOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   5
      Top             =   5520
      Width           =   2685
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PARTES DE MOTOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   4
      Top             =   4755
      Width           =   1875
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BOMBAS DE AGUA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   3
      Top             =   3975
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ARTÍCULOS ELÉCTRICOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   2
      Top             =   3165
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ARTÍCULOS PESADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   1
      Top             =   2355
      Width           =   2130
   End
End
Attribute VB_Name = "ProductosUbicacionHelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Atras_Click()
Unload Me
End Sub

Private Sub lbl_NormasBombas_Click()
ProductosNormasForm.Picture = LoadPicture(App.Path & "\Interfaz\NormasFormBombas.jpg")
ProductosNormasForm.lbl_Clasificacion.Caption = "BOMBAS DE AGUA"
ProductosNormasForm.txt_Descripcion.Text = "Esta clasificación de segmentos del almacén, tiene la finalidad de almacenar productos cuyos rubros apilados sean Bombas de Agua, Bombas de Aceite, Bombas de Gasolina, Etc."
ProductosNormasForm.txt_Norma1.Text = "1.) Sólo se debe usar el segmento S7 para dichos productos"
ProductosNormasForm.txt_Norma2.Text = "2.) Se recomienda utilizar las cajas originales de los productos, en caso de que sean muchos apilarlas de forma vertical u horizontalmente según sea el caso."
ProductosNormasForm.txt_Norma3.Text = "3.) Utilizar las etiquetas tanto para los bordes de los estantes como las destinadas a las cajas (contenedores) de los productos."
ProductosNormasForm.txt_Norma4.Text = "4.) En caso de que alguna bomba contenga una cantidad significativa (+10) y no exista espacio disponible para tal cantidad, se recomienda dejar algunas en el estante (4) y las demás guardarlas en una caja en el otro almacén."
ProductosNormasForm.txt_Norma5.Text = "5.) No mezclar productos de esta clasificación con otro tipo de productos que sean Electricos"

ProductosNormasForm.Show vbModal
End Sub

Private Sub lbl_NormasCables_Click()
ProductosNormasForm.Picture = LoadPicture(App.Path & "\Interfaz\NormasFormCables.jpg")
ProductosNormasForm.lbl_Clasificacion.Caption = "CABLES"
ProductosNormasForm.txt_Descripcion.Text = "Esta clasificación de segmentos del almacén, tiene la finalidad de almacenar productos cuyos rubros apilados contengan cables automotrices. Por ejemplo: Cables de Bujías."
ProductosNormasForm.txt_Norma1.Text = "1.) Sólo se debe usar el segmento S15."
ProductosNormasForm.txt_Norma2.Text = "2.) Almacenar los productos por docenas, y hacerle una abertura por la parte de al frente para sacar detallados (si se requiere)."
ProductosNormasForm.txt_Norma3.Text = "3.) Utilizar las etiquetas tanto para los bordes de los estantes como las destinadas a las cajas (contenedores) de los productos."
ProductosNormasForm.txt_Norma4.Text = "4.) No apilar más de 4 columnas de docenas de cables a la vez. Si existe muchos cables del mismo tipo, almacenar sólo 4 columnas en este segmento y lo demás enviar a otro almacén."
ProductosNormasForm.txt_Norma5.Text = "5.) No mezclar productos de esta clasificación con otro tipo de productos que manejen o contenga liquídos. (Aceites, Lubricantes, Liga de Frenos, Etc.)"

ProductosNormasForm.Show vbModal
End Sub

Private Sub lbl_NormasElectricos_Click()
ProductosNormasForm.Picture = LoadPicture(App.Path & "\Interfaz\NormasFormElectricos.jpg")
ProductosNormasForm.lbl_Clasificacion.Caption = "ARTÍCULOS ELÉCTRICOS"
ProductosNormasForm.txt_Descripcion.Text = "Esta clasificación de segmentos del almacén, tiene la finalidad de almacenar productos cuyos rubros apilados contengan partes/piezas eléctricas de automóvil. Por ejemplo: Fusibles, Bobinas, Bujías, Etc."
ProductosNormasForm.txt_Norma1.Text = "1.) Sólo se deben usar los segmentos S5, S6, S8 y S9."
ProductosNormasForm.txt_Norma2.Text = "2.) Al utilizar cajas para resguardar los productos se debe prestar especial atención de que se encuentren bien firmes y selladas, de tal forma que no permita separación de los productos."
ProductosNormasForm.txt_Norma3.Text = "3.) Utilizar las etiquetas tanto para los bordes de los estantes como las destinadas a las cajas (contenedores) de los productos."
ProductosNormasForm.txt_Norma4.Text = "4.) Productos cuya cantidad sea sobresaliente al de los demás, preferiblemente ubicar en la parte de abajo de los estantes."
ProductosNormasForm.txt_Norma5.Text = "5.) No mezclar productos de esta clasificación con otro tipo de productos que sean Pesados, Líquidos, o cualquier tipo de Partes de motor. "

ProductosNormasForm.Show vbModal
End Sub

Private Sub lbl_NormasEmpacaduras_Click()
ProductosNormasForm.Picture = LoadPicture(App.Path & "\Interfaz\NormasFormEmpacaduras.jpg")
ProductosNormasForm.lbl_Clasificacion.Caption = "EMPACADURAS"
ProductosNormasForm.txt_Descripcion.Text = "Esta clasificación de segmentos del almacén, tiene la finalidad de almacenar productos cuyos rubros apilados contengan Empacaturas de Camara (EC), Empacaduras Tapa Valvulas (TV) y Juegos de Empacaduras."
ProductosNormasForm.txt_Norma1.Text = "1.) Sólo se deben usar los segmentos S16, S17, S18 y S19."
ProductosNormasForm.txt_Norma2.Text = "2.) NO MEZCLAR los diferentes tipos de empacaduras por ningún motivo. Empacaduras  EC, TV o juegos pdeben de estar separadas unas de otras aunque sean para el mismo vehículo."
ProductosNormasForm.txt_Norma3.Text = "3.) SELLAR las cajas/recipicientes/Envolotorios de las empacaduras una vez sacada la cantidad utilizada prestando especial atención en que no se saldrá ninguna de dicho recipiente."
ProductosNormasForm.txt_Norma4.Text = "4.) Identificar los tipos de empacaduras con las etiquetas correspondientes a su tipo (Prefiriblemente usar SOLO las de las cajas) o la más grande según sea el caso."
ProductosNormasForm.txt_Norma5.Text = "5.) EVITAR en lo posible colocar los juegos de empacaduras junto a artículos pesados."

ProductosNormasForm.Show vbModal
End Sub

Private Sub lbl_NormasFiltros_Click()
ProductosNormasForm.Picture = LoadPicture(App.Path & "\Interfaz\NormasFormFiltros.jpg")
ProductosNormasForm.lbl_Clasificacion.Caption = "FILTROS"
ProductosNormasForm.txt_Descripcion.Text = "Esta clasificación de segmentos del almacén, tiene la finalidad de almacenar productos cuyos rubros apilados contengan Filtros de Aire, Aceite, Cabina."
ProductosNormasForm.txt_Norma1.Text = "1.) Sólo se deben usar los segmentos S12, S13 y S14."
ProductosNormasForm.txt_Norma2.Text = "2.) Utilizar las cajas originales de los productos para apilarlos, tratando en lo posible que los de mayor cantidad se encuentren abajo."
ProductosNormasForm.txt_Norma3.Text = "3.) Utilizar las etiquetas tanto para los bordes de los estantes como las destinadas a las cajas (contenedores) de los productos."
ProductosNormasForm.txt_Norma4.Text = "4.) Utilizar el Segmento S12 para filtros de Aceites, El Segmento S13 para Filtros de Aire y el Segmento S14 para Filstros de Cabina, aunque éstos dos últimos se pueden mezclar."
ProductosNormasForm.txt_Norma5.Text = "5.) No mezclar productos de esta clasificación con otro tipo, por ejemplo: Artículos pesados, Partes de motor ó Bombas."

ProductosNormasForm.Show vbModal
End Sub

Private Sub lbl_NormasHot_Click()
ProductosNormasForm.Picture = LoadPicture(App.Path & "\Interfaz\NormasFormHot.jpg")
ProductosNormasForm.lbl_Clasificacion.Caption = "ARTÍCULOS MÁS VENDIDOS"
ProductosNormasForm.txt_Descripcion.Text = "Esta clasificación de segmento del almacén, tiene la finalidad de almacenar aquellos productos que son muy solicitados y se necesitan colocar en un lugar estratégico para agilizar búsquedas. Por ejemplo: Reservorios, Accesorios, Limpiaparabrisas, Etc."
ProductosNormasForm.txt_Norma1.Text = "1.) Sólo se debe usar el segmento S10"
ProductosNormasForm.txt_Norma2.Text = "2.) Apilar los productos mas pesados en la parte de abajo del segmeto, y los más ligeros arriba. Productos como reservorios o  limpiaparabrisas utilizar tirrap para colgarlos alrededor."
ProductosNormasForm.txt_Norma3.Text = "3.) Productos cuya salida se vea reducida y tenga su clasificación en otro segmento, se debe enviar al mismo para dar lugar a otros posibles productos que necesiten ubicarse más estratégicamente."
ProductosNormasForm.txt_Norma4.Text = "4.) Se aceptan mezclar varios tipos de productos. PERO se debe de tener ESPECIAL CUIDADO en esta actividad, puesto que se tiene cierto margen de riesgo. Mezclarlos con precaución y sentido común."
ProductosNormasForm.txt_Norma5.Text = ""

ProductosNormasForm.Show vbModal
End Sub

Private Sub lbl_NormasPesados_Click()
ProductosNormasForm.Picture = LoadPicture(App.Path & "\Interfaz\NormasFormPesados.jpg")
ProductosNormasForm.lbl_Clasificacion.Caption = "ARTÍCULOS PESADOS"
ProductosNormasForm.txt_Descripcion.Text = "Esta clasificación de segmentos del almacén, tiene la finalidad de almacenar productos cuyos rubros apilados tengan un peso o magnitud considerable. Por ejemplo: Bujes, Bases de Motor, Bases de Caja, Etc."
ProductosNormasForm.txt_Norma1.Text = "1.) Sólo se deben usar los segmentos S1, S2, S3 y S4."
ProductosNormasForm.txt_Norma2.Text = "2.) Se recomienda usar o armar cajas para almacenar un mismo tipo de bases o bujes del tamaño que se requiera y prestar mucha atención al sellado de las mismas."
ProductosNormasForm.txt_Norma3.Text = "3.) Utilizar las etiquetas tanto para los bordes de los estantes como las destinadas a las cajas (contenedores) de los productos."
ProductosNormasForm.txt_Norma4.Text = "4.) Si de un producto en particular se tiene una cantidad significativa (+250) y el mismo es de un tamaño pequeño, se recomienda sellar el empaque cada vez que se descuente (despache) ese tipo de mercancía."
ProductosNormasForm.txt_Norma5.Text = "5.) No mezclar productos de esta clasificación con otro tipo de productos  como Artículos Eléctricos, Bombas de Agua, Partes de Motor, Empacaduras, Etc. "

ProductosNormasForm.Show vbModal
End Sub

Private Sub lbl_NormasPMotor_Click()
ProductosNormasForm.Picture = LoadPicture(App.Path & "\Interfaz\NormasFormPMotor.jpg")
ProductosNormasForm.lbl_Clasificacion.Caption = "PARTES DE MOTOR"
ProductosNormasForm.txt_Descripcion.Text = "Esta clasificación de segmentos del almacén, tiene la finalidad de almacenar productos cuyos rubros apilados  constituyan piezas de motor de un vehículo. Por ejemplo: Correas, Engranajes, Pistones, Pilas, Etc.  "
ProductosNormasForm.txt_Norma1.Text = "1.) Sólo se debe utilizar el segmento S11."
ProductosNormasForm.txt_Norma2.Text = "2.) Al utilizar cajas para resguardar los productos se debe prestar especial atención de que se encuentren bien firmes y selladas, de tal forma que evite que algún producto se caiga/escape de dichas cajas."
ProductosNormasForm.txt_Norma3.Text = "3.) Utilizar las etiquetas tanto para los bordes de los estantes como las destinadas a las cajas (contenedores) de los productos."
ProductosNormasForm.txt_Norma4.Text = "4.) Productos cuyos peso/magnitud sea sobresaliente al de los demás, preferiblemente ubicar en la parte de abajo del estante."
ProductosNormasForm.txt_Norma5.Text = "5.) No mezclar productos de esta clasificación con otro tipo de productos cómo Filtros, Artículos Eléctricos, Bombas, Etc."

ProductosNormasForm.Show vbModal
End Sub

Private Sub lbl_S1_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 0
End Sub

Private Sub lbl_S10_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 9
End Sub

Private Sub lbl_S11_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 10
End Sub

Private Sub lbl_S12_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 11
End Sub

Private Sub lbl_S13_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 12
End Sub

Private Sub lbl_S14_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 13
End Sub

Private Sub lbl_S15_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 14
End Sub

Private Sub lbl_S16_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 15
End Sub

Private Sub lbl_S17_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 16
End Sub

Private Sub lbl_S18_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 17
End Sub

Private Sub lbl_S19_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 18
End Sub

Private Sub lbl_S2_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 1
End Sub

Private Sub lbl_S3_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 2
End Sub

Private Sub lbl_S4_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 3
End Sub

Private Sub lbl_S5_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 4
End Sub

Private Sub lbl_S6_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 5
End Sub

Private Sub lbl_S7_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 6
End Sub

Private Sub lbl_S8_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 7
End Sub

Private Sub lbl_S9_Click()
          Unload Me
          ProductoNewForm.Cbo_Ubicacion.ListIndex = 8
End Sub
