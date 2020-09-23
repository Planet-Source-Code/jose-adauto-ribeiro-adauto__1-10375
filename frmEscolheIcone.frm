VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEscolheIcone 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Escolha do Ícone do Programa"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPGM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   0
      Left            =   165
      Picture         =   "frmEscolheIcone.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   33.032
      TabIndex        =   3
      Top             =   360
      Width           =   510
   End
   Begin Threed.SSCommand cmdValida 
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "VALIDAR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblNumIcone 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " N. 01 "
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   165
      TabIndex        =   2
      Top             =   960
      Width           =   525
   End
   Begin VB.Label lblIcone 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Número do ícone a validar: nn "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3420
      TabIndex        =   1
      Top             =   3000
      Width           =   2775
   End
End
Attribute VB_Name = "frmEscolheIcone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IAnt As Integer
Private Sub cmdValida_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Ic&, J As Integer, L As Integer
Ic& = 0
'Ic& = TabelaDeProgramasAtivos(NumeroDoItem, 5)
Call CarregaIcone(0, Ic&)
L = 1
For J = 1 To IIf(TabelaDeProgramasAtivos(NumeroDoItem, 6) - 1 > MaximoComIconesMaiores, MaximoComIconesMaiores, TabelaDeProgramasAtivos(NumeroDoItem, 6) - 1)
    Ic& = J
    Load picPGM(J)
    L = L + 1
    
    picPGM(J).Left = picPGM(J - 1).Left + 60 + picPGM(J - 1).Width
    Load lblNumIcone(J)
    lblNumIcone(J).Left = picPGM(J).Left
    lblNumIcone(J).Caption = " N. " + Format(J + 1, "00")
    Call CarregaIcone(J, Ic&)
    picPGM(J).Visible = True
    lblNumIcone(J).Visible = True
 Next J
lblNumIcone(TabelaDeProgramasAtivos(NumeroDoItem, 5)).BackColor = RGB(0, 0, 255)
IAnt = CInt(TabelaDeProgramasAtivos(NumeroDoItem, 5))
frmEscolheIcone.lblIcone.Caption = " Número do ícone a validar: " + Format(IAnt + 1, "00")
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
End Sub
Private Sub CarregaIcone(i As Integer, L&)
Dim hIcon&
    
    ' clear picturebox (cls won't work)
    picPGM(i).Picture = LoadPicture()
    ' get a handle on the icon associated with o módulo ...
    hIcon& = ExtractAssociatedIcon(App.hInstance, TabelaDeProgramasAtivos(NumeroDoItem, 1), L&)
    ' draw the icon on the dc (device context) of picturebox
    DrawIcon picPGM(i).hdc, 0, 0, hIcon&
    ' destroy the icon (necessary)
    DestroyIcon hIcon&

End Sub

Private Sub picPGM_Click(Index As Integer)
TabelaDeProgramasAtivos(NumeroDoItem, 5) = Index
frmEscolheIcone.lblIcone.Caption = " Número do ícone a validar: " + Format(Index + 1, "00")
frmEscolheIcone.lblNumIcone(Index).BackColor = RGB(0, 0, 255)
If Index <> IAnt Then
  frmEscolheIcone.lblNumIcone(IAnt).BackColor = RGB(255, 255, 255)
End If
IAnt = Index
End Sub
