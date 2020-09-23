VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmFunção 
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand cmdOpção 
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   4
      ToolTipText     =   "Clique aqui para alterar a posição do programa na lista ..."
      Top             =   240
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Alterar Posição do Programa na Lista"
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
   Begin Threed.SSCommand cmdOpção 
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   3
      ToolTipText     =   "Clique aqui para alterar o nome fantasia do programa ou parâmetro passado para ativá-lo ..."
      Top             =   960
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Alterar Nome ou Parâmetros do Programa"
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
   Begin Threed.SSCommand cmdOpção 
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "Clique aqui para eliminar o programa da lista ..."
      Top             =   1680
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Excluir o Programa da Lista"
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
   Begin Threed.SSCommand cmdOpção 
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   1
      ToolTipText     =   "Clique aqui para escolher outro ícone para mostrar na barra ..."
      Top             =   2400
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Escolher Outro Ícone "
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
   Begin Threed.SSCommand cmdOpção 
      Height          =   375
      Index           =   4
      Left            =   480
      TabIndex        =   0
      ToolTipText     =   "Clique aqui para retornar sem alterar nada na lista ..."
      Top             =   3120
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Retornar Sem Alteração"
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
End
Attribute VB_Name = "frmFunção"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim First As Boolean
Private Sub cmdOpção_Click(Index As Integer)
Select Case Index
   Case 0
     frmManutencao.flxLista.Width = frmManutencao.f3dPosicao.Left - 5
     frmManutencao.f3dPosicao.Visible = True
     Funcao = Posicionar
     frmManutencao.UpDown.Value = NumeroDoItem
     Me.Hide
   Case 1
     Funcao = Alterar
     frmOpcoes.Show 1
     Call TornaBrancaACelula(LinhaAnterior)
     Me.Hide
   Case 2
     Funcao = Excluir
     frmOpcoes.Show 1
     Call TornaBrancaACelula(LinhaAnterior)
     Me.Hide
   Case 3
     Funcao = AlterarÍcone
     Call TornaBrancaACelula(LinhaAnterior)
     Me.Hide
     frmEscolheIcone.Show 1
     Call frmActive.AtivaPGMNaBarra(frmActive, frmActive.picPGM, NumeroDoItem)
     frmActive.imgBotao(NumeroDoItem).Picture = frmActive.picPGM.Image
     Chave = PGMIcon + Format(NumeroDoItem, "00")
     Valor = Format(TabelaDeProgramasAtivos(NumeroDoItem, 5), "00")
     X = EscreveIni(Seção, Chave, Valor)
   Case 4
     Funcao = Desistir
     Call TornaBrancaACelula(LinhaAnterior)
     Me.Hide
End Select
End Sub

Private Sub Form_Activate()
Dim L As Integer
L = ExtractIcon(App.hInstance, TabelaDeProgramasAtivos(NumeroDoItem, 1), -1)
If L > 1 Then
   frmFunção.cmdOpção(3).Font.Bold = True
   frmFunção.cmdOpção(3).Enabled = True
   TabelaDeProgramasAtivos(NumeroDoItem, 6) = L

Else
   frmFunção.cmdOpção(3).Font.Bold = False
   frmFunção.cmdOpção(3).Enabled = False
End If
If First Then
   cmdOpção(4).SetFocus
   First = False
End If
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
First = True
End Sub
