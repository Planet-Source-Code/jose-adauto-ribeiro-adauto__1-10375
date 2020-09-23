VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmManutencao 
   Caption         =   "Manutenção dos Programas da Lista"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9180
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame f3dPosicao 
      Height          =   2775
      Left            =   5880
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   4895
      _StockProps     =   14
      Caption         =   " Posiciona o Programa "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.VScrollBar UpDown 
         Height          =   1335
         Left            =   120
         Max             =   1
         Min             =   1
         TabIndex        =   6
         Top             =   360
         Value           =   1
         Width           =   255
      End
      Begin Threed.SSCommand cmdValidaPosicao 
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Valida Posição"
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
      Begin Threed.SSCommand cmdDesistePosicao 
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   2280
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Desiste"
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
      Begin VB.Label Label2 
         Caption         =   "Para Baixo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Para Cima"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxLista 
      Height          =   5295
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Para alterar nome ou parâmetro, excluir programa ou alterar ícone, selecione a linha e tecle ENTER ou dê um DOUBLE CLICK ... "
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   4
      BackColorSel    =   8454143
      ForeColorSel    =   8421631
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "  Total de programas ativos: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Left            =   1800
      TabIndex        =   0
      Top             =   5400
      Width           =   3975
   End
   Begin VB.Menu Pop02 
      Caption         =   "PopMenu2"
      Visible         =   0   'False
      Begin VB.Menu mnuCustomizacao 
         Caption         =   "&Alterar Nome/Parâmetros ou Excluir da Lista ..."
      End
      Begin VB.Menu mnuPosicao 
         Caption         =   "&Modificar ordem do programa na Lista ... "
      End
   End
   Begin VB.Menu mnuRetornar 
      Caption         =   "&RETORNAR"
   End
End
Attribute VB_Name = "frmManutencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flxListaWidth As Integer
Dim flxListaEnabled As Boolean
Private Sub flxLista_Clickx()
   flxLista.Col = 0
   flxLista.ColSel = 0
   flxLista.RowSel = flxLista.Row
   NumeroDoItem = flxLista.TextMatrix(flxLista.Row, 0)
   frmOpcoes.Show 1

End Sub

Private Sub cmdDesistePosicao_Click()
f3dPosicao.Visible = False
flxLista.Width = flxListaWidth
If LinhaAnterior <> 0 Then
   Call TornaBrancaACelula(LinhaAnterior)
End If
LinhaAnterior = 0
Call TornaBrancaACelula(NumeroDoItem)
Funcao = ""
flxListaEnabled = True
End Sub

Private Sub cmdValidaPosicao_Click()
Dim X1, X2, X3, X4
If LinhaAnterior < NumeroDoItem Then
   i = LinhaAnterior
   J = NumeroDoItem
   X1 = TabelaDeProgramasAtivos(J, 1)
   X2 = TabelaDeProgramasAtivos(J, 2)
   X3 = TabelaDeProgramasAtivos(J, 3)
   X4 = TabelaDeProgramasAtivos(J, 4)
   For L = J To i + 1 Step -1
      TabelaDeProgramasAtivos(L, 1) = TabelaDeProgramasAtivos(L - 1, 1)
      TabelaDeProgramasAtivos(L, 2) = TabelaDeProgramasAtivos(L - 1, 2)
      TabelaDeProgramasAtivos(L, 3) = TabelaDeProgramasAtivos(L - 1, 3)
      TabelaDeProgramasAtivos(L, 4) = TabelaDeProgramasAtivos(L - 1, 4)
   Next L
   TabelaDeProgramasAtivos(i, 1) = X1
   TabelaDeProgramasAtivos(i, 2) = X2
   TabelaDeProgramasAtivos(i, 3) = X3
   TabelaDeProgramasAtivos(i, 4) = X4
Else
   i = LinhaAnterior
   J = NumeroDoItem
   X1 = TabelaDeProgramasAtivos(J, 1)
   X2 = TabelaDeProgramasAtivos(J, 2)
   X3 = TabelaDeProgramasAtivos(J, 3)
   X4 = TabelaDeProgramasAtivos(J, 4)
   For L = J + 1 To i Step 1
      TabelaDeProgramasAtivos(L - 1, 1) = TabelaDeProgramasAtivos(L, 1)
      TabelaDeProgramasAtivos(L - 1, 2) = TabelaDeProgramasAtivos(L, 2)
      TabelaDeProgramasAtivos(L - 1, 3) = TabelaDeProgramasAtivos(L, 3)
      TabelaDeProgramasAtivos(L - 1, 4) = TabelaDeProgramasAtivos(L, 4)
   Next L
   TabelaDeProgramasAtivos(i, 1) = X1
   TabelaDeProgramasAtivos(i, 2) = X2
   TabelaDeProgramasAtivos(i, 3) = X3
   TabelaDeProgramasAtivos(i, 4) = X4

End If
Call RegerarActiveINI
AlterouLista = True
InicializarBarra = True
f3dPosicao.Visible = False
Funcao = ""
NumeroDoItem = 0
LinhaAnterior = 0
flxLista.Width = flxListaWidth
flxListaEnabled = True
Call Form_Activate
End Sub

Private Sub flxLista_DblClick()
If flxListaEnabled Then
   NumeroDoItem = flxLista.TextMatrix(flxLista.Row, 0)
   flxLista.Col = 1
   flxLista.ColSel = 1
   flxLista.CellBackColor = RGB(255, 0, 0)
   If LinhaAnterior <> 0 Then
      flxLista.ColSel = 1
      flxLista.Row = LinhaAnterior
      flxLista.CellBackColor = RGB(255, 255, 255)
   End If
   LinhaAnterior = NumeroDoItem
   frmFunção.Show 1
End If

End Sub

Private Sub flxLista_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call flxLista_DblClick
End If
End Sub

Private Sub Form_Activate()
Dim i As Integer, mItem As Object
Dim X1, X2, X3

If AlterouLista Then
   AlterouLista = False
   flxLista.Rows = 1
   UpDown.Max = NumeroDeProgramasAtivos
   For i = 1 To NumeroDeProgramasAtivos
        X1 = TabelaDeProgramasAtivos(i, 2) ' Nome fantasia do programa
        X2 = TabelaDeProgramasAtivos(i, 1) ' Path & nome do executável
        X3 = TabelaDeProgramasAtivos(i, 4) ' Parâmetros que serão passados ao programa na ativação
        flxLista.AddItem Format(i, "00") & vbTab & X1 & vbTab & X2 & vbTab & X3
   Next i
   Label1.Caption = "Total de programas ativos: " + Str(NumeroDeProgramasAtivos)
End If
If Funcao = Posicionar Then
   LinhaAnterior = 0
   cmdValidaPosicao.Enabled = False
   flxListaEnabled = False
End If
End Sub

Private Sub Form_Load()
    AlterouLista = True
    Me.flxLista.Cols = 4
    Me.flxLista.Width = Screen.Width
    i = Me.flxLista.Width '- 200
    Me.flxLista.ColWidth(0) = Me.flxLista.Width * (1 / 9)
    Me.flxLista.ColWidth(1) = Me.flxLista.Width * (2 / 9)
    Me.flxLista.ColWidth(2) = Me.flxLista.Width * (4 / 9)
    Me.flxLista.ColWidth(3) = Me.flxLista.Width * (2 / 9)
    Me.flxLista.ColAlignment(0) = flexAlignCenterCenter
    Me.flxLista.ColAlignment(1) = flexAlignCenterCenter
    Me.flxLista.ColAlignment(2) = flexAlignLeftCenter
    Me.flxLista.ColAlignment(3) = flexAlignLeftCenter
    Me.flxLista.TextMatrix(0, 0) = "Número"
    Me.flxLista.TextMatrix(0, 1) = "Nome"
    Me.flxLista.TextMatrix(0, 2) = "Diretório e Executável"
    Me.flxLista.TextMatrix(0, 3) = "Parâmetros p/ Ativação"
    Funcao = ""
    flxListaWidth = flxLista.Width
    flxListaEnabled = True
    Load frmFunção
End Sub



Private Sub mnuRetornar_Click()
Unload frmFunção
Unload Me
End Sub

Private Sub mnuCustomizacao_Click()
f3dPosicao.Visible = False
frmOpcoes.Show 1
End Sub

Private Sub mnuPosicao_Click()
   f3dPosicao.Visible = True
   UpDown.Value = flxLista.Row


End Sub

Private Sub UpDown_Change()
If Funcao = Posicionar Then
   If LinhaAnterior <> NumeroDoItem And UpDown.Value <> NumeroDoItem Then
      flxLista.Col = 1
      flxLista.RowSel = UpDown.Value
      flxLista.Row = UpDown.Value
      flxLista.CellBackColor = RGB(0, 0, 255)
      cmdValidaPosicao.Enabled = True
      If LinhaAnterior <> 0 Then
         Call TornaBrancaACelula(LinhaAnterior)
      End If
      LinhaAnterior = UpDown.Value
   ElseIf UpDown.Value = NumeroDoItem And LinhaAnterior <> NumeroDoItem Then
      cmdValidaPosicao.Enabled = False
      Call TornaBrancaACelula(LinhaAnterior)
   End If
   
   Exit Sub
End If
   If LinhaAnterior <> 0 Then
      flxLista.ColSel = 1
      flxLista.Row = LinhaAnterior
      flxLista.CellBackColor = RGB(255, 255, 255)
   End If
   LinhaAnterior = UpDown.Value
   flxLista.Row = UpDown.Value
   flxLista.RowSel = UpDown.Value
   flxLista.Col = 1
   flxLista.ColSel = 1
   flxLista.CellBackColor = RGB(255, 0, 0)
End Sub
