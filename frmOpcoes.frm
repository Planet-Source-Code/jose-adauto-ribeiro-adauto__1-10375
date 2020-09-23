VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmOpcoes 
   Caption         =   "Atualiza Informações do Programa"
   ClientHeight    =   4110
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8760
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "ELIMINAR ENTRADA DA LISTA"
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
      Left            =   2520
      TabIndex        =   6
      Top             =   3360
      Width           =   3735
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   735
      Left            =   960
      TabIndex        =   10
      Top             =   0
      Width           =   7095
      _Version        =   65536
      _ExtentX        =   12515
      _ExtentY        =   1296
      _StockProps     =   14
      Caption         =   " Caminho e nome do arquivo do programa "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.TextBox txtArquivo 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   360
         Width           =   6855
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   735
      Left            =   4320
      TabIndex        =   7
      Top             =   840
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   1296
      _StockProps     =   14
      Caption         =   " Parâmetro Anterior "
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
      Begin VB.TextBox txtParmAnterior 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdValidar 
      Caption         =   "VALIDAR ALTERAÇÕES"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   2760
      Width           =   3015
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   1296
      _StockProps     =   14
      Caption         =   " Nome Anterior "
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
      Begin VB.TextBox txtNomeAnterior 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   2415
      End
   End
   Begin Threed.SSFrame f3dNovoNome 
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   1296
      _StockProps     =   14
      Caption         =   " Novo Nome "
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
      Begin VB.TextBox txtNomeNovo 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
   End
   Begin Threed.SSFrame f3dNovoParametro 
      Height          =   735
      Left            =   4320
      TabIndex        =   9
      Top             =   1680
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   1296
      _StockProps     =   14
      Caption         =   " Novo Parâmetro "
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
      Begin VB.TextBox txtParmNovo 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Menu mnuRetornar 
      Caption         =   "&RETORNAR"
   End
End
Attribute VB_Name = "frmOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExcluir_Click()
If MsgBox("Tem certez que deseja excluir este programa da lista ? ", vbYesNo + vbQuestion, "Atenção !") = vbYes Then
   Dim btnX As Button
J = 0
'debug.Print "Rotina excluir:" + Str(NumeroDeProgramasAtivos)
For i = 1 To NumeroDeProgramasAtivos
    If i > NumeroDoItem Then
       Chave = PGMName + Format(J, "00")
       Valor = TabelaDeProgramasAtivos(i, 1)
       TabelaDeProgramasAtivos(J, 1) = Valor
       x = EscreveIni(Seção, Chave, Valor)
       Chave = PGMTitle + Format(J, "00")
       Valor = TabelaDeProgramasAtivos(i, 2)
       TabelaDeProgramasAtivos(J, 2) = Valor
       x = EscreveIni(Seção, Chave, Valor)
       'Chave = Left(TabelaDeProgramasAtivos(I, 3), Len(TabelaDeProgramasAtivos(I, 3)) - 4)
       'Valor = Format(J, "00") + Chave
       'X = EscreveIni(Seção, Chave, Valor)
       Chave = PGMbmp + Format(J, "00")
       Valor = TabelaDeProgramasAtivos(i, 3)
       TabelaDeProgramasAtivos(J, 3) = Valor
       x = EscreveIni(Seção, Chave, Valor)
       Chave = PGMParm + Format(J, "00")
       Valor = TabelaDeProgramasAtivos(i, 4)
       TabelaDeProgramasAtivos(J, 4) = Valor
       x = EscreveIni(Seção, Chave, Valor)
       Chave = PGMIcon + Format(J, "00")
       Valor = Format(TabelaDeProgramasAtivos(i, 5), "00")
       x = EscreveIni(Seção, Chave, Valor)
       J = J + 1
    Else
       J = J + 1
    End If
       
Next i
   Chave = PGMName + Format(NumeroDeProgramasAtivos, "00")
   Valor = ""
   x = EscreveIni(Seção, Chave, Valor)
   Chave = PGMTitle + Format(NumeroDeProgramasAtivos, "00")
   Valor = ""
   x = EscreveIni(Seção, Chave, Valor)
   Chave = PGMbmp + Format(NumeroDeProgramasAtivos, "00")
   Valor = ""
   x = EscreveIni(Seção, Chave, Valor)
   Chave = PGMParm + Format(NumeroDeProgramasAtivos, "00")
   Valor = ""
   x = EscreveIni(Seção, Chave, Valor)
    Chave = PGMIcon + Format(NumeroDeProgramasAtivos, "00")
    Valor = ""
    x = EscreveIni(Seção, Chave, Valor)
   AlterouLista = True
   InicializarBarra = True
   NumeroDeProgramasAtivos = NumeroDeProgramasAtivos - 1
   Unload Me

End If
End Sub
Private Sub cmdValidar_Click()
Dim Erro As Boolean
If Trim(txtNomeNovo.Text) <> TitleAnterior Then
   If Trim(txtNomeNovo.Text) = "" Then
      txtNomeNovo.Text = txtNomeAnterior.Text
      txtNomeNovo.SelLength = Len(txtNomeAnterior.Text)
      txtNomeNovo.SelStart = 0
      Erro = True
      txtNomeNovo.SetFocus
   End If
End If
If Not Erro Then
   If Trim(txtNomeNovo.Text) <> TitleAnterior Or Trim(txtParmNovo.Text) <> TabelaDeProgramasAtivos(NumeroDoItem, 4) Then
      'If Trim(txtNomeNovo.Text) <> TitleAnterior Then
      '   InicializarBarra = True
      'End If
      TabelaDeProgramasAtivos(NumeroDoItem, 2) = Trim(txtNomeNovo.Text)
      Chave = PGMTitle + Format(NumeroDoItem, "00")
      Valor = TabelaDeProgramasAtivos(NumeroDoItem, 2)
      x = EscreveIni(Seção, Chave, Valor)
      frmactive.imgBotao(NumeroDoItem).ToolTipText = Format(NumeroDoItem, "00") + " - " + Valor
      TabelaDeProgramasAtivos(NumeroDoItem, 4) = Trim(txtParmNovo.Text)
      Chave = PGMParm + Format(NumeroDoItem, "00")
      Valor = TabelaDeProgramasAtivos(NumeroDoItem, 4)
      x = EscreveIni(Seção, Chave, Valor)
      AlterouLista = True
      Unload Me
   Else
      Erro = True
      MsgBox "Novo nome não pode ser branco ou omitido."
      txtAnterior.SetFocus
   End If
Unload Me
End If

End Sub

Private Sub Form_Load()
   'debug.Print NumeroDeProgramasAtivos
txtNomeAnterior.Text = TabelaDeProgramasAtivos(NumeroDoItem, 2)
txtNomeNovo.Text = txtNomeAnterior.Text
txtParmAnterior.Text = TabelaDeProgramasAtivos(NumeroDoItem, 4)
txtParmNovo.Text = txtParmAnterior.Text
txtArquivo.Text = TabelaDeProgramasAtivos(NumeroDoItem, 1)
txtNomeAnterior.Enabled = False
txtParmAnterior.Enabled = False
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
If Funcao = Alterar Then
   frmOpcoes.cmdExcluir.Visible = False
Else
   frmOpcoes.cmdValidar.Visible = False
   frmOpcoes.f3dNovoNome.Visible = False
   frmOpcoes.f3dNovoParametro.Visible = False
End If
End Sub

Private Sub mnuRetornar_Click()
Unload Me
End Sub

Private Sub txtNomeNovo_GotFocus()
Call TextSelected
End Sub

Private Sub txtParmNovo_GotFocus()
Call TextSelected

End Sub

