VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmActive 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   330
   ClientLeft      =   120
   ClientTop       =   15
   ClientWidth     =   9195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmActive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picNotifier 
      Height          =   630
      Left            =   4080
      ScaleHeight     =   570
      ScaleWidth      =   435
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   252
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picPGM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1320
      ScaleHeight     =   23
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog cdgPGM 
      Left            =   360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgBotao 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   350
      Index           =   0
      Left            =   0
      Picture         =   "frmActive.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   350
   End
   Begin VB.Image imgBarraAnterior 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   175
      Left            =   840
      Picture         =   "frmActive.frx":0F4C
      Stretch         =   -1  'True
      ToolTipText     =   "Clique aqui para ver os programas da barra anterior ... "
      Top             =   175
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgProximaBarra 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   170
      Left            =   840
      Picture         =   "frmActive.frx":1B8E
      Stretch         =   -1  'True
      ToolTipText     =   "Clique aqui para ver os programas da próxima barra ... "
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Menu Pop01 
      Caption         =   "PopMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuIncluir 
         Caption         =   "&Incluir programa na lista ..."
      End
      Begin VB.Menu mnuCustomizar 
         Caption         =   "&Customizar programa na lista ..."
      End
      Begin VB.Menu mnutopo 
         Caption         =   "Ficar sempre no &Topo"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCaption 
         Caption         =   "Permite deslocar na tela"
      End
      Begin VB.Menu mnuBotoes 
         Caption         =   "&Botões Maiores "
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSobre 
         Caption         =   "&Sobre o Programa ..."
      End
      Begin VB.Menu mnuEncerrar 
         Caption         =   "&Encerrar"
      End
   End
End
Attribute VB_Name = "frmActive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const cIconPixelWidth As Integer = 32
Private Const cIconPixelHeight As Integer = 32
Dim First As Boolean, AlturaOriginalDoFormFrmActive As Long
Dim NomePGM As String, NomeBMP As String, TitlePGM As String
Dim m_hinst As Double
'
' Extract icons class file.
'
'Dim IconExtract As New CExecIcons
Dim res
      'Declare a user-defined variable to pass to the Shell_NotifyIcon
      'function.
      Private Type NOTIFYICONDATA
         cbSize As Long
         hWnd As Long
         uId As Long
         uFlags As Long
         uCallBackMessage As Long
         hIcon As Long
         szTip As String * 64
      End Type

      'Declare the constants for the API function. These constants can be
      'found in the header file Shellapi.h.

      'The following constants are the messages sent to the
      'Shell_NotifyIcon function to add, modify, or delete an icon from the
      'taskbar status area.
      Private Const NIM_ADD = &H0
      Private Const NIM_MODIFY = &H1
      Private Const NIM_DELETE = &H2

      'The following constant is the message sent when a mouse event occurs
      'within the rectangular boundaries of the icon in the taskbar status
      'area.
      Private Const WM_MOUSEMOVE = &H200

      'The following constants are the flags that indicate the valid
      'members of the NOTIFYICONDATA data type.
      Private Const NIF_MESSAGE = &H1
      Private Const NIF_ICON = &H2
      Private Const NIF_TIP = &H4

      'The following constants are used to determine the mouse input on the
      'the icon in the taskbar status area.

      'Left-click constants.
      Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Private Const WM_LBUTTONDOWN = &H201     'Button down
      Private Const WM_LBUTTONUP = &H202       'Button up

      'Right-click constants.
      Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
      Private Const WM_RBUTTONDOWN = &H204     'Button down
      Private Const WM_RBUTTONUP = &H205       'Button up

      'Declare the API function call.
      Private Declare Function Shell_NotifyIcon Lib "shell32" _
         Alias "Shell_NotifyIconA" _
         (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      'Dimension a variable as the user-defined data type.
      Dim nid As NOTIFYICONDATA
Private Sub AtiveSeNovamente()
      Call Form_Terminate
      X = Shell(DiretorioInicial + "\active.exe", vbNormalFocus)
      End
End Sub

Private Sub Form_Activate()
Dim msg1 As String
On Error GoTo ErroNaCarga
   If InicializarBarra Then
      Call AtiveSeNovamente
   End If
If First Then
   msg1 = "Rotina RecuperaStatusTOPEMaxPGMNaBarra"
   Call RecuperaStatusTOPEMaxPGMNaBarra
   If NumeroMaximoDeProgramasNaBarra = MaximoComIconesMaiores Then
         Me.imgBotao(i).Width = 500
         Me.imgBotao(i).Height = 500
         Me.imgBarraAnterior.Width = Me.imgBarraAnterior.Width * (500 / 350)
         Me.imgProximaBarra.Width = Me.imgProximaBarra.Width * (500 / 350)
         Me.imgBarraAnterior.Height = Me.imgBarraAnterior.Height * (500 / 350)
         Me.imgProximaBarra.Height = Me.imgProximaBarra.Height * (500 / 350)
         Me.Height = 510
   End If
   Me.Width = Me.imgBotao(0).Width
   AlturaOriginalDoFormFrmActive = frmActive.Height
   msg1 = "Rotina AtualizaProgramasNaBarra"
   Call AtualizaProgramasNaBarra
   i = IIf(NumeroDeProgramasAtivos < 5, 5, NumeroDeProgramasAtivos)
   Me.Width = (Me.imgBotao(0).Width * 1.01) * (IIf(i > NumeroMaximoDeProgramasNaBarra, NumeroMaximoDeProgramasNaBarra, i) + 1) + IIf(i > NumeroMaximoDeProgramasNaBarra, Me.imgBarraAnterior.Width, 0)
   Me.Left = (Screen.Width - Me.Width) / 2
   First = False
End If
Exit Sub
ErroNaCarga:
MsgBox "Erro na ativação do programa: " + msg1
End Sub
Private Sub Form_Load()
Dim msg1 As String
On Error GoTo ErroNaCarga
PGMVersao = "Versão " & App.Major & "." & App.Minor & "." & App.Revision
DiretorioInicial = App.Path
ChDir DiretorioInicial
CrLf = Chr(10) + Chr(13)
Seção = "PGMS_Barra"
Active_INI = DiretorioInicial + "\Active.INI"
picPGM.Height = picPGM.Height * (cIconPixelHeight / picPGM.ScaleHeight)
picPGM.Width = picPGM.Height
HScroll1.SmallChange = cIconPixelWidth
HScroll1.LargeChange = cIconPixelWidth * 5
First = True
frmActive.Top = 0
NumeroDeProgramasAtivos = 0
msg1 = "Rotina FileExists"
If FileExists(Active_INI) Then
   Chave = CTLNumPGM
   msg1 = "Rotina LeIni"
   X = LeIni(Seção, Chave)
   If X = "" Then
      MsgBox "Arquivo de Controle ACTIVE.INI inválido. Será reiniciado."
      Kill Active_INI
   Else
      NumeroDeProgramasAtivos = CInt(X)
   End If
End If
Me.imgBotao(0).ToolTipText = "Clique aqui para personalizar sua barra de programas ..."
NumeroDoPrimeiroProgramaDaBarra = 1
msg1 = "Rotina ColocaNaBandeja"
Call ColocaIconeNaBandeja
Exit Sub
ErroNaCarga:
MsgBox "Erro na carga do programa: " + msg1
End Sub

Private Sub Form_Terminate()
         'Delete the added icon from the taskbar status area when the
         'program ends.
         Shell_NotifyIcon NIM_DELETE, nid

End Sub

Private Sub imgBarraAnterior_Click()
If NumeroDoPrimeiroProgramaDaBarra - NumeroMaximoDeProgramasNaBarra >= 1 Then
   NumeroDoPrimeiroProgramaDaBarra = NumeroDoPrimeiroProgramaDaBarra - NumeroMaximoDeProgramasNaBarra
   Call MostraProgramasAnterioresDaBarra
   If NumeroDoPrimeiroProgramaDaBarra = 1 Then
      imgBarraAnterior.Visible = False
   End If
   imgProximaBarra.Visible = True
End If
End Sub

Private Sub imgBotao_Click(Index As Integer)
  
  Select Case Index
     Case 0
     ' Trata-se do botao que ativa o MenuPop de funções
       PopupMenu Pop01
     Case Else
     ' Qualquer outro caso aciona o botao correspondente ...
     ' Use the Index aa a pointer to program's table ...
        Dim ParmPgm
        If TabelaDeProgramasAtivos(Index, 4) = "" Then
           ParmPgm = ""
        Else
           ParmPgm = " """ + TabelaDeProgramasAtivos(Index, 4) + """ "
           'ParmPgm = "  " + TabelaDeProgramasAtivos(Index, 4) '+ " '"
        End If
        On Error Resume Next
        X = Shell(TabelaDeProgramasAtivos(Index, 1) + ParmPgm, vbNormalFocus)
        If Err <> 0 Then
           MsgBox "Erro na execução do programa: " + TabelaDeProgramasAtivos(Index, 1) + CrLf + IIf(ParmPgm = "", "Mensagem de erro: " + Error$, " com PARM: " + ParmPgm + CrLf + "Será acionado sem parâmetros." + CrLf + "Mensagem de erro: " + Error$)
           If ParmPgm <> "" Then
              X = Shell(TabelaDeProgramasAtivos(Index, 1), vbNormalFocus)
           End If
        End If
  End Select
End Sub


Private Sub imgProximaBarra_Click()
If NumeroDoPrimeiroProgramaDaBarra + NumeroMaximoDeProgramasNaBarra <= NumeroDeProgramasAtivos Then
   Call EscondeProgramasAtuaisDaBarra
   NumeroDoPrimeiroProgramaDaBarra = NumeroDoPrimeiroProgramaDaBarra + NumeroMaximoDeProgramasNaBarra
   For i = NumeroDoPrimeiroProgramaDaBarra To IIf(NumeroDeProgramasAtivos >= NumeroDoPrimeiroProgramaDaBarra + NumeroMaximoDeProgramasNaBarra, NumeroDoPrimeiroProgramaDaBarra + NumeroMaximoDeProgramasNaBarra - 1, NumeroDeProgramasAtivos)
      imgBotao(i).Visible = True
   Next i
   imgBarraAnterior.Visible = True
   If NumeroDoPrimeiroProgramaDaBarra + NumeroMaximoDeProgramasNaBarra > NumeroDeProgramasAtivos Then
      imgProximaBarra.Visible = False
   End If
End If
End Sub

Private Sub mnuBotoes_Click()
If Not mnuBotoes.Checked Then
   mnuBotoes.Checked = True
   NumeroMaximoDeProgramasNaBarra = MaximoComIconesMaiores
   Call GravaMaxPGMNaBarraNoINI
   If NumeroDeProgramasAtivos <= MaximoComIconesMaiores Then
      For i = 0 To NumeroDeProgramasAtivos
         Me.imgBotao(i).Width = 500
         Me.imgBotao(i).Height = 500
         If i > 0 Then
            Me.imgBotao(i).Left = Me.imgBotao(i - 1).Left + Me.imgBotao(i - 1).Width
         End If
         Me.Height = 510
      Next i
      Me.imgBarraAnterior.Width = Me.imgBarraAnterior.Width * (500 / 350)
      Me.imgProximaBarra.Width = Me.imgProximaBarra.Width * (500 / 350)
      Me.imgBarraAnterior.Height = Me.imgBarraAnterior.Height * (500 / 350)
      Me.imgProximaBarra.Height = Me.imgProximaBarra.Height * (500 / 350)
      Me.Width = (NumeroDeProgramasAtivos + 1) * (Me.imgBotao(0).Width * 1.01)
      Me.Left = (Screen.Width - Me.Width) / 2
   Else
      Call AtiveSeNovamente
   End If
Else
   mnuBotoes.Checked = False
   NumeroMaximoDeProgramasNaBarra = MaximoComIconesMenores
   Call GravaMaxPGMNaBarraNoINI
   If NumeroDeProgramasAtivos <= MaximoComIconesMenores Then
      For i = 0 To NumeroDeProgramasAtivos
         Me.imgBotao(i).Width = 350
         Me.imgBotao(i).Height = 350
         Me.Height = 360
         If i > 0 Then
            Me.imgBotao(i).Left = Me.imgBotao(i - 1).Left + Me.imgBotao(i - 1).Width
         End If
      Next i
      Me.imgBarraAnterior.Width = 270
      Me.imgProximaBarra.Width = 270
      Me.imgBarraAnterior.Height = 175
      Me.imgProximaBarra.Height = 175
      Me.Width = (NumeroDeProgramasAtivos + 1) * (Me.imgBotao(0).Width * 1.01)
      Me.Left = (Screen.Width - Me.Width) / 2
   Else
      Call AtiveSeNovamente
   End If
End If
End Sub

Private Sub mnuCaption_Click()
If Not mnuCaption.Checked Then
   frmActive.Caption = "Barra de programas"
   frmActive.Height = frmActive.Height - 250
   frmActive.Top = frmActive.Top + 350
Else
   frmActive.Caption = ""
   frmActive.Height = AlturaOriginalDoFormFrmActive
   frmActive.Top = frmActive.Top + 600
End If
frmActive.mnuCaption.Checked = Not frmActive.mnuCaption.Checked
End Sub

Private Sub mnuCustomizar_Click()
If mnutopo.Checked Then
   'To turn off topmost (make the form act normal again):
   res = SetWindowPos(frmActive.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End If
frmManutencao.Show 1
If mnutopo.Checked Then
   'To set Form1 as a TopMost form, do the following:
   res = SetWindowPos(frmActive.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End If
End Sub

Private Sub mnuSobre_Click()
If mnutopo.Checked Then
   Call ColocaFormNormal
End If
frmSplash.Show 1
If mnutopo.Checked Then
   Call ColocaFormONTOP
 End If

End Sub

Private Sub mnutopo_Click()
If Not frmActive.mnutopo.Checked Then
   Call ColocaFormONTOP
Else
   Call ColocaFormNormal
End If
frmActive.mnutopo.Checked = Not frmActive.mnutopo.Checked
Call SalvarStatusTOP
End Sub



Private Sub mnuEncerrar_Click()
Call Form_Terminate
 End
End Sub
Private Sub mnuIncluir_Click()
If NumeroDeProgramasAtivos < NumeroMaximoDeProgramasAtivos Then
   If IndicouPGM() Then
       Call GravaProximoPGM
    End If
Else
   Beep
   MsgBox "Número máximo de programas na barra já atingdido: " + Str(NumeroMaximoDeProgramasAtivos)
End If
End Sub
Public Function IndicouPGM()
    cdgPGM.CancelError = True
    On Error GoTo cancelfile
    IndicouPGM = False
pedearquivoexe:
    cdgPGM.FLAGS = &H2&
    cdgPGM.Filter = "PROGRAMA (*.EXE)|*.EXE"
    cdgPGM.Action = 1
    ' em CMDIALOG1.FILENAME vem o nome do ARQUIVO escolhido
    ' se estiver nulo e' porque foi dado cancel
    ArquivoDePrograma = cdgPGM.FileName
    If Not FileExists(ArquivoDePrograma) Then
       Beep
       MsgBox "ACT901I Arquivo " + UCase(ArquivoDePrograma) + " não existe. Reentre ..."
       GoTo pedearquivoexe
    Else
       IndicouPGM = True
    End If
Exit Function
cancelfile:
    ChDir DiretorioInicial
    IndicouPGM = False
End Function
Public Sub AtualizaProgramasNaBarra()
Dim J As Integer, TodosCarregados As Boolean
J = 0
ProximoBotaoDaBarra = 1
J = 0
For i = 1 To NumeroDeProgramasAtivos
   Chave = PGMName + Format(i, "00")
   X = LeIni(Seção, Chave)
   If X = "" Then
      Exit For
   Else
      TabelaDeProgramasAtivos(J + 1, 1) = X
      NomePGM = PegaNomeDoPrograma(X)
         Chave = PGMbmp + Format(i, "00")
         X = LeIni(Seção, Chave)
         NomeBMP = DiretorioInicial + "\" + X
         J = J + 1
         TabelaDeProgramasAtivos(J, 3) = X
         Chave = PGMTitle + Format(i, "00")
         X = LeIni(Seção, Chave)
         TitlePGM = X
         TabelaDeProgramasAtivos(J, 2) = X
         Chave = PGMParm + Format(i, "00")
         X = LeIni(Seção, Chave)
         TabelaDeProgramasAtivos(J, 4) = X
         X = TitlePGM
         Chave = PGMIcon + Format(i, "00")
         X = LeIni(Seção, Chave)
         If X = "" Then
            TabelaDeProgramasAtivos(J, 5) = 0
         Else
            TabelaDeProgramasAtivos(J, 5) = CInt(X)
         End If
         Call InsereBotaoNaBarra(J)
   End If
Next i
If NumeroDeProgramasAtivos > 0 And NumeroDeProgramasAtivos <> J Then
   NumeroDeProgramasAtivos = J
   Call RegerarActiveINI
Else
   NumeroDeProgramasAtivos = J
End If
If NumeroDeProgramasAtivos > NumeroMaximoDeProgramasNaBarra Then
   imgProximaBarra.Visible = True
End If
End Sub
Private Sub EscolheIconeParaABarra(Npgm)
   Dim ICount As Integer
   ICount = ExtractIcon(App.hInstance, TabelaDeProgramasAtivos(Npgm, 1), -1)
   If ICount > 1 Then
      TabelaDeProgramasAtivos(Npgm, 5) = 0
      TabelaDeProgramasAtivos(Npgm, 6) = ICount
      If mnutopo.Checked Then
        'To turn off topmost (make the form act normal again):
        res = SetWindowPos(frmActive.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
      End If
      frmActive.Hide
      NumeroDoItem = Npgm
      frmEscolheIcone.Show 1
      frmActive.Show
      If mnutopo.Checked Then
         'To set Form1 as a TopMost form, do the following:
         res = SetWindowPos(frmActive.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
      End If
   Else
      TabelaDeProgramasAtivos(Npgm, 5) = 0
      TabelaDeProgramasAtivos(Npgm, 6) = 1
   End If

End Sub

Public Sub AtivaPGMNaBarra(F As Form, picPGM As PictureBox, NUMPGM As Integer)
Dim hIcon&, L&
    
    ' clear picturebox (cls won't work)
    F.picPGM.Picture = LoadPicture()
    ' get a handle on the icon associated with o módulo ...
    L& = TabelaDeProgramasAtivos(NUMPGM, 5)
    hIcon& = ExtractAssociatedIcon(App.hInstance, TabelaDeProgramasAtivos(NUMPGM, 1), L&)
    ' draw the icon on the dc (device context) of picturebox
    DrawIcon F.picPGM.hdc, 0, 0, hIcon&
    ' destroy the icon (necessary)
    DestroyIcon hIcon&
End Sub

Public Sub GravaProximoPGM()
 NomePGM = PegaNomeDoPrograma(ArquivoDePrograma)
 Chave = NomePGM
 If VerificaSeEmDuplicidade(NomePGM) Then
    NumeroDeProgramasAtivos = NumeroDeProgramasAtivos + 1
    Chave = PGMName + Format(NumeroDeProgramasAtivos, "00")
    TabelaDeProgramasAtivos(NumeroDeProgramasAtivos, 1) = ArquivoDePrograma
    NomeBMP = NomePGM + ".bmp"
    TabelaDeProgramasAtivos(NumeroDeProgramasAtivos, 2) = NomePGM
    TabelaDeProgramasAtivos(NumeroDeProgramasAtivos, 3) = NomeBMP
    Valor = ArquivoDePrograma
    X = EscreveIni(Seção, Chave, Valor)
    TabelaDeProgramasAtivos(NumeroDeProgramasAtivos, 4) = ""
    Chave = PGMTitle + Format(NumeroDeProgramasAtivos, "00")
    Valor = TabelaDeProgramasAtivos(NumeroDeProgramasAtivos, 2)
    X = EscreveIni(Seção, Chave, Valor)
    Chave = PGMbmp + Format(NumeroDeProgramasAtivos, "00")
    Valor = NomeBMP
    X = EscreveIni(Seção, Chave, Valor)
    Chave = CTLNumPGM
    Valor = Format(NumeroDeProgramasAtivos, "00")
    X = EscreveIni(Seção, Chave, Valor)
    Call EscolheIconeParaABarra(NumeroDeProgramasAtivos)
    Chave = PGMIcon + Format(i, "00")
    Valor = Format(TabelaDeProgramasAtivos(NumeroDeProgramasAtivos, 5), "00")
    X = EscreveIni(Seção, Chave, Valor)
    Call InsereBotaoNaBarra(NumeroDeProgramasAtivos)
 End If

End Sub
Public Function VerificaSeEmDuplicidade(NomePGM As String)
   VerificaSeEmDuplicidade = True
   For i = 1 To NumeroDeProgramasAtivos
       If UCase(NomePGM) + ".BMP" = UCase(TabelaDeProgramasAtivos(i, 3)) Then
          If MsgBox("O programa '" + NomePGM + "' já está na barra (Número " + Format(i, "00") + ") Deseja incluir mais um ícone para êle ? ", vbYesNo, "Atênção !") = vbNo Then
             VerificaSeEmDuplicidade = False
             Exit For
          Else
             Exit For
          End If
       End If
   Next i
   
End Function
Public Function PegaNomeDoPrograma(Arq As String)
Dim i As Integer
J = 0
For i = Len(Arq) To 3 Step -1
    If Mid(Arq, i, 1) = "\" Then
       Exit For
    Else
       J = J + 1
    End If
Next i
X = Right(Arq, J)
PegaNomeDoPrograma = Left(X, J - 4)
End Function


Public Sub InsereBotaoNaBarra(NumeroDoPrograma As Integer)
    Call AtivaPGMNaBarra(frmActive, picPGM, NumeroDoPrograma)
    SavePicture picPGM.Image, DiretorioInicial + "\" + TabelaDeProgramasAtivos(NumeroDoPrograma, 3)
    Load imgBotao(NumeroDoPrograma)
    imgBotao(NumeroDoPrograma).Left = imgBotao((NumeroDoPrograma - 1) Mod NumeroMaximoDeProgramasNaBarra).Left + imgBotao(0).Width
    imgBotao(NumeroDoPrograma).Picture = LoadPicture(DiretorioInicial + "\" + TabelaDeProgramasAtivos(NumeroDoPrograma, 3))
    imgBotao(NumeroDoPrograma).ToolTipText = Format(NumeroDoPrograma, "00") + " - " + TabelaDeProgramasAtivos(NumeroDoPrograma, 2)
    Kill DiretorioInicial + "\" + TabelaDeProgramasAtivos(NumeroDoPrograma, 3)
    If NumeroDoPrimeiroProgramaDaBarra <= NumeroDoPrograma And NumeroDoPrograma < NumeroDoPrimeiroProgramaDaBarra + NumeroMaximoDeProgramasNaBarra Then
       imgBotao(NumeroDoPrograma).Visible = True
       If Not imgProximaBarra.Visible And Not imgBarraAnterior.Visible Then
          If NumeroDoPrograma <= 5 Then
             Me.Width = Me.Width + Me.imgBotao(0).Width
             Me.Left = (Screen.Width - Me.Width) / 2
          Else
             Me.Width = Me.imgBotao(0).Width * 1.01 * (NumeroDoPrograma + 1)
             Me.Left = (Screen.Width - Me.Width) / 2
          End If
       End If
    Else
       imgBotao(NumeroDoPrograma).Visible = False
       
    End If
   'I = IIf(NumeroDeProgramasAtivos < 5, 5, NumeroDeProgramasAtivos)
   'Me.Width = (Me.imgBotao(0).Width * 1.01) * (IIf(I > NumeroMaximoDeProgramasNaBarra, NumeroMaximoDeProgramasNaBarra, I) + 1) + IIf(I > NumeroMaximoDeProgramasNaBarra, Me.imgBarraAnterior.Width, 0)
    If NumeroDoPrograma > NumeroMaximoDeProgramasNaBarra And (Not imgProximaBarra.Visible And Not imgBarraAnterior.Visible) Then
       'Debug.Print NumeroDoPrograma
       imgProximaBarra.Visible = True
       Me.Width = Me.Width + Me.imgProximaBarra.Width
       imgProximaBarra.Left = Me.Width - imgProximaBarra.Width '- 200
       imgBarraAnterior.Left = imgProximaBarra.Left
    End If
    ProximoBotaoDaBarra = ProximoBotaoDaBarra + 1
End Sub


Public Sub RecuperaStatusTOPEMaxPGMNaBarra()
Chave = PGMStatus
X = LeIni(Seção, Chave)
If X = TopoYES Then
   Call ColocaFormONTOP
   frmActive.mnutopo.Checked = True
Else
   Call ColocaFormNormal
   frmActive.mnutopo.Checked = False
End If
Chave = CTLMaximoPGMNaBarra
X = LeIni(Seção, Chave)
If X = "" Then
   NumeroMaximoDeProgramasNaBarra = MaximoComIconesMaiores
   Call GravaMaxPGMNaBarraNoINI
   frmActive.mnuBotoes.Checked = True
Else
   NumeroMaximoDeProgramasNaBarra = IIf(CInt(X) = MaximoComIconesMaiores, MaximoComIconesMaiores, MaximoComIconesMenores)
   Call GravaMaxPGMNaBarraNoINI
   frmActive.mnuBotoes.Checked = IIf(NumeroMaximoDeProgramasNaBarra = MaximoComIconesMenores, False, True)
End If
End Sub
Public Sub GravaMaxPGMNaBarraNoINI()
Chave = CTLMaximoPGMNaBarra
Valor = Format(NumeroMaximoDeProgramasNaBarra, "00")
X = EscreveIni(Seção, Chave, Valor)
End Sub
Public Sub EscondeProgramasAtuaisDaBarra()
Dim Limite As Integer
Limite = IIf(NumeroDoPrimeiroProgramaDaBarra + NumeroMaximoDeProgramasNaBarra >= NumeroDeProgramasAtivos, NumeroDoPrimeiroProgramaDaBarra + NumeroMaximoDeProgramasNaBarra - 1, NumeroDeProgramasAtivos)
For i = NumeroDoPrimeiroProgramaDaBarra To NumeroDoPrimeiroProgramaDaBarra + NumeroMaximoDeProgramasNaBarra - 1 Step 1
    imgBotao(i).Visible = False
Next i
End Sub
Public Sub MostraProgramasAnterioresDaBarra()
'Pressupoe que o NumeroDoPrimeiroProgramaDaBarra já foi atualizado
For i = NumeroDoPrimeiroProgramaDaBarra To NumeroDoPrimeiroProgramaDaBarra + NumeroMaximoDeProgramasNaBarra - 1 Step 1
    imgBotao(i).Visible = True
Next i
End Sub


Public Sub ColocaFormONTOP()
   'To set Form1 as a TopMost form, do the following:
   res = SetWindowPos(frmActive.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   'if res%=0, there is an error
End Sub
Public Sub ColocaFormNormal()
   'To turn off topmost (make the form act normal again):
   res = SetWindowPos(frmActive.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Public Sub ColocaIconeNaBandeja()
    nid.cbSize = Len(nid)
    nid.hWnd = picNotifier.hWnd
    ' Change the uId to 1&
    nid.uId = 1&
    ' Use the respective Flags that should be used,
    ' so it works properly, just like any other.
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    ' When there's WM_MOUSEMOVE, we'll need to see
    ' if there's mouse clicking, or whatever.
    nid.uCallBackMessage = WM_MOUSEMOVE
    ' Use the Main Form's Icon for the process.
    nid.hIcon = Me.Icon
    ' Use the Tip "Visual Basic Island" as our tooltip
    nid.szTip = "Programa Barra Ativa " + PGMVersao & vbNullChar
    ' Now, we actually add nid to the Taskbar.
    Shell_NotifyIcon NIM_ADD, nid

End Sub

Private Sub picNotifier_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static Rec As Boolean, Msg As Long
    ' Msg is the current X divided by the Screen's
    ' X in TwipsPerPixel Measurement's, so it's the
    ' same as the picNotifier.
    Msg = X / Screen.TwipsPerPixelX
    
    ' If Rec is False
    If Rec = False Then
        ' Make Rec True.
        Rec = True
        ' Determine what Msg really was:
        Select Case Msg
            ' If DoubleClick
            Case WM_LBUTTONDBLCLK:
            ' If Button is Down
            Case WM_LBUTTONDOWN:
            
            'If Button is Up
            Case WM_LBUTTONUP:
            
            'If the RightButton is clicked
            Case WM_RBUTTONDBLCLK:
            
            'If the RightBurron is Down
            Case WM_RBUTTONDOWN:
            
            'If RightButton is Up
            Case WM_RBUTTONUP:
                PopupMenu Pop01
        'End Determination
        End Select
        
        'Change Rec Back to False.
        Rec = False
    End If


End Sub

