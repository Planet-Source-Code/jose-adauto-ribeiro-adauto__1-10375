Attribute VB_Name = "Module1"
Option Explicit
 Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
     Global Const SWP_NOMOVE = 2
     Global Const SWP_NOSIZE = 1
     Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
     Global Const HWND_TOPMOST = -1
     Global Const HWND_NOTOPMOST = -2

  Public Const INVALID_HANDLE_VALUE = -1
  Public Const MAX_PATH = 260

  Type FILETIME
     dwLowDateTime As Long
     dwHighDateTime As Long
  End Type

  Type WIN32_FIND_DATA
     dwFileAttributes As Long
     ftCreationTime As FILETIME
     ftLastAccessTime As FILETIME
     ftLastWriteTime As FILETIME
     nFileSizeHigh As Long
     nFileSizeLow As Long
     dwReserved0 As Long
     dwReserved1 As Long
     cFileName As String * MAX_PATH
     cAlternate As String * 14
  End Type
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function ExtractIcon Lib "shell32" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

Global NumeroDeProgramasAtivos As Integer, ArquivoDePrograma As String, NextPGM As Integer
Global DiretorioInicial As String, UltimoProgramaAtivo As Integer, ProximoBotaoDaBarra As Integer
Global Active_INI As String, Chave As String, Seção As String, Valor As String
Global i As Integer, J As Integer, K As Integer, L As Integer, X As String, CrLf As String
Global TabelaDeProgramasAtivos(1 To 44, 1 To 6) As String, NumeroDoPrimeiroProgramaDaBarra As Integer
Global TitleAnterior As String, NumeroDoItem As Integer, AlterouLista As Boolean, InicializarBarra As Boolean
Global NumeroMaximoDeProgramasNaBarra As Integer ', NumeroMaximoDeProgramasAtivos As Integer
Global Const MaximoComIconesMenores = 20, MaximoComIconesMaiores = 14
Global Const NumeroMaximoDeProgramasAtivos = 44
Global Const Posicionar = "P", Alterar = "A", Excluir = "E", Desistir = "D", AlterarÍcone = "I"
Global Funcao As String, LinhaAnterior As Integer, PGMVersao As String
Global Const MB_OK = 0, MB_OKCANCEL = 1      ' Define buttons.
Global Const MB_YESNOCANCEL = 3, MB_YESNO = 4
Global Const MB_ICONSTOP = 16, MB_ICONquestion = 32            ' Define Icons.
Global Const MB_ICONEXCLAMATION = 48, MB_ICONINFORMATION = 64
Global Const MB_DEFBUTTON2 = 256, IDYES = 6, IDNO = 7          ' Define other.
Global Const PGMName = "PGM_Name_", PGMTitle = "PGM_Title_", PGMbmp = "PGM_bmp_", PGMParm = "PGM-Parm_"
Global Const CTLNumPGM = "Numero_De_Programas_Ativos", PGMStatus = "PGMStatus_Topo", PGMIcon = "PGM_Icone_"
Global Const CTLPrimeiroPrograma = "PrimeiroProgramaDaBarra", CTLProximoPrograma = "Proximo_Programa_"
Global Const CTLMaximoPGMNaBarra = "MaximoDeProgramasNaBarra", TopoYES = "Sim", TopoNO = "Não"
Global DgDef, Msg, Response, Title      ' Declare variables.
' ------------ Declarações necesárias para a função de extrair icon de um módulo ---------

Private Type PicBmp
  Size As Long
  tType As Long   '* 245 - 2 100 + 75 5 90+4*45
  hBmp As Long
  hPal As Long
  Reserved As Long
End Type

Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

'Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
'Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
'Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
'---
Public Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long


' ------------ Final das declarações necesárias para a função de extrair icon de um módulo ---------
Public Sub RegerarActiveINI()
Kill Active_INI
Chave = CTLNumPGM
Valor = Format(NumeroDeProgramasAtivos, "00")
X = EscreveIni(Seção, Chave, Valor)
Call SalvarStatusTOP
For i = 1 To NumeroDeProgramasAtivos
    Chave = PGMName + Format(i, "00")
    Valor = TabelaDeProgramasAtivos(i, 1)
    X = EscreveIni(Seção, Chave, Valor)
    Chave = PGMTitle + Format(i, "00")
    Valor = TabelaDeProgramasAtivos(i, 2)
    X = EscreveIni(Seção, Chave, Valor)
    Chave = Left(TabelaDeProgramasAtivos(i, 3), Len(TabelaDeProgramasAtivos(i, 3)) - 4)
    Valor = Format(i, "00") + Valor
    X = EscreveIni(Seção, Chave, Valor)
    Chave = PGMbmp + Format(i, "00")
    Valor = TabelaDeProgramasAtivos(i, 3)
    X = EscreveIni(Seção, Chave, Valor)
    Chave = PGMParm + Format(i, "00")
    Valor = TabelaDeProgramasAtivos(i, 4)
    X = EscreveIni(Seção, Chave, Valor)
Next i

End Sub
Public Sub SalvarStatusTOP()
Chave = PGMStatus
Valor = IIf(frmActive.mnutopo.Checked, TopoYES, TopoNO)
X = EscreveIni(Seção, Chave, Valor)
End Sub

Public Sub TextSelected()
Dim i As Integer
Dim oMyTextBox As Object

Set oMyTextBox = Screen.ActiveControl
If TypeName(oMyTextBox) = "TextBox" Then
i = Len(oMyTextBox.Text)

oMyTextBox.SelStart = 0
oMyTextBox.SelLength = i
End If
End Sub
'----------------------------------------------------------
' Check for the existence of a file by attempting an OPEN.
'----------------------------------------------------------
Public Function FileExists(sSource As String) As Boolean

   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   
   hFile = FindFirstFile(sSource, WFD)
   FileExists = hFile <> INVALID_HANDLE_VALUE
   
   Call FindClose(hFile)
   
End Function



Public Function EscreveIni(Seção As String, Chave As String, Valor As String)
  
  ' Escreve Valor na seção e chave indicadas
  Dim lpFileName As String, X  As Integer
  lpFileName = Active_INI

  X = WritePrivateProfileString(Seção, Chave, Valor, lpFileName)
  
End Function
Public Function LeIni(Seção As String, Chave As String)
  ' Lê o valor na seção e chave indicadas

  Dim lpAppName As String, lpKeyName As String
  Dim lpDefault As String, lpReturnedString As String
  Dim nSize As Integer, lpFileName As String, X As Integer

  lpFileName = Active_INI
  lpAppName = Seção
  lpKeyName = Chave
  lpDefault = ""
  lpReturnedString = Space$(512)
  nSize = 512

  X = GetPrivateProfileString(lpAppName, lpKeyName, lpDefault, lpReturnedString, nSize, lpFileName)
  LeIni = Left$(lpReturnedString, X)
End Function
Public Sub TornaBrancaACelula(N)
'Torna a coluna 01 da linha N da Lista de programas branco novamente ...
      frmManutencao.flxLista.ColSel = 1
      frmManutencao.flxLista.Row = N
      frmManutencao.flxLista.CellBackColor = RGB(255, 255, 255)
End Sub
