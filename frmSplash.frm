VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      ToolTipText     =   "Para sair, tecle ESC ou clique em qualquer parte deste painel ..."
      Top             =   60
      Width           =   7080
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         ToolTipText     =   "Para sair, tecle ESC ou clique em qualquer parte deste painel ..."
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblDataLiberação 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         ToolTipText     =   "Para sair, tecle ESC ou clique em qualquer parte deste painel ..."
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "Company Adauto Software House"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         ToolTipText     =   "Para sair, tecle ESC ou clique em qualquer parte deste painel ..."
         Top             =   3270
         Width           =   2535
      End
      Begin VB.Label lblWarning 
         Caption         =   "Direitos reservados para uso no Brasil de 2000 a 2004."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   2
         ToolTipText     =   "Para sair, tecle ESC ou clique em qualquer parte deste painel ..."
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version  "
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
         Left            =   5850
         TabIndex        =   5
         ToolTipText     =   "Para sair, tecle ESC ou clique em qualquer parte deste painel ..."
         Top             =   2700
         Width           =   1005
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Windows 95/98 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4485
         TabIndex        =   6
         ToolTipText     =   "Para sair, tecle ESC ou clique em qualquer parte deste painel ..."
         Top             =   2340
         Width           =   2370
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Active Bar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2400
         TabIndex        =   8
         ToolTipText     =   "Para sair, tecle ESC ou clique em qualquer parte deste painel ..."
         Top             =   1140
         Width           =   3180
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "LicenseTo José ADAUTO Ribeiro         "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Adauto Software House"
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
         Left            =   2355
         TabIndex        =   7
         ToolTipText     =   "Para sair, tecle ESC ou clique em qualquer parte deste painel ..."
         Top             =   705
         Width           =   4065
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = PGMVersao
    lblProductName.Caption = App.Title
    lblDataLiberação = "Liberado em: " + Format(FileDateTime(App.Path + "\" + App.EXEName + ".exe"), "dd / mm / yyyy")
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub imgLogo_Click()
    Unload Me

End Sub

Private Sub lblCompany_Click()
    Unload Me

End Sub

Private Sub lblCompanyProduct_Click()
    Unload Me

End Sub

Private Sub lblDataLiberação_Click()
    Unload Me

End Sub

Private Sub lblLicenseTo_Click()
    Unload Me

End Sub

Private Sub lblPlatform_Click()
    Unload Me

End Sub

Private Sub lblProductName_Click()
    Unload Me

End Sub

Private Sub lblVersion_Click()
    Unload Me

End Sub

Private Sub lblWarning_Click()
    Unload Me

End Sub


