VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmMenu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2205
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   737
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraMenu 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   11055
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   9360
         OleObjectBlob   =   "FrmMenu.frx":0000
         Top             =   1200
      End
      Begin VB.Image Image10 
         Height          =   240
         Left            =   7200
         Picture         =   "FrmMenu.frx":4A505
         Top             =   1200
         Width           =   195
      End
      Begin VB.Label LblManutJogos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Jogos"
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
         Height          =   255
         Left            =   7440
         MouseIcon       =   "FrmMenu.frx":4A7C7
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   1200
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   240
         Left            =   5400
         Picture         =   "FrmMenu.frx":4B091
         Top             =   480
         Width           =   240
      End
      Begin VB.Label LblMaqSit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Situação"
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
         Height          =   255
         Left            =   5640
         MouseIcon       =   "FrmMenu.frx":4B3D3
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Sobre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   10200
         TabIndex        =   27
         Top             =   120
         Width           =   855
      End
      Begin VB.Shape Shape15 
         Height          =   255
         Left            =   10200
         Top             =   120
         Width           =   855
      End
      Begin VB.Label LblSistema 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sistema"
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
         Height          =   255
         Left            =   10320
         MouseIcon       =   "FrmMenu.frx":4BC9D
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   480
         Width           =   675
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Acesso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Cartão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   23
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Caixa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8880
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Manutenção"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7080
         TabIndex        =   11
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Máquina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5280
         TabIndex        =   15
         Top             =   120
         Width           =   1815
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   1440
         Picture         =   "FrmMenu.frx":4C567
         Top             =   480
         Width           =   240
      End
      Begin VB.Image Image5 
         Height          =   225
         Left            =   1440
         Picture         =   "FrmMenu.frx":4C8A9
         Top             =   840
         Width           =   240
      End
      Begin VB.Shape Shape3 
         Height          =   255
         Left            =   1320
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label LblAcesAlt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Alterar"
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
         Height          =   255
         Left            =   1680
         MouseIcon       =   "FrmMenu.frx":4CBBB
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   840
         Width           =   735
      End
      Begin VB.Label LblAcesInc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Novo"
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
         Height          =   255
         Left            =   1680
         MouseIcon       =   "FrmMenu.frx":4D485
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   480
         Width           =   855
      End
      Begin VB.Label LblCartCanc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Cancelar"
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
         Height          =   255
         Left            =   3000
         MouseIcon       =   "FrmMenu.frx":4DD4F
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label LblCartCons 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Consultar"
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
         Height          =   255
         Left            =   3000
         MouseIcon       =   "FrmMenu.frx":4E619
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   840
         Width           =   900
      End
      Begin VB.Label LblCartInc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Novo"
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
         Height          =   255
         Left            =   3000
         MouseIcon       =   "FrmMenu.frx":4EEE3
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00808080&
         Height          =   255
         Left            =   0
         Top             =   120
         Width           =   1335
      End
      Begin VB.Shape Shape5 
         Height          =   255
         Left            =   2640
         Top             =   120
         Width           =   1335
      End
      Begin VB.Image Image9 
         Height          =   240
         Left            =   2760
         Picture         =   "FrmMenu.frx":4F7AD
         Top             =   1200
         Width           =   210
      End
      Begin VB.Image Image8 
         Height          =   240
         Left            =   2760
         Picture         =   "FrmMenu.frx":4FAAF
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Image7 
         Height          =   240
         Left            =   2760
         Picture         =   "FrmMenu.frx":4FDF1
         Top             =   480
         Width           =   240
      End
      Begin VB.Label LblCredArq 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Arquivo"
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
         Height          =   255
         Left            =   4320
         MouseIcon       =   "FrmMenu.frx":50133
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label LblCredCons 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Consultar"
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
         Height          =   255
         Left            =   4320
         MouseIcon       =   "FrmMenu.frx":509FD
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   840
         Width           =   900
      End
      Begin VB.Label LblCredInc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Novo"
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
         Height          =   255
         Left            =   4320
         MouseIcon       =   "FrmMenu.frx":512C7
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
      Begin VB.Shape Shape7 
         Height          =   255
         Left            =   3960
         Top             =   120
         Width           =   1335
      End
      Begin VB.Image Image13 
         Height          =   240
         Left            =   4080
         Picture         =   "FrmMenu.frx":51B91
         Top             =   1200
         Width           =   210
      End
      Begin VB.Image Image12 
         Height          =   240
         Left            =   4080
         Picture         =   "FrmMenu.frx":51E93
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Image11 
         Height          =   240
         Left            =   4080
         Picture         =   "FrmMenu.frx":521D5
         Top             =   480
         Width           =   240
      End
      Begin VB.Label LblMaqArq 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Arquivo de uso"
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
         Height          =   255
         Left            =   5640
         MouseIcon       =   "FrmMenu.frx":52517
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   1560
         Width           =   1350
      End
      Begin VB.Label LblMaqCon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Conectados"
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
         Height          =   255
         Left            =   5640
         MouseIcon       =   "FrmMenu.frx":52DE1
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label LblMaqVis 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Visualização"
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
         Height          =   255
         Left            =   5640
         MouseIcon       =   "FrmMenu.frx":536AB
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.Image Image16 
         Height          =   240
         Left            =   5400
         Picture         =   "FrmMenu.frx":53F75
         Top             =   1560
         Width           =   210
      End
      Begin VB.Image Image15 
         Height          =   240
         Left            =   5400
         Picture         =   "FrmMenu.frx":54277
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image Image14 
         Height          =   240
         Left            =   5400
         Picture         =   "FrmMenu.frx":545B9
         Top             =   840
         Width           =   240
      End
      Begin VB.Shape Shape9 
         Height          =   255
         Left            =   5280
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label LblManutNiver 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Aniversariantes"
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
         Height          =   255
         Left            =   7440
         MouseIcon       =   "FrmMenu.frx":548FB
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1560
         Width           =   1395
      End
      Begin VB.Label LblManutPre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Preços"
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
         Height          =   255
         Left            =   7440
         MouseIcon       =   "FrmMenu.frx":551C5
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.Label LblManutSen 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Senhas"
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
         Height          =   255
         Left            =   7440
         MouseIcon       =   "FrmMenu.frx":55A8F
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.Image Image19 
         Height          =   195
         Left            =   7200
         Picture         =   "FrmMenu.frx":56359
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Image18 
         Height          =   240
         Left            =   7200
         Picture         =   "FrmMenu.frx":5660B
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Image17 
         Height          =   240
         Left            =   7200
         Picture         =   "FrmMenu.frx":5694D
         Top             =   480
         Width           =   195
      End
      Begin VB.Shape Shape11 
         Height          =   255
         Left            =   7080
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label LblCxCons 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Consultar"
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
         Height          =   255
         Left            =   9240
         MouseIcon       =   "FrmMenu.frx":56C0F
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   840
         Width           =   900
      End
      Begin VB.Label LblCxInc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Novo"
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
         Height          =   255
         Left            =   9240
         MouseIcon       =   "FrmMenu.frx":574D9
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.Image Image21 
         Height          =   240
         Left            =   9000
         Picture         =   "FrmMenu.frx":57DA3
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Image20 
         Height          =   240
         Left            =   9000
         Picture         =   "FrmMenu.frx":580E5
         Top             =   480
         Width           =   240
      End
      Begin VB.Shape Shape13 
         Height          =   255
         Left            =   8880
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label LblCliCons 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Consultar"
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
         Height          =   255
         Left            =   360
         MouseIcon       =   "FrmMenu.frx":58427
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label LblCliAlt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Alterar"
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
         Height          =   255
         Left            =   360
         MouseIcon       =   "FrmMenu.frx":58CF1
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.Label LblCliInc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Novo"
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
         Height          =   255
         Left            =   360
         MouseIcon       =   "FrmMenu.frx":595BB
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   120
         Picture         =   "FrmMenu.frx":59E85
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "FrmMenu.frx":5A1C7
         Top             =   480
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   225
         Left            =   120
         Picture         =   "FrmMenu.frx":5A509
         Top             =   840
         Width           =   240
      End
      Begin VB.Shape Shape14 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   1575
         Left            =   8880
         Top             =   360
         Width           =   1335
      End
      Begin VB.Shape Shape12 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   1575
         Left            =   7080
         Top             =   360
         Width           =   1815
      End
      Begin VB.Shape Shape10 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   1575
         Left            =   5280
         Top             =   360
         Width           =   1815
      End
      Begin VB.Shape Shape8 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   1575
         Left            =   3960
         Top             =   360
         Width           =   1335
      End
      Begin VB.Shape Shape6 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   1575
         Left            =   2640
         Top             =   360
         Width           =   1335
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   1575
         Left            =   1320
         Top             =   360
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   1575
         Left            =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.Shape Shape16 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   1575
         Left            =   10200
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  Height = 1845
  Top = 20
  Width = 11085
  
  Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
  Skin1.ApplySkin (FrmMenu.hwnd)
  Skin1.RemoveSkin (FraMenu.hwnd)
  
End Sub

Private Sub Form_Resize()
  Me.Left = (MDIPrincipal.Width / 2) - (Me.Width / 1.93)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbHourglass
    FrmMenuVis.Show
End Sub

Private Sub LblAcesAlt_Click()
    Screen.MousePointer = vbHourglass
    VGStrAlt = "acesso"
    FrmCadCliAlt.Show
End Sub

Private Sub LblAcesInc_Click()
    Screen.MousePointer = vbHourglass
    FrmAcesso.Show
End Sub

Private Sub LblCartCanc_Click()
    Screen.MousePointer = vbHourglass
    FrmCancel.Show
End Sub

Private Sub LblCartCons_Click()
    Screen.MousePointer = vbHourglass
    FrmConsultCart.Show
End Sub

'Private Sub LblCartImp_Click()
'    FrmCredito.Show
'End Sub

Private Sub LblCartInc_Click()
    Screen.MousePointer = vbHourglass
    FrmCartao.Show
End Sub

Private Sub LblCliAlt_Click()
    Screen.MousePointer = vbHourglass
    VGStrAlt = "cliente"
    FrmCadCliAlt.Show
End Sub

Private Sub LblCliCons_Click()
    Screen.MousePointer = vbHourglass
    FrmConsultCli.Show
End Sub

Private Sub LblCliInc_Click()
    Screen.MousePointer = vbHourglass
    FrmCadCli.Show
End Sub

Private Sub LblCredArq_Click()
    Screen.MousePointer = vbHourglass
    FrmConsultCred.Show
End Sub

Private Sub LblCredCons_Click()
    Screen.MousePointer = vbHourglass
    FrmConsultCart.Show
End Sub

Private Sub LblCredInc_Click()
    Screen.MousePointer = vbHourglass
    FrmCredito.Show
End Sub

Private Sub LblCxCons_Click()
    Screen.MousePointer = vbHourglass
    FrmConsCaixa.Show
End Sub

Private Sub LblCxInc_Click()
    Screen.MousePointer = vbHourglass
    FrmCodProd.Show
End Sub

Private Sub LblManutJogos_Click()
    Screen.MousePointer = vbHourglass
    FrmJogos.Show
End Sub

Private Sub LblManutNiver_Click()
    Screen.MousePointer = vbHourglass
    FrmNiver.Show
End Sub

Private Sub LblManutPre_Click()
    Screen.MousePointer = vbHourglass
    FrmPreco.Show
End Sub

Private Sub LblManutSen_Click()
    Screen.MousePointer = vbHourglass
    FrmSenha.Show
End Sub

Private Sub LblMaqArq_Click()
    Screen.MousePointer = vbHourglass
    FrmMaqCli.Show
End Sub

Private Sub LblMaqCon_Click()
    Screen.MousePointer = vbHourglass
    FrmConect.Show
End Sub

Private Sub LblMaqSit_Click()
    Screen.MousePointer = vbHourglass
    FrmMaqSituacao.Show
End Sub

Private Sub LblMaqVis_Click()
    Screen.MousePointer = vbHourglass
    FrmMaquina.Show
End Sub

Private Sub LblSistema_Click()
    Screen.MousePointer = vbHourglass
    frmSistema.Show
End Sub
