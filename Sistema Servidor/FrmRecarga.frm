VERSION 5.00
Begin VB.Form FrmRecarga 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recarga de Cartão"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "FrmRecarga.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraRecarga 
      BackColor       =   &H00C0E0FF&
      Height          =   2535
      Left            =   2040
      TabIndex        =   5
      Top             =   480
      Width           =   4095
      Begin VB.TextBox TxtDtRec 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         ToolTipText     =   "Data da recarga"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox TxtCodCli 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         ToolTipText     =   "Código de cadastro do cliente"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox TxtTempRec 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         ToolTipText     =   "Tempo de recarga"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton CmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox TxtNumCartao 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         ToolTipText     =   "Número do cartão"
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label LblDtRec 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Data da Recarga:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label LblCodCli 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cód. Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label LblTempRec 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Tempo de Recarga:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1590
      End
      Begin VB.Label LblNumCartao 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Nº do Cartão:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1050
      End
   End
   Begin VB.Image ImgGuardiao 
      Height          =   3465
      Left            =   240
      Picture         =   "FrmRecarga.frx":08CA
      Top             =   120
      Width           =   1545
   End
End
Attribute VB_Name = "FrmRecarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdIncluir_Click()
    VGStrBox = MsgBox("Recarga efetuada.", vbInformation, "Informação")
    Unload Me
End Sub

Private Sub Form_Load()
    Height = 3825
    Width = 6375
    Top = 1605
    Left = 2985
    
    Unload FrmAcesso
    
    If VGStrForm = "Acesso" Then
        TxtCodCli.Text = FormataNum(VGIntCodCli)
        VGStrForm = ""
    End If
    
    TxtDtRec.Text = FormataDataUS(Date)
    
End Sub

