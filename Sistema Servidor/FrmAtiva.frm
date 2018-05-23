VERSION 5.00
Begin VB.Form FrmAtiva 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ativação de cartão (Descancelamento)"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "FrmAtiva.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      ToolTipText     =   "Responsável pelo cancelamento"
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      ToolTipText     =   "Responsável pelo cancelamento"
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton CmdVoltar 
      Caption         =   "Voltar"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox TxtMotivo 
      Height          =   885
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Motivo do cancelamento"
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox TxtResp 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      ToolTipText     =   "Responsável pelo cancelamento"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label LblMotivo 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Motivo:"
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
      TabIndex        =   5
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label LblResp 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Responsável:"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   1110
   End
   Begin VB.Label LblDtCancel 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Data Cancel.:"
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
      TabIndex        =   2
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label LblNumCart 
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
      TabIndex        =   1
      Top             =   240
      Width           =   1050
   End
End
Attribute VB_Name = "FrmAtiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public MsgResp As String
Public MsgMotivo As String

Private Sub CmdCancel_Click()

    If TxtResp.Text <> "" And TxtMotivo.Text <> "" Then
        'gravar informações de cancelamento no banco
        
        Conecta
        
        Dim RecCart As ADODB.Recordset
                           
        StrSql = "Update tb_cartao set Cancelado=1,Motivo='" & TxtMotivo.Text & "',Resp='" & TxtResp.Text & "',DtCancel='" & FormataDataUS(Date) & "' where NumCartao=" & LblResNumCart
        Set RecCart = vgCon.Execute(StrSql)
        
        VPStrBox = MsgBox("Cancelamento efetuado.", vbInformation, "Informação")
                
        Desconecta
        
        Unload Me
        Unload FrmCancel
        
    Else    'campos em branco
        
        If TxtResp.Text = "" Then
            MsgResp = " - Responsável" & Chr(13)
        End If

        If TxtMotivo.Text = "" Then
            MsgMotivo = " - Motivo" & Chr(13)
        End If
        
        VPStrBox = MsgBox("Campo(s) não pode(m) estar em branco:" & Chr(13) & Chr(13) & MsgResp & MsgMotivo, vbCritical, "Aviso de erro")
        
        MsgResp = ""
        MsgMotivo = ""
        TxtResp.SetFocus
    End If

End Sub

Private Sub CmdVoltar_Click()
    Unload Me
    FrmCancel.Enabled = True
End Sub

Private Sub Form_Load()
    Height = 3480
    Width = 4725
    Top = 2025
    Left = 4725
    
    FrmCancel.Enabled = False
    
    LblResNumCart.Caption = FrmCancel.TxtNumCart.Text
    LblResDtCancel.Caption = FormataData(Date)
End Sub
