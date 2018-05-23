VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmSenha 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liberação de acesso"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3345
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FrmSenha.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraAcesso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton CmdOk 
         Caption         =   "Ok"
         Height          =   375
         Left            =   600
         TabIndex        =   1
         ToolTipText     =   "Confirma senha"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox TxtSenha 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   0
         ToolTipText     =   "Senha"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         ToolTipText     =   "Sai da opção"
         Top             =   720
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   120
         OleObjectBlob   =   "FrmSenha.frx":000C
         Top             =   600
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblSenha 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmSenha.frx":0240
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MsgSenha As String
Public VPStrBox As String

Private Sub CmdOK_Click()
    Screen.MousePointer = vbHourglass
    
    If TxtSenha.Text = "" Then
        VPStrBox = MsgBox("Preencha a senha", vbCritical, "Guide System - Aviso de erro")
        TxtSenha.SetFocus
    Else
        Conecta
        
        Dim RecSenha As New ADODB.Recordset
                           
        StrSql = "Select Senha,Tipo from tb_controle where Senha='" & TxtSenha.Text & "'"
        RecSenha.Open StrSql, vgCon, 1, 3
               
        If RecSenha.EOF Then    'não achou nada
            VPStrBox = MsgBox("Senha inexistente.", vbCritical, "Guide System - Aviso de erro")
            TxtSenha.SetFocus
        Else
            If VGStrLocalTemp = "conscaixa" Then
            
                If RecSenha.Fields.Item(1).Value = "caixa" Then
                    'liberado para fazer a modificação
                    VGStrSenha = "sim"
                    Unload Me
                    
                    FrmConsCaixa.Enabled = True
                    VGStrLocalTemp = ""
                    FrmConsCaixa.Exclui_Item
                Else
                    VPStrBox = MsgBox("Senha inexistente.", vbCritical, "Guide System - Aviso de erro")
                End If
            
            ElseIf VGStrLocalTemp = "senhaadm" Then
                
                If RecSenha.Fields.Item(1).Value = "senhaadm" Then
                    'liberado para fazer a modificação
                    VGStrSenha = "sim"
                    Unload Me
                    
                    'FrmSenhaSistem.Enabled = True
                    VGStrLocalTemp = ""
                    FrmSenhaSistem.Show
                Else
                    VPStrBox = MsgBox("Senha inexistente.", vbCritical, "Guide System - Aviso de erro")
                End If
            
            End If
            
        End If
       
    End If
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdSair_Click()
    Unload Me
    
    If VGStrLocalTemp = "caixa" Then
        VGIntCodItem = 0
        FrmConsCaixa.Enabled = True
    End If

End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmSenha.hwnd)
    
    Height = 1785
    Width = 3705
    
''    If VGStrLocalTemp = "conscaixa" Then
''        FrmConsCaixa.Enabled = False
''        Top = 1275
''        Left = 3795
''    ElseIf VGStrLocalTemp = "senhaadm" Then
        'Top = 1275
        'Left = 3695

''    End If

    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  FrmSenha.Left = (MDIPrincipal.Width / 2) - (FrmSenha.Width / 1.93)
  FrmSenha.Top = (MDIPrincipal.Height / 3) - (FrmSenha.Height / 5)
End Sub

Private Sub TxtSenha_GotFocus()
    TxtSenha.SelStart = 0
    TxtSenha.SelLength = Len(TxtSenha.Text)
End Sub

Private Sub TxtSenha_KeyPress(KeyAscii As Integer)
    
    '============ Símbolos ======================
    If KeyAscii >= 33 And KeyAscii <= 47 Then
        KeyAscii = 0
    
    ElseIf KeyAscii >= 58 And KeyAscii <= 64 Then
        KeyAscii = 0

    ElseIf KeyAscii >= 91 And KeyAscii <= 96 Then
        KeyAscii = 0
    
    ElseIf KeyAscii >= 123 And KeyAscii <= 126 Then
        KeyAscii = 0
    
    End If
    '=========================================
    
    '======= Combinações de teclas com CTRL ========
    If KeyAscii >= 1 And KeyAscii <= 7 Then
        KeyAscii = 0
    
    ElseIf KeyAscii >= 9 And KeyAscii <= 29 Then
        KeyAscii = 0

    ElseIf KeyAscii = 127 Then
        KeyAscii = 0
    
    End If
    '=========================================

End Sub

