VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmAltAces 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Alteração Cadastro de Acessos"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   ClipControls    =   0   'False
   Icon            =   "FrmAltAces.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3600
   Begin VB.Frame FraAcesso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3135
      Begin VB.TextBox TxtSenhaNova 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         ToolTipText     =   "Senha nova"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox TxtCodCli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Código do cliente"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtSenhaAtual 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "Senha atual"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton CmdAlterar 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   960
         TabIndex        =   4
         ToolTipText     =   "Altera acesso no sistema"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox TxtLogin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         ToolTipText     =   "Login"
         Top             =   720
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCodCli 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmAltAces.frx":000C
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   240
         OleObjectBlob   =   "FrmAltAces.frx":0082
         Top             =   1800
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblLogin 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmAltAces.frx":02B6
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblSenhaAtual 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmAltAces.frx":0320
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblSenhaNova 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmAltAces.frx":0396
         TabIndex        =   9
         Top             =   1440
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmAltAces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MsgLogin As String
Public MsgSenhaAtual As String
Public MsgSenhaNova As String
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdAlterar_Click()
    Screen.MousePointer = vbHourglass
    
    If TxtLogin.Text = "" Or TxtSenhaAtual.Text = "" Or TxtSenhaNova.Text = "" Then
        VPStrBox = MsgBox("Preencha os campos em branco", vbCritical, "Guide System - Aviso de erro")
        TxtSenhaAtual.SetFocus
    Else
        Conecta
        
        Dim RecPesq As New ADODB.Recordset
        Dim RecVerif As New ADODB.Recordset
        Dim RecLog As New ADODB.Recordset
                           
        StrSql = "Select Login,Senha from tb_acesso where CodCli=" & VGIntCodCliTemp & " and Senha='" & TxtSenhaAtual.Text & "'"
        RecPesq.Open StrSql, vgCon, 1, 3
               
        If RecPesq.EOF Then     'não achou registro
            VPStrBox = MsgBox("Senha Atual não confere.", vbCritical, "Guide System - Aviso de erro")
            
            TxtSenhaAtual.Text = ""
            TxtSenhaNova.Text = ""
            TxtSenhaAtual.SetFocus
        Else
            
            StrSql = "Select Login from tb_acesso where Login='" & TxtLogin.Text & "' and CodCli <>" & VGIntCodCliTemp
            RecLog.Open StrSql, vgCon, 1, 3
            
            If Not RecLog.EOF Then      'achou registro
                
                VPStrBox = MsgBox("Este Login já existe.", vbCritical, "Guide System - Aviso de erro")
                TxtLogin.SetFocus
            Else
                
                StrSql = "Select * from tb_acesso where CodCli=" & VGIntCodCliTemp
                RecVerif.Open StrSql, vgCon, 1, 3
                
                RecVerif("Login") = TxtLogin.Text
                RecVerif("Senha") = TxtSenhaNova.Text
                RecVerif.Update
                
                VPStrBox = MsgBox("Acesso modificado.", vbInformation, "Guide System - Informação")
                VGIntCodCliTemp = 0
                
                Unload Me
            
            End If
        End If
        
        Desconecta
    End If
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmAltAces.hwnd)
    
    Height = 2970
    Width = 3720
    'Top = 2160
    'Left = 4155
    
    Unload FrmCadCli
    Unload FrmCadCliAlt
    
    Conecta

    Dim RecAces As New ADODB.Recordset

    StrSql = "Select Login,Senha from tb_acesso where CodCli=" & VGIntCodCliTemp
    RecAces.Open StrSql, vgCon, 1, 3

    TxtCodCli.Text = FormataNum(VGIntCodCliTemp)
    TxtLogin.Text = RecAces.Fields.Item(0).Value

    Desconecta
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  FrmAltAces.Left = (MDIPrincipal.Width / 2) - (FrmAltAces.Width / 1.93)
  FrmAltAces.Top = (MDIPrincipal.Height / 3) - (FrmAltAces.Height / 5)
End Sub

Private Sub TxtCodCli_GotFocus()
    TxtCodCli.SelStart = 0
    TxtCodCli.SelLength = Len(TxtCodCli.Text)
End Sub

Private Sub TxtCodCli_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtLogin_GotFocus()
    TxtLogin.SelStart = 0
    TxtLogin.SelLength = Len(TxtLogin.Text)
End Sub

Private Sub TxtLogin_KeyPress(KeyAscii As Integer)
    
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

Private Sub TxtSenhaAtual_GotFocus()
    TxtSenhaAtual.SelStart = 0
    TxtSenhaAtual.SelLength = Len(TxtSenhaAtual.Text)
End Sub

Private Sub TxtSenhaAtual_KeyPress(KeyAscii As Integer)
    
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

Private Sub TxtSenhaNova_GotFocus()
    TxtSenhaNova.SelStart = 0
    TxtSenhaNova.SelLength = Len(TxtSenhaNova.Text)
End Sub

Private Sub TxtSenhaNova_KeyPress(KeyAscii As Integer)
    
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
