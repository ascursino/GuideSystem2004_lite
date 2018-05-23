VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmAcesso 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Acessos"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   Icon            =   "FrmAcesso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraAcesso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3135
      Begin VB.TextBox TxtSenha2 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         ToolTipText     =   "Confirmação de senha"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox TxtCodCli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         ToolTipText     =   "Código do cliente"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtSenha 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "Senha"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton CmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   960
         TabIndex        =   4
         ToolTipText     =   "Inclui acesso no sistema"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox TxtLogin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         MaxLength       =   10
         TabIndex        =   1
         ToolTipText     =   "Login "
         Top             =   720
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCodCli 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmAcesso.frx":000C
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   360
         OleObjectBlob   =   "FrmAcesso.frx":0082
         Top             =   1800
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblLogin 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmAcesso.frx":02B6
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblSenha 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmAcesso.frx":0320
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblSenha2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmAcesso.frx":038A
         TabIndex        =   6
         Top             =   1440
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MsgLogin As String
Public MsgSenha As String
Public MsgSenha2 As String
Public MsgCodCli As String
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdIncluir_Click()
    Screen.MousePointer = vbHourglass
    
    If TxtCodCli.Text = "" Or TxtLogin.Text = "" Or TxtSenha.Text = "" Or TxtSenha2.Text = "" Then
        VPStrBox = MsgBox("Preencha os campos em branco", vbCritical, "Guide System - Aviso de erro")
        TxtLogin.SetFocus
    
    ElseIf TxtSenha.Text <> TxtSenha2.Text Then
        VPStrBox = MsgBox("Senhas não conferem.", vbCritical, "Guide System - Aviso de erro")
        TxtSenha.Text = ""
        TxtSenha2.Text = ""
        TxtSenha.SetFocus
    
    Else
        'If VGIntCodCli = 0 Then
            VGIntCodCli = TxtCodCli.Text
        'End If
        
        Conecta
        
        Dim RecAc As New ADODB.Recordset
        Dim RecPesq As New ADODB.Recordset
        Dim RecPesq2 As New ADODB.Recordset
        Dim RecVerif As New ADODB.Recordset
                           
        StrSql = "Select Nome from tb_cliente where CodCli=" & VGIntCodCli
        RecPesq.Open StrSql, vgCon, 1, 3
               
        If RecPesq.EOF Then
            VPStrBox = MsgBox("Esse Código de Cliente não existe.", vbCritical, "Guide System - Aviso de erro")
            VGIntCodCli = 0
            TxtCodCli.SetFocus
        Else
            
            StrSql = "Select Login from tb_acesso where CodCli=" & VGIntCodCli
            RecVerif.Open StrSql, vgCon, 1, 3
            
            If Not RecVerif.EOF Then    'achou um cliente
                VPStrBox = MsgBox("Esse cliente já possui acesso cadastrado.", vbInformation, "Guide System - Informação")
            
            Else
                StrSql = "Select Login from tb_acesso where login='" & TxtLogin.Text & "'"
                RecPesq2.Open StrSql, vgCon, 1, 3
                
                If RecPesq2.EOF Then    'não achou registro
                
                    VPStrResponse = MsgBox("Confirma o cadastro de acesso do cliente" & Chr(13) & RecPesq.Fields.Item(0).Value & " ?", vbYesNo)
                    
                    If VPStrResponse = vbYes Then
                        StrSql = "Select * from tb_acesso"
                        RecAc.Open StrSql, vgCon, 1, 3
                        
                        RecAc.AddNew
                        RecAc("CodCli") = TxtCodCli.Text
                        RecAc("Login") = TxtLogin.Text
                        RecAc("Senha") = TxtSenha.Text
                        RecAc("DtAcesso") = FormataDataUS(Date)
                        RecAc.Update
                        
                        VPStrBox = MsgBox("Acesso cadastrado.", vbInformation, "Guide System - Informação")
                        
                        VGStrResponse = MsgBox("Deseja criar um cartão para o cliente?", vbYesNo)
                        
                        If VGStrResponse = vbYes Then
                            VGStrForm = "Acesso"
                            FrmCartao.Show
                        Else
                            Unload Me
                        End If
                    Else
                        Unload Me
                    End If
                Else
                    VPStrBox = MsgBox("Esse login já existe." & Chr(13) & "Favor escolher outro.", vbInformation, "Guide System - Informação")
                    TxtLogin.Text = ""
                End If
            End If
        End If
        
    End If
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmAcesso.hwnd)
    
    Height = 3000
    Width = 3720
    'Top = 1275
    'Left = 3960
    
    Unload FrmCadCli
    
    If VGStrForm = "Cliente" Then
        TxtCodCli.Text = FormataNum(VGIntCodCli)
        VGStrForm = ""
        VGIntCodCli = 0
    End If
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  FrmAcesso.Left = (MDIPrincipal.Width / 2) - (FrmAcesso.Width / 1.93)
  FrmAcesso.Top = (MDIPrincipal.Height / 3) - (FrmAcesso.Height / 5)
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

Private Sub TxtSenha2_GotFocus()
    TxtSenha2.SelStart = 0
    TxtSenha2.SelLength = Len(TxtSenha2.Text)
End Sub

Private Sub TxtSenha2_KeyPress(KeyAscii As Integer)
    
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
