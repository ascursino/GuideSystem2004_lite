VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCartao 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Criar Cartão"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "FrmCartao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraRecarga 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4095
      Begin VB.TextBox TxtDtCart 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "__/__/____"
         ToolTipText     =   "Data de criação do cartão"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox TxtCodCli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         ToolTipText     =   "Código do cliente"
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton CmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         ToolTipText     =   "Inclui cadastro do cartão"
         Top             =   1560
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   480
         OleObjectBlob   =   "FrmCartao.frx":000C
         Top             =   1440
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDtCart 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmCartao.frx":0240
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCodCli 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmCartao.frx":02BC
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmCartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String
Public VPStrGrava As String

Private Sub CmdIncluir_Click()
    Screen.MousePointer = vbHourglass
    
    If TxtDtCart.Text = "" Or TxtCodCli.Text = "" Then
        VPStrBox = MsgBox("Preencha os campos em branco", vbCritical, "Guide System - Aviso de erro")
    Else
        Conecta
        
        Dim RecCar As New ADODB.Recordset
        Dim RecNumCar As New ADODB.Recordset
        Dim RecCli As New ADODB.Recordset
        Dim RecCred As New ADODB.Recordset
        
        StrSql = "Select CodCli from tb_cliente where CodCli=" & TxtCodCli.Text
        RecCli.Open StrSql, vgCon, 1, 3
        
        If Not RecCli.EOF Then  'achou o cliente
            
            StrSql = "Select NumCartao from tb_cartao where Cancelado = '0' and CodCli=" & TxtCodCli.Text
            RecNumCar.Open StrSql, vgCon, 1, 3
            
            If Not RecNumCar.EOF Then   'achou cartao
                VPStrResponse = MsgBox("Esse cliente já tem um cartão." & Chr(13) & "Deseja consultá-lo?", vbYesNo)
            
                If VPStrResponse = vbYes Then
                    VGStrConsultCart = "sim"
                    VGIntCodCli = TxtCodCli.Text
                    'VGIntNumCartao = TxtNumCartao.Text
                    FrmConsultCart.Show
                Else
                    VPStrResponse = MsgBox("Criar outro cartão para este cliente?", vbYesNo)
                
                    If VPStrResponse = vbYes Then
                        'StrSql = "Insert into tb_cartao VALUES (" & TxtCodCli.Text & "," & _
                        '"'" & FormataDataUS(TxtDtCart.Text) & "','0',null,null,null)"
                        StrSql = "Select * from tb_cartao"
                        RecCar.Open StrSql, vgCon, 1, 3
                        
                        RecCar.AddNew
                        RecCar("CodCli") = TxtCodCli.Text
                        RecCar("DtCartao") = FormataDataUS(TxtDtCart.Text)
                        RecCar("Cancelado") = "0"
                        RecCar("Motivo") = Null
                        RecCar("Resp") = Null
                        RecCar("DtCancel") = Null
                        RecCar.Update
                        
                        RecCar.Close
                        
                        StrSql = "Select max(NumCartao) as NumCartao from tb_cartao where CodCli=" & TxtCodCli.Text
                        RecCar.Open StrSql, vgCon, 1, 3
                          
                        VGIntNumCartao = RecCar("NumCartao")
                        
                        Desconecta
                        
                        VPStrBox = MsgBox("Cartão nº " & FormataNum(VGIntNumCartao) & " criado.", vbInformation, "Guide System - Informação")
                        
                        VPStrResponse = MsgBox("Deseja inserir créditos neste cartão ?", vbYesNo)
                            
                        If VPStrResponse = vbYes Then
                            VGStrForm = "Cartao"
                            VGIntCodCli = TxtCodCli.Text
                            FrmCredito.Show
                        Else
                            Unload Me
                        End If
                    Else
                        Unload Me
                    End If
                    
                End If
                
            Else    'não achou cartão
                StrSql = "Select * from tb_cartao"
                RecCar.Open StrSql, vgCon, 1, 3
                
                RecCar.AddNew
                RecCar("CodCli") = TxtCodCli.Text
                RecCar("DtCartao") = FormataDataUS(TxtDtCart.Text)
                RecCar("Cancelado") = "0"
                RecCar("Motivo") = Null
                RecCar("Resp") = Null
                RecCar("DtCancel") = Null
                RecCar.Update
                
                RecCar.Close
                
                StrSql = "Select max(NumCartao) as NumCartao from tb_cartao where CodCli=" & TxtCodCli.Text
                RecCar.Open StrSql, vgCon, 1, 3
                
                VGIntNumCartao = RecCar("NumCartao")
                
                Desconecta
                
                VPStrBox = MsgBox("Cartão nº " & FormataNum(VGIntNumCartao) & " criado.", vbInformation, "Guide System - Informação")
                
                VPStrResponse = MsgBox("Deseja inserir créditos neste cartão ?", vbYesNo)
                    
                If VPStrResponse = vbYes Then
                    VGStrForm = "Cartao"
                    VGIntCodCli = TxtCodCli.Text
                    FrmCredito.Show
                Else
                    Unload Me
                End If
    
            End If
        Else    'não achou o cliente
            VPStrBox = MsgBox("Esse Código de Cliente não existe.", vbCritical, "Guide System - Aviso de erro")
        End If
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmCartao.hwnd)
    
    Height = 2610
    Width = 4680
    'Top = 1275
    'Left = 3450
    
    Unload FrmAcesso
    
    TxtDtCart.Text = FormataData(Date)

    If VGStrForm = "Acesso" Then
        TxtCodCli.Text = FormataNum(VGIntCodCli)
        VGStrForm = ""
    End If
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  FrmCartao.Left = (MDIPrincipal.Width / 2) - (FrmCartao.Width / 1.93)
  FrmCartao.Top = (MDIPrincipal.Height / 3) - (FrmCartao.Height / 5)
End Sub

Private Sub TxtCodCli_GotFocus()
    TxtCodCli.SelStart = 0
    TxtCodCli.SelLength = Len(TxtCodCli.Text)
End Sub

Private Sub TxtCodCli_KeyPress(KeyAscii As Integer)

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

    ElseIf KeyAscii = 127 Or KeyAscii = 168 Then
        KeyAscii = 0
    
    End If
    '=========================================

    '===== Letras em maiúsculo e minúsculo ===
    If KeyAscii >= 65 And KeyAscii <= 90 Then
        KeyAscii = 0
    
    ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = 0

    ElseIf KeyAscii = 199 Or KeyAscii = 231 Then
        KeyAscii = 0
    
    End If
    '=========================================

End Sub

Private Sub TxtDtCart_GotFocus()
    If TxtDtCart.Text = "__/__/____" Then
        TxtDtCart.Text = ""
    End If
    
    TxtDtCart.SelStart = 0
    TxtDtCart.SelLength = Len(TxtDtCart.Text)
End Sub

Private Sub TxtDtCart_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    
    If TxtDtCart.Text = "__/__/____" Then
        TxtDtCart.Text = ""
    End If
End Sub

Private Sub TxtDtCart_LostFocus()
    
    Dim VLStrData As String
    
    If TxtDtCart.Text <> "" Then
        VLStrData = VerificaData(TxtDtCart.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtCart.SetFocus
        Else
            TxtDtCart.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtCart.Text = "__/__/____"
    End If
End Sub
