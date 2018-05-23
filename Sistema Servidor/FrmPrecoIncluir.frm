VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmPrecoIncAlt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inclusão de Produtos & Preços"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   ControlBox      =   0   'False
   Icon            =   "FrmPrecoIncluir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraProdPreco 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton CmdFechar 
         Caption         =   "Fechar"
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         ToolTipText     =   "Fecha janela"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton CmdIncAlt 
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TxtPreco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   1
         ToolTipText     =   "Preço"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox TxtProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   0
         ToolTipText     =   "Descrição do produto"
         Top             =   360
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   120
         OleObjectBlob   =   "FrmPrecoIncluir.frx":000C
         Top             =   1200
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblProd 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmPrecoIncluir.frx":0240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblPreco 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmPrecoIncluir.frx":02AE
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblExemp 
         Height          =   255
         Left            =   3000
         OleObjectBlob   =   "FrmPrecoIncluir.frx":0318
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmPrecoIncAlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public MsgProd As String
Public MsgPreco As String

Private Sub CmdFechar_Click()
    Unload Me
    FrmPreco.Enabled = True
End Sub

Private Sub CmdIncAlt_Click()
    Screen.MousePointer = vbHourglass
    
    If TxtProd.Text <> "" And TxtPreco.Text <> "" Then
    
        If VGStrPreco = "incluir" Then
            Conecta
            
            Dim RecPrInc As New ADODB.Recordset
            
            StrSql = "Select * from tb_preco"
            RecPrInc.Open StrSql, vgCon, 1, 3
            
            RecPrInc.AddNew
            RecPrInc("Prod") = TxtProd.Text
            RecPrInc("Preco") = TxtPreco.Text
            RecPrInc.Update
            
            Desconecta
        
            VPStrBox = MsgBox("Produto cadastrado", vbInformation, "Guide System - Informação")
            
            VGStrPreco = ""
            Unload Me
            FrmPreco.Enabled = True
            
        ElseIf VGStrPreco = "alterar" Then
            
            Conecta
            
            Dim RecPrAlt As New ADODB.Recordset
            
            StrSql = "Select * from tb_preco where CodProd=" & VGIntCodProd
            RecPrAlt.Open StrSql, vgCon, 1, 3
            
            RecPrAlt("Prod") = TxtProd.Text
            RecPrAlt("Preco") = TxtPreco.Text
            RecPrAlt.Update
            
            Desconecta
        
            VPStrBox = MsgBox("Produto atualizado", vbInformation, "Guide System - Informação")
            
            VGStrPreco = ""
            'VGIntCodProd = 0
            Unload Me
            FrmPreco.Enabled = True
        
        End If
        
        Call FrmPreco.MontaGridPreco
    Else
        VPStrBox = MsgBox("Preencha os campos em branco", vbCritical, "Guide System - Aviso de erro")
    End If
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmPrecoIncAlt.hwnd)
    
    Height = 2535
    Width = 4905
    'Top = 2325
    'Left = 4875
    
    FrmPreco.Enabled = False

    If VGStrPreco = "incluir" Then
        CmdIncAlt.Caption = "Incluir"
        CmdIncAlt.ToolTipText = "Inclui produto"

    ElseIf VGStrPreco = "alterar" Then

        Conecta

        Dim RecPr As New ADODB.Recordset

        StrSql = "Select Prod,Preco from tb_preco where CodProd=" & VGIntCodProd
        RecPr.Open StrSql, vgCon, 1, 3

        TxtProd.Text = RecPr.Fields.Item(0).Value
        TxtPreco.Text = RecPr.Fields.Item(1).Value

        Desconecta

        CmdIncAlt.Caption = "Alterar"
        CmdIncAlt.ToolTipText = "Altera produto"

    End If
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  FrmPrecoIncAlt.Left = (MDIPrincipal.Width / 2) - (FrmPrecoIncAlt.Width / 1.93)
  FrmPrecoIncAlt.Top = (MDIPrincipal.Height / 3) - (FrmPrecoIncAlt.Height / 5)
End Sub

Private Sub TxtPreco_GotFocus()
    TxtPreco.SelStart = 0
    TxtPreco.SelLength = Len(TxtPreco.Text)
End Sub

Private Sub TxtPreco_KeyPress(KeyAscii As Integer)
    
    '============ Símbolos ======================
    If KeyAscii >= 33 And KeyAscii <= 43 Then
        KeyAscii = 0
    
    ElseIf KeyAscii >= 45 And KeyAscii <= 47 Then
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

    '====== Letras em maiúsculo e minúsculo ==========
    If KeyAscii >= 65 And KeyAscii <= 90 Then
        KeyAscii = 0
    
    ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = 0
    
    ElseIf KeyAscii = 199 Or KeyAscii = 231 Then
        KeyAscii = 0
    
    End If
    '=========================================

End Sub

Private Sub TxtPreco_LostFocus()
    If TxtPreco.Text <> "" Then
        TxtPreco.Text = Trim(Replace(FormataMoeda(TxtPreco.Text), "R$", ""))
    End If
End Sub

Private Sub TxtProd_GotFocus()
    TxtProd.SelStart = 0
    TxtProd.SelLength = Len(TxtProd.Text)
End Sub

Private Sub TxtProd_KeyPress(KeyAscii As Integer)

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

End Sub
