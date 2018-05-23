VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCadCliAlt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "teste"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   ControlBox      =   0   'False
   Icon            =   "FrmCadCliAlt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton CmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         ToolTipText     =   "Sai da opção"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   720
         TabIndex        =   1
         ToolTipText     =   "Confirma código do cliente"
         Top             =   840
         Width           =   615
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
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   120
         OleObjectBlob   =   "FrmCadCliAlt.frx":08CA
         Top             =   720
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCodCli 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmCadCliAlt.frx":0AFE
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmCadCliAlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String

Private Sub CmdOK_Click()

    Screen.MousePointer = vbHourglass
    
    If TxtCodCli.Text = "" Then
        VPStrBox = MsgBox("Preencha o código do cliente", vbCritical, "Guide System - Aviso de erro")
    Else
        Conecta
        
        Dim RecCli As New ADODB.Recordset
                           
        StrSql = "Select CodCli from tb_cliente where CodCli=" & TxtCodCli.Text
        RecCli.Open StrSql, vgCon, 1, 3
        
        If RecCli.EOF Then  'não achou nada
            VPStrBox = MsgBox("Esse Código de Cliente não existe.", vbCritical, "Guide System - Aviso de erro")
            Desconecta
        Else
            VGIntCodCliTemp = RecCli.Fields.Item(0).Value
            
            Desconecta
            Unload Me
            
            If VGStrAlt = "cliente" Then
                VGStrAlt = ""
                FrmAltCli.Show
            
            ElseIf VGStrAlt = "acesso" Then
                VGStrAlt = ""
                FrmAltAces.Show
            
            ElseIf VGStrAlt = "espera" Then
                VGStrAlt = ""
                FrmMaquina.Enabled = True
                FrmMaquina.Inclui_Lista
                
            End If
                    
        End If
    End If
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub CmdSair_Click()
    Unload Me
    If VGStrAlt = "" Or VGStrAlt = "espera" Then
        FrmMaquina.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmCadCliAlt.hwnd)
    
    Height = 1905
    Width = 3720
    
    If VGStrAlt = "cliente" Then
        'Top = 1275
        'Left = 3795
        Me.Caption = "Alteração Cadastro de Cliente"
    
    ElseIf VGStrAlt = "acesso" Then
        'Top = 1275
        'Left = 3795
        Me.Caption = "Alteração Cadastro de Acesso"
    
    ElseIf VGStrAlt = "espera" Then
        'Top = 3855
        'Left = 3975
        Me.Caption = "Inclusão na Lista de Espera"
        FrmMaquina.Enabled = False
    End If
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub Form_Resize()
  FrmCadCliAlt.Left = (MDIPrincipal.Width / 2) - (FrmCadCliAlt.Width / 1.93)
  FrmCadCliAlt.Top = (MDIPrincipal.Height / 3) - (FrmCadCliAlt.Height / 5)
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
