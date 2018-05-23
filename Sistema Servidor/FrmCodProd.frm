VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCodProd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lançamento de caixa"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   ControlBox      =   0   'False
   Icon            =   "FrmCodProd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   3135
      Begin VB.TextBox TxtQtde 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         ToolTipText     =   "Quantidade do produto"
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         ToolTipText     =   "Sai da opção"
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   720
         TabIndex        =   2
         ToolTipText     =   "Confirma código do cliente"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox TxtCodProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         ToolTipText     =   "Código do produto"
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   120
         OleObjectBlob   =   "FrmCodProd.frx":08CA
         Top             =   1200
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCodProd 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCodProd.frx":0AFE
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblQtde 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCodProd.frx":0B76
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmCodProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String

Private Sub CmdOK_Click()

    Screen.MousePointer = vbHourglass
    
    If TxtCodProd.Text = "" Or TxtQtde.Text = "" Then
        VPStrBox = MsgBox("Preencha os campos em branco", vbCritical, "Guide System - Aviso de erro")
    Else
        Conecta
        
        Dim RecProd As New ADODB.Recordset
                           
        StrSql = "Select CodProd from tb_preco where CodProd=" & TxtCodProd.Text
        RecProd.Open StrSql, vgCon, 1, 3
        
        If RecProd.EOF Then  'não achou nada
            VPStrBox = MsgBox("Esse Código do Produto não existe.", vbCritical, "Guide System - Aviso de erro")
            Desconecta
        Else
            VGIntCodProd = RecProd.Fields.Item(0).Value
            VGIntQtde = TxtQtde.Text
            
            Desconecta
            Unload Me
            FrmCaixa.Enabled = True
            FrmCaixa.MontaForm
        End If
    End If
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub CmdSair_Click()
    Unload Me
    FrmCaixa.Enabled = True
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmCodProd.hwnd)
    
    Height = 2385
    Width = 3705
    'Top = 2730
    'Left = 4845
    
    FrmCaixa.Enabled = False
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  FrmCodProd.Left = (MDIPrincipal.Width / 2) - (FrmCodProd.Width / 1.93)
  FrmCodProd.Top = (MDIPrincipal.Height / 3) - (FrmCodProd.Height / 5)
End Sub

Private Sub TxtCodProd_GotFocus()
    TxtCodProd.SelStart = 0
    TxtCodProd.SelLength = Len(TxtCodProd.Text)
End Sub

Private Sub TxtCodProd_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtQtde_GotFocus()
    TxtQtde.SelStart = 0
    TxtQtde.SelLength = Len(TxtQtde.Text)
End Sub

Private Sub TxtQtde_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtQtde_LostFocus()
    If TxtQtde.Text = "" Then
        TxtQtde.Text = "1"
    End If
End Sub
