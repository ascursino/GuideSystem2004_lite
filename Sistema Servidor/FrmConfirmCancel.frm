VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmConfirmCancel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirmação de Cancelamento"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   ControlBox      =   0   'False
   Icon            =   "FrmConfirmCancel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton CmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   840
         TabIndex        =   2
         ToolTipText     =   "Confirma cancelamento do cartão"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox TxtResp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         ToolTipText     =   "Responsável pelo cancelamento"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox TxtMotivo 
         Appearance      =   0  'Flat
         Height          =   885
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Motivo do cancelamento"
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CommandButton CmdVoltar 
         Caption         =   "Voltar"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         ToolTipText     =   "Volta a tela anterior"
         Top             =   2520
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   360
         OleObjectBlob   =   "FrmConfirmCancel.frx":08CA
         Top             =   1920
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumCart 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmConfirmCancel.frx":0AFE
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDtCancel 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmConfirmCancel.frx":0B76
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblResp 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmConfirmCancel.frx":0BEE
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblMotivo 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmConfirmCancel.frx":0C64
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblResNumCart 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmConfirmCancel.frx":0CD0
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblResDtCancel 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmConfirmCancel.frx":0D3A
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmConfirmCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public MsgResp As String
Public MsgMotivo As String

Private Sub CmdOK_Click()
    Screen.MousePointer = vbHourglass
    
    If TxtResp.Text <> "" And TxtMotivo.Text <> "" Then
        'gravar informações de cancelamento no banco
        
        Conecta
        
        Dim RecCart As New ADODB.Recordset
                           
        StrSql = "Select * from tb_cartao where NumCartao=" & LblResNumCart
        RecCart.Open StrSql, vgCon, 1, 3
        
        RecCart("Cancelado") = "1"
        RecCart("Motivo") = TxtMotivo.Text
        RecCart("Resp") = TxtResp.Text
        RecCart("DtCancel") = FormataDataUS(Date)
        RecCart.Update
        
        VPStrBox = MsgBox("Cancelamento efetuado.", vbInformation, "Guide System - Informação")
                
        Desconecta
        
        Unload Me
        Unload FrmCancel
        
    Else    'campos em branco
        VPStrBox = MsgBox("Preencha os campos em branco", vbCritical, "Guide System - Aviso de erro")
    End If
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub CmdVoltar_Click()
    Unload Me
    FrmCancel.Enabled = True
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmConfirmCancel.hwnd)
    
    Height = 3585
    Width = 5250
    'Top = 2280
    'Left = 4920
    
    FrmCancel.Enabled = False
    
    LblResNumCart.Caption = FrmCancel.TxtNumCart.Text
    LblResDtCancel.Caption = FormataData(Date)
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  FrmConfirmCancel.Left = (MDIPrincipal.Width / 2) - (FrmConfirmCancel.Width / 1.93)
  FrmConfirmCancel.Top = (MDIPrincipal.Height / 3) - (FrmConfirmCancel.Height / 5)
End Sub

Private Sub TxtMotivo_GotFocus()
    TxtMotivo.SelStart = 0
    TxtMotivo.SelLength = Len(TxtMotivo.Text)
End Sub

Private Sub TxtMotivo_KeyPress(KeyAscii As Integer)
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

Private Sub TxtResp_GotFocus()
    TxtResp.SelStart = 0
    TxtResp.SelLength = Len(TxtResp.Text)
End Sub

Private Sub TxtResp_KeyPress(KeyAscii As Integer)
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
