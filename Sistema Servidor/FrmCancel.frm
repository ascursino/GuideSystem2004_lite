VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCancel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelamento de Cartão"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   Icon            =   "FrmCancel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCancel 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Confirma cancelamento"
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame FraCartao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Dados do cartão"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton CmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   4320
         TabIndex        =   1
         ToolTipText     =   "Confirma número do cartão"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox TxtNumCart 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         ToolTipText     =   "Número do cartão"
         Top             =   360
         Width           =   2655
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumCart 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCancel.frx":000C
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame FraCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Dados do Cliente"
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   5175
      Begin ACTIVESKINLibCtl.SkinLabel LblCodCli 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCancel.frx":0084
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNome 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCancel.frx":00FC
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDtNasc 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCancel.frx":0164
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblIdent 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCancel.frx":01D6
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCpf 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCancel.frx":024A
         TabIndex        =   10
         Top             =   1800
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblResCodCli 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "FrmCancel.frx":02B0
         TabIndex        =   11
         Top             =   360
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblResNome 
         Height          =   255
         Left            =   840
         OleObjectBlob   =   "FrmCancel.frx":031A
         TabIndex        =   12
         Top             =   720
         Width           =   4215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblResDtNasc 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmCancel.frx":0384
         TabIndex        =   13
         Top             =   1080
         Width           =   3975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblResIdent 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "FrmCancel.frx":03EE
         TabIndex        =   14
         Top             =   1440
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblResCpf 
         Height          =   255
         Left            =   600
         OleObjectBlob   =   "FrmCancel.frx":0458
         TabIndex        =   15
         Top             =   1800
         Width           =   4455
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4200
      OleObjectBlob   =   "FrmCancel.frx":04C2
      Top             =   3720
   End
End
Attribute VB_Name = "FrmCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String

Private Sub CmdCancel_Click()
    Screen.MousePointer = vbHourglass
    
    VGStrResponse = MsgBox("Deseja realmente cancelar este cartão?", vbYesNo)
   
    If VGStrResponse = vbYes Then
        FrmConfirmCancel.Show
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdOK_Click()
    Screen.MousePointer = vbHourglass
    
    If TxtNumCart.Text <> "" Then
        
        Conecta
        
        Dim RecCart As New ADODB.Recordset
        Dim RecCli As New ADODB.Recordset
                           
        StrSql = "Select CodCli,Cancelado from tb_cartao where NumCartao=" & TxtNumCart.Text
        RecCart.Open StrSql, vgCon, 1, 3
        
        If RecCart.EOF Then     'não achou nenhum cartão
            VPStrBox = MsgBox("Este cartão não existe.", vbCritical, "Guide System - Aviso de erro")
            CmdCancel.Enabled = False
        
        Else        'achou cartão
            If RecCart.Fields.Item(1).Value = True Then    'cartão está cancelado
                VPStrBox = MsgBox("Este cartão já está cancelado.", vbInformation, "Guide System - Informação")
                
                LblResCodCli.Caption = ""
                LblResNome.Caption = ""
                LblResDtNasc.Caption = ""
                LblResIdent.Caption = ""
                LblResCpf.Caption = ""
                CmdCancel.Enabled = False
            
            Else    'NÃO está cancelado.
                
                StrSql = "Select CodCli,Nome,NascDia,NascMes,NascAno,Ident,Cpf from tb_cliente where CodCli=" & RecCart.Fields.Item(0).Value
                RecCli.Open StrSql, vgCon, 1, 3
                
                LblResCodCli.Caption = RecCli.Fields.Item(0).Value
                LblResNome.Caption = RecCli.Fields.Item(1).Value
                
                If RecCli.Fields.Item(2).Value <> "0" And RecCli.Fields.Item(3).Value <> "0" And RecCli.Fields.Item(4).Value <> "0" Then
                    LblResDtNasc.Caption = FormataNum(RecCli.Fields.Item(2).Value) & "/" & FormataNum(RecCli.Fields.Item(3).Value) & "/" & RecCli.Fields.Item(4).Value
                Else
                    LblResDtNasc.Caption = ""
                End If
                
                LblResIdent.Caption = RecCli.Fields.Item(5).Value
                LblResCpf.Caption = RecCli.Fields.Item(6).Value
                
                LblResCodCli.Visible = True
                LblResNome.Visible = True
                LblResDtNasc.Visible = True
                LblResIdent.Visible = True
                LblResCpf.Visible = True
                CmdCancel.Enabled = True
            End If
        End If
        Desconecta
        
    Else
        VPStrBox = MsgBox("Digite o número do cartão.", vbCritical, "Guide System - Aviso de erro")
        TxtNumCart.SetFocus
    End If
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmCancel.hwnd)
    
    Height = 4845
    Width = 5775
    'Top = 1275
    'Left = 2940

    LblResCodCli.Visible = False
    LblResNome.Visible = False
    LblResDtNasc.Visible = False
    LblResIdent.Visible = False
    LblResCpf.Visible = False
    CmdCancel.Enabled = False
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  FrmCancel.Left = (MDIPrincipal.Width / 2) - (FrmCancel.Width / 1.93)
  FrmCancel.Top = (MDIPrincipal.Height / 3) - (FrmCancel.Height / 5)
End Sub

Private Sub TxtNumCart_GotFocus()
    TxtNumCart.SelStart = 0
    TxtNumCart.SelLength = Len(TxtNumCart.Text)
End Sub

Private Sub TxtNumCart_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub
