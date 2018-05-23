VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCaixa 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lançamentos do Caixa"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   Icon            =   "FrmCaixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraCaixa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Lançamentos"
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2775
      Begin VB.TextBox TxtDataItem 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "__/__/____"
         ToolTipText     =   "Data do item"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox TxtCred 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """R$""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   2
         ToolTipText     =   "Valor do crédito"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox TxtDeb 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   3
         ToolTipText     =   "Valor do débito"
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton CmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   840
         TabIndex        =   4
         ToolTipText     =   "Inclui lançamento no caixa"
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox TxtDescr 
         Appearance      =   0  'Flat
         Height          =   765
         Left            =   240
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Descrição do item"
         Top             =   1080
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   240
         OleObjectBlob   =   "FrmCaixa.frx":000C
         Top             =   3000
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDataItem 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCaixa.frx":0240
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDescr 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCaixa.frx":02A8
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCred 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCaixa.frx":031A
         TabIndex        =   8
         Top             =   2040
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDeb 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCaixa.frx":0388
         TabIndex        =   9
         Top             =   2520
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblExpCred 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "FrmCaixa.frx":03F4
         TabIndex        =   10
         Top             =   2040
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblExpDeb 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "FrmCaixa.frx":0460
         TabIndex        =   11
         Top             =   2520
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String
Public VPStrGrava As String
Public VPStrCred As String
Public VPStrDeb As String
Public MsgDataItem As String
Public MsgDescr As String
Public MsgDeb As String
Public MsgCred As String

Private Sub CmdIncluir_Click()
    Screen.MousePointer = vbHourglass
    
    If TxtDataItem.Text = "" Or TxtDescr.Text = "" Or (TxtCred.Text = "" And TxtDeb.Text = "") Then
        VPStrBox = MsgBox("Preencha o(s) campo(s) em branco", vbCritical, "Guide System - Aviso de erro")
    Else
        If TxtCred.Text <> "" And TxtDeb.Text <> "" Then
            VPStrBox = MsgBox("Lançamento não pode assumir valor de crédito" & Chr(13) & "e débito ao mesmo tempo", vbCritical, "Guide System - Aviso de erro")
        Else
            Conecta
            
            Dim RecCx As New ADODB.Recordset
            
            If TxtDeb.Text = "" Then
                VPStrDeb = "0"
            Else
                VPStrDeb = TxtDeb.Text
            End If
                
            If TxtCred.Text = "" Then
                VPStrCred = "0"
            Else
                VPStrCred = TxtCred.Text
            End If
            
            StrSql = "Select * from tb_caixa "
            RecCx.Open StrSql, vgCon, 1, 3
                  
            RecCx.AddNew
            RecCx("Descr") = TxtDescr.Text
            RecCx("Vldeb") = VPStrDeb
            RecCx("Vlcred") = VPStrCred
            RecCx("DtItem") = FormataDataUS(TxtDataItem.Text)
            RecCx.Update
            
            Desconecta
                            
            VPStrBox = MsgBox("Lançamento cadastrado.", vbInformation, "Guide System - Informação")
                    
            TxtDescr.Text = ""
            TxtCred.Text = ""
            TxtDeb.Text = ""
            
            FrmCodProd.Show
        End If
    End If
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmCaixa.hwnd)
    
    Height = 4320
    Width = 3345
    'Top = 1275
    'Left = 3960
    
    TxtDataItem.Text = FormataData(Date)
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  FrmCaixa.Left = (MDIPrincipal.Width / 2) - (FrmCaixa.Width / 1.93)
  FrmCaixa.Top = (MDIPrincipal.Height / 3) - (FrmCaixa.Height / 5)
End Sub

Private Sub TxtCred_GotFocus()
    TxtCred.SelStart = 0
    TxtCred.SelLength = Len(TxtCred.Text)
End Sub

Private Sub TxtCred_KeyPress(KeyAscii As Integer)

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

Private Sub TxtCred_LostFocus()
    If TxtCred.Text <> "" Then
        TxtCred.Text = Trim(Replace(FormataMoeda(TxtCred.Text), "R$", ""))
    End If
End Sub

Private Sub TxtDataItem_GotFocus()
    If TxtDataItem.Text = "__/__/____" Then
        TxtDataItem.Text = ""
    End If
    
    TxtDataItem.SelStart = 0
    TxtDataItem.SelLength = Len(TxtDataItem.Text)
End Sub

Private Sub TxtDataItem_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    
    If TxtDataItem.Text = "__/__/____" Then
        TxtDataItem.Text = ""
    End If
End Sub

Private Sub TxtDataItem_LostFocus()
    
    Dim VLStrData As String
    
    If TxtDataItem.Text <> "" Then
        VLStrData = VerificaData(TxtDataItem.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDataItem.SetFocus
        Else
            TxtDataItem.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDataItem.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDeb_GotFocus()
    TxtDeb.SelStart = 0
    TxtDeb.SelLength = Len(TxtDeb.Text)
End Sub

Private Sub TxtDeb_KeyPress(KeyAscii As Integer)
    
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

Sub MontaForm()

    If VGIntCodProd <> 0 Then
        
        Conecta
        
        Dim RecProd As New ADODB.Recordset
        Dim preco As Currency
        
        StrSql = "Select Prod, Preco from tb_preco where CodProd=" & VGIntCodProd
        RecProd.Open StrSql, vgCon, 1, 3
        
        preco = RecProd.Fields.Item(1).Value * VGIntQtde
        
        TxtDescr.Text = RecProd.Fields.Item(0).Value
        TxtCred.Text = Trim(Replace(FormataMoeda(preco), "R$", ""))
        VGIntCodProd = 0
        VGIntQtde = 0
    End If

End Sub

Private Sub TxtDeb_LostFocus()
    If TxtDeb.Text <> "" Then
        TxtDeb.Text = Trim(Replace(FormataMoeda(TxtDeb.Text), "R$", ""))
    End If
End Sub

Private Sub TxtDescr_GotFocus()
    TxtDescr.SelStart = 0
    TxtDescr.SelLength = Len(TxtDescr.Text)
End Sub

Private Sub TxtDescr_KeyPress(KeyAscii As Integer)
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
