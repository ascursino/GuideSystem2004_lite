VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCredito 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inserir Créditos"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "FrmCredito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraRecarga 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4695
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   2535
         Left            =   3000
         TabIndex        =   13
         Top             =   120
         Width           =   15
      End
      Begin VB.CommandButton CmdCalcPreco 
         Caption         =   "Calcular preço"
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         ToolTipText     =   "Calcular preço do crédito"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox TxtCredMin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   3
         ToolTipText     =   "Minutos para crédito"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox TxtDtCred 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "__/__/____"
         ToolTipText     =   "Data do crédito"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TxtCredHora 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   2
         ToolTipText     =   "Horas para crédito"
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton CmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         ToolTipText     =   "Inclui créditos"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox TxtNumCartao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         ToolTipText     =   "Número do cartão"
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   360
         OleObjectBlob   =   "FrmCredito.frx":000C
         Top             =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumCartao 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCredito.frx":0240
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDtCred 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCredito.frx":02B8
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCred 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCredito.frx":0336
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblMin 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "FrmCredito.frx":03A4
         TabIndex        =   10
         Top             =   1200
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "FrmCredito.frx":040C
         TabIndex        =   11
         Top             =   1560
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "FrmCredito.frx":0478
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblHora 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "FrmCredito.frx":04F4
         TabIndex        =   14
         Top             =   960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblPreco 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "FrmCredito.frx":0564
         TabIndex        =   15
         Top             =   1440
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrCred As String
Public VPStrRest As String
Public DataHora As String
Public VPStrTemp As String
Public VPIntHora As Integer
Public VPIntMin As Integer
Public VPIntSeg As Integer
Public VPIntHoraTemp As Integer
Public VPIntMinTemp As Integer
Public VPIntSegTemp As Integer
Public MsgHora As String
Public MsgMin As String
Public MsgSeg As String
Public VPStrGrava As String
Public VPStrDescr As String
Public VPStrValCred As String
Public VPStrTempRest As String
Public VPStrTempCartao As String

Private Sub CmdCalcPreco_Click()
    If TxtCredHora.Text = "" Or TxtCredMin.Text = "" Then
        VPStrBox = MsgBox("Preencha o(s) campo(s) de crédito", vbCritical, "Guide System - Aviso de erro")
    Else
        Dim VLStrCredito As String
        VLStrCredito = TxtCredHora.Text & ":" & TxtCredMin.Text & ":00"
        
        LblHora.Caption = TxtCredHora.Text & "h e " & TxtCredMin.Text & "m"
        
        Conecta
        LblPreco.Caption = FormataMoeda(Calcula_Preco(VLStrCredito))
        Desconecta
    End If
End Sub

Private Sub CmdIncluir_Click()
    Dim VLStrCredito As String
    
    Screen.MousePointer = vbHourglass
    
    'Call VerificaCred
    
    If TxtNumCartao.Text = "" Or TxtDtCred.Text = "" Or TxtCredHora.Text = "" Or TxtCredHora.Text = "00" Then
        VPStrBox = MsgBox("Preencha os campos em branco", vbCritical, "Guide System - Aviso de erro")
    Else
        VLStrCredito = TxtCredHora.Text & ":" & TxtCredMin.Text & ":00"
        
        If TxtCredMin.Text = "" Then
            TxtCredMin.Text = "00"
        End If
        
        Conecta
        
        Dim RecVerif As New ADODB.Recordset
        Dim RecCre As New ADODB.Recordset
        Dim RecGrCre As New ADODB.Recordset
        Dim RecCart As New ADODB.Recordset
        Dim RecEsp As New ADODB.Recordset
        Dim RecCxa As New ADODB.Recordset
        
        StrSql = "Select Cancelado,CodCli from tb_cartao where NumCartao=" & TxtNumCartao.Text
        RecCart.Open StrSql, vgCon, 1, 3
        
        If RecCart.EOF Then 'não achou cartão
            VPStrBox = MsgBox("Este cartão não existe.", vbInformation, "Guide System - Informação")

        ElseIf RecCart.Fields.Item(0).Value = True Then  'cartão está cancelado
            VPStrBox = MsgBox("Este cartão está cancelado." & Chr(13) & "Os créditos não poderão ser inseridos.", vbInformation, "Guide System - Informação")
                
        Else    'cartão ativo
        
            StrSql = "Select TempoCred,TempoRest from tb_credito where NumCartao=" & TxtNumCartao.Text
            RecVerif.Open StrSql, vgCon, 1, 3
            
            If RecVerif.EOF Then
                VPStrTempRest = "00:00:00"
                VPStrTempCartao = "insert"
            Else
                VPStrTempRest = RecVerif.Fields.Item(1).Value
            End If
            
            If VPStrTempRest = "00:00:00" Then    'esse cartão não tem créditos
                
                If VPStrTempCartao = "insert" Then
                    'insere créditos
                    StrSql = "Select * from tb_credito"
                    RecCre.Open StrSql, vgCon, 1, 3
                    
                    RecCre.AddNew
                    RecCre("NumCartao") = TxtNumCartao.Text
                    RecCre("TempoCred") = VLStrCredito
                    RecCre("TempoRest") = VLStrCredito
                    RecCre("DtCred") = FormataDataUS(TxtDtCred.Text)
                    RecCre.Update
                    
                    VPStrTempCartao = ""
                Else
                    'atualiza créditos restantes
                    VPIntHora = Mid(VLStrCredito, 1, 2)
                    VPIntMin = Mid(VLStrCredito, 4, 2)
                    VPIntSeg = Mid(VLStrCredito, 7, 2)
    
                    VPStrCred = VLStrCredito
                    VPStrRest = TimeSerial(Hour(RecVerif.Fields.Item(1).Value) + VPIntHora, Minute(RecVerif.Fields.Item(1).Value) + VPIntMin, Second(RecVerif.Fields.Item(1).Value) + VPIntSeg)
    
                    StrSql = "Select * from tb_credito where NumCartao=" & TxtNumCartao.Text
                    RecCre.Open StrSql, vgCon, 1, 3
                    
                    RecCre("TempoCred") = VPStrCred
                    RecCre("TempoRest") = VPStrRest
                    RecCre("DtCred") = FormataDataUS(Date)
                    RecCre.Update
                    
                End If
                
                'insere dados na tabela que guarda todos os créditos
                StrSql = "Select * from tb_guardacredito"
                RecGrCre.Open StrSql, vgCon, 1, 3
                
                RecGrCre.AddNew
                RecGrCre("NumCartao") = TxtNumCartao.Text
                RecGrCre("TempoCred") = VLStrCredito
                RecGrCre("DtCred") = FormataDataUS(TxtDtCred.Text)
                RecGrCre.Update
                
                'insere item na tabela de caixa
                VPStrDescr = "Crédito de " & VLStrCredito & " para cartão " & FormataNum(TxtNumCartao.Text)
                VPStrValCred = Trim(Replace(FormataMoeda(Calcula_Preco(VLStrCredito)), "R$", ""))
                
              
                StrSql = "Select * from tb_caixa"
                RecCxa.Open StrSql, vgCon, 1, 3
                
                RecCxa.AddNew
                RecCxa("Descr") = VPStrDescr
                RecCxa("Vldeb") = "0"
                RecCxa("Vlcred") = VPStrValCred
                RecCxa("DtItem") = FormataDataUS(TxtDtCred.Text)
                RecCxa.Update
                
                VPStrTemp = "insert"
            
            Else   'cartão ainda tem crédito
                VPStrResponse = MsgBox("Esse cartão ainda possui créditos." & Chr(13) & "Deseja adicionar?", vbYesNo)
                
                If VPStrResponse = vbYes Then
                    VPIntHora = Mid(VLStrCredito, 1, 2)
                    VPIntMin = Mid(VLStrCredito, 4, 2)
                    VPIntSeg = Mid(VLStrCredito, 7, 2)
                    
                    'VPStrCred = TimeSerial(Hour(RecVerif.Fields.Item(0).Value) + VPIntHora, Minute(RecVerif.Fields.Item(0).Value) + VPIntMin, Second(RecVerif.Fields.Item(0).Value) + VPIntSeg)
                    VPStrCred = VLStrCredito
                    VPStrRest = TimeSerial(Hour(RecVerif.Fields.Item(1).Value) + VPIntHora, Minute(RecVerif.Fields.Item(1).Value) + VPIntMin, Second(RecVerif.Fields.Item(1).Value) + VPIntSeg)
                    
                    'atualiza créditos restantes
                    StrSql = "Select * from tb_credito where NumCartao=" & TxtNumCartao.Text
                    RecCre.Open StrSql, vgCon, 1, 3
                    
                    RecCre("TempoCred") = VPStrCred
                    RecCre("TempoRest") = VPStrRest
                    RecCre("DtCred") = FormataDataUS(Date)
                    RecCre.Update
                    
                    'insere dados na tabela que guarda todos os creditos
                    StrSql = "Select * from tb_guardacredito"
                    RecGrCre.Open StrSql, vgCon, 1, 3
                    
                    RecGrCre.AddNew
                    RecGrCre("NumCartao") = TxtNumCartao.Text
                    RecGrCre("TempoCred") = VLStrCredito
                    RecGrCre("DtCred") = FormataDataUS(TxtDtCred.Text)
                    RecGrCre.Update
                    
                    'insere item na tabela de caixa
                    VPStrDescr = "Crédito de " & VLStrCredito & " para cartão " & FormataNum(TxtNumCartao.Text)
                    VPStrValCred = FormataMoeda(Calcula_Preco(VLStrCredito))
                    
                    StrSql = "Select * from tb_caixa"
                    RecCxa.Open StrSql, vgCon, 1, 3
                    
                    RecCxa.AddNew
                    RecCxa("Descr") = VPStrDescr
                    RecCxa("Vldeb") = "0"
                    RecCxa("Vlcred") = VPStrValCred
                    RecCxa("DtItem") = FormataDataUS(TxtDtCred.Text)
                    RecCxa.Update
                    
                    VPStrTemp = "update"
                Else
                    Unload Me
                End If
            
            End If
                'inserir cliente na tabela de lista de espera
                
                VPStrResponse = MsgBox("Inserir cliente na lista de espera?", vbYesNo)
                
                If VPStrResponse = vbYes Then
                    DataHora = FormataDataUS(Date) & " " & Time
                
                    StrSql = "Select * from tb_espera"
                    RecEsp.Open StrSql, vgCon, 1, 3
                    
                    RecEsp.AddNew
                    RecEsp("CodCli") = RecCart.Fields.Item(1).Value
                    RecEsp("Entrada") = DataHora
                    RecEsp.Update
                    
                Else
                    Unload Me
                End If
                
                If VPStrTemp = "insert" Then
                    VPStrBox = MsgBox("Créditos inseridos.", vbInformation, "Guide System - Informação")
                
                ElseIf VPStrTemp = "update" Then
                    VPStrBox = MsgBox("Créditos adicionados.", vbInformation, "Guide System - Informação")
                End If
                
                VPStrTemp = ""
                
                Unload Me
                Desconecta
                
                If VPStrResponse = vbYes Then
                    FrmMaquina.Show
                Else
                    VGStrTempCred = "naoespera"
                End If
                
        End If
        
        If VGStrTempCred = "" Then
            Desconecta
        End If
        
        VGStrTempCred = ""
    End If
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmCredito.hwnd)
    
    Height = 3390
    Width = 5295
    'Top = 1275
    'Left = 3465
    
    Unload FrmCartao
    
    If VGStrForm = "Cartao" Then
        'TxtCodCli.Text = FormataNum(VGIntCodCli)
        TxtNumCartao.Text = FormataNum(VGIntNumCartao)
        VGStrForm = ""
    End If
    
    TxtDtCred.Text = FormataData(Date)
    
    Screen.MousePointer = vbNormal
End Sub

'Sub VerificaCred()
'
'    If TxtCred.Text <> "" Then
'        VPIntHoraTemp = Mid(TxtCred.Text, 1, 2)
'        VPIntMinTemp = Mid(TxtCred.Text, 4, 2)
'        VPIntSegTemp = Mid(TxtCred.Text, 7, 2)
'
'        If VPIntHoraTemp > 99 Or VPIntMinTemp > 59 Or VPIntSegTemp > 59 Then
'            If VPIntHoraTemp > 99 Then
'                MsgHora = "- Hora acima do permitido. Máximo de 99 horas." & Chr(13)
'            End If
'
'            If VPIntMinTemp > 59 Then
'                MsgMin = "- Minuto acima do permitido. Máximo de 59 minutos." & Chr(13)
'            End If
'
'            If VPIntSegTemp > 59 Then
'                MsgSeg = "- Segundo acima do permitido. Máximo de 59 segundos." & Chr(13)
'            End If
'
'            VPStrBox = MsgBox("Erro(s) ocorrido(s):" & Chr(13) & Chr(13) & MsgHora & MsgMin & MsgSeg, vbCritical, "Guide System - Aviso de erro")
'            MsgHora = ""
'            MsgMin = ""
'            MsgSeg = ""
'        Else
'            VPStrGrava = "sim"
'        End If
'    End If
'End Sub

Private Sub Form_Resize()
  FrmCredito.Left = (MDIPrincipal.Width / 2) - (FrmCredito.Width / 1.93)
  FrmCredito.Top = (MDIPrincipal.Height / 3) - (FrmCredito.Height / 5)
End Sub

Private Sub TxtCred_GotFocus()
    TxtCred.SelStart = 0
    TxtCred.SelLength = Len(TxtCred.Text)
End Sub

Private Sub TxtCred_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtCred_LostFocus()
    TxtCred.Text = FormataCred(TxtCred.Text)
'    Dim VLStrData As String
'
'    If TxtCred.Text <> "" Then
'        VLStrData = VerificaData(TxtDtCred.Text)
'
'        If VGStrDataErro = "sim" Then
'            TxtDtCred.SetFocus
'        Else
'            TxtDtCred.Text = VLStrData
'        End If
'
'        VGStrDataErro = ""
'    Else
'        VPStrBox = MsgBox("Erro(s) ocorrido(s):" & Chr(13) & Chr(13) & MsgHora & MsgMin & MsgSeg, vbCritical, "Guide System - Aviso de erro")
'    End If


End Sub

Private Sub TxtCredHora_LostFocus()
    If TxtCredHora.Text = "" Or TxtCredHora.Text = "0" Then
        TxtCredHora.Text = "00"
    Else
        TxtCredHora.Text = FormataNum(TxtCredHora.Text)
    End If
End Sub

Private Sub TxtCredMin_LostFocus()
    If TxtCredMin.Text = "" Or TxtCredMin.Text = "0" Then
        TxtCredMin.Text = "00"
    Else
        TxtCredMin.Text = FormataNum(TxtCredMin.Text)
    End If
End Sub

Private Sub TxtDtCred_GotFocus()
    If TxtDtCred.Text = "__/__/____" Then
        TxtDtCred.Text = ""
    End If
    
    TxtDtCred.SelStart = 0
    TxtDtCred.SelLength = Len(TxtDtCred.Text)
End Sub

Private Sub TxtDtCred_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    
    If TxtDtCred.Text = "__/__/____" Then
        TxtDtCred.Text = ""
    End If
End Sub

Private Sub TxtDtCred_LostFocus()
    Dim VLStrData As String
    
    If TxtDtCred.Text <> "" Then
        VLStrData = VerificaData(TxtDtCred.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtCred.SetFocus
        Else
            TxtDtCred.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtCred.Text = "__/__/____"
    End If
End Sub

Private Sub TxtNumCartao_GotFocus()
    TxtNumCartao.SelStart = 0
    TxtNumCartao.SelLength = Len(TxtNumCartao.Text)
End Sub

Private Sub TxtNumCartao_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub
