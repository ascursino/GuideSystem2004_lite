VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmMaquina 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Visualização de máquinas"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10425
   Icon            =   "FrmMaquina.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   10425
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   7680
      TabIndex        =   23
      Top             =   120
      Width           =   2415
      Begin VB.Frame FraSit 
         Caption         =   "Máquina 1"
         Height          =   1215
         Left            =   120
         TabIndex        =   35
         Top             =   2880
         Width           =   2175
         Begin VB.CommandButton CmdTrocar 
            Caption         =   "Trocar"
            Height          =   255
            Left            =   1200
            TabIndex        =   39
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton OptLivre 
            Caption         =   "Livre"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton OptOcup 
            Caption         =   "Ocupada"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton OptFora 
            Caption         =   "Fora de uso"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Height          =   15
         Left            =   120
         TabIndex        =   34
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CommandButton CmdFechaTodas 
         Caption         =   "Fechar todos os jogos"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Atualizar máquinas"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Atualiza status das máquinas"
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton CmdEspera 
         Caption         =   "Incluir na lista de espera"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Inclui na lista de espera"
         Top             =   2280
         Width           =   2175
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "Ok"
         Height          =   255
         Left            =   1680
         TabIndex        =   29
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox TxtFecharJogo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   28
         Top             =   600
         Width           =   495
      End
      Begin VB.Frame Frame4 
         Height          =   15
         Left            =   120
         TabIndex        =   26
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Height          =   15
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Height          =   15
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblFecharJogoInd 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmMaquina.frx":000C
         TabIndex        =   32
         Top             =   240
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblMaq 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmMaquina.frx":0096
         TabIndex        =   33
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame FraMaq 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Máquinas"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.PictureBox Img08 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   5160
         MouseIcon       =   "FrmMaquina.frx":0104
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaquina.frx":09CE
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   21
         Top             =   360
         Width           =   540
      End
      Begin VB.PictureBox Img09 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   5880
         MouseIcon       =   "FrmMaquina.frx":1942
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaquina.frx":220C
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   20
         Top             =   360
         Width           =   540
      End
      Begin VB.PictureBox Img10 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   6600
         MouseIcon       =   "FrmMaquina.frx":3180
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaquina.frx":3A4A
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   19
         Top             =   360
         Width           =   540
      End
      Begin VB.PictureBox Img07 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   4440
         MouseIcon       =   "FrmMaquina.frx":49BE
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaquina.frx":5288
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   15
         Top             =   360
         Width           =   540
      End
      Begin VB.PictureBox Img05 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   3000
         MouseIcon       =   "FrmMaquina.frx":61FC
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaquina.frx":6AC6
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   14
         Top             =   360
         Width           =   540
      End
      Begin VB.PictureBox Img04 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   2280
         MouseIcon       =   "FrmMaquina.frx":7A3A
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaquina.frx":8304
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   13
         Top             =   360
         Width           =   540
      End
      Begin VB.PictureBox Img03 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   1560
         MouseIcon       =   "FrmMaquina.frx":9278
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaquina.frx":9B42
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   12
         Top             =   360
         Width           =   540
      End
      Begin VB.PictureBox Img06 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   3720
         MouseIcon       =   "FrmMaquina.frx":AAB6
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaquina.frx":B380
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   11
         Top             =   360
         Width           =   540
      End
      Begin VB.PictureBox Img01 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   120
         MouseIcon       =   "FrmMaquina.frx":C2F4
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaquina.frx":CBBE
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   10
         Top             =   360
         Width           =   540
      End
      Begin VB.PictureBox Img02 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   840
         MouseIcon       =   "FrmMaquina.frx":DB32
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaquina.frx":E3FC
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   9
         Top             =   360
         Width           =   540
      End
      Begin ACTIVESKINLibCtl.SkinLabel Lbl02 
         Height          =   255
         Left            =   840
         OleObjectBlob   =   "FrmMaquina.frx":F370
         TabIndex        =   2
         Top             =   960
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel Lbl04 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "FrmMaquina.frx":F3D2
         TabIndex        =   3
         Top             =   960
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel Lbl06 
         Height          =   255
         Left            =   3720
         OleObjectBlob   =   "FrmMaquina.frx":F434
         TabIndex        =   4
         Top             =   960
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel Lbl01 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmMaquina.frx":F496
         TabIndex        =   5
         Top             =   960
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel Lbl03 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FrmMaquina.frx":F4F8
         TabIndex        =   6
         Top             =   960
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel Lbl05 
         Height          =   255
         Left            =   3000
         OleObjectBlob   =   "FrmMaquina.frx":F55A
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel Lbl07 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "FrmMaquina.frx":F5BC
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel Lbl08 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "FrmMaquina.frx":F61E
         TabIndex        =   16
         Top             =   960
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel Lbl10 
         Height          =   255
         Left            =   6600
         OleObjectBlob   =   "FrmMaquina.frx":F680
         TabIndex        =   17
         Top             =   960
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel Lbl09 
         Height          =   255
         Left            =   5880
         OleObjectBlob   =   "FrmMaquina.frx":F6E2
         TabIndex        =   18
         Top             =   960
         Width           =   495
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5640
      OleObjectBlob   =   "FrmMaquina.frx":F744
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6360
      Top             =   2400
   End
   Begin VB.Frame FraEspera 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Lista de espera"
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   7215
      Begin FPSpread.vaSpread GrdEspera 
         Height          =   2415
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   6990
         _Version        =   393216
         _ExtentX        =   12330
         _ExtentY        =   4260
         _StockProps     =   64
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   0
         MaxCols         =   3
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   12632256
         SpreadDesigner  =   "FrmMaquina.frx":F978
         UserResize      =   1
      End
   End
End
Attribute VB_Name = "FrmMaquina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPIntLinha As Integer
Public VPStrBox As String
Public VPStrNomCli As String
Public VPIntMaq As Integer
Public VPStrMaq1 As String
Public VPStrMaq2 As String
Public VPStrMaq3 As String
Public VPStrMaq4 As String
Public VPStrMaq5 As String
Public VPStrMaq6 As String
Public VPStrMaq7 As String
Public VPStrMaq8 As String
Public VPStrMaq9 As String
Public VPStrMaq10 As String

Sub MontaGridEspera()
    Screen.MousePointer = vbHourglass
    
    'criar lista de espera
    
    Conecta
    
    Dim RecResult As New ADODB.Recordset
    Dim RecCli As New ADODB.Recordset
    
    StrSql = "Select CodCli,Entrada from tb_espera order by Entrada"
    RecResult.Open StrSql, vgCon, 1, 3
        
    'If RecResult.EOF Then
    '       VPStrBox = MsgBox("Lista de espera vazia.", vbInformation, "Guide System - Informação")
    'End If
   
    VPIntLinha = 1
    
    GrdEspera.MaxRows = VPIntLinha
           
    Do While Not RecResult.EOF
        
        StrSql = "Select Nome from tb_cliente where CodCli=" & RecResult.Fields.Item(0).Value
        RecCli.Open StrSql, vgCon, 1, 3
        
        GrdEspera.Row = VPIntLinha
        GrdEspera.Lock = True
                        
        GrdEspera.Col = 1
        GrdEspera.Text = FormataData(RecResult.Fields.Item(1).Value) & "  " & FormataHora(RecResult.Fields.Item(1).Value)
        GrdEspera.Lock = True
        
        GrdEspera.Col = 2
        GrdEspera.Text = FormataNum(RecResult.Fields.Item(0).Value)
        GrdEspera.Lock = True
        
        If RecCli.EOF Then  'não achou nada
            VPStrNomCli = "CLIENTE EXCLUÍDO"
        Else
            VPStrNomCli = RecCli.Fields.Item(0).Value
        End If
        
        GrdEspera.Col = 3
        GrdEspera.Text = VPStrNomCli
        GrdEspera.Lock = True
        
        VPIntLinha = VPIntLinha + 1
        
        GrdEspera.MaxRows = GrdEspera.MaxRows + 1
        RecResult.MoveNext
        RecCli.Close
    Loop
    GrdEspera.MaxRows = GrdEspera.MaxRows - 1
    RecResult.Close
    
    Desconecta
    
    VGStrTempCred = "credito"
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdEspera_Click()
    Screen.MousePointer = vbHourglass
    VGStrAlt = "espera"
    FrmCadCliAlt.Show
    Screen.MousePointer = vbNormal
End Sub

Sub Inclui_Lista()
    Screen.MousePointer = vbHourglass
    
    Conecta
    
    Dim RecEsp As New ADODB.Recordset
    Dim DataHora As String
    
    DataHora = FormataDataUS(Date) & " " & Time

    StrSql = "Insert into tb_espera VALUES (" & VGIntCodCliTemp & "," & _
                 "'" & DataHora & "')"
    RecEsp.Open StrSql, vgCon, 1, 3
    
    Desconecta
    
    Me.MontaGridEspera
    
End Sub

Private Sub CmdFechaTodas_Click()
    Screen.MousePointer = vbHourglass
    
    Conecta
    
    Dim RecVerif As New ADODB.Recordset
    Dim RecFecha As New ADODB.Recordset
    
    StrSql = "Select NumMaq from tb_conect"
    RecVerif.Open StrSql, vgCon, 1, 3
    
    If Not RecVerif.EOF Then    'achou alguma máquina
        
        Do While Not RecVerif.EOF
        
            StrSql = "Select * from tb_comandos"
            RecFecha.Open StrSql, vgCon, 1, 3
            
            RecFecha.AddNew
            RecFecha("Maq") = RecVerif.Fields.Item(0).Value
            RecFecha.Update
            
            RecVerif.MoveNext
            'RecFecha.Clone
        Loop
    
        Desconecta
        
        VPStrBox = MsgBox("Concluído.", vbInformation, "Guide System - Informação")
    Else
        Desconecta
        VPStrBox = MsgBox("Máquinas não estão conectadas ao sistema.", vbInformation, "Guide System - Informação")
    End If
    
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub CmdOK_Click()
    Screen.MousePointer = vbHourglass
    
    If TxtFecharJogo.Text = "" Then
        VPStrBox = MsgBox("Preencha o número da máquina.", vbInformation, "Guide System - Informação")
    Else
        Conecta
        
        Dim RecLista As New ADODB.Recordset
        Dim RecVerif As New ADODB.Recordset
        
        StrSql = "Select NumMaq from tb_conect where NumMaq=" & TxtFecharJogo.Text
        RecVerif.Open StrSql, vgCon, 1, 3
        
        If Not RecVerif.EOF Then    'achou a máquina
            StrSql = "Select * from tb_comandos"
            RecLista.Open StrSql, vgCon, 1, 3
            
            RecLista.AddNew
            RecLista("Maq") = TxtFecharJogo.Text
            RecLista.Update
            
            TxtFecharJogo.Text = ""
        Else
            VPStrBox = MsgBox("Máquina não está conectada ao sistema.", vbInformation, "Guide System - Informação")
        End If
        Desconecta
    End If
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdRefresh_Click()
    Screen.MousePointer = vbHourglass
    
    Conecta
    
    Dim RecMaq As New ADODB.Recordset
    
    StrSql = "Select NumMaq,Situacao from tb_maquina"
    RecMaq.Open StrSql, vgCon, 1, 3

    Do While Not RecMaq.EOF
    
        If RecMaq.Fields.Item(0).Value = 1 Then     'máquina
            If RecMaq.Fields.Item(1).Value = "livre" Then       'situação
                Img01.Picture = LoadPicture(App.Path & "\MaqLivre.bmp")
                VPStrMaq1 = "livre"
            ElseIf RecMaq.Fields.Item(1).Value = "ocupado" Then
                Img01.Picture = LoadPicture(App.Path & "\MaqOcup.bmp")
                VPStrMaq1 = "ocupado"
            ElseIf RecMaq.Fields.Item(1).Value = "fora" Then
                Img01.Picture = LoadPicture(App.Path & "\MaqFora.bmp")
                VPStrMaq1 = "fora"
            End If
                    
        ElseIf RecMaq.Fields.Item(0).Value = 2 Then
            If RecMaq.Fields.Item(1).Value = "livre" Then
                Img02.Picture = LoadPicture(App.Path & "\MaqLivre.bmp")
                VPStrMaq2 = "livre"
            ElseIf RecMaq.Fields.Item(1).Value = "ocupado" Then
                Img02.Picture = LoadPicture(App.Path & "\MaqOcup.bmp")
                VPStrMaq2 = "ocupado"
            ElseIf RecMaq.Fields.Item(1).Value = "fora" Then
                Img02.Picture = LoadPicture(App.Path & "\MaqFora.bmp")
                VPStrMaq2 = "fora"
            End If
            
        ElseIf RecMaq.Fields.Item(0).Value = 3 Then
            If RecMaq.Fields.Item(1).Value = "livre" Then
                Img03.Picture = LoadPicture(App.Path & "\MaqLivre.bmp")
                VPStrMaq3 = "livre"
            ElseIf RecMaq.Fields.Item(1).Value = "ocupado" Then
                Img03.Picture = LoadPicture(App.Path & "\MaqOcup.bmp")
                VPStrMaq3 = "ocupado"
            ElseIf RecMaq.Fields.Item(1).Value = "fora" Then
                Img03.Picture = LoadPicture(App.Path & "\MaqFora.bmp")
                VPStrMaq3 = "fora"
            End If
        
        ElseIf RecMaq.Fields.Item(0).Value = 4 Then
            If RecMaq.Fields.Item(1).Value = "livre" Then
                Img04.Picture = LoadPicture(App.Path & "\MaqLivre.bmp")
                VPStrMaq4 = "livre"
            ElseIf RecMaq.Fields.Item(1).Value = "ocupado" Then
                Img04.Picture = LoadPicture(App.Path & "\MaqOcup.bmp")
                VPStrMaq4 = "ocupado"
            ElseIf RecMaq.Fields.Item(1).Value = "fora" Then
                Img04.Picture = LoadPicture(App.Path & "\MaqFora.bmp")
                VPStrMaq4 = "fora"
            End If
        
        ElseIf RecMaq.Fields.Item(0).Value = 5 Then
            If RecMaq.Fields.Item(1).Value = "livre" Then
                Img05.Picture = LoadPicture(App.Path & "\MaqLivre.bmp")
                VPStrMaq5 = "livre"
            ElseIf RecMaq.Fields.Item(1).Value = "ocupado" Then
                Img05.Picture = LoadPicture(App.Path & "\MaqOcup.bmp")
                VPStrMaq5 = "ocupado"
            ElseIf RecMaq.Fields.Item(1).Value = "fora" Then
                Img05.Picture = LoadPicture(App.Path & "\MaqFora.bmp")
                VPStrMaq5 = "fora"
            End If
        
        ElseIf RecMaq.Fields.Item(0).Value = 6 Then
            If RecMaq.Fields.Item(1).Value = "livre" Then
                Img06.Picture = LoadPicture(App.Path & "\MaqLivre.bmp")
                VPStrMaq6 = "livre"
            ElseIf RecMaq.Fields.Item(1).Value = "ocupado" Then
                Img06.Picture = LoadPicture(App.Path & "\MaqOcup.bmp")
                VPStrMaq6 = "ocupado"
            ElseIf RecMaq.Fields.Item(1).Value = "fora" Then
                Img06.Picture = LoadPicture(App.Path & "\MaqFora.bmp")
                VPStrMaq6 = "fora"
            End If
        
        ElseIf RecMaq.Fields.Item(0).Value = 7 Then
            If RecMaq.Fields.Item(1).Value = "livre" Then
                Img07.Picture = LoadPicture(App.Path & "\MaqLivre.bmp")
                VPStrMaq7 = "livre"
            ElseIf RecMaq.Fields.Item(1).Value = "ocupado" Then
                Img07.Picture = LoadPicture(App.Path & "\MaqOcup.bmp")
                VPStrMaq7 = "ocupado"
            ElseIf RecMaq.Fields.Item(1).Value = "fora" Then
                Img07.Picture = LoadPicture(App.Path & "\MaqFora.bmp")
                VPStrMaq7 = "fora"
            End If
        
        ElseIf RecMaq.Fields.Item(0).Value = 8 Then
            If RecMaq.Fields.Item(1).Value = "livre" Then
                Img08.Picture = LoadPicture(App.Path & "\MaqLivre.bmp")
                VPStrMaq8 = "livre"
            ElseIf RecMaq.Fields.Item(1).Value = "ocupado" Then
                Img08.Picture = LoadPicture(App.Path & "\MaqOcup.bmp")
                VPStrMaq8 = "ocupado"
            ElseIf RecMaq.Fields.Item(1).Value = "fora" Then
                Img08.Picture = LoadPicture(App.Path & "\MaqFora.bmp")
                VPStrMaq8 = "fora"
            End If
        
        ElseIf RecMaq.Fields.Item(0).Value = 9 Then
            If RecMaq.Fields.Item(1).Value = "livre" Then
                Img09.Picture = LoadPicture(App.Path & "\MaqLivre.bmp")
                VPStrMaq9 = "livre"
            ElseIf RecMaq.Fields.Item(1).Value = "ocupado" Then
                Img09.Picture = LoadPicture(App.Path & "\MaqOcup.bmp")
                VPStrMaq9 = "ocupado"
            ElseIf RecMaq.Fields.Item(1).Value = "fora" Then
                Img09.Picture = LoadPicture(App.Path & "\MaqFora.bmp")
                VPStrMaq9 = "fora"
            End If
        
        ElseIf RecMaq.Fields.Item(0).Value = 10 Then
            If RecMaq.Fields.Item(1).Value = "livre" Then
                Img10.Picture = LoadPicture(App.Path & "\MaqLivre.bmp")
                VPStrMaq10 = "livre"
            ElseIf RecMaq.Fields.Item(1).Value = "ocupado" Then
                Img10.Picture = LoadPicture(App.Path & "\MaqOcup.bmp")
                VPStrMaq10 = "ocupado"
            ElseIf RecMaq.Fields.Item(1).Value = "fora" Then
                Img10.Picture = LoadPicture(App.Path & "\MaqFora.bmp")
                VPStrMaq10 = "fora"
            End If
        
        End If
        
        RecMaq.MoveNext
    Loop
    
    Desconecta
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdTrocar_Click()
    Screen.MousePointer = vbHourglass
    Dim VLStrSit As String
    
    Conecta
    
    If OptLivre.Value = True Then
        VLStrSit = "livre"
    ElseIf OptOcup.Value = True Then
        VLStrSit = "ocupado"
    ElseIf OptFora.Value = True Then
        VLStrSit = "fora"
    End If
        
    vgCon.Execute ("Update tb_maquina set Situacao='" & VLStrSit & "' where NumMaq=" & VPIntMaq)
    
    Desconecta
    
    CmdRefresh.Value = True
    
    FraSit.Visible = False
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmMaquina.hwnd)
    
    Height = 5145
    Width = 10545
    'Top = 1275
    'Left = 30
    
    FraSit.Visible = False
    
    CmdRefresh.Value = True
    
    Me.MontaGridEspera
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  FrmMaquina.Left = (MDIPrincipal.Width / 2) - (FrmMaquina.Width / 1.93)
  FrmMaquina.Top = (MDIPrincipal.Height / 3) - (FrmMaquina.Height / 5)
End Sub

Private Sub GrdEspera_DblClick(ByVal Col As Long, ByVal Row As Long)

    VPStrResponse = MsgBox("Retirar da lista?", vbYesNo)
    
    If VPStrResponse = vbYes Then
        Conecta
        
        Dim RecLista As New ADODB.Recordset
        
        GrdEspera.Row = Row
        GrdEspera.Col = 2
        
        StrSql = "Delete from tb_espera where CodCli=" & GrdEspera.Text
        RecLista.Open StrSql, vgCon, 1, 3
        
        Desconecta
        
        Me.MontaGridEspera

    End If

End Sub

Private Sub Img01_Click()
    FraSit.Visible = True
    FraSit.Caption = "Máquina 1"
    VPIntMaq = 1
    If VPStrMaq1 = "livre" Then
        OptLivre = True
    ElseIf VPStrMaq1 = "ocupado" Then
        OptOcup = True
    ElseIf VPStrMaq1 = "fora" Then
        OptFora = True
    End If
End Sub

Private Sub Img02_Click()
    FraSit.Visible = True
    FraSit.Caption = "Máquina 2"
    VPIntMaq = 2
    If VPStrMaq2 = "livre" Then
        OptLivre = True
    ElseIf VPStrMaq2 = "ocupado" Then
        OptOcup = True
    ElseIf VPStrMaq2 = "fora" Then
        OptFora = True
    End If
End Sub

Private Sub Img03_Click()
    FraSit.Visible = True
    FraSit.Caption = "Máquina 3"
    VPIntMaq = 3
    If VPStrMaq3 = "livre" Then
        OptLivre = True
    ElseIf VPStrMaq3 = "ocupado" Then
        OptOcup = True
    ElseIf VPStrMaq3 = "fora" Then
        OptFora = True
    End If
End Sub

Private Sub Img04_Click()
    FraSit.Visible = True
    FraSit.Caption = "Máquina 4"
    VPIntMaq = 4
    If VPStrMaq4 = "livre" Then
        OptLivre = True
    ElseIf VPStrMaq4 = "ocupado" Then
        OptOcup = True
    ElseIf VPStrMaq4 = "fora" Then
        OptFora = True
    End If
End Sub

Private Sub Img05_Click()
    FraSit.Visible = True
    FraSit.Caption = "Máquina 5"
    VPIntMaq = 5
    If VPStrMaq5 = "livre" Then
        OptLivre = True
    ElseIf VPStrMaq5 = "ocupado" Then
        OptOcup = True
    ElseIf VPStrMaq5 = "fora" Then
        OptFora = True
    End If
End Sub

Private Sub Img06_Click()
    FraSit.Visible = True
    FraSit.Caption = "Máquina 6"
    VPIntMaq = 6
    If VPStrMaq6 = "livre" Then
        OptLivre = True
    ElseIf VPStrMaq6 = "ocupado" Then
        OptOcup = True
    ElseIf VPStrMaq6 = "fora" Then
        OptFora = True
    End If
End Sub

Private Sub Img07_Click()
    FraSit.Visible = True
    FraSit.Caption = "Máquina 7"
    VPIntMaq = 7
    If VPStrMaq7 = "livre" Then
        OptLivre = True
    ElseIf VPStrMaq7 = "ocupado" Then
        OptOcup = True
    ElseIf VPStrMaq7 = "fora" Then
        OptFora = True
    End If
End Sub

Private Sub Img08_Click()
    FraSit.Visible = True
    FraSit.Caption = "Máquina 8"
    VPIntMaq = 8
    If VPStrMaq8 = "livre" Then
        OptLivre = True
    ElseIf VPStrMaq8 = "ocupado" Then
        OptOcup = True
    ElseIf VPStrMaq8 = "fora" Then
        OptFora = True
    End If
End Sub

Private Sub Img09_Click()
    FraSit.Visible = True
    FraSit.Caption = "Máquina 9"
    VPIntMaq = 9
    If VPStrMaq9 = "livre" Then
        OptLivre = True
    ElseIf VPStrMaq9 = "ocupado" Then
        OptOcup = True
    ElseIf VPStrMaq9 = "fora" Then
        OptFora = True
    End If
End Sub

Private Sub Img10_Click()
    FraSit.Visible = True
    FraSit.Caption = "Máquina 10"
    VPIntMaq = 10
    If VPStrMaq10 = "livre" Then
        OptLivre = True
    ElseIf VPStrMaq10 = "ocupado" Then
        OptOcup = True
    ElseIf VPStrMaq10 = "fora" Then
        OptFora = True
    End If
End Sub

Private Sub TxtFecharJogo_GotFocus()
    TxtFecharJogo.SelStart = 0
    TxtFecharJogo.SelLength = Len(TxtFecharJogo.Text)
End Sub

'Private Sub Timer1_Timer()

'    CmdRefresh.Value = True

'End Sub

Private Sub TxtFecharJogo_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub
