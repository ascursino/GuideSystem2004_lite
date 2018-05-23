VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmNiver 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aniversariantes"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   Icon            =   "FrmNiver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraConsulta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Consulta"
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2535
      Begin VB.CommandButton CmdPesquisar 
         Caption         =   "Pesquisar"
         Height          =   375
         Left            =   720
         TabIndex        =   4
         ToolTipText     =   "Pesquisar aniversariantes"
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox CboAteMes 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Text            =   "CboAteMes"
         ToolTipText     =   "Mês final"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox CboDeMes 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Text            =   "CboDeMes"
         ToolTipText     =   "Mês inicial"
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox CboAteDia 
         Height          =   315
         Left            =   600
         TabIndex        =   2
         Text            =   "CboAteDia"
         ToolTipText     =   "Dia final"
         Top             =   1080
         Width           =   615
      End
      Begin VB.ComboBox CboDeDia 
         Height          =   315
         ItemData        =   "FrmNiver.frx":000C
         Left            =   600
         List            =   "FrmNiver.frx":000E
         TabIndex        =   0
         Text            =   "CboDeDia"
         ToolTipText     =   "Dia inicial"
         Top             =   600
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDia 
         Height          =   255
         Left            =   600
         OleObjectBlob   =   "FrmNiver.frx":0010
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblMes 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "FrmNiver.frx":0074
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDe 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmNiver.frx":00D8
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblPara 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmNiver.frx":013C
         TabIndex        =   5
         Top             =   1080
         Width           =   495
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   480
      OleObjectBlob   =   "FrmNiver.frx":01A4
      Top             =   720
   End
   Begin FPSpread.vaSpread GrdNiver 
      Height          =   4935
      Left            =   2760
      TabIndex        =   10
      Top             =   600
      Width           =   5190
      _Version        =   393216
      _ExtentX        =   9155
      _ExtentY        =   8705
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
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
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   12632256
      SpreadDesigner  =   "FrmNiver.frx":03D8
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblMesNiver 
      Height          =   375
      Left            =   2760
      OleObjectBlob   =   "FrmNiver.frx":0791
      TabIndex        =   11
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "FrmNiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrMes As String
Public VPStrData As String
Public VPStrAno As String
Public VPStrBox As String
Public VPIntLinha As Integer
Public RecResult As New ADODB.Recordset

Private Sub CmdPesquisar_Click()
    Screen.MousePointer = vbHourglass
    
    Conecta
    
    StrSql = "Select Nome,NascDia,NascMes,NascAno from tb_cliente where 0=0"
        
    If CboDeDia.Text <> "" And CboDeMes.Text <> "" Then
        StrSql = StrSql + " and NascDia >= " & CboDeDia.Text & " and NascMes >= " & FormataNumMes(CboDeMes.Text) & ""
    End If

    If CboAteDia.Text <> "" And CboAteMes.Text <> "" Then
        StrSql = StrSql + " and NascDia <= " & CboAteDia.Text & " and NascMes <= " & FormataNumMes(CboAteMes.Text) & ""
    End If

    StrSql = StrSql + " order by NascMes,NascDia asc"
    
    RecResult.Open StrSql, vgCon, 1, 3
    
    If (CboDeMes.Text <> CboAteMes.Text) And (CboDeMes.Text <> "" And CboAteMes.Text <> "") Then
        LblMesNiver.Caption = CboDeMes.Text & " a " & CboAteMes.Text & "/" & Year(Now())
    ElseIf CboDeDia.Text = "" And CboDeMes.Text = "" And CboAteDia.Text = "" And CboAteMes.Text = "" Then
        LblMesNiver.Caption = "Ano de " & Year(Now())
    Else
        LblMesNiver.Caption = CboDeMes.Text & "/" & Year(Now())
    End If
    
    Call MontaGrid
    
    Desconecta
    
    Screen.MousePointer = vbNormal
End Sub

Sub MontaGrid()
    
    If RecResult.EOF Then
           VPStrBox = MsgBox("Não existem aniversariantes para esta data.", vbInformation, "Guide System - Informação")
    End If
   
    VPIntLinha = 1
    
    GrdNiver.MaxRows = VPIntLinha
           
    Do While Not RecResult.EOF
        
        GrdNiver.Row = VPIntLinha
        GrdNiver.Lock = True
                        
        GrdNiver.Col = 1
        If RecResult.Fields.Item(1).Value <> "0" And RecResult.Fields.Item(2).Value <> "0" Then
            GrdNiver.Text = FormataNum(RecResult.Fields.Item(1).Value) & "/" & FormataNum(RecResult.Fields.Item(2).Value)
        Else
            GrdNiver.Text = ""
        End If
        GrdNiver.Lock = True
        
        GrdNiver.Col = 2
        GrdNiver.Text = RecResult.Fields.Item(0).Value
        GrdNiver.Lock = True
        
        GrdNiver.Col = 3
        GrdNiver.Text = Val(Calcula_Idade(RecResult.Fields.Item(1).Value, RecResult.Fields.Item(2).Value, RecResult.Fields.Item(3).Value))
        GrdNiver.Lock = True
           
        VPIntLinha = VPIntLinha + 1
        
        GrdNiver.MaxRows = GrdNiver.MaxRows + 1
        RecResult.MoveNext
    Loop
    GrdNiver.MaxRows = GrdNiver.MaxRows - 1
    RecResult.Close
    
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmNiver.hwnd)
    
    Height = 6240
    Width = 8520
    'Top = 1275
    'Left = 1650
   
    Call PegaMes

    LblMesNiver.Caption = VPStrMes

    Call MontaCboDia
    CboDeDia.Text = FormataDia(Date)
    CboAteDia.Text = FormataDia(Date)

    Call MontaCboMes
    CboDeMes.Text = FormataNomeMes(FormataMes(Date))
    CboAteMes.Text = FormataNomeMes(FormataMes(Date))

    CmdPesquisar.Value = True
    
    Screen.MousePointer = vbNormal
End Sub

Sub PegaMes()
    VPStrData = DatePart("m", Date)
    VPStrAno = DatePart("yyyy", Date)
    
    If VPStrData = "1" Then
        VPStrMes = "Janeiro/" & VPStrAno
        
    ElseIf VPStrData = "2" Then
        VPStrMes = "Fevereiro/" & VPStrAno
        
    ElseIf VPStrData = "3" Then
        VPStrMes = "Março/" & VPStrAno
        
    ElseIf VPStrData = "4" Then
        VPStrMes = "Abril/" & VPStrAno
        
    ElseIf VPStrData = "5" Then
        VPStrMes = "Maio/" & VPStrAno
        
    ElseIf VPStrData = "6" Then
        VPStrMes = "Junho/" & VPStrAno
        
    ElseIf VPStrData = "7" Then
        VPStrMes = "Julho/" & VPStrAno
        
    ElseIf VPStrData = "8" Then
        VPStrMes = "Agosto/" & VPStrAno
        
    ElseIf VPStrData = "9" Then
        VPStrMes = "Setembro/" & VPStrAno
        
    ElseIf VPStrData = "10" Then
        VPStrMes = "Outubro/" & VPStrAno
        
    ElseIf VPStrData = "11" Then
        VPStrMes = "Novembro/" & VPStrAno
        
    ElseIf VPStrData = "12" Then
        VPStrMes = "Dezembro/" & VPStrAno
    End If
        
End Sub

Sub MontaCboDia()
    
    A = 1
    Do While A < 32
         CboDeDia.AddItem (FormataNum(A))
         CboAteDia.AddItem (FormataNum(A))
         A = A + 1
    Loop

End Sub

Sub MontaCboMes()
        
    A = 1
    Do While A < 13
         CboDeMes.AddItem (FormataNomeMes(A))
         CboAteMes.AddItem (FormataNomeMes(A))
         A = A + 1
    Loop

End Sub

Private Sub Form_Resize()
  FrmNiver.Left = (MDIPrincipal.Width / 2) - (FrmNiver.Width / 1.93)
  FrmNiver.Top = (MDIPrincipal.Height / 3) - (FrmNiver.Height / 5)
End Sub
