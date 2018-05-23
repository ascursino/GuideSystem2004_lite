VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmMaqSituacao 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alteração da situação das máquinas"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   Icon            =   "FrmMaqSituacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAlterar 
      Caption         =   "Confirmar alteração"
      Height          =   495
      Left            =   2880
      TabIndex        =   40
      Top             =   3480
      Width           =   2655
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8040
      OleObjectBlob   =   "FrmMaqSituacao.frx":000C
      Top             =   120
   End
   Begin VB.Frame Frame10 
      Caption         =   "Máquina 10"
      Height          =   1455
      Left            =   6840
      TabIndex        =   36
      Top             =   1800
      Width           =   1455
      Begin VB.OptionButton OptMaq10L 
         Caption         =   "Livre"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptMaq10O 
         Caption         =   "Ocupada"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptMaq10F 
         Caption         =   "Fora de uso"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Máquina 8"
      Height          =   1455
      Left            =   3480
      TabIndex        =   32
      Top             =   1800
      Width           =   1455
      Begin VB.OptionButton OptMaq8L 
         Caption         =   "Livre"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptMaq8O 
         Caption         =   "Ocupada"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptMaq8F 
         Caption         =   "Fora de uso"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Máquina 7"
      Height          =   1455
      Left            =   1800
      TabIndex        =   28
      Top             =   1800
      Width           =   1455
      Begin VB.OptionButton OptMaq7L 
         Caption         =   "Livre"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptMaq7O 
         Caption         =   "Ocupada"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptMaq7F 
         Caption         =   "Fora de uso"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Máquina 6"
      Height          =   1455
      Left            =   120
      TabIndex        =   24
      Top             =   1800
      Width           =   1455
      Begin VB.OptionButton OptMaq6L 
         Caption         =   "Livre"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptMaq6O 
         Caption         =   "Ocupada"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptMaq6F 
         Caption         =   "Fora de uso"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Máquina 5"
      Height          =   1455
      Left            =   6840
      TabIndex        =   20
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton OptMaq5F 
         Caption         =   "Fora de uso"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton OptMaq5O 
         Caption         =   "Ocupada"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptMaq5L 
         Caption         =   "Livre"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Máquina 4"
      Height          =   1455
      Left            =   5160
      TabIndex        =   16
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton OptMaq4F 
         Caption         =   "Fora de uso"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton OptMaq4O 
         Caption         =   "Ocupada"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptMaq4L 
         Caption         =   "Livre"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Máquina 3"
      Height          =   1455
      Left            =   3480
      TabIndex        =   12
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton OptMaq3F 
         Caption         =   "Fora de uso"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton OptMaq3O 
         Caption         =   "Ocupada"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptMaq3L 
         Caption         =   "Livre"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Máquina 2"
      Height          =   1455
      Left            =   1800
      TabIndex        =   8
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton OptMaq2F 
         Caption         =   "Fora de uso"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton OptMaq2O 
         Caption         =   "Ocupada"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptMaq2L 
         Caption         =   "Livre"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Máquina 1"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton OptMaq1F 
         Caption         =   "Fora de uso"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton OptMaq1O 
         Caption         =   "Ocupada"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptMaq1L 
         Caption         =   "Livre"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Máquina 9"
      Height          =   1455
      Left            =   5160
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
      Begin VB.OptionButton OptMaq9L 
         Caption         =   "Livre"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptMaq9O 
         Caption         =   "Ocupada"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptMaq9F 
         Caption         =   "Fora de uso"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmMaqSituacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
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

Private Sub CmdAlterar_Click()
    Screen.MousePointer = vbHourglass
    Conecta

    For i = 1 To 10
        If i = 1 Then
            vgCon.Execute ("Update tb_maquina set Situacao='" & VPStrMaq1 & "' where NumMaq=" & i)
        ElseIf i = 2 Then
            vgCon.Execute ("Update tb_maquina set Situacao='" & VPStrMaq2 & "' where NumMaq=" & i)
        ElseIf i = 3 Then
            vgCon.Execute ("Update tb_maquina set Situacao='" & VPStrMaq3 & "' where NumMaq=" & i)
        ElseIf i = 4 Then
            vgCon.Execute ("Update tb_maquina set Situacao='" & VPStrMaq4 & "' where NumMaq=" & i)
        ElseIf i = 5 Then
            vgCon.Execute ("Update tb_maquina set Situacao='" & VPStrMaq5 & "' where NumMaq=" & i)
        ElseIf i = 6 Then
            vgCon.Execute ("Update tb_maquina set Situacao='" & VPStrMaq6 & "' where NumMaq=" & i)
        ElseIf i = 7 Then
            vgCon.Execute ("Update tb_maquina set Situacao='" & VPStrMaq7 & "' where NumMaq=" & i)
        ElseIf i = 8 Then
            vgCon.Execute ("Update tb_maquina set Situacao='" & VPStrMaq8 & "' where NumMaq=" & i)
        ElseIf i = 9 Then
            vgCon.Execute ("Update tb_maquina set Situacao='" & VPStrMaq9 & "' where NumMaq=" & i)
        ElseIf i = 10 Then
            vgCon.Execute ("Update tb_maquina set Situacao='" & VPStrMaq10 & "' where NumMaq=" & i)
        End If
    Next
    
    Desconecta
    
    Unload Me
    
    FrmMaquina.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmMaqSituacao.hwnd)
    
    Height = 4680
    Width = 8730
    
    Conecta

    Dim RecMaq As New ADODB.Recordset

    StrSql = "Select * from tb_maquina order by NumMaq"
    RecMaq.Open StrSql, vgCon, 1, 3

    Do While Not RecMaq.EOF
        If RecMaq("NumMaq") = 1 Then
            If RecMaq("Situacao") = "livre" Then
                OptMaq1L.Value = True
            ElseIf RecMaq("Situacao") = "ocupado" Then
                OptMaq1O.Value = True
            ElseIf RecMaq("Situacao") = "fora" Then
                OptMaq1F.Value = True
            End If
        
        ElseIf RecMaq("NumMaq") = 2 Then
            If RecMaq("Situacao") = "livre" Then
                OptMaq2L.Value = True
            ElseIf RecMaq("Situacao") = "ocupado" Then
                OptMaq2O.Value = True
            ElseIf RecMaq("Situacao") = "fora" Then
                OptMaq2F.Value = True
            End If
        
        ElseIf RecMaq("NumMaq") = 3 Then
            If RecMaq("Situacao") = "livre" Then
                OptMaq3L.Value = True
            ElseIf RecMaq("Situacao") = "ocupado" Then
                OptMaq3O.Value = True
            ElseIf RecMaq("Situacao") = "fora" Then
                OptMaq3F.Value = True
            End If
        
        ElseIf RecMaq("NumMaq") = 4 Then
            If RecMaq("Situacao") = "livre" Then
                OptMaq4L.Value = True
            ElseIf RecMaq("Situacao") = "ocupado" Then
                OptMaq4O.Value = True
            ElseIf RecMaq("Situacao") = "fora" Then
                OptMaq4F.Value = True
            End If
        
        ElseIf RecMaq("NumMaq") = 5 Then
            If RecMaq("Situacao") = "livre" Then
                OptMaq5L.Value = True
            ElseIf RecMaq("Situacao") = "ocupado" Then
                OptMaq5O.Value = True
            ElseIf RecMaq("Situacao") = "fora" Then
                OptMaq5F.Value = True
            End If
        
        ElseIf RecMaq("NumMaq") = 6 Then
            If RecMaq("Situacao") = "livre" Then
                OptMaq6L.Value = True
            ElseIf RecMaq("Situacao") = "ocupado" Then
                OptMaq6O.Value = True
            ElseIf RecMaq("Situacao") = "fora" Then
                OptMaq6F.Value = True
            End If
        
        ElseIf RecMaq("NumMaq") = 7 Then
            If RecMaq("Situacao") = "livre" Then
                OptMaq7L.Value = True
            ElseIf RecMaq("Situacao") = "ocupado" Then
                OptMaq7O.Value = True
            ElseIf RecMaq("Situacao") = "fora" Then
                OptMaq7F.Value = True
            End If
        
        ElseIf RecMaq("NumMaq") = 8 Then
            If RecMaq("Situacao") = "livre" Then
                OptMaq8L.Value = True
            ElseIf RecMaq("Situacao") = "ocupado" Then
                OptMaq8O.Value = True
            ElseIf RecMaq("Situacao") = "fora" Then
                OptMaq8F.Value = True
            End If
        
        ElseIf RecMaq("NumMaq") = 9 Then
            If RecMaq("Situacao") = "livre" Then
                OptMaq9L.Value = True
            ElseIf RecMaq("Situacao") = "ocupado" Then
                OptMaq9O.Value = True
            ElseIf RecMaq("Situacao") = "fora" Then
                OptMaq9F.Value = True
            End If
        
        ElseIf RecMaq("NumMaq") = 10 Then
            If RecMaq("Situacao") = "livre" Then
                OptMaq10L.Value = True
            ElseIf RecMaq("Situacao") = "ocupado" Then
                OptMaq10O.Value = True
            ElseIf RecMaq("Situacao") = "fora" Then
                OptMaq10F.Value = True
            End If
        
        End If
    
        RecMaq.MoveNext
    Loop
    
    Desconecta
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
    FrmMaqSituacao.Left = (MDIPrincipal.Width / 2) - (FrmMaqSituacao.Width / 1.93)
    FrmMaqSituacao.Top = (MDIPrincipal.Height / 3) - (FrmMaqSituacao.Height / 5)
End Sub

Private Sub OptMaq1L_Click()
    VPStrMaq1 = "livre"
End Sub

Private Sub OptMaq1O_Click()
    VPStrMaq1 = "ocupado"
End Sub

Private Sub OptMaq1F_Click()
    VPStrMaq1 = "fora"
End Sub

Private Sub OptMaq2L_Click()
    VPStrMaq2 = "livre"
End Sub

Private Sub OptMaq2O_Click()
    VPStrMaq2 = "ocupado"
End Sub

Private Sub OptMaq2F_Click()
    VPStrMaq2 = "fora"
End Sub

Private Sub OptMaq3L_Click()
    VPStrMaq3 = "livre"
End Sub

Private Sub OptMaq3O_Click()
    VPStrMaq3 = "ocupado"
End Sub

Private Sub OptMaq3F_Click()
    VPStrMaq3 = "fora"
End Sub

Private Sub OptMaq4L_Click()
    VPStrMaq4 = "livre"
End Sub

Private Sub OptMaq4O_Click()
    VPStrMaq4 = "ocupado"
End Sub

Private Sub OptMaq4F_Click()
    VPStrMaq4 = "fora"
End Sub

Private Sub OptMaq5L_Click()
    VPStrMaq5 = "livre"
End Sub

Private Sub OptMaq5O_Click()
    VPStrMaq5 = "ocupado"
End Sub

Private Sub OptMaq5F_Click()
    VPStrMaq5 = "fora"
End Sub

Private Sub OptMaq6L_Click()
    VPStrMaq6 = "livre"
End Sub

Private Sub OptMaq6O_Click()
    VPStrMaq6 = "ocupado"
End Sub

Private Sub OptMaq6F_Click()
    VPStrMaq6 = "fora"
End Sub

Private Sub OptMaq7L_Click()
    VPStrMaq7 = "livre"
End Sub

Private Sub OptMaq7O_Click()
    VPStrMaq7 = "ocupado"
End Sub

Private Sub OptMaq7F_Click()
    VPStrMaq7 = "fora"
End Sub

Private Sub OptMaq8L_Click()
    VPStrMaq8 = "livre"
End Sub

Private Sub OptMaq8O_Click()
    VPStrMaq8 = "ocupado"
End Sub

Private Sub OptMaq8F_Click()
    VPStrMaq8 = "fora"
End Sub

Private Sub OptMaq9L_Click()
    VPStrMaq9 = "livre"
End Sub

Private Sub OptMaq9O_Click()
    VPStrMaq9 = "ocupado"
End Sub

Private Sub OptMaq9F_Click()
    VPStrMaq9 = "fora"
End Sub

Private Sub OptMaq10L_Click()
    VPStrMaq10 = "livre"
End Sub

Private Sub OptMaq10O_Click()
    VPStrMaq10 = "ocupado"
End Sub

Private Sub OptMaq10F_Click()
    VPStrMaq10 = "fora"
End Sub


