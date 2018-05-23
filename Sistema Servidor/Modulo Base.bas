Attribute VB_Name = "ModuloBase"
Global vgCon As New ADODB.Connection
Global font0 As String
Global font1 As String
Global font2 As String
Global font3 As String
Global font4 As String
Global font5 As String
Global font6 As String

Global VGStrBox As String
Global VGStrResponse As String
Global VGStrForm As String
Global VGStrConsultCart As String
Global VGStrPreco As String
Global VGStrTempCred As String
Global StrCon As String
Global StrSql As String
Global StrSql1 As String
Global VGStrAlt As String
Global VGStrSenha As String
Global VGStrLocalTemp As String
Global VGStrImprimir As String

Global VGIntCodCli As Integer
Global VGIntCodItem As Integer
Global VGIntNumCartao As Integer
Global VGIntCodProd As Integer
Global VGIntCodCliTemp As Integer
Global VGIntQtde As Integer

Public Function Decipher(ByVal from_text As String) As String

Const MIN_ASC = 32  ' Space.
Const MAX_ASC = 126 ' ~.
Const NUM_ASC = MAX_ASC - MIN_ASC + 1

Dim offset As Long
Dim str_len As Integer
Dim i As Integer
Dim ch As Integer

    ' Initialize the random number generator.
    offset = 123
    Rnd -1
    Randomize offset

    ' Encipher the string.
    str_len = Len(from_text)
    For i = 1 To str_len
        ch = Asc(Mid$(from_text, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            Decipher = Decipher & Chr$(ch)
        End If
    Next i

End Function

Sub Conecta()
    'Dim vlFso, vlArquivo
    'Dim vgStr As String

    '--- Abre arquivo criptografado com a string de conexão
    'Set vlFso = CreateObject("Scripting.FileSystemObject")
    'Set vlArquivo = vlFso.OpenTextFile(App.Path & "\protocolo.dll", ForReading, False)
    
    '--- Descriptografa string de conexão
    'vgStr = Decipher(vlArquivo.ReadLine)
    
    '--- Abre conexão com SQL
    'vgCon.Open vgStr
        
    Set vgCon = New ADODB.Connection
    
    'Acesso a banco SQL Server
    'StrCon = "driver={SQL Server}; server=CAIXA;uid=sa;pwd=cursino;database=TribusVictrix"
    
    'Acesso a banco Access
    StrCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\guidesystem.mdb;Persist Security Info=False"
    
    vgCon.Open StrCon
End Sub

Sub Desconecta()
    vgCon.Close
End Sub

Function FormataDataUS(vData)
 
    Dim Dia As String
    Dim Mes As String
    Dim Ano As String
    
    Dia = DatePart("D", vData)
    Mes = DatePart("M", vData)
    Ano = DatePart("YYYY", vData)
     
    If Dia < 10 Then
     Dia = "0" & Dia
    End If
     
    If Mes < 10 Then
     Mes = "0" & Mes
    End If
     
    FormataDataUS = Ano & "/" & Mes & "/" & Dia
  
End Function

Function FormataData(vData)
 
    Dim Dia As String
    Dim Mes As String
    Dim Ano As String
    
    If IsDate(vData) = False Or IsNull(vdate) = True Then
        FormataData = ""
    Else
        Dia = DatePart("D", vData)
        Mes = DatePart("M", vData)
        Ano = DatePart("YYYY", vData)
         
        If Dia < 10 Then
         Dia = "0" & Dia
        End If
         
        If Mes < 10 Then
         Mes = "0" & Mes
        End If
         
        FormataData = Dia & "/" & Mes & "/" & Ano
    End If
End Function

Function FormataHora(vHora)
 
    Dim Hora As String
    Dim Min As String
    Dim Seg As String
    
    Hora = DatePart("H", vHora)
    Min = DatePart("N", vHora)
    Seg = DatePart("S", vHora)
    
    If Hora < 10 Then
     Hora = "0" & Hora
    End If
    
    If Min < 10 Then
     Min = "0" & Min
    End If
    
    If Seg < 10 Then
     Seg = "0" & Seg
    End If
     
    FormataHora = Hora & ":" & Min & ":" & Seg
  
End Function

Function FormataDataHora(vDataHora)
 
    Dim Dia As String
    Dim Mes As String
    Dim Ano As String
    
    Dim Hora As String
    Dim Min As String
    Dim Seg As String
    
    Hora = DatePart("H", vDataHora)
    Min = DatePart("N", vDataHora)
    Seg = DatePart("S", vDataHora)
    
    Dia = DatePart("D", vDataHora)
    Mes = DatePart("M", vDataHora)
    Ano = Mid(DatePart("YYYY", vDataHora), 3, 4)
     
    If Dia < 10 Then
     Dia = "0" & Dia
    End If
     
    If Mes < 10 Then
     Mes = "0" & Mes
    End If
     
    If Hora < 10 Then
     Hora = "0" & Hora
    End If
    
    If Min < 10 Then
     Min = "0" & Min
    End If
    
    If Seg < 10 Then
     Seg = "0" & Seg
    End If
     
    FormataDataHora = Dia & "/" & Mes & "/" & Ano & "    " & Hora & ":" & Min & ":" & Seg
  
End Function

Function FormataNum(vnum)
 
    Dim num As String
    
    If IsNull(vnum) = True Then
        num = 0
    Else
        If vnum > 0 And vnum < 10 Then
            FormataNum = 0 & vnum
        Else
            FormataNum = vnum
        End If
    End If
End Function

Function FormataDia(vDia)
 
    Dim Dia As String
    
    Dia = DatePart("D", vDia)
     
    If Dia < 10 Then
        Dia = "0" & Dia
    End If
    
    FormataDia = Dia
  
End Function

Function FormataMes(vMes)
 
    Dim Mes As Integer
    
    Mes = DatePart("M", vMes)
     
    'If Mes < 10 Then
    ' Mes = "0" & Mes
    'End If
     
    FormataMes = Mes
  
End Function

Function FormataNomeMes(vNomMes)
    
    If vNomMes = 1 Then
        FormataNomeMes = "Janeiro"
    ElseIf vNomMes = 2 Then
        FormataNomeMes = "Fevereiro"
    ElseIf vNomMes = 3 Then
        FormataNomeMes = "Março"
    ElseIf vNomMes = 4 Then
        FormataNomeMes = "Abril"
    ElseIf vNomMes = 5 Then
        FormataNomeMes = "Maio"
    ElseIf vNomMes = 6 Then
        FormataNomeMes = "Junho"
    ElseIf vNomMes = 7 Then
        FormataNomeMes = "Julho"
    ElseIf vNomMes = 8 Then
        FormataNomeMes = "Agosto"
    ElseIf vNomMes = 9 Then
        FormataNomeMes = "Setembro"
    ElseIf vNomMes = 10 Then
        FormataNomeMes = "Outubro"
    ElseIf vNomMes = 11 Then
        FormataNomeMes = "Novembro"
    ElseIf vNomMes = 12 Then
        FormataNomeMes = "Dezembro"
    End If
  
End Function

Function FormataNumMes(vNumMes)
    
    If vNumMes = "Janeiro" Then
        FormataNumMes = 1
    ElseIf vNumMes = "Fevereiro" Then
        FormataNumMes = 2
    ElseIf vNumMes = "Março" Then
        FormataNumMes = 3
    ElseIf vNumMes = "Abril" Then
        FormataNumMes = 4
    ElseIf vNumMes = "Maio" Then
        FormataNumMes = 5
    ElseIf vNumMes = "Junho" Then
        FormataNumMes = 6
    ElseIf vNumMes = "Julho" Then
        FormataNumMes = 7
    ElseIf vNumMes = "Agosto" Then
        FormataNumMes = 8
    ElseIf vNumMes = "Setembro" Then
        FormataNumMes = 9
    ElseIf vNumMes = "Outubro" Then
        FormataNumMes = 10
    ElseIf vNumMes = "Novembro" Then
        FormataNumMes = 11
    ElseIf vNumMes = "Dezembro" Then
        FormataNumMes = 12
    End If
  
End Function

'Function MontaIdade(vAno)
'    Dim AnoAtual As Integer
   
'    AnoAtual = DatePart("yyyy", Date)
    
'    MontaIdade = AnoAtual - vAno

'End Function

Function Calcula_Idade(vDia, vMes, vAno)
    
    Dim DiaAtual As Integer
    Dim MesAtual As Integer
    Dim AnoAtual As Integer
    Dim IdadeTemp As Integer
    Dim Idade As Integer
    
    DiaAtual = DatePart("d", Date)
    MesAtual = DatePart("m", Date)
    AnoAtual = DatePart("yyyy", Date)
    
    IdadeTemp = AnoAtual - vAno
    
    If vMes < MesAtual Then
        Idade = IdadeTemp
        
    ElseIf vMes > MesAtual Then
        Idade = IdadeTemp - 1
        
    ElseIf vMes = MesAtual Then
        
        If vDia < DiaAtual Then
            Idade = IdadeTemp
            
        ElseIf vDia > DiaAtual Then
            Idade = IdadeTemp - 1
        
        ElseIf vDia = DiaAtual Then
            Idade = IdadeTemp
            
        End If
    
    End If
    
    Calcula_Idade = Idade
    
End Function

Function Calcula_Preco(vCred)
    
    Dim Hora As String
    Dim Min As String
    Dim RecPr As New ADODB.Recordset
    Dim Moeda As Currency
    Dim Moeda2 As Currency
    Dim valorprop As Currency
    Dim valormin As Currency
    
    Hora = Mid(vCred, 1, 2)
    Min = Mid(vCred, 4, 2)
    
    StrSql = "Select Preco from tb_preco where CodProd=1"
    RecPr.Open StrSql, vgCon, 1, 3
    
    Moeda = RecPr.Fields.Item(0).Value
    
    If Min = "00" Or Min = "0" Then
        Calcula_Preco = CInt(Hora) * Moeda
    Else
        If Hora = "00" Or Hora = "0" Then
            If Min = "30" Then
                Calcula_Preco = Moeda / 2
            Else
                valorprop = Moeda / 60
                Calcula_Preco = Min * valorprop
            End If
        Else
            If Min = "30" Then
                valormin = Moeda / 2
            Else
                valormin = Min * (Moeda / 60)
            End If
            Calcula_Preco = (Hora * Moeda) + valormin
        End If
    End If
    
End Function

Function VerificaData(vData)
    Dim DiaTemp As String
    Dim MesTemp As String
    Dim AnoTemp As String

    If Len(vData) = 8 Then
        DiaTemp = Mid(vData, 1, 2)
        MesTemp = Mid(vData, 3, 2)
        AnoTemp = Mid(vData, 5)

        If DiaTemp > 31 Or MesTemp > 12 Then
            VPStrBox = MsgBox("Data inválida.", vbCritical, "Guide System - Aviso de erro")
            VGStrDataErro = "sim"
        
        ElseIf IsDate(DiaTemp & "/" & MesTemp & "/" & AnoTemp) = False Then
            VPStrBox = MsgBox("Data inválida.", vbCritical, "Guide System - Aviso de erro")
            VGStrDataErro = "sim"
            
        Else
            VerificaData = DiaTemp & "/" & MesTemp & "/" & AnoTemp
        End If
        
    ElseIf Len(vData) = 10 And InStr(vData, "/") And vData <> "__/__/____" Then
        DiaTemp = Mid(vData, 1, 2)
        MesTemp = Mid(vData, 4, 2)
        AnoTemp = Mid(vData, 7)

        If DiaTemp > 31 Or MesTemp > 12 Then
            VPStrBox = MsgBox("Data inválida.", vbCritical, "Guide System - Aviso de erro")
            VGStrDataErro = "sim"
        
        ElseIf IsDate(DiaTemp & "/" & MesTemp & "/" & AnoTemp) = False Then
            VPStrBox = MsgBox("Data inválida.", vbCritical, "Guide System - Aviso de erro")
            VGStrDataErro = "sim"
        
        Else
            VerificaData = DiaTemp & "/" & MesTemp & "/" & AnoTemp
        End If

    Else
        VPStrBox = MsgBox("Formato da data está incorreto.", vbCritical, "Guide System - Aviso de erro")
        VGStrDataErro = "sim"
    End If
End Function

Function FormataMoeda(pvalor)
    Dim valor As String
    Dim centavo As String
    Dim centavotemp As String
    Dim poscentavo As Integer
    Dim posreal As Integer
    Dim realtemp As String
    Dim real As String
    Dim lenreal As Integer
    Dim ponto As String
    Dim sinal As String
    
    If InStr(pvalor, "R$") = 0 Then
        pvalor = pvalor
    Else
        pvalor = Trim(Mid(pvalor, 3))
    End If
    
    If InStr(pvalor, "-") <> 0 Then
        sinal = Mid(pvalor, 1, 1)
        pvalor = Mid(pvalor, 2)
    End If
    
    valor = pvalor
    poscentavo = InStr(valor, ",")
    
    If poscentavo <> 0 Then
        centavotemp = Mid(valor, poscentavo)
        If Len(centavotemp) = 2 Then
            centavo = centavotemp & "0"
        ElseIf Len(centavotemp) > 2 Then
            If Len(centavotemp) > 3 Then
                ultcentavo = CInt(Mid(centavotemp, 4, 1))
                If ultcentavo > 5 Then
                    centavo = "," & CInt(Trim(Replace(Mid(centavotemp, 1, 3), ",", ""))) + 1
                Else
                    centavo = Mid(centavotemp, 1, 3)
                End If
            Else
                centavo = centavotemp
            End If
        End If
        
        realtemp = Mid(valor, 1, poscentavo - 1)
    Else
        centavo = ",00"
        realtemp = valor
    End If
    
    posreal = Len(realtemp)
    lenreal = 3

    Do While posreal <> 0
        real = Mid(realtemp, posreal, 1) & real
        
        If Len(real) = lenreal Then
            If Len(realtemp) <> 3 Then
                If Mid(realtemp, Len(realtemp) - 3, 1) <> "." Then
                    real = "." & real
                End If
            Else
                real = "." & real
            End If
            lenreal = Len(real) + 3
        End If
        
        posreal = posreal - 1
    Loop
    
    ponto = Mid(real, 1, 1)
    If ponto = "." Then
        real = Mid(real, 2)
    End If
        
    If sinal = "-" Then
        FormataMoeda = "R$ " & sinal & real & centavo
    Else
        FormataMoeda = "R$ " & real & centavo
    End If
End Function
