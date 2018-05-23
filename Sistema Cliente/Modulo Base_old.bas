Attribute VB_Name = "ModuloBase"
Global vgCon As New ADODB.Connection
Global StrCon As String
Global StrSql As String

Global VGStrJogo As String
Global VGIntCodCli As Integer
Global VGIntNumCartao As Integer
Global VGStrCredito As String
Global VGStrCredRest As String
Global VGStrDataEntr As String
Global VGStrHoraEntr As String
Global VGStrDataSaida As String
Global VGStrHoraSaida As String
Global VGIntMaq As Integer
Global VGIntIdade As Integer
Global VGStrBox As String
Global VGStrIP As String
Global VGStrRestIdade As String
Global VGStrExecJogo As String
Global VGStrExecJogo1 As String
Global Contador As Double

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
    Dim vlFso, vlArquivo
    Dim vgStr As String

    '--- Abre arquivo criptografado com a string de conexão
    'Set vlFso = CreateObject("Scripting.FileSystemObject")
    'Set vlArquivo = vlFso.OpenTextFile(App.Path & "\client.dll", ForReading, False)
    
    '--- Descriptografa string de conexão
    'vgStr = Decipher(vlArquivo.ReadLine)
    
    '--- Abre conexão com SQL
    'vgCon.Open vgStr
        
    Set vgCon = New ADODB.Connection
    
    StrCon = "driver={SQL Server};" & _
      "server=CAIXA;uid=sa;pwd=cursino;database=TribusVictrix"
      
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
    
    If vnum < 10 Then
        num = 0 & vnum
    Else
        num = Val(vnum)
    End If
    
    FormataNum = num
  
End Function

Function FormataDia(vDia)
 
    Dim Dia As Integer
    
    Dia = DatePart("D", vDia)
     
    If Dia < 10 Then
        Dia = "0" & Dia
    End If
    
    FormataDia = Dia
  
End Function

Function FormataMes(vMes)
 
    Dim Mes As Integer
    
    Mes = DatePart("M", vData)
     
    If Mes < 10 Then
     Mes = "0" & Mes
    End If
     
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

Function Calcula_Maq(vIP)

    If vIP = "192.168.0.39" Then
        Calcula_Maq = "01"
    
    ElseIf vIP = "192.168.0.2" Then
        Calcula_Maq = "02"

    ElseIf vIP = "192.168.0.3" Then
        Calcula_Maq = "03"

    ElseIf vIP = "192.168.0.4" Then
        Calcula_Maq = "04"

    ElseIf vIP = "192.168.0.5" Then
        Calcula_Maq = "05"

    ElseIf vIP = "192.168.0.6" Then
        Calcula_Maq = "06"

    ElseIf vIP = "192.168.0.7" Then
        Calcula_Maq = "07"

    ElseIf vIP = "192.168.0.8" Then
        Calcula_Maq = "08"

    ElseIf vIP = "192.168.0.9" Then
        Calcula_Maq = "09"

    ElseIf vIP = "192.168.0.10" Then
        Calcula_Maq = "10"

    ElseIf vIP = "192.168.0.11" Then
        Calcula_Maq = "11"

    ElseIf vIP = "192.168.0.12" Then
        Calcula_Maq = "12"

    ElseIf vIP = "192.168.0.13" Then
        Calcula_Maq = "13"

    ElseIf vIP = "192.168.0.14" Then
        Calcula_Maq = "14"

    ElseIf vIP = "192.168.0.15" Then
        Calcula_Maq = "15"

    ElseIf vIP = "192.168.0.16" Then
        Calcula_Maq = "16"

    ElseIf vIP = "192.168.0.17" Then
        Calcula_Maq = "17"

    ElseIf vIP = "192.168.0.18" Then
        Calcula_Maq = "18"

    ElseIf vIP = "192.168.0.19" Then
        Calcula_Maq = "19"

    ElseIf vIP = "192.168.0.20" Then
        Calcula_Maq = "20"

    ElseIf vIP = "192.168.0.21" Then
        Calcula_Maq = "21"

    ElseIf vIP = "192.168.0.22" Then
        Calcula_Maq = "22"

    ElseIf vIP = "192.168.0.23" Then
        Calcula_Maq = "23"
    
    ElseIf vIP = "192.168.0.24" Then
        Calcula_Maq = "24"
    
    ElseIf vIP = "192.168.0.25" Then
        Calcula_Maq = "25"
    
    ElseIf vIP = "192.168.0.26" Then
        Calcula_Maq = "26"
    
    'ElseIf vIP = "192.168.0.27" Then
    ElseIf vIP = "192.168.0.253" Then
        Calcula_Maq = "27"
    
    ElseIf vIP = "192.168.0.28" Then
        Calcula_Maq = "28"
    
    ElseIf vIP = "192.168.0.29" Then
        Calcula_Maq = "29"
    
    ElseIf vIP = "192.168.0.30" Then
        Calcula_Maq = "30"
    
    ElseIf vIP = "192.168.0.31" Then
        Calcula_Maq = "31"
    
    ElseIf vIP = "192.168.0.32" Then
        Calcula_Maq = "32"
    
    ElseIf vIP = "192.168.0.33" Then
        Calcula_Maq = "33"
    
    Else
        VGStrBox = MsgBox("Máquina não foi encontrada na rede." & Chr(13) & "IP inexistente", vbInformation, "Informação")
        VGStrIP = "inexistente"
    End If

End Function

Function Calcula_Hora(vHAtual, vHEntr)

    Dim HoraAtual As String
    Dim MinAtual As String
    Dim SegAtual As String
    
    Dim HoraEntr As String
    Dim MinEntr As String
    Dim SegEntr As String
    
    Dim Hora As String
    Dim Min As String
    Dim Seg As String
    
    HoraAtual = DatePart("H", vHAtual)
    MinAtual = DatePart("N", vHAtual)
    SegAtual = DatePart("S", vHAtual)
    
    HoraEntr = DatePart("H", vHEntr)
    MinEntr = DatePart("N", vHEntr)
    SegEntr = DatePart("S", vHEntr)
    
    Hora = HoraAtual - HoraEntr
    Min = MinAtual - MinEntr
    Seg = SegAtual - SegEntr
    
    If Hora < 10 Then
     Hora = "0" & Hora
    End If
    
    If Min < 10 Then
     Min = "0" & Min
    End If
    
    If Seg < 10 Then
     Seg = "0" & Seg
    End If
     
    Calcula_Hora = Hora & ":" & Min & ":" & Seg

End Function

Public Sub MakeMeService()
  Dim pid As Long
  Dim reserv As Long
  pid = GetCurrentProcessId()
  reserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub
