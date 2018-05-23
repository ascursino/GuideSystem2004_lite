Attribute VB_Name = "ModuleBase2"
'No módulo:
Public Declare Function ExitWindowsEx Lib "user32" _
       (ByVal uFlags As Long, _
       ByVal dwReserved As Long) As Long

'Public Const EWX_LOGOFF As Long = 0 'Faz Logoff do usuário.
'Public Const EWX_SHUTDOWN As Long = 1 'Desligar o computador.
Public Const EWX_REBOOT As Long = 2 'Reiniciar o computador.
'Public Const EWX_FORCE As Long = 4 'Força a ação desejada.



'para arredondar canto do formulário
Private Declare Function CreateRoundRectRgn Lib _
        "gdi32" (ByVal X1 As Long, ByVal Y1 As _
        Long, ByVal X2 As Long, ByVal Y2 As Long, _
        ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" _
        (ByVal hwnd As Long, ByVal hRgn As Long, _
        ByVal bRedraw As Boolean) As Long
Private Declare Function GetClientRect Lib "user32" _
        (ByVal hwnd As Long, lpRect As Rect) As Long
Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Sub Retangulo(m_hWnd As Long, Fator As Byte)
  Dim RGN As Long
  Dim RC As Rect
  Call GetClientRect(m_hWnd, RC)
  RGN = CreateRoundRectRgn(RC.Left, RC.Top, RC.Right, _
                           RC.Bottom, Fator, Fator)
  SetWindowRgn m_hWnd, RGN, True
End Sub
