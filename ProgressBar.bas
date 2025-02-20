Attribute VB_Name = "BarraProgreso"
Option Explicit
  
  
' Constantes para SendMessage
Private Const PBM_SETBKCOLOR  As Long = (&H2000& + 1)
Private Const PBM_SETBARCOLOR As Long = (&H400 + 9)
  
' Delaración de SendMessage
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                         (ByVal hwnd As Long, ByVal wMsg As Long, _
                          ByVal wParam As Long, lParam As Any) As Long
  
'Cambia el color del Value de la barra, si no se especifica _
  el color por defecto utiliza el color Vederde
  
Public Sub Color_Progreso(ByVal HWND_Prog As Long, _
                          Optional ByVal color As Long = vbGreen)
      
    Call SendMessage(HWND_Prog, PBM_SETBARCOLOR, 0&, ByVal color)
  
End Sub
  
' Cambia el color del fondo del Progress, si no se especifica _
  el color por defecto utiliza el color Rojo
  
Public Sub Color_Fondo(ByVal HWND_Prog As Long, _
                       Optional ByVal color As Long = vbRed)
      
    Call SendMessage(HWND_Prog, PBM_SETBKCOLOR, 0&, ByVal color)
  
End Sub

