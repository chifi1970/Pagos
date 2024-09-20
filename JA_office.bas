Attribute VB_Name = "Module1"
'Global AI(50, 6)
'Global JA(50, 6)
Global ruta$, ruta_pdf$
Global tipomsg
Global IDdisc$
Global cliente_ID$
Global poliza_buscada$
Global user$, cargo$
Global transfiere$
Global oficina_guardada$(5)
Global valido1
Global asunto$
Global bloqueado As Integer
Global usuario_DMV$
Global pantalla As Integer
Global fecha1$, fecha2$
Global seg_end As Integer
Global posicion As Integer
Global forma$
Global Modificado As Integer

Global administrador As Integer, name_admon$

Global fecha_rango1$, fecha_rango2$, op_searchx As Integer

Global id_agente$, ID_manager$, correo_agente$, correo_manager$, correo_admin$

Global grupo$


Type underwriting
  cust_id As String * 6
  recibo As String * 7
  poliza As String * 15
  bf As Single
  Error1 As String * 40
  Error2 As String * 40
  Error3 As String * 40
  Error4 As String * 40
  Error5 As String * 40
  Error6 As String * 40
  Error7 As String * 40
  Error8 As String * 40
  Error9 As String * 40
  Error10 As String * 40
  tag1 As String * 40
  tag2 As String * 40
  tag3 As String * 40
  tag4 As String * 40
  tag5 As String * 40
    
  agente As String * 30
End Type

Global uw As underwriting

  
Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer ' As Long para que funcione en Windows XP con VB6
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type


Public Const FO_DELETE = &H3
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_SILENT = &H4


Declare Function PathIsURL Lib "shlwapi.dll" Alias "PathIsURLA" (ByVal pszPath As String) As Long


Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
    (lpFileOp As SHFILEOPSTRUCT) As Long
    


Type registro_pass
   id_employee As Integer
   programa As Integer
   usuario As String * 60
   pass As String * 30
End Type

Global Const tam_pass = 94
Global pass As registro_pass



    ' esto es para obtener el IP

Public Const MAX_WSADescription As Long = 256
Public Const MAX_WSASYSStatus As Long = 128
' Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

Public Type HOSTENT
   hName      As Long
   hAliases   As Long
   hAddrType  As Integer
   hLen       As Integer
   hAddrList  As Long
End Type

Public Type WSADATA
   wVersion      As Integer
   wHighVersion  As Integer
   szDescription(0 To MAX_WSADescription)   As Byte
   szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
   wMaxSockets   As Integer
   wMaxUDPDG     As Integer
   dwVendorInfo  As Long
End Type

Public Declare Function WSAGetLastError Lib "wsock32" () As Long

Public Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "wsock32" () As Long

Public Declare Function gethostname Lib "wsock32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
   
Public Declare Function gethostbyname Lib "wsock32" _
  (ByVal szHost As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" (hpvDest As Any, _
   ByVal hpvSource As Long, _
   ByVal cbCopy As Long)
   
Private Function RandomInt(lowerbound As Long, upperbound As Long) As Long
    RandomInt = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function





Public Function ShellDelete(ParamArray vntFileName() As Variant) As Long

    Dim i As Integer
    Dim sFileNames As String
    Dim SHFileOp As SHFILEOPSTRUCT

    For i = LBound(vntFileName) To UBound(vntFileName)
    sFileNames = sFileNames & vntFileName(i) & vbNullChar
    Next
    sFileNames = sFileNames & vbNullChar

    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = sFileNames
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION

 'FOF_ALLOWUNDO
    End With

    ShellDelete = SHFileOperation(SHFileOp)

End Function

Sub Resize_For_Resolution(ByVal SFX As Single, _
       ByVal SFY As Single, MyForm As Form)
       On Error Resume Next
       
      Dim i As Integer
      Dim SFFont As Single

      SFFont = (SFX + SFY) / 2  ' average scale
      ' Size the Controls for the new resolution
      On Error Resume Next  ' for read-only or nonexistent properties
      With MyForm
        For i = 0 To .Count - 1
         If TypeOf .Controls(i) Is ComboBox Then   ' cannot change Height
           .Controls(i).Left = .Controls(i).Left * SFX
           .Controls(i).Top = .Controls(i).Top * SFY
           .Controls(i).Width = .Controls(i).Width * SFX
         Else
           .Controls(i).Move .Controls(i).Left * SFX, _
            .Controls(i).Top * SFY, _
            .Controls(i).Width * SFX, _
            .Controls(i).Height * SFY
         End If
           .Controls(i).FontSize = .Controls(i).FontSize * SFFont
        Next i
        If RePosForm Then
          ' Now size the Form
          .Move .Left * SFX, .Top * SFY, .Width * SFX, .Height * SFY
        End If
      End With
End Sub

