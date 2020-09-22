Attribute VB_Name = "Module1"
Option Explicit


Public Enum StandardIconEnum
    IDI_ASTERISK = 32516&       ' like vbInformation
    IDI_EXCLAMATION = 32515&    ' like vbExlamation
    IDI_HAND = 32513&           ' like vbCritical
    IDI_QUESTION = 32514&       ' like vbQuestion
End Enum

Public Declare Function LoadStandardIcon Lib "user32" Alias _
    "LoadIconA" (ByVal hInstance As Long, ByVal lpIconNum As _
    StandardIconEnum) As Long
    
Public Declare Function DrawIcon Lib "user32" (ByVal hDC _
    As Long, ByVal x As Long, ByVal y As Long, _
    ByVal hIcon As Long) As Long
    


'\\  M E S S A G E   D I A L O G   T Y P E S

Public Const BX_OK = 1&
Public Const BX_OKCANCEL = 2&
Public Const BX_YESNO = 3&
Public Const BX_YESNOCANCEL = 4&

' \\ I C O N   C O N S T A N T S

Public Const IC_NOICON = 0&
Public Const IC_EXCLAME = 5&
Public Const IC_CRITICAL = 6&
Public Const IC_INFO = 7&
Public Const IC_CONFIRM = 8&
Public Const IC_CUSTOM = 9&

' \\ M E S S A G E   B O X   R E T U R N   C O N S T A N T S

Public Const mbNo = 0&
Public Const mbYes = 1&
Public Const mbCancel = -1&




Public Declare Function _
    GetOpenFileName _
      Lib "comdlg32.dll" _
      Alias "GetSaveFileNameA" ( _
      pOpenfilename As OPENFILENAME _
) As Long


' // Type declarations
Public Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

'----------------------------------------------------------------------
' Name:         Open_File
' Author:       DCroft
' Created:      Sunday, May 07,2000 @ 10:43:06 pm (Vers: 1.0.0000)
'
' Description:  Calls the Open Dialog
'----------------------------------------------------------------------

' open FIle Function
Function Open_File( _
           hwnd As Long, _
           ByVal Title As String, _
           ByVal Filter As String _
         ) As String
   '

   Dim OpenFileDialog As OPENFILENAME
   Dim rv As Long
   
   ' // init dialog
   With OpenFileDialog
     .lStructSize = Len(OpenFileDialog)
     .hwndOwner = hwnd&
     .hInstance = App.hInstance
     .lpstrFilter = Filter$  '+ "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
     .lpstrFile = Space$(254)
     .nMaxFile = 255
     .lpstrFileTitle = Space$(254)
     .nMaxFileTitle = 255
     .lpstrInitialDir = CurDir
     .lpstrTitle = Title$
     .flags = 0
   End With
  
   ' // call API to show the dialog that was just initialized
   rv& = GetOpenFileName(OpenFileDialog)
   
   If (rv&) Then
      Open_File = Trim$(OpenFileDialog.lpstrFile)
   Else
      Open_File = ""
   End If
   
End Function


