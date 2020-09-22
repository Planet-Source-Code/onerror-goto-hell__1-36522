Attribute VB_Name = "FileOp"
Option Explicit

Public Const FO_MOVE As Long = &H1
Public Const FO_COPY As Long = &H2
Public Const FO_DELETE As Long = &H3
Public Const FO_RENAME As Long = &H4

Public Const FOF_MULTIDESTFILES As Long = &H1
Public Const FOF_CONFIRMMOUSE As Long = &H2
Public Const FOF_SILENT As Long = &H4
Public Const FOF_RENAMEONCOLLISION As Long = &H8
Public Const FOF_NOCONFIRMATION As Long = &H10
Public Const FOF_WANTMAPPINGHANDLE As Long = &H20
Public Const FOF_CREATEPROGRESSDLG As Long = &H0
Public Const FOF_ALLOWUNDO As Long = &H40
Public Const FOF_FILESONLY As Long = &H80
Public Const FOF_SIMPLEPROGRESS As Long = &H100
Public Const FOF_NOCONFIRMMKDIR As Long = &H200

Type SHFILEOPSTRUCT
   hwnd As Long
   wFunc As Long
   pFrom As String
   pTo As String
   fFlags As Long
   fAnyOperationsAborted As Long
   hNameMappings As Long
   lpszProgressTitle As String
End Type

Declare Function SHFileOperation Lib "Shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Function FileOperation(frm As Form, Funcao As String, Origem As String, Destino As String) As Boolean

    Dim iRet As Long
    Dim fileop As SHFILEOPSTRUCT
    
    fileop.hwnd = frm.hwnd
    Select Case UCase(Funcao)
        Case "MOVE": fileop.wFunc = FO_MOVE
            fileop.fFlags = FOF_FILESONLY
        Case "COPY": fileop.wFunc = FO_COPY
            fileop.fFlags = FOF_FILESONLY
            
            'The files to copy separated by Nulls and terminated by 2 null
            '.pFrom = "C:\Arq1" & vbNullChar & "C:\Arq2" & vbNullChar & vbNullChar
            
            'or to copy all files use this line
            '.pFrom = "C:\*.*" & vbNullChar & vbNullChar

            'The directory or filename(s) to copy into terminated in 2 nulls.
            '.pTo = "C:\testfolder\" & vbNullChar & vbNullChar
            
        Case "DELETE": fileop.wFunc = FO_DELETE
            fileop.fFlags = FOF_ALLOWUNDO
            'Allow undo--in other words, place the files into the Recycle Bin
            
        Case "RENAME": fileop.wFunc = FO_RENAME
            fileop.fFlags = FOF_FILESONLY
        Case Else
            MsgBox "Função inválida"
            Stop
    End Select
    fileop.pFrom = Origem & vbNullChar & vbNullChar
    fileop.pTo = Destino & vbNullChar & vbNullChar
    
    iRet = SHFileOperation(fileop)

    If iRet <> 0 Then          'Operation failed
       MsgBox Err.LastDllError, vbInformation, "Operation failed"
       FileOperation = False
       'Msgbox the error that occurred in the API.
    ElseIf fileop.fAnyOperationsAborted <> 0 Then
        MsgBox "Operation Failed"
        FileOperation = False
    Else
        FileOperation = True
    End If
    
End Function
