Attribute VB_Name = "ErrorModule"
Option Explicit

Public BaseMDI As Form

Public RunOnceVerificaERRO As Boolean

'--- CAPTURA DE TELA
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const SRCCOPY = &HCC0020 ' (DWORD) destination = source

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

'API

Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long

Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Declare Function PatBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long

Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long

Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
'if you have problems with this function add the Alias "SetClipboardDataA"
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
'já declarado acima ... Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
'já declarado acima ... Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'já declarado acima ... Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long


Public Sub InformaErro(MeuErr As Integer, MeuError As String, Rotina As String)

    'Exibe mensagem de erro e grava .LOG da ocorrência
    Dim mErro As String
    Dim Texto As String
    Dim VideoFileName As String

    On Error Resume Next
    
    mErro = "Módulo: " & App.EXEName & vbCrLf & _
            "Erro: " & MeuErr & " - " & MeuError & vbCrLf & _
            "Rotina: " & Replace(Rotina, "_", vbNullString)
    
    'maybe you don't wanna log this errors
    'If MeuErr = 91 Then GoTo NaoGravaLog 'no loaded
    'If MeuErr = 484 Then GoTo NaoGravaLog '*** printer not found
    'If MeuErr = 3045 Then GoTo NaoGravaLog 'exclusive use by another user
    'If MeuErr = 3186 Then GoTo NaoGravaLog
    'If MeuErr = 3356 Then GoTo NaoGravaLog
        
    '(Error 3186) Couldn't save; currently locked by user <name> on machine <name>
    'NESTE erro, nao grava no arquivo de erro
    
    '* * * Documenta o erro e a tela atual em disco
    
    VideoFileName = App.Path & "\TELA_DE_ERRO_" & Format$(Now, "yyyymmdd hhmmss") & ".BMP"
    
    CapturaTodaTela VideoFileName
    
    Texto = "Modulo.: " & App.EXEName & " " & App.ProductName & vbCrLf
    Texto = Texto & "Erro...: " & MeuErr & " " & MeuError & vbCrLf
    Texto = Texto & "Rotina.: " & Rotina & vbCrLf
    Texto = Texto & "Video..: " & VideoFileName & vbCrLf
    Texto = Texto & "Rota...: " & CurDir$ & vbCrLf
    Texto = Texto & String$(78, 45) & vbCrLf
    
    Open Get_AppPath() & "ERROR.LOG" For Append As #200
    Print #200, "Data...: " & Format$(Now, "dddd, dd-mmmm-yyyy ttttt AM/PM")
    Print #200, Texto
    Close #200
    
    GravaRelOperacoes Texto
    
NaoGravaLog:

    VBA.MsgBox mErro, vbCritical, "ERRO NO SISTEMA"
    
End Sub


Public Sub CapturaTodaTela(NomeDoArquivo_BMP_ParaGravar As String)
 
    'Usage: to capture the entire screen
    
    If VerificaERRO() = True Then On Error GoTo Erro
    
    Const PicName As String = "picTemp123"
    
    BaseMDI.Controls.Add "VB.PictureBox", PicName

    CaptureScreen 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY
    
    BaseMDI.Controls(PicName) = Clipboard.GetData()
    
    If NomeDoArquivo_BMP_ParaGravar > vbNullString Then
        SavePicture BaseMDI.Controls(PicName).Picture, NomeDoArquivo_BMP_ParaGravar
    End If
    
    BaseMDI.Controls.Remove BaseMDI.Controls(PicName)
    
Saida:
    Exit Sub

Erro:
    InformaErro Err.Number, Err.Description, "CapturaTodaTela"
    Resume Saida

End Sub



 'Screen Capture Procedure, coordinates are expressed in pixels
Public Sub CaptureScreen(Left As Long, Top As Long, Width As Long, Height As Long)
    
    If VerificaERRO() = True Then On Error GoTo Erro

    Dim srcDC As Long
    Dim trgDC As Long
    Dim BMPHandle As Long
    Dim dm As DEVMODE
    
    srcDC = CreateDC("DISPLAY", vbNullString, vbNullString, dm)
    trgDC = CreateCompatibleDC(srcDC)
    BMPHandle = CreateCompatibleBitmap(srcDC, Width, Height)
    SelectObject trgDC, BMPHandle
    BitBlt trgDC, 0, 0, Width, Height, srcDC, Left, Top, SRCCOPY
    OpenClipboard Screen.ActiveForm.hwnd
    EmptyClipboard
    SetClipboardData 2, BMPHandle
    CloseClipboard
    DeleteDC trgDC
    ReleaseDC BMPHandle, srcDC
Saida:
    Exit Sub

Erro:
    InformaErro Err.Number, Err.Description, "CaptureScreen"
    Resume Saida

End Sub
 
Private Sub CapturaForm(frm As Form, pic As PictureBox)
    
    'Usage: to capture a form
    
    If VerificaERRO() = True Then On Error GoTo Erro

    CaptureScreen frm.Left \ Screen.TwipsPerPixelX, frm.Top \ Screen.TwipsPerPixelY, frm.Width \ Screen.TwipsPerPixelX, frm.Height \ Screen.TwipsPerPixelY
    pic = Clipboard.GetData()

Saida:
    Exit Sub

Erro:
    InformaErro Err.Number, Err.Description, "CapturaForm"
    Resume Saida

End Sub

Public Function Get_AppPath() As String

    If VerificaERRO() = True Then On Error GoTo Erro

    'Retorna o diretório do programa em uso
    
    Dim Aux As String
    Aux = App.Path
    If Right$(Aux, 1) <> "\" Then Aux = Aux & "\"
    Get_AppPath = Aux

Saida:
    Exit Function

Erro:
    InformaErro Err.Number, Err.Description, "GetAppPath"
    Resume Saida

End Function

Public Sub GravaRelOperacoes(Texto As String)

    If VerificaERRO() = True Then On Error GoTo Erro
    
    'Data / competencia / usuario / rotina
    
    Dim fl As Integer
        
    fl = FreeFile
    
    Open Get_AppPath & "RelOpr.LOG" For Append Shared As #fl
    Print #fl, Format$(Now, "General Date") & " " & Trim$(Texto)
    Close #fl
    
Saida:
    Exit Sub

Erro:
    InformaErro Err.Number, Err.Description, "GravaRelOperacoes"
    Resume Saida


End Sub

Public Function VerificaERRO() As Boolean

    ' Liga ou não ON ERROR GOTO
    ' no vb, debug é executado
    ' no compilado, Debug.Print  NÃO é executado
    Dim VerificaERROValor As Boolean
    If RunOnceVerificaERRO = True Then GoTo Saida
    
    On Error Resume Next
    Debug.Print 2 / 0

    If Err.Number = 11 Then
        VerificaERROValor = False
    Else
        VerificaERROValor = True
    End If
    Err.Number = 0
    
    
    RunOnceVerificaERRO = True

Saida:
    VerificaERRO = VerificaERROValor

End Function


