VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert ON ERROR into your project"
   ClientHeight    =   5385
   ClientLeft      =   945
   ClientTop       =   1455
   ClientWidth     =   10020
   Icon            =   "frmOnError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   10020
   Begin VB.CommandButton cmdTexto2 
      Caption         =   "Error routine"
      Height          =   615
      Left            =   300
      Picture         =   "frmOnError.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Edit the error routine"
      Top             =   3750
      UseMaskColor    =   -1  'True
      Width           =   1245
   End
   Begin VB.CommandButton cmdTexto1 
      Caption         =   "OnError routine"
      Height          =   615
      Left            =   300
      Picture         =   "frmOnError.frx":03C2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Edit ON ERROR text"
      Top             =   3030
      UseMaskColor    =   -1  'True
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Run"
      Height          =   615
      Left            =   300
      Picture         =   "frmOnError.frx":0778
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1470
      UseMaskColor    =   -1  'True
      Width           =   1245
   End
   Begin VB.ListBox lst 
      Height          =   4155
      Left            =   1920
      TabIndex        =   4
      ToolTipText     =   "Project files"
      Top             =   240
      Width           =   7905
   End
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   9210
      Top             =   4530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSeleciona 
      Caption         =   "Select .VBP"
      Height          =   615
      Left            =   300
      Picture         =   "frmOnError.frx":08A2
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Select the .vbp"
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   1245
   End
   Begin VB.Label lblRot 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Don't forget: ADD ErrorModule.BAS to your project"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   30
      TabIndex        =   6
      Top             =   4860
      Width           =   9945
   End
   Begin VB.Label lblMsg 
      Height          =   225
      Left            =   1920
      TabIndex        =   5
      Top             =   4440
      Width           =   7905
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RotaArq As String

Private Sub cmdOk_Click()

    Dim k As Integer
    Dim k1 As Integer
    Dim k2 As Integer
    Dim k3 As Integer
    
    Dim Arq As String
    Dim tbl() As String
    Dim Canal As Integer
    Dim Linha As String
    Dim ix As Integer
    Dim NomeProcedure As String
    Dim Rotina_ON_ERROR As String
    Dim Rotina_ERRO As String
    Dim OnErrorIncluidos As Long
    Dim AchouOnError As Boolean
    Dim RotaBKP As String
    Dim ixBKP As Integer
    
    Const NrProcedures As Integer = 15
    
    Dim pTipo(1 To NrProcedures) As String
    Dim pEnd(1 To NrProcedures) As String
     
    pTipo(1) = "Sub ":                pEnd(1) = "End Sub"
    pTipo(2) = "Private Sub ":        pEnd(2) = "End Sub"
    pTipo(3) = "Public Sub ":         pEnd(3) = "End Sub"
    pTipo(4) = "Friend Sub ":         pEnd(4) = "End Sub"
    pTipo(5) = "Static Sub ":         pEnd(5) = "End Sub"
    
    pTipo(6) = "Function ":           pEnd(6) = "End Function"
    pTipo(7) = "Private Function ":   pEnd(7) = "End Function"
    pTipo(8) = "Public Function ":    pEnd(8) = "End Function"
    pTipo(9) = "Friend Function ":    pEnd(9) = "End Function"
    pTipo(10) = "Static Function ":  pEnd(10) = "End Function"
    
    pTipo(11) = "Property ":         pEnd(11) = "End Property"
    pTipo(12) = "Private Property ": pEnd(12) = "End Property"
    pTipo(13) = "Public Property ":  pEnd(13) = "End Property"
    pTipo(14) = "Friend Property ":  pEnd(14) = "End Property"
    pTipo(15) = "Static Property ":  pEnd(15) = "End Property"
        
    RotaBKP = RotaArq & "BKP0000\"
    
    While Dir$(RotaBKP & "*.*") > vbNullString
        ixBKP = ixBKP + 1
        RotaBKP = RotaArq & "BKP" & Format$(ixBKP, "0000") & "\"
    Wend
    
    If FileOperation(Me, "COPY", RotaArq & "*.*", RotaBKP) = False Then
        If MsgBox("Continua?", vbYesNo) = vbNo Then Exit Sub
    End If
    
    Canal = FreeFile
    Me.MousePointer = vbHourglass
    
    For k = 0 To lst.ListCount - 1
        Arq = lst.List(k)
        lst.ListIndex = k - 1
        Mensagem "Processando: " & Arq
        
        If Dir$(Arq) = vbNullString Then
            MsgBox Arq & " não encontrado", vbCritical
        Else
            Open Arq For Input As #Canal
            ix = 0
            While Not EOF(Canal)
                Line Input #Canal, Linha
                ix = ix + 1
                ReDim Preserve tbl(1 To ix) As String
                tbl(ix) = Linha
            Wend
            Close #Canal
            
            AchouOnError = False
            
            Open Arq For Output As #Canal
            For k1 = 1 To ix
                For k2 = 1 To NrProcedures
                    If Left$(tbl(k1), Len(pTipo(k2))) = pTipo(k2) Then
                        NomeProcedure = ParseNomeProcedure(Len(pTipo(k2)), tbl(k1))
                        Rotina_ON_ERROR = FormataRotina_ON_ERROR()
                        Rotina_ERRO = FormataRotina_ERRO(pEnd(k2), NomeProcedure)
                        For k3 = k1 + 1 To ix
                            If tbl(k3) > vbNullString Then
                                If InStr(tbl(k3), "On Error") > 0 Then
                                    Rotina_ON_ERROR = vbNullString
                                    Rotina_ERRO = vbNullString
                                    AchouOnError = True
                                    Exit For
                                End If
                                If InStr(tbl(k3), pEnd(k2)) > 0 Then Exit For 'Fim da procedure
                            End If
                        Next k3
                        Exit For
                    End If
                    If AchouOnError = True Then
                        Rotina_ON_ERROR = vbNullString
                        Rotina_ERRO = vbNullString
                    End If
                    If Left$(Trim$(tbl(k1)), Len(pEnd(k2))) = pEnd(k2) Then
                        If AchouOnError = True Then
                            AchouOnError = False
                            Exit For
                        End If
                        Print #Canal, Rotina_ERRO
                        Rotina_ERRO = vbNullString
                        OnErrorIncluidos = OnErrorIncluidos + 1
                        Exit For
                    End If
                Next k2
                Print #Canal, tbl(k1)
                If Rotina_ON_ERROR > vbNullString And Right$(tbl(k1), 1) <> "_" Then
                    Print #Canal, Rotina_ON_ERROR
                    Rotina_ON_ERROR = vbNullString
                End If
                
            Next k1
            Close #Canal
            
            Erase tbl
        End If
        
    Next k


    Me.MousePointer = vbArrow
    
    MsgBox vbCrLf & "Finished!" & vbCrLf & vbCrLf & "OnError included: " & OnErrorIncluidos & vbCrLf, vbInformation
    Unload Me
    
End Sub

Private Sub cmdSeleciona_Click()
    
    Dim Arq As String
    Dim Linha As String
    Dim Canal As Integer
    Dim Posicao As Integer
    Dim Arquivo As String
    
    On Error GoTo Erro
    
    lst.Clear
    cmDialog.CancelError = False
    cmDialog.FileName = "*.vbp"
    cmDialog.ShowOpen
    
    Arq = cmDialog.FileTitle
    
    RotaArq = cmDialog.FileName
    RotaArq = Left$(RotaArq, Len(RotaArq) - Len(Arq))
    
    Canal = FreeFile
    Open RotaArq & Arq For Input As #Canal
    While Not EOF(Canal)
        Line Input #Canal, Linha
        Linha = UCase$(Linha)
        'MsgBox Linha
        If Left$(Linha, 6) = "MODULE" Or Left$(Linha, 5) = "CLASS" Then
            Posicao = InStr(Linha, "; ") + 2
            Arquivo = Mid$(Linha, Posicao)
            lst.AddItem RotaArq & Arquivo
        ElseIf Left$(Linha, 4) = "FORM" Then
            Posicao = InStr(Linha, "=") + 1
            Arquivo = Mid$(Linha, Posicao)
            lst.AddItem RotaArq & Arquivo
        End If
    Wend
    Close #Canal
    
    If lst.ListCount > 0 Then
        cmdOk.Enabled = True
        cmdOk.SetFocus
    End If
    
Saida:
    Exit Sub

Erro:
    If Err.Number = 32755 Then GoTo Saida
GoTo Saida

End Sub

Private Sub cmdSeleciona_GotFocus()

    cmdOk.Enabled = False
    
End Sub
Private Sub Mensagem(msg As String)
        
    lblMsg.Caption = msg
    lblMsg.Refresh

End Sub
Private Function ParseNomeProcedure(PosI As Integer, Linha As String) As String

    Dim PosF As Integer
    
    PosF = InStr(Linha, "(")

    ParseNomeProcedure = Mid$(Linha, PosI + 1, PosF - PosI - 1)
    
End Function

Private Sub cmdTexto1_Click()

    frmTxt.NumTxt = 1
    frmTxt.Show vbModal

End Sub

Private Sub cmdTexto2_Click()

    frmTxt.NumTxt = 2
    frmTxt.Show vbModal

End Sub

Private Sub Form_Load()

    Set BaseMDI = Me
    Dim Canal As Integer
    
    Canal = FreeFile
    
    Open App.Path & "\Texto1.TXT" For Input As #Canal
    Texto1 = Input(LOF(Canal), Canal)
    Close #Canal

    Open App.Path & "\Texto2.TXT" For Input As #Canal
    Texto2 = Input(LOF(Canal), Canal)
    Close #Canal
    
    If InStr(Texto1, C_ON_ERROR_GOTO) = 0 Then
        MsgBox "Missing: " & C_ON_ERROR_GOTO & " in file Texto1.TXT"
    End If
    
    If InStr(Texto2, C_EXIT_PROCEDURE) = 0 Then
        MsgBox "Missing: " & C_EXIT_PROCEDURE & " in file Texto2.TXT"
    End If
    
    If InStr(Texto2, C_NOME_PROCEDURE) = 0 Then
        MsgBox "Missing: " & C_NOME_PROCEDURE & " in file Texto2.TXT"
    End If

End Sub
Private Function FormataRotina_ON_ERROR() As String

    FormataRotina_ON_ERROR = Texto1
    
End Function

Private Function FormataRotina_ERRO(TxtEND As String, txtNomeProcedure As String) As String

    Dim txt As String
    Dim PosI As Integer
    Dim PosF As Integer
    Dim txtEXIT As String
    
    txtEXIT = "Exit " & Mid$(TxtEND, 4)
    
    txt = Texto2
    PosI = InStr(txt, C_EXIT_PROCEDURE)
    PosF = PosI + Len(C_EXIT_PROCEDURE)
    txt = Left$(txt, PosI - 1) & txtEXIT & Mid$(txt, PosF)
    
    PosI = InStr(txt, C_NOME_PROCEDURE)
    PosF = PosI + Len(C_NOME_PROCEDURE)
    txt = Left$(txt, PosI - 1) & txtNomeProcedure & Mid$(txt, PosF)
    
    'MsgBox txt
    
    If Mid$(TxtEND, 5) = "Function" Then
        txt = txt & vbCrLf & "!!!TO_DO__FUNCTION_MUST_RETURN_SOME_VALUE_!!!"
    End If
    
    FormataRotina_ERRO = txt
    
End Function
