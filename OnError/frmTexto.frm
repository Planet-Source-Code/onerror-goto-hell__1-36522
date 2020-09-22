VERSION 5.00
Begin VB.Form frmTxt 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5430
   ClientLeft      =   525
   ClientTop       =   1545
   ClientWidth     =   10425
   Icon            =   "frmTexto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   210
      Width           =   10185
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   645
      Left            =   6240
      Picture         =   "frmTexto.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   1245
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Save"
      Height          =   645
      Left            =   3240
      Picture         =   "frmTexto.frx":00FE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   1245
   End
End
Attribute VB_Name = "frmTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NumTxt As Integer

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdOk_Click()

    Dim Canal As Integer
    Dim Texto As String
    
    Canal = FreeFile
    
    Texto = txt.Text
    
    Select Case NumTxt
        Case 1
            If InStr(Texto, C_ON_ERROR_GOTO) = 0 Then
                MsgBox "Missing: VARIABLE " & C_ON_ERROR_GOTO & " in text."
                txt.SetFocus
                Exit Sub
            End If
            Texto1 = Texto
            Open App.Path & "\Texto1.TXT" For Output As #Canal
            Print #Canal, Texto1
            Close #Canal
        
        Case 2
            If InStr(Texto, C_EXIT_PROCEDURE) = 0 Then
                MsgBox "Missing: VARIABLE " & C_EXIT_PROCEDURE & " in text."
                txt.SetFocus
                Exit Sub
            End If
            
            If InStr(Texto, C_NOME_PROCEDURE) = 0 Then
                MsgBox "Missing: VARIABLE " & C_NOME_PROCEDURE & " in text."
                txt.SetFocus
                Exit Sub
            End If
            
            Texto2 = Texto
            Open App.Path & "\Texto2.TXT" For Output As #Canal
            Print #Canal, Texto2
            Close #Canal
        Case Else: Stop
    
    End Select
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    Select Case NumTxt
        Case 1
            Me.Caption = "Type the ON ERROR text"
            txt.Text = Texto1
        Case 2
            Me.Caption = "Type the text of error routine"
            txt.Text = Texto2
        Case Else: Stop
    End Select

End Sub
