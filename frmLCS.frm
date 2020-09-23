VERSION 5.00
Begin VB.Form frmLCS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Longest Common Substring Example"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLCS 
      Height          =   315
      Left            =   2940
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "&Compute"
      Height          =   435
      Left            =   3600
      TabIndex        =   2
      Top             =   1080
      Width           =   1515
   End
   Begin VB.TextBox txtB 
      Height          =   315
      Left            =   660
      TabIndex        =   1
      Text            =   "ALOHA"
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtA 
      Height          =   315
      Left            =   660
      TabIndex        =   0
      Text            =   "HELLO"
      Top             =   180
      Width           =   2055
   End
   Begin VB.Label lblB 
      Caption         =   "Text B:"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   660
      Width           =   555
   End
   Begin VB.Label lblA 
      Caption         =   "Text A:"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   240
      Width           =   555
   End
   Begin VB.Label lblLCS 
      Caption         =   "Longest Common Substring is:"
      Height          =   195
      Left            =   2940
      TabIndex        =   4
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmLCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''
'    A L O H A '
' H  0 0 0 1 0 '
' E  0 0 0 0 0 '
' L  0 1 0 0 0 '
' L  0 1 0 0 0 '
' O  0 0 2 0 0 '
''''''''''''''''

'Note that the function is case insensitive, also that
'it will only return the first longest common substring
Public Function LongestCommonSubstring(A As String, B As String) As String
    Dim sLenA, sLenB, sLen, i, j, idxA, idxB As Integer
    
    sLenA = Len(A): sLenB = Len(B)
    ReDim ArrLCS(0 To sLenA, 0 To sLenB) As Integer
    
    For i = 1 To sLenA
        For j = 1 To sLenB
            If UCase$(Mid$(A, i, 1)) <> UCase$(Mid$(B, j, 1)) Then
                ArrLCS(i, j) = 0
            Else
                ArrLCS(i, j) = 1 + ArrLCS(i - 1, j - 1)
                If ArrLCS(i, j) > sLen Then
                    sLen = ArrLCS(i, j)
                    idxA = i: idxB = j
                End If
            End If
        Next
    Next
    
    LongestCommonSubstring = Mid$(A, idxA - sLen + 1, ArrLCS(idxA, idxB))
    Erase ArrLCS
End Function

Private Sub cmdCompute_Click()
    txtLCS.Text = LongestCommonSubstring(txtA.Text, txtB.Text)
End Sub
