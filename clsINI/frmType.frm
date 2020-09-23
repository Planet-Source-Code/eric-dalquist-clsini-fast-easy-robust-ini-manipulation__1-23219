VERSION 5.00
Begin VB.Form frmType 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1530
      TabIndex        =   5
      Top             =   720
      Width           =   675
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   690
      TabIndex        =   4
      Top             =   720
      Width           =   675
   End
   Begin VB.OptionButton optType 
      Caption         =   "Key"
      Height          =   195
      Index           =   1
      Left            =   1020
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.OptionButton optType 
      Caption         =   "Value"
      Height          =   195
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.OptionButton optType 
      Caption         =   "Section"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblTitle 
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2775
   End
End
Attribute VB_Name = "frmType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private okClick As Boolean
Private cancelClick As Boolean

Private Sub Form_Unload(Cancel As Integer)
    cancelClick = True
    Cancel = 1
End Sub

Private Sub btnOK_Click()
    okClick = True
End Sub

Private Sub btnCancel_Click()
    cancelClick = True
End Sub

Public Function getType(title As String, caption As String) As Integer
    okClick = False
    cancelClick = False
    Me.caption = title
    lblTitle.caption = caption
    
    If title = "Rename" Then
        optType(2).Enabled = False
    ElseIf title = "Add" Then
        optType(0).Enabled = False
    Else
        optType(2).Enabled = True
    End If
    
    Me.Show
    
    While Not okClick And Not cancelClick
        Sleep 10
        DoEvents
    Wend
    
    If okClick Then
        If optType(0).Value Then
            getType = 0
        ElseIf optType(1).Value Then
            getType = 1
        ElseIf optType(2).Value Then
            getType = 2
        End If
    Else
        getType = -1
    End If
    
    Me.Hide
End Function
