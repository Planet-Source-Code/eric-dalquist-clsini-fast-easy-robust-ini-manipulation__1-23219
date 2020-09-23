VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "clsINI Demo dll"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnDelete 
      Caption         =   "Delete"
      Height          =   315
      Left            =   2040
      TabIndex        =   13
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton btnRename 
      Caption         =   "Rename"
      Height          =   315
      Left            =   1080
      TabIndex        =   12
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   1980
      TabIndex        =   10
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   1020
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtSection 
      Height          =   285
      Left            =   60
      TabIndex        =   8
      Top             =   3000
      Width           =   975
   End
   Begin VB.ListBox lstValues 
      Height          =   2205
      ItemData        =   "frmMain.frx":0000
      Left            =   1980
      List            =   "frmMain.frx":0002
      TabIndex        =   7
      Top             =   780
      Width           =   975
   End
   Begin VB.ListBox lstKeys 
      Height          =   2205
      ItemData        =   "frmMain.frx":0004
      Left            =   1020
      List            =   "frmMain.frx":0006
      TabIndex        =   6
      Top             =   780
      Width           =   975
   End
   Begin VB.ListBox lstSections 
      Height          =   2205
      ItemData        =   "frmMain.frx":0008
      Left            =   60
      List            =   "frmMain.frx":000A
      TabIndex        =   5
      Top             =   780
      Width           =   975
   End
   Begin VB.TextBox txtINIFile 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   660
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblValues 
      Caption         =   "Values:"
      Height          =   195
      Left            =   2040
      TabIndex        =   4
      Top             =   540
      Width           =   855
   End
   Begin VB.Label lblKeys 
      Caption         =   "Keys:"
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   540
      Width           =   855
   End
   Begin VB.Label lblSections 
      Caption         =   "Sections:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   855
   End
   Begin VB.Label lblINIFile 
      Alignment       =   1  'Right Justify
      Caption         =   "INI File:"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
' Author:   Eric Dalquist
' email:    ebdalqui@mtu.edu
'
' Comments: Well since I whipped this up to demo my INI class and I HATE
'           commenting I'm not going to do it. If you do have any questions on
'           how I did anything don't hesitate to email me. I'll try to get back
'           to you within a day or two.
'
' Bugs:     It's a demo program ... so who knows :-)
'********************************************************************************

Option Explicit

Private INI As New clsINI

Private Sub Form_Load()
    If Right$(App.Path, 1) = "\" Then
        INI.INIFile = App.Path & "clsINIDemo.ini"
    Else
        INI.INIFile = App.Path & "\" & "clsINIDemo.ini"
    End If
    
    txtINIFile.Text = INI.INIFile
    Load frmType
    frmType.Hide
    loadSections
    loadKeysAndVals
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmType
    End
End Sub

Private Sub btnAdd_Click()
    Dim action As Integer
    Dim section As String
    Dim key As String
    Dim value As String
    
    action = frmType.getType("Add", "Select type to add.")
    
    If action = 1 Then
        section = InputBox("Enter section:", "Add - Section")
        If section <> "" Then
            key = InputBox("Enter key:", "Add - Key")
            If key <> "" Then INI.CreateKey section, key
        End If
    ElseIf action = 2 Then
        section = InputBox("Enter section:", "Add - Section")
        If section <> "" Then
            key = InputBox("Enter key:", "Add - Key")
            If key <> "" Then
                value = InputBox("Enter value:", "Add - Value")
                If value <> "" Then INI.CreateKeyValue section, key, value
            End If
        End If
    End If
    
    loadSections
End Sub

Private Sub btnRename_Click()
    Dim sIndex As Integer
    Dim index As Integer
    Dim action As Integer
    
    If lstSections.ListCount > 0 Then
        btnAdd.Enabled = False
        btnRename.Enabled = False
        btnDelete.Enabled = False
        
        action = frmType.getType("Rename", "Select level to rename.")
        
        If action = 0 Then
            For index = 0 To lstSections.ListCount - 1
                If lstSections.Selected(index) Then
                    INI.RenameSection lstSections.List(index), InputBox("Rename " & lstSections.List(index) & " to:", "Rename")
                    Exit For
                End If
            Next index
            
        ElseIf action = 1 Then
            For index = 0 To lstSections.ListCount - 1
                If lstSections.Selected(index) Then
                    sIndex = index
                    Exit For
                End If
            Next index
            
            For index = 0 To lstKeys.ListCount - 1
                If lstKeys.Selected(index) Then
                    INI.RenameKey lstSections.List(sIndex), lstKeys.List(index), InputBox("Rename " & lstKeys.List(index) & " to:", "Rename")
                    Exit For
                End If
            Next index
        End If
        
        loadSections
        loadKeysAndVals
        
        btnAdd.Enabled = True
        btnRename.Enabled = True
        btnDelete.Enabled = True
    End If
End Sub

Private Sub btnDelete_Click()
    Dim sIndex As Integer
    Dim index As Integer
    Dim action As Integer
    
    If lstSections.ListCount > 0 Then
        btnAdd.Enabled = False
        btnRename.Enabled = False
        btnDelete.Enabled = False
        
        action = frmType.getType("Delete", "Select level to delete at.")
        
        If action = 0 Then
            For index = 0 To lstSections.ListCount - 1
                If lstSections.Selected(index) Then
                    INI.DeleteSection lstSections.List(index)
                    Exit For
                End If
            Next index
            
        ElseIf action = 1 Then
            For index = 0 To lstSections.ListCount - 1
                If lstSections.Selected(index) Then
                    sIndex = index
                    Exit For
                End If
            Next index
            
            For index = 0 To lstKeys.ListCount - 1
                If lstKeys.Selected(index) Then
                    INI.DeleteKey lstSections.List(sIndex), lstKeys.List(index)
                    Exit For
                End If
            Next index
            
        ElseIf action = 2 Then
            For index = 0 To lstSections.ListCount - 1
                If lstSections.Selected(index) Then
                    sIndex = index
                    Exit For
                End If
            Next index
            
            For index = 0 To lstKeys.ListCount - 1
                If lstKeys.Selected(index) Then
                    INI.DeleteKeyValue lstSections.List(sIndex), lstKeys.List(index)
                    Exit For
                End If
            Next index
        End If
        
        loadSections
        loadKeysAndVals
        
        btnAdd.Enabled = True
        btnRename.Enabled = True
        btnDelete.Enabled = True
    End If
End Sub

Private Sub loadSections()
    Dim sectionList() As String
    Dim sectionCount As Integer
    Dim index As Integer
    Dim oldSel As Integer
    
    oldSel = 0
    
    sectionList = Split(INI.GetSections, Chr$(0))
    sectionCount = UBound(sectionList)
    
    If lstSections.SelCount > 0 Then
        For index = 0 To lstSections.ListCount - 1
            If lstSections.Selected(index) Then
                oldSel = index
                Exit For
            End If
        Next index
    End If
    
    lstSections.Clear
    For index = 0 To sectionCount
        lstSections.AddItem sectionList(index)
    Next index
    
    If lstSections.ListCount > 0 Then
        If oldSel >= lstSections.ListCount Then oldSel = lstSections.ListCount - 1
        lstSections.Selected(oldSel) = True
    End If
End Sub

Private Sub loadKeysAndVals()
    Dim KeyList() As String
    Dim KeyCount As Integer
    Dim index As Integer
    Dim section As Integer
    Dim oldSel As Integer
    
    oldSel = 0
    
    If lstSections.SelCount > 0 Then
        For index = 0 To lstSections.ListCount - 1
            If lstSections.Selected(index) Then
                section = index
                Exit For
            End If
        Next index
        
        KeyList = Split(INI.GetKeysInSection(lstSections.List(section)), Chr$(0))
        KeyCount = UBound(KeyList)
        
        If lstKeys.SelCount > 0 Then
            For index = 0 To lstKeys.ListCount - 1
                If lstKeys.Selected(index) Then
                    oldSel = index
                    Exit For
                End If
            Next index
        End If
        
        lstKeys.Clear
        lstValues.Clear
        For index = 0 To KeyCount
            lstKeys.AddItem KeyList(index)
            lstValues.AddItem INI.GetKeyValue(lstSections.List(section), KeyList(index))
        Next index
        
        If lstKeys.ListCount > 0 Then
            If oldSel >= lstKeys.ListCount Then oldSel = lstKeys.ListCount - 1
            lstKeys.Selected(oldSel) = True
        End If
    End If
End Sub

Private Sub setFields()
    Dim index As Integer
    
    txtSection.Text = ""
    txtKey.Text = ""
    txtValue.Text = ""
    
    If lstSections.SelCount > 0 Then
        For index = 0 To lstSections.ListCount - 1
            If lstSections.Selected(index) = True Then
                txtSection.Text = lstSections.List(index)
                Exit For
            End If
        Next index
        
        For index = 0 To lstKeys.ListCount - 1
            If lstKeys.Selected(index) = True Then
                txtKey.Text = lstKeys.List(index)
                Exit For
            End If
        Next index
        
        For index = 0 To lstValues.ListCount - 1
            If lstValues.Selected(index) = True Then
                txtValue.Text = lstValues.List(index)
                Exit For
            End If
        Next index
    End If
End Sub

Private Sub lstSections_Click()
    loadKeysAndVals
    setFields
End Sub

Private Sub lstKeys_Click()
    Dim index As Integer
    
    If lstKeys.SelCount > 0 Then
        For index = 0 To lstKeys.ListCount - 1
            If lstKeys.Selected(index) Then
                lstValues.Selected(index) = True
                Exit For
            End If
        Next index
    End If
    setFields
End Sub

Private Sub lstValues_Click()
    Dim index As Integer
    
    If lstValues.SelCount > 0 Then
        For index = 0 To lstValues.ListCount - 1
            If lstValues.Selected(index) Then
                lstKeys.Selected(index) = True
                Exit For
            End If
        Next index
    End If
    setFields
End Sub
