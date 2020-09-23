VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************************
' Author:   Eric Dalquist
' email:    ebdalqui@mtu.edu (feel free to email me with questions)
'
' Credit:   Well I found a similar INI class on planet source code about a year
'           ago and search for it so I could give though athor some credit but
'           couldn't find it. So if you think this is based on your code it may
'           well be and I give you credit for it.
'
' Comments: Well the INI API calls are fairly simple and easy to use directly in
'           a program. I felt I wanted something with a bit more power and ease
'           of use. This INI class can be compiled on it's own and included into
'           a project as a dll or you can add the cslINI.cls file to your
'           project so you don't have to keep track of another dll. This class
'           includes all the functions I could think of uses for and actually get
'           to work. One thing that may seem missing is a createSection method.
'           I tried to get this to work and id does kind of if you don't mind a
'           null entry in your section. To do it just take the same idea from
'           the createKey method and replace Key with "" in the API call.
'
' Bugs:     Well I don't know if I would call this a bug but it sure fits under
'           quirks. You might notice in some spots I use "" and in others
'           vbNullString which are supposedly the same thing. When passed to the
'           PrivateProfileString API calls they behave much differently. For
'           example if you create a key with the value "" the key will exist in
'           the INI but have no value. If you create the same key with
'           vbNullString the key will not be written to the file. This is how I
'           ended up implementing the delete methods. Whatever works right?
'********************************************************************************

'*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!'
' REMEMBER to make the cslINI modual 6 - GlobalMultiUse under instancing in it's '
' properties if compiling this as a standalone DLL! Using the program            '
' cslINIDemo_src changes this setting to private.                                '
'*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!'

Option Explicit

'The two API calls to read and write from the INI file
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)

'the INI file's path
Private var_INIFile As String
'What to return if no data is found
Private var_DefaultRetVal As String

Private Sub Class_Initialize()
    'when the object is created we set the default file name
    ResetINIFilePath
    'our default for no data is "" (vbNullString)
    var_DefaultRetVal = ""
End Sub

Private Sub primeINIFile()
    'Priming the INI file stops the API calls from complaining if the file doesn't exist
    'I'm pretty sure there is a better way to do this but for the sake of
    'ease I just open the INIFile for append (so we don't screw with the data
    'already in it) and then close it right away. If the file was already there nothing
    'happens otherwise the file gets created
    Open var_INIFile For Append As #1
    Close #1
End Sub

'****************************** Start Let Properties ******************************
Public Property Let INIFile(ByVal INIFileIn As String)
    'set the INI path
    var_INIFile = INIFileIn
End Property

Public Property Let DefaultReturnValue(ByVal RetVal As String)
    'set the default return value
    var_DefaultRetVal = RetVal
End Property
'****************************** End Let Properties ******************************

'****************************** Start Get Properties ******************************
Public Property Get INIFile() As String
    'return the INI path
    INIFile = var_INIFile
End Property

Public Property Get DefaultReturnValue() As String
    'return the default retun value
    DefaultReturnValue = var_DefaultRetVal
End Property
'****************************** End Get Properties ******************************

Public Sub ResetINIFilePath()
    'this sets the INI file to the programs path and exe name but adds a .ini on the end
    'ex: the program c:\iniTest\test.exe would have the INI file c:\iniTest\test.ini
    'this set up makes for easy refrence to the file
    
    'if we are running in the root folder app.path will have a \ on the end.
    If Right$(App.Path, 1) = "\" Then
        var_INIFile = App.Path & App.EXEName & ".ini"
    Else 'Otherwise we need to add our own.
        var_INIFile = App.Path & "\" & App.EXEName & ".ini"
    End If
End Sub

'****************************** Start Create Methods ******************************
Public Sub CreateKeyValue(ByVal section As String, ByVal key As String, ByVal value As String, Optional ByVal INIFileLoc As String)
'************************************************************
'This little block of code is in all the sub and functions
'that use the INI file. Since there is a file variable that
'is set by default and can be modified by the user the
'INIFileLoc parameter is optional. This block of code checks
'to see if the argument was passed. If it was the UseFile var
'is set to the passed file if not it is set to the local file
'variable.
    Dim UseFile As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
'************************************************************
    
    'call the write profile string with all the parameters
    WritePrivateProfileString section, key, value, UseFile
End Sub

Public Sub CreateKey(ByVal section As String, ByVal key As String, Optional ByVal INIFileLoc As String)
    Dim UseFile As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    
    'call the write profile string but since we only want to create a key but
    'not give it a value we pass "" (vbNullString) for the value
    WritePrivateProfileString section, key, "", UseFile
End Sub
'****************************** End Create Methods ******************************

'****************************** Start Get Methods ******************************
Public Function GetKeyValue(ByVal section As String, ByVal key As String, Optional ByVal INIFileLoc As String) As String
    Dim RetVal As String
    Dim KeyLen As Integer
    Dim UseFile As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    
    'call our priming function to be sure the file exists
    primeINIFile
    
    'ok I just choose 255 as a nice number for the max length of data being retrieved
    'I think you can change it and the API calls will still work. It just creates a
    'string of that length in memory for the API call to write the results of it's action
    'to if the key is not found the contents of var_DefaultRetVal are returned
    RetVal = String$(255, 0)
    KeyLen = GetPrivateProfileString(section, key, "", RetVal, Len(RetVal), UseFile)
    
    'return a null string if the return has 0 length
    If KeyLen = 0 Then
        GetKeyValue = ""
    Else ' or the key's value
        GetKeyValue = Left$(RetVal, KeyLen)
    End If
End Function

Public Function GetKeysInSection(ByVal section As String, Optional ByVal INIFileLoc As String) As String
    Dim RetVal As String
    Dim KeyLen As Integer
    Dim UseFile As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    
    'prime the INI to make sure it exists
    primeINIFile
    
    'since we could have a lot of keys in one section we are going to use a 1024 char buffer
    RetVal = String$(1024, 0)
    KeyLen = GetPrivateProfileString(section, vbNullString, var_DefaultRetVal, RetVal, Len(RetVal), UseFile)
    
    'if no keys are found then return "" (vbNullString)
    If KeyLen = 0 Then
        GetKeysInSection = ""
    Else
        'if the retval is > 0 then return the results
        'since we are getting multiple keys but returning them as one string the
        'programer should use the split() function in the returned value with
        'chr$(0) being the delimiter
        GetKeysInSection = Left$(RetVal, KeyLen - 1)
    End If
End Function

Public Function GetSections(Optional ByVal INIFileLoc As String) As String
    Dim RetVal As String
    Dim KeyLen As Integer
    Dim UseFile As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    
    'prime the INI to make sure it exists
    primeINIFile
    
    'since we could have a lot of sections in one file we are going to use a 1024 char buffer
    RetVal = String$(1024, 0)
    KeyLen = GetPrivateProfileString(vbNullString, vbNullString, var_DefaultRetVal, RetVal, Len(RetVal), UseFile)
    
    'if no sections are found then return "" (vbNullString)
    If KeyLen = 0 Then
        GetSections = ""
    Else
        'if the retval is > 0 then return the results
        'since we are getting multiple sections but returning them as one string the
        'programer should use the split() function in the returned value with
        'chr$(0) being the delimiter
        GetSections = Left$(RetVal, KeyLen - 1)
    End If
End Function
'****************************** End Get Methods ******************************

'****************************** Start Delete Methods ******************************
Public Sub DeleteKeyValue(ByVal section As String, ByVal key As String, Optional ByVal INIFileLoc As String)
    Dim UseFile As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    
    'to delete the key's value just write a null string to it
    WritePrivateProfileString section, key, "", UseFile
End Sub

Public Sub DeleteKey(ByVal section As String, ByVal key As String, Optional ByVal INIFileLoc As String)
    Dim UseFile As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    
    'to delete a key use the actual vbnullstring object ... for some reason
    'it behaves differently that passing the API ""
    WritePrivateProfileString section, key, vbNullString, UseFile
End Sub

Public Sub DeleteSection(ByVal section As String, Optional ByVal INIFileLoc As String)
    Dim UseFile As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    
    'to delete a section it's the same thing as a key but we write a vbNullString
    'to the key name and value
    WritePrivateProfileString section, vbNullString, vbNullString, UseFile
End Sub
'****************************** End Delete Methods ******************************

'****************************** Start Rename Methods ******************************
Public Sub RenameKey(ByVal section As String, ByVal CurrentKey As String, ByVal NewKey As String, Optional ByVal INIFileLoc As String)
    Dim UseFile As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    
    'make sure the file exists
    primeINIFile
    
    'woohoo code reuse!
    'get the value and create the new key
    CreateKeyValue section, NewKey, GetKeyValue(section, CurrentKey, UseFile), UseFile
    'delete the old key
    DeleteKey section, CurrentKey, UseFile
End Sub

Public Sub RenameSection(ByVal CurrentSection As String, ByVal NewSection As String, Optional ByVal INIFileLoc As String)
    Dim KeyList() As String
    Dim KeyCount As Integer
    Dim index As Integer
    Dim UseFile As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    
    'makle sure the file exists
    primeINIFile
    
    'renaming a section works the same way.
    'we get all the key names into an array
    KeyList = Split(GetKeysInSection(CurrentSection, UseFile), Chr$(0))
    'get the number of keys
    KeyCount = UBound(KeyList)
    
    'for each key in the array
    For index = 0 To KeyCount
        'create the new key in the new section
        CreateKeyValue NewSection, KeyList(index), GetKeyValue(CurrentSection, KeyList(index), UseFile), UseFile
    Next index
    
    'delete the old section
    DeleteSection CurrentSection, UseFile
End Sub
'****************************** End Rename Methods ******************************
