Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const MAX_PATH = 255
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const OPEN_EXISTING = 3

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_MULTI_SZ = 7

Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0

Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Sub Main()

'===============================================
'Worm engine by keef
'---------------------
'07/2021
'===============================================



'===============================================
'GetDisks()
'---------------
'this function lists the disks in the system
'===============================================
Dim i, X
i = Split(GetDisks, "-")
For Each X In i
MsgBox X, 0, "GetDisks"
Next

'===============================================
'CK
'--------
'this function creates an entry in regedit
'===============================================
Call CK(2, "SOFTWARE", "test", "asdasdasd", 1)
Call CK(2, "SOFTWARE", 123, 1234, 2)

'===============================================
'RK
'-------
'this function reads a key from the registry
'===============================================
Call MsgBox(RK(2, "SOFTWARE", "test", 1), 0, "RK")
Call MsgBox(RK(2, "SOFTWARE", 123, 2), 0, "RK")

'===============================================
'VAL_DIR
'------------
'this functions checks attributes, if the attribute
'is from a folder the function will return TRUE
'===============================================
Call CreateDirectory("C:\New folder", 0)
MsgBox VAL_DIR(GetAttr("C:\New folder")), 0, "VAL_DIR"

'===============================================
'SF
'----------
'this function returns some special folders
'sf(0) = c:\windows
'sf(1) = c:\windows\system32
'sf(2) = c:\users\username\appdata\local\temp
'sf(3) = the route of the .exe
'sf(4) = c:\ or the disk where windows its installed
'===============================================
MsgBox SF(0), 0, "SF"
MsgBox SF(1), 0, "SF"
MsgBox SF(2), 0, "SF"
MsgBox SF(3), 0, "SF"
MsgBox SF(4), 0, "SF"

'===============================================
'PE
'--------
'this function gets the route of the .exe
'just like sf(3)
'===============================================
MsgBox PE, 0, "PE"

'===============================================
'FILE_EXISTS
'-------------
'this function returns TRUE if the file mentioned
'exists, if not, it returns FALSE
'===============================================
MsgBox FILE_EXISTS(SF(1) + "rundll32.exe"), 0, "FILE_EXISTS"

'===============================================
'GET_NAME
'-----------
'this function gets the route without the extension
'of a determined file (ex: c:\windows\regedit instead of
'c:\windows\regedit.exe)
'===============================================
MsgBox GET_NAME(SF(1) + "rundll32.exe"), 0, "GET_NAME"

'===============================================
'UF
'--------
'this function returns some special folders
'located in the users directory (like appdata,
'or the desktop, etc.)
'===============================================
MsgBox UF(16), 0, "UF"

'===============================================
'VAL_PATH
'---------
'this function validates paths
'c:\windows -> c:\windows\
'or
'c: -> c:\
'this is useful for doing things with files
'avoiding errors
'===============================================
MsgBox VAL_PATH("C"), 0, "VAL_PATH"
MsgBox VAL_PATH("C:"), 0, "VAL_PATH"
MsgBox VAL_PATH("C:\Windows"), 0, "VAL_PATH"
MsgBox VAL_PATH("C:\Windows\"), 0, "VAL_PATH"

'===============================================
'FileSize
'-----------
'this function gets the size of a file
'in bytes
'===============================================
MsgBox FileSize(SF(1) + "rundll32.exe"), 0, "FileSize"

'
'
'
'
'
'
'
'
'
'
'===============================================
' made with love <3
' by keef
'===============================================
End Sub
Public Function GetDisks()
On Error Resume Next
Dim S As Long, X, z, n, i
        S = GetLogicalDrives()
            For i = 0 To 25
                If (S And 2 ^ i) Then
                    X = X & Chr$(65 + i) & ":\" & "$"
                    z = Split(X, "$")
                        For Each n In z
                            If Len(n) > 1 Then
                                If InStr(GetDisks, n) = 0 Then
                                        GetDisks = GetDisks + n + "-"
                                    End If
                                End If
                        Next
                End If
Next
End Function
Public Sub CK(ByVal lRegKey As Long, ByVal lKey As String, ByVal lValueName As String, lValue As Variant, ByVal lValueType As Long)
On Error Resume Next
    Dim H       As Long
    Dim Key     As Long

    Select Case lRegKey

        Case 0
        lRegKey = HKEY_CLASSES_ROOT
    
        Case 1
        lRegKey = HKEY_LOCAL_MACHINE

        Case 2
        lRegKey = HKEY_CURRENT_USER
    
    End Select

        H = RegCreateKeyEx(lRegKey, lKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, Key, H)
    
            If lValueType = 1 Then
                Dim szVal$
                szVal = lValue
                H = RegSetValueEx(Key, lValueName, 0&, REG_SZ, ByVal szVal, Len(szVal) + 1)
                Else
                Dim lVal&
                lVal = lValue
                H = RegSetValueEx(Key, lValueName, 0&, REG_DWORD, lVal, 4)
                End If

H = RegCloseKey(Key)
End Sub
Public Function RK(lRegKey As Long, ByVal lKey As String, ByVal lName As String, ByVal lType As Long) As Variant
On Error Resume Next
    Dim Key     As Long
    Dim zKey    As Long
    Dim hKey    As String
    Dim sx      As Long
    Select Case lRegKey

        Case 0
        lRegKey = HKEY_CLASSES_ROOT
        
        Case 1
        lRegKey = HKEY_LOCAL_MACHINE

        Case 2
        lRegKey = HKEY_CURRENT_USER
    
    End Select

    If lType = 1 Then
        lType = REG_SZ
        Else
        lType = REG_DWORD
        End If

    Call RegOpenKeyEx(lRegKey, lKey, 0, KEY_QUERY_VALUE, zKey)
    Call RegQueryValueEx(zKey, lName, 0, lType, ByVal 0&, Key)
    
        If lType = REG_SZ Then
                hKey = String(Key, Chr$(0))
                Call RegQueryValueEx(zKey, lName, 0, 0, ByVal hKey, Key)
                hKey = Left$(hKey, Key - 1)
                RK = hKey
                Call RegCloseKey(zKey)
            ElseIf lType = REG_DWORD Then
                Call RegQueryValueEx(zKey, lName, 0, 0, sx, Key)
                    RK = sx
                    Call RegCloseKey(zKey)
                End If
End Function
Public Function VAL_DIR(ByVal lAttributes As Long) As Boolean
On Error Resume Next
    If lAttributes = 16 Then VAL_DIR = True
    If lAttributes = 16 + 2 Then VAL_DIR = True
    If lAttributes = 16 + 4 Then VAL_DIR = True
    If lAttributes = 16 + 6 Then VAL_DIR = True
    If lAttributes = 16 + 1 Then VAL_DIR = True
    If lAttributes = 16 + 7 Then VAL_DIR = True
    If lAttributes = 16 + 5 Then VAL_DIR = True
    If lAttributes = 16 + 1 + 7 Then VAL_DIR = True
    If lAttributes = 16 + 7 + 5 Then VAL_DIR = True
End Function
Public Function SF(ByRef v1 As Byte) As String
On Error Resume Next

    Dim A$
    Dim B&
    
        Select Case v1
        
            Case 0
            A = String$(255, Chr$(0))
            B = GetWindowsDirectory(A, MAX_PATH)
            SF = VAL_PATH(Mid$(A, 1, InStr(A, Chr$(0)) - 1))
            A = "": B = 0
            
            Case 1
            A = String$(255, Chr$(0))
            B = GetSystemDirectory(A, MAX_PATH)
            SF = VAL_PATH(Mid$(A, 1, InStr(A, Chr$(0)) - 1))
            A = "": B = 0
            
            Case 2
            A = String$(255, Chr$(0))
            B = GetTempPath(MAX_PATH, A)
            SF = Mid$(A, 1, InStr(A, Chr$(0)) - 1)
            A = "": B = 0
            
            Case 3
            Dim c, i As Integer
            c = Array(".exe", ".com", ".cmd", ".bat", ".scr", ".dll", ".bat", ".pif", ".dat", ".db", ".sys")
                For i = 0 To UBound(c)
                    If FILE_EXISTS(GET_NAME(PE) + c(i)) Then
                        SF = GET_NAME(PE) + c(i)
                        End If
                Next
                
            Case 4
            SF = VAL_PATH(Environ$("SystemDrive"))
            
        End Select

End Function
Private Function PE() As String
On Error Resume Next
Dim z$
    
    z = Space$(255)
        
        Call GetModuleFileName(0, z, MAX_PATH)
    
PE = Mid$(z, 1, InStr(z, Chr$(0)) - 1)
z = ""
End Function
Public Function FILE_EXISTS(ByVal lFile As String) As Boolean
On Error Resume Next

FILE_EXISTS = IIf(Dir(lFile, vbArchive + vbHidden + vbNormal _
+ vbSystem) <> "", True, False)

End Function
Public Function GET_NAME(ByVal lFile As String) As String
On Error Resume Next

GET_NAME = Left$(lFile, InStr(lFile, ".") - 1)

End Function
Public Function UF(ByVal lFolder As Byte) As String
On Error Resume Next
Dim Rt      As Long
Dim IDL     As ITEMIDLIST
Dim Path As String

Rt = SHGetSpecialFolderLocation(100, lFolder, IDL)
       Path = Space$(512)
       Rt = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path)
       UF = Left$(Path, InStr(Path, Chr$(0)) - 1)
       UF = VAL_PATH(UF)
       Exit Function
End Function

Public Function VAL_PATH(ByVal v1 As String) As String
On Error Resume Next
If Len(v1) = 1 Then
    VAL_PATH = v1 + ":\"
    ElseIf Len(v1) = 2 Then
        If Right(v1, 1) = ":" Then VAL_PATH = v1 + "\"
        If Right(v1, 1) <> ":" Then VAL_PATH = Left$(v1, 1) + ":\"
    ElseIf Len(v1) = 3 Then
        If Right$(v1, 2) = ":\" Then VAL_PATH = v1
        If InStr(v1, ":") = 0 Then VAL_PATH = Left$(v1, 1) + ":\"
        If Right$(v1, 1) <> "\" Then VAL_PATH = v1 + "\"
        If InStr(v1, "\") = 0 Then VAL_PATH = Left$(v1, 1) + ":\"
        Else
        If Right$(v1, 1) <> "\" Then VAL_PATH = v1 + "\"
        If Right$(v1, 1) = "\" Then VAL_PATH = v1
    End If
End Function
Public Function FileSize(ByVal lFile As String) As Long
On Error Resume Next
    
    Dim c As Long
    
    c = CreateFile(lFile, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, &H80, ByVal 0&)
    FileSize = GetFileSize(c, ByVal 0&)

Call CloseHandle(c)
End Function


