Attribute VB_Name = "Module1"
Private Const FO_COPY = &H2&
Private Const FOF_ALLOWUNDO = &H40&
Private Const FOF_CONFIRMMOUSE = &H2&
Private Const FOF_CREATEPROGRESSDLG = &H0&
Private Const FOF_FILESONLY = &H80&
Private Const FOF_MULTIDESTFILES = &H1&
Private Const FOF_NOCONFIRMATION = &H10&
Private Const FOF_NOCONFIRMMKDIR = &H200&
Private Const FOF_RENAMEONCOLLISION = &H8&
Private Const FOF_SILENT = &H4& 'Progress not visible
Private Const FOF_SIMPLEPROGRESS = &H100& 'do not show filenames
Private Const FOF_WANTMAPPINGHANDLE = &H20&

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SHFileOperation Lib "Shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any) As Long
    
Public Function CopyFile(Source As String, Dest As String, AskOverwrite, Visib) As Boolean
    Dim lenFileop As Long
    Dim foBuf() As Byte
    Dim fileop As SHFILEOPSTRUCT
    lenFileop = LenB(fileop)
    ReDim foBuf(1 To lenFileop)
    With fileop
        .hwnd = Form1.hwnd
        .wFunc = FO_COPY
        .pFrom = Source & vbNullChar & vbNullChar & vbNullChar
        .pTo = Dest & vbNullChar & vbNullChar
        
        If AskOverwrite = False Then .fFlags = FOF_NOCONFIRMATION 'no ask to overwrite
        If Visib = False Then .fFlags = .fFlags Or FOF_SILENT 'not visible
            
        .lpszProgressTitle = "Kopiere " & Dest & vbNullChar & vbNullChar
    End With
    
    Call CopyMemory(foBuf(1), fileop, lenFileop)
    Call CopyMemory(foBuf(19), foBuf(21), 12)
    
    
    CopyFile = SHFileOperation(foBuf(1)) = 0
    

End Function


    

