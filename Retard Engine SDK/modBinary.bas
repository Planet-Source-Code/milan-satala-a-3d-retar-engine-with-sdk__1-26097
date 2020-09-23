Attribute VB_Name = "modBinary"
'--------------------------------------------------------
'Binary extractor by TOMAS BANOVEC
'
'(c) 2001 by Mental Soft
'--------------------------------------------------------

'Sample of calling EXTRACT function:
'Extract SourceFile, ExtractPath, FileNameToExtract
'if FileNameToExtract is "" then will procedure extract all files in archive

Public Function GetField(field As String, FieldPos As Long) As String
On Error Resume Next
Dim FieldCounter As Long
Dim IPPositionStart As Long
Dim IPPositionEnde As Long
Dim TempPosition As Long
Dim OpenedID As String
    TempPosition = 1
    
    For FieldCounter = 1 To FieldPos - 1 Step 1
        IPPositionStart = InStr(TempPosition, field, "|", vbTextCompare)
        TempPosition = IPPositionStart + 1
    Next FieldCounter
    IPPositionStart = IPPositionStart + 1
    IPPositionEnde = InStr(IPPositionStart, field, "|", vbTextCompare)
    If IPPositionEnde >= IPPositionStart Then GetField = Mid(field, IPPositionStart, IPPositionEnde - IPPositionStart)
End Function

Function Extract(ByVal FileToExtract As String, ByVal ExtractPath As String, Optional ByVal TextureName As String) As Boolean
Dim t, W As Variant
Dim m_Buffer As String
Dim m_BufferLen As String
Dim m_Temp As String
Dim FilesCount As Integer
Dim m_Names() As String
Dim m_Size() As Long
Dim i As Integer
Dim blCorrect As Boolean
Dim m_FileCursor As Long 'get position in text file
blCorrect = False
'TextureName - name of file to extract. If this option is not selected, function will extract all textures
' check for path
If Dir(FileToExtract) = "" Then
Extract = False 'if source file does not exist then exit function
Exit Function
End If

If Not Right(ExtractPath, 1) = "\" Then ExtractPath = ExtractPath & "\"
If Dir(ExtractPath, vbDirectory) = "" Then MkDir ExtractPath

'------------EXTRACT NAMES AND FILE SIZE--------------------
t = FreeFile
Open FileToExtract For Binary Access Read As #t
    m_Buffer = Space(200) 'select grabbed bytes to 200
    Get #t, , m_Buffer 'grab bytes
    m_BufferLen = Val(GetField(m_Buffer, 2)) + Len(GetField(m_Buffer, 1)) + Len(GetField(m_Buffer, 3)) + 5
    FilesCount = Val(GetField(m_Buffer, 3))
    ReDim m_Names(FilesCount) As String
    ReDim m_Size(FilesCount) As Long
    m_Buffer = Space(m_BufferLen)
    Get #t, 1, m_Buffer
        For i = 1 To FilesCount
            m_Names(i) = GetField(m_Buffer, 2 * i + 2)
            If LCase(m_Names(i)) = LCase(TextureName) Then blCorrect = True
            m_Size(i) = Val(GetField(m_Buffer, 2 * i + 3))
        Next i
        
If blCorrect = False And Not TextureName = "" Then Extract = False: Exit Function
'----------------EXTRACT FILES------------------------------
m_FileCursor = m_BufferLen + 1 'set start byte to 1 byte after header

    For i = 1 To FilesCount
    FileSize = m_Size(i) 'get size of file to extract
    If LCase(TextureName) = LCase(m_Names(i)) Or TextureName = "" Then
     W = FreeFile
      Open ExtractPath & m_Names(i) For Binary Access Write As #W
        
        Do Until FileSize <= 0
            If FileSize >= CHUNK_SIZE Then
                m_Temp = Space(CHUNK_SIZE) 'this select how many bytes will be grabbed from source file
            Else
                m_Temp = Space(FileSize)
            End If
        FileSize = FileSize - CHUNK_SIZE
          Get #t, m_FileCursor, m_Temp
          Put #W, , m_Temp
          m_FileCursor = m_FileCursor + Len(m_Temp)
        Loop
      Close #W
    Else
    m_FileCursor = m_FileCursor + m_Size(i)
    End If
    
    Next i

Close #t
End Function
Function AddFile(ByVal PACFileName As String, ByVal FileToAdd As String) As Boolean
Dim t, W As Variant
Dim m_Buffer As String
Dim m_BufferLen As String
Dim m_Temp As String
Dim FilesCount As Integer
Dim m_Names() As String
Dim m_Size() As Long
Dim i As Integer
Dim blCorrect As Boolean
Dim m_FileCursor As Long 'get position in text file
Dim ExtractPath As String
AddFile = False

If FileToAdd = "" Then AddFile = False: Exit Function
ExtractPath = App.Path
If Not Right(ExtractPath, 1) = "\" Then ExtractPath = ExtractPath & "\"
ExtractPath = ExtractPath & "Temp\"
If Dir(ExtractPath, vbDirectory) = "" Then MkDir ExtractPath

'------------EXTRACT NAMES AND FILE SIZE--------------------
t = FreeFile
Open PACFileName For Binary Access Read As #t
    m_Buffer = Space(200) 'select grabbed bytes to 200
    Get #t, , m_Buffer 'grab bytes
    m_BufferLen = Val(GetField(m_Buffer, 2)) + Len(GetField(m_Buffer, 1)) + Len(GetField(m_Buffer, 3)) + 5
    FilesCount = Val(GetField(m_Buffer, 3))
    m_Buffer = Space(m_BufferLen) 'select grabbed bytes to 100
    Get #t, 1, m_Buffer 'grab bytes
    ReDim m_Names(FilesCount + 1) As String
    ReDim m_Size(FilesCount + 1) As Long
        For i = 1 To FilesCount
                m_Names(i) = GetField(m_Buffer, 2 * i + 2)
                m_Size(i) = Val(GetField(m_Buffer, 2 * i + 3))
        Next i
'--------------------ADD FILE TO LIST--------------------

          m_Names(FilesCount + 1) = Dir(FileToAdd)
          m_Size(FilesCount + 1) = FileLen(FileToAdd)


'----------------EXTRACT FILES------------------------------
W = FreeFile
    Open ExtractPath & "temp.PAC" For Binary Access Write As #W
Temp = ""
    For i = 1 To FilesCount + 1
        Temp = Temp & m_Names(i) & "|" & m_Size(i) & "|"
    Next i
Temp = "RETARD ENGINE PAC file|" & Len(Temp) & "|" & FilesCount + 1 & "|" & Temp
Put #W, 1, Temp
Dim m_FileCursor3 As Long
m_FileCursor3 = Val(m_BufferLen) + 1
m_FileCursor = Len(Temp) + 5 'set start byte to 1 byte after header
    For i = 1 To FilesCount
    FileSize = m_Size(i) 'get size of file to extract
        Do Until FileSize <= 0
            If FileSize >= CHUNK_SIZE Then
                m_Temp = Space(CHUNK_SIZE) 'this select how many bytes will be grabbed from source file
            Else
                m_Temp = Space(FileSize)
            End If
        FileSize = FileSize - CHUNK_SIZE
          Get #t, m_FileCursor3, m_Temp
          Put #W, m_FileCursor, m_Temp
          m_FileCursor3 = m_FileCursor3 + Len(m_Temp)
          m_FileCursor = m_FileCursor + Len(m_Temp)
        Loop
    Next i
Close #t
'-----------------NOW ADD FILE INTO ARCHIVE-------------
Dim m_FileCursor2 As Long
    m_FileCursor2 = 1
    t = FreeFile
    Open FileToAdd For Binary Access Read As #t
    FileSize = m_Size(FilesCount + 1) 'get size of file to extract
    
        Do Until FileSize <= 0
            If FileSize >= CHUNK_SIZE Then
                m_Temp = Space(CHUNK_SIZE) 'this select how many bytes will be grabbed from source file
            Else
                m_Temp = Space(FileSize)
            End If
        FileSize = FileSize - CHUNK_SIZE
          Get #t, m_FileCursor2, m_Temp
          Put #W, m_FileCursor, m_Temp
          m_FileCursor2 = m_FileCursor2 + Len(m_Temp)
          m_FileCursor = m_FileCursor + Len(m_Temp)
        Loop
       
    
Close #W
Close #t

Kill PACFileName
FileCopy ExtractPath & "temp.PAC", PACFileName
Kill ExtractPath & "temp.PAC"
RmDir Left(ExtractPath, Len(ExtractPath) - 1)
AddFile = True
End Function

Function DeleteFile(ByVal PACFileName As String, ByVal FileToDelete As String) As Boolean
Dim t, W As Variant
Dim m_BufferLen As String
Dim m_Buffer As String
Dim m_Temp As String
Dim FilesCount As Integer
Dim m_Names() As String
Dim m_Size() As Long
Dim i As Integer
Dim blCorrect As Boolean
Dim m_FileCursor As Long 'get position in text file
Dim ExtractPath As String

ExtractPath = App.Path
If Not Right(ExtractPath, 1) = "\" Then ExtractPath = ExtractPath & "\"
ExtractPath = ExtractPath & "Temp\"
If Dir(ExtractPath, vbDirectory) = "" Then MkDir ExtractPath

'------------EXTRACT NAMES AND FILE SIZE--------------------
t = FreeFile
Open PACFileName For Binary Access Read As #t
    m_Buffer = Space(200) 'select grabbed bytes to 200
    Get #t, , m_Buffer 'grab bytes
    m_BufferLen = Val(GetField(m_Buffer, 2)) + Len(GetField(m_Buffer, 1)) + Len(GetField(m_Buffer, 3)) + 5
    FilesCount = Val(GetField(m_Buffer, 3))
    m_Buffer = Space(m_BufferLen) 'select grabbed bytes to 100
    Get #t, 1, m_Buffer 'grab bytes
        ReDim m_Names(FilesCount) As String
    ReDim m_Size(FilesCount) As Long
        For i = 1 To FilesCount
                m_Names(i) = GetField(m_Buffer, 2 * i + 2)
                m_Size(i) = Val(GetField(m_Buffer, 2 * i + 3))
        Next i


'----------------EXTRACT FILES------------------------------
W = FreeFile
      Open ExtractPath & "temp.PAC" For Binary Access Write As #W

    For i = 1 To FilesCount
        If Not LCase(m_Names(i)) = LCase(FileToDelete) Then
            Temp = Temp & m_Names(i) & "|" & m_Size(i) & "|"
        End If
    Next i
    
Temp = "PAC file created by PAC WRITER v 1.01 - (c) 2001 by MET@L SOFT|" & Len(Temp) & "|" & FilesCount - 1 & "|" & Temp
Put #W, 1, Temp
Dim m_FileCursor2 As Long
m_FileCursor = Val(m_BufferLen) + 1 'set start byte to 1 byte after header
m_FileCursor2 = Len(Temp) + 5
    For i = 1 To FilesCount
    FileSize = m_Size(i) 'get size of file to extract
    If Not LCase(m_Names(i)) = LCase(FileToDelete) Then
        
        Do Until FileSize <= 0
            If FileSize >= CHUNK_SIZE Then
                m_Temp = Space(CHUNK_SIZE) 'this select how many bytes will be grabbed from source file
            Else
                m_Temp = Space(FileSize)
            End If
        FileSize = FileSize - CHUNK_SIZE
          Get #t, m_FileCursor, m_Temp
          Put #W, m_FileCursor2, m_Temp
          m_FileCursor = m_FileCursor + Len(m_Temp)
          m_FileCursor2 = m_FileCursor2 + Len(m_Temp)
        Loop
      
    Else
    m_FileCursor = m_FileCursor + m_Size(i)
    End If
    
    Next i
Close #W
Close #t
Close
Kill PACFileName
Close
FileCopy ExtractPath & "temp.PAC", PACFileName
Kill ExtractPath & "temp.PAC"
RmDir Left(ExtractPath, Len(ExtractPath) - 1)

End Function

Function CoDec(ByVal CoDe As String) As String
Dim R As Long
Dim z_Buff As String
For R = 1 To Len(CoDe)

z_Buff = z_Buff & Mid(CoDe, R, 1)
Next R
End Function
