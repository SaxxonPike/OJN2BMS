VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOJNFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OJN file class
' saxxonpike 2006-2007

Private Const BMEC = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const TitleFilter = "/\?:*|<>"

Private Type InternalNote
    Offs As Double
    NoteValue As Double
    MeasureNum As Long
    Lane As Integer
    Properties(0 To 1) As Byte
End Type

Private Type OJNHeader
    SongID As Long
    ojnID(0 To 1) As Long
    Genre As Long
    InitialBPM As Single
    Difficulty(0 To 3) As Integer
    TotalNoteCount(0 To 2) As Long
    PlayedNoteCount(0 To 2) As Long
    MeasureCount(0 To 2) As Long
    LineCount(0 To 2) As Long
    Unk(0 To 31) As Byte
    Title(0 To 63) As Byte
    Artist(0 To 31) As Byte
    Manufacturer(0 To 31) As Byte
    OJMFile(0 To 31) As Byte
    ImageSize As Long
    TimeInSeconds(0 To 2) As Long
    Offsets(0 To 3) As Long
End Type

Private Type OJMHeader
    M30ID1 As Long
    M30ID2(0 To 1) As Byte
    d1(0 To 2) As Integer
    SampleCount As Long
    d2(0 To 2) As Long
End Type

Private Type OJMSample
    SampleName(0 To 31) As Byte
    SampleSize As Long
    SampleType As Integer
    d1 As Integer
    d2 As Long
    SampleNumber As Long
    UncompressedSize As Long
End Type

Private Type xLoadedSample
    Info As OJMSample
    Dat() As Byte
End Type

Private Type NoteHead
    Measure As Long
    Lane As Integer
    Linesize As Integer
End Type

Private Type NoteType
    NoteID As Integer
    Flags(0 To 1) As Byte
End Type

Private bLoaded As Boolean
Private LoadedHeader As OJNHeader
Private OJMSampleInfo() As OJMSample
Private LoadedOJMHeader As OJMHeader

Private Notes() As InternalNote
Private LoadedSamples() As xLoadedSample

Private Function StrConv2(inArr() As Byte) As String
    StrConv2 = StrConv(inArr(), vbUnicode)
    If InStr(StrConv2, Chr(0)) > 0 Then
        StrConv2 = Left(StrConv2, InStr(StrConv2, Chr(0)) - 1)
    End If
End Function

Public Function LoadOJN(ByVal FName As String, ByVal DFolder As String, ByVal DoFreeze As Boolean, ByVal ShowErrors As Boolean) As Boolean
    Dim z As Long
    Dim f As Long
    Dim j As Long
    Dim s As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long
    Dim ThisLine As String
    Dim BPMList As String
    Dim aF() As Single
    Dim aL() As NoteType
    Dim aH As NoteHead
    Dim ImageMem() As Byte
    Dim ThumbMem() As Byte
    Dim TitleTag As String
    Dim FileName As String
    Dim Title As String
    Dim Genre As String
    Dim Artist As String
    Dim RealTitle As String
    Dim SampleUsed(0 To 1295) As Boolean
    
    If DFolder = "" Or FName = "" Then
        Exit Function
    End If
    If Right(DFolder, 1) <> "\" Then
        DFolder = DFolder + "\"
    End If
    
    
    ' ====== INITIAL READ ======
    
    Dim OJMFile As String
    bLoaded = False
    f = FreeFile
    Open FName For Binary As #f
    Get #f, 1, LoadedHeader
    If LoadedHeader.ojnID(0) <> 7236207 Then
        If ShowErrors Then
            MsgBox "Could not load OJN file. Identification mismatch."
        End If
        Close #f
        Exit Function
    End If
    OJMFile = Left(FName, InStrRev(FName, "\")) + StrConv2(LoadedHeader.OJMFile())
    Title = StrConv2(LoadedHeader.Title())
    For x = 1 To Len(TitleFilter) + 1
        Title = Replace(Title, Mid(TitleFilter + Chr(34), x, 1), "_")
    Next x
    Title = Trim(Title)
    Do While Right(Title, 1) = "."
        Title = Left(Title, Len(Title) - 1)
    Loop
    Title = Trim(Title)
    DFolder = DFolder + "[" + CStr(LoadedHeader.SongID) + "] " + Title
    On Error Resume Next
    MkDir DFolder
    On Error GoTo 0
    DFolder = DFolder + "\"
    
    
    '====== SAMPLE CONVERSION ======
    
    j = FreeFile
    Open OJMFile For Binary As #j
    Get #j, 1, LoadedOJMHeader
    s = 28
    n = 0
    If LoadedOJMHeader.M30ID1 <> 3158861 Then
        If ShowErrors Then
            MsgBox "Could not decode sample data: incorrect format.", vbCritical
        End If
        Close #j
        Close #f
        Exit Function
    End If
    ReDim LoadedSamples(0 To LoadedOJMHeader.SampleCount) As xLoadedSample
    Do While s < LOF(j)
        Get #j, s + 1, LoadedSamples(n).Info
        With LoadedSamples(n)
            ReDim .Dat(0 To .Info.SampleSize - 1) As Byte
            Get #j, s + 1 + 52, .Dat
            s = s + 52 + .Info.SampleSize
            For x = 0 To UBound(.Dat) - 3 Step 4
                .Dat(x + 0) = .Dat(x + 0) Xor 110
                .Dat(x + 1) = .Dat(x + 1) Xor 97
                .Dat(x + 2) = .Dat(x + 2) Xor 109
                .Dat(x + 3) = .Dat(x + 3) Xor 105
            Next x
            x = FreeFile
            If .Info.SampleType = 5 Then
                Open DFolder + BMEString(.Info.SampleNumber + 1) + ".ogg" For Binary As #x
                Put #x, 1, .Dat
                Close #x
                SampleUsed(.Info.SampleNumber + 1) = True
            ElseIf .Info.SampleType = 0 Then
                Open DFolder + BMEString(1294 - .Info.SampleNumber) + ".ogg" For Binary As #x
                SampleUsed(1294 - .Info.SampleNumber) = True
                Put #x, 1, .Dat
                Close #x
            Else
                x = x
            End If
        End With
        n = n + 1
    Loop
    
    Close #j
    
    
    '====== CHART CONVERSION ======
    
    RealTitle = StrConv2(LoadedHeader.Title())
    Artist = StrConv2(LoadedHeader.Artist())
    Select Case LoadedHeader.Genre
        Case 0: Genre = "Ballad"
        Case 1: Genre = "Rock"
        Case 2: Genre = "Dance"
        Case 3: Genre = "Techno"
        Case 4: Genre = "Hip-Hop"
        Case 5: Genre = "Soul"
        Case 6: Genre = "Jazz"
        Case 7: Genre = "Funk"
        Case 8: Genre = "Classic"
        Case 9: Genre = "Traditional"
        Case 10: Genre = "Etc."
    End Select
    
    j = FreeFile
    With LoadedHeader
        For n = 0 To 2
            z = 0
            BPMList = ""
            Select Case n
                Case 0
                    FileName = DFolder + Title + " [Easy]"
                    TitleTag = RealTitle + " [Easy]"
                Case 1
                    FileName = DFolder + Title + " [Normal]"
                    TitleTag = RealTitle + " [Normal]"
                Case 2
                    FileName = DFolder + Title + " [Hard]"
                    TitleTag = RealTitle + " [Hard]"
            End Select
            Open FileName + ".bme" For Output As #j
            Print #j, ";maxcombo "; CStr(.PlayedNoteCount(n) - 1)
            Print #j, "#TITLE "; TitleTag
            Print #j, "#ARTIST "; Artist
            Print #j, "#GENRE "; Genre
            Print #j, "#PLAYLEVEL "; CStr(.Difficulty(n))
            Print #j, "#BPM "; CStr(.InitialBPM)
            If DoFreeze = True Then
                Print #j, "#LNTYPE 1"
                Print #j, "#LNOBJ ZZ"
            End If
            For x = 0 To 1294
                If SampleUsed(x) = True Then
                    Print #j, "#WAV" + BMEString(x + 0) + " " + BMEString(x + 0) + ".WAV"
                End If
            Next x
            s = LoadedHeader.Offsets(n) + 1
            For x = 1 To LoadedHeader.LineCount(n)
                Get #f, s, aH
                ThisLine = "#" + AddZeros(CStr(aH.Measure + 1), 3)
                s = s + 8
                Select Case aH.Lane
                    Case 0
                        'do nothing
                    Case 1
                        ThisLine = ThisLine + "08:"
                        ReDim aF(0 To aH.Linesize - 1) As Single
                        Get #f, s, aF
                        For y = 0 To aH.Linesize - 1
                            If aF(y) <> 0 Then
                                z = z + 1
                                BPMList = BPMList + "#BPM" + BMEString2(z + 0) + " " + CStr(aF(y)) + vbCrLf
                                ThisLine = ThisLine + BMEString2(z + 0)
                            Else
                                ThisLine = ThisLine + "00"
                            End If
                        Next y
                    Case Else
                        Select Case aH.Lane
                            Case 2: ThisLine = ThisLine + "11:"
                            Case 3: ThisLine = ThisLine + "12:"
                            Case 4: ThisLine = ThisLine + "13:"
                            Case 5: ThisLine = ThisLine + "14:"
                            Case 6: ThisLine = ThisLine + "15:"
                            Case 7: ThisLine = ThisLine + "18:"
                            Case 8: ThisLine = ThisLine + "19:"
                            Case Else: ThisLine = ThisLine + "01:"
                        End Select
                        ReDim aL(0 To aH.Linesize - 1) As NoteType
                        Get #f, s, aL
                        For y = 0 To aH.Linesize - 1
                            Select Case aL(y).Flags(1)
                                Case 0, 2
                                    ThisLine = ThisLine + BMEString(aL(y).NoteID)
                                Case 1
                                    ThisLine = ThisLine + "00"
                                Case 3
                                    If DoFreeze = True Then
                                        ThisLine = ThisLine + BMEString(1295)
                                    Else
                                        ThisLine = ThisLine + "00"
                                    End If
                                Case 4
                                    ThisLine = ThisLine + BMEString(1295 - aL(y).NoteID)
                            End Select
                        Next y
                End Select
                Print #j, ThisLine
                s = s + (aH.Linesize * 4)
            Next x
            Print #j, BPMList
            Close #j
        Next n
    End With
    
    
    '====== IMAGE EXPORT ======
    ReDim ImageMem(0 To LoadedHeader.ImageSize - 1) As Byte
    ReDim ThumbMem(0 To (LOF(f) - (LoadedHeader.ImageSize + LoadedHeader.Offsets(3))) - 1) As Byte
    Get #f, LoadedHeader.Offsets(3) + 1, ImageMem
    Get #f, , ThumbMem
    j = FreeFile
    Open DFolder + "Image.jpg" For Binary As #j
    Put #j, 1, ImageMem
    Close #j
    Open DFolder + "Thumb.bmp" For Binary As #j
    Put #j, 1, ThumbMem
    Close #j
    
    
    
    Close #f
    LoadOJN = True
    bLoaded = True
End Function

Private Sub Class_Initialize()
    ReDim Notes(0) As InternalNote
    ReDim LoadedSamples(0) As xLoadedSample
End Sub

Private Function BMEString(inVal As Integer) As String
    BMEString = Mid(BMEC, (inVal \ 36) + 1, 1) + Mid(BMEC, (inVal Mod 36) + 1, 1)
End Function

Private Function BMEString2(inVal As Integer) As String
    BMEString2 = Mid(BMEC, (inVal \ 10) + 1, 1) + Mid(BMEC, (inVal Mod 10) + 1, 1)
End Function

Private Function AddZeros(iString As String, iNum As Long) As String
    AddZeros = iString
    Do While Len(AddZeros) < iNum
        AddZeros = "0" + AddZeros
    Loop
End Function
