Imports System
Imports System.IO
Imports System.Collections

Public Module HMTUtility
    '---------------------------------------------------------------------------
    ' Global Utility Functions
    '---------------------------------------------------------------------------
    Public Function SwapENDIAN(ByVal strString As String) As String
        Dim strNew As String
        Dim intA, intB As Integer

        If Len(strString) Mod 2 <> 0 Then strString = strString & "0"
        If Len(strString) >= 2 Then
            For intA = Len(strString) To 1 Step -2
                intB = intB + 2
                strNew = strNew & strString.Substring(Len(strString) - intB, 2)
            Next intA
        End If
        Return strNew
    End Function

    'Use a nifty little trick to work arount VB.net's lack of unsigned data types.
    Public Function Unsigned(ByVal i As Integer) As Long
        Unsigned = CLng("&H" & Hex(i))
    End Function

    'Get a specified subdirectory level from a path string.
    Public Function getLevel(ByVal s As String, ByVal intChunk As Integer) As String
        'Redundant - the NumberOfLevels can be calculated in this function.
        Dim tmpStr() As String
        Dim intnumLevels As Integer = NumberOfLevels(s)
        Dim x As Integer

        ReDim tmpStr(intnumLevels + 1)
        For x = 1 To intnumLevels
            tmpStr(x) = Mid(s, 1, s.IndexOf("\") + 1) 'Get the string up to the "\"
            tmpStr(x) = Mid(tmpStr(x), 1, tmpStr(x).Length - 1) 'Trim the "\" off of the end.
            s = Mid(s, s.IndexOf("\") + 2, (s.Length + 2) - s.IndexOf("\"))
        Next
        tmpStr(0) = s 'Set the 0 element to the filename itself.
        getLevel = tmpStr(intChunk)
    End Function

    'Count the number of subdirectories in the path string.
    Public Function NumberOfLevels(ByVal s As String) As Integer
        Dim x As Integer = 0
        Dim y As Integer = 0

        For x = 1 To s.Length
            If Mid(s, x, 1) = "\" Then
                y += 1
            End If
        Next
        NumberOfLevels = y
    End Function

    Public Function reverseString(ByVal s As String) As String
        Dim s2 As String = ""
        For x As Integer = Len(s) To 1 Step -1
            s2 &= s.Chars(x - 1)
        Next
        Return s2
    End Function

    Public Enum OperandEnum
        Add = 1
        Subtract = 2
    End Enum
    Public Function UMath(ByVal o As OperandEnum, ByVal NumberOne As UInt32, ByVal NumberTwo As Integer) As UInt32
        Try
            Dim i As Int64 = Convert.ToInt64(NumberOne)
            Select Case o
                Case OperandEnum.Add
                    Return Convert.ToUInt32(i + NumberTwo)
                Case OperandEnum.Subtract
                    Return Convert.ToUInt32(i - NumberTwo)
            End Select
        Catch
            Return Convert.ToUInt32(0)
        End Try

    End Function

    Public Function qwordTodword(ByVal l As Long) As Integer
        Dim signBit As Short
        Dim coreNumber As Integer
        'Get the leftmost bit
        If l > 4294967295 Then Return -1 'This is the max that you can hold in an unsigned dword
        signBit = l >> 31
        coreNumber = (l << 33) >> 33
        Return coreNumber * (signBit * -1)
    End Function

    Public Function intToString(ByVal i As Long) As String
        Dim t As String = Hex(i)
        Dim s As String
        'Make the Hex string an even number of digits
        t.PadLeft(t.Length + (t.Length Mod 2), "0")
        For x As Integer = t.Length - 1 To 1 Step -2
            s &= Chr(Val("&H" & Mid(t, x, 2)))
        Next
        Return s
    End Function

    'Convert a C-style (0x) hex string to a long integer
    Public Function strHexToDec(ByVal strHex As String) As Long
        strHex = strHex.Remove(0, 2) 'Take off the "0x"
        Return Convert.ToInt64(strHex, 16)
    End Function
End Module

Public Class WINAPI
    ' Declare the Windows API function CopyMemory
    ' Note that all variables are ByVal. pDst is passed ByVal because we want
    ' CopyMemory to go to that location and modify the data that is pointed to
    ' by the IntPtr, and not the pointer itself.                   
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDst As IntPtr, _
                                                                 ByVal pSrc() As Byte, _
                                                                 ByVal ByteLen As Long)
    Declare Sub PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal data() As Byte, _
                                                              ByVal hMod As IntPtr, _
                                                              ByVal dwFlags As Int32)
End Class

Public Class RecursiveFileProcessor

    Public Structure FILE_LIST_STRUCT
        Implements IComparable
        Public filename As String
        Public tag As String
        Public processed As Boolean
        Public Function CompareTo(ByVal obj As Object) As Integer Implements System.IComparable.CompareTo
            Dim f As FILE_LIST_STRUCT = CType(obj, FILE_LIST_STRUCT)
            Return Me.filename.CompareTo(f.filename)
        End Function
    End Structure
    Public FileList() As FILE_LIST_STRUCT
    Public fileNumber As Integer = 0

    ' Process all files in the directory passed in, and recurse on any directories 
    ' that are found to process the files they contain
    Public Sub ProcessDirectory(ByVal targetDirectory As String, ByVal strExtension As String)
        Dim fileEntries As String() = Directory.GetFiles(targetDirectory)
        ' Process the list of files found in the directory
        Dim fileName As String
        For Each fileName In fileEntries
            If Right(fileName, Len(strExtension)) = strExtension Then ProcessFile(fileName)
        Next fileName

        Dim subdirectoryEntries As String() = Directory.GetDirectories(targetDirectory)
        ' Recurse into subdirectories of this directory
        Dim subdirectory As String
        For Each subdirectory In subdirectoryEntries
            ProcessDirectory(subdirectory, strExtension)
        Next subdirectory

    End Sub      'ProcessDirectory

    ' Real logic for processing found files would go here.
    Public Sub ProcessFile(ByVal path As String)
        Dim s As String
        fileNumber += 1
        ReDim Preserve FileList(fileNumber)
        FileList(fileNumber).filename = path
        s = path.Remove(path.Length - 5, 5)
        s = s.Remove(0, s.Length - 4)
        FileList(fileNumber).tag = s
        FileList(fileNumber).processed = False
    End Sub 'ProcessFile
End Class 'RecursiveFileProcessor

