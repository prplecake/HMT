'///////////////////////////////////////////////////////////
'// Halo Map Tools
'// Map File Base Class
'// File Structure Reverse-Engineered by PfhorSlayer and N-
'///////////////////////////////////////////////////////////

Imports System.IO

Public Class HaloMap

#Region "Structures"

    '// Main File Header Structure
    Public Structure FILE_HEADER_STRUCT
        Const HEADER_EMPTY_INTS As Integer = 485
        Public id As Integer
        Public version As Integer
        Public decomp_len As Integer
        Public unknown1 As Integer
        Public offset_to_index_decomp As Int32
        Public unknown2 As Integer
        <VBFixedString(8)> Public zeros As String
        <VBFixedString(32)> Public name As String
        <VBFixedString(32)> Public builddate As String
        Public maptype As Integer
        Public unknown3 As Integer
        Public empty_ints() As Integer 'Length will be dimensioned as HEADER_EMPTY_INTS
        Public footer As Integer
        Public Sub ReadStruct(ByRef bs As BinaryReader)
            Dim x As Integer 'Loop Counter
            id = bs.ReadInt32
            version = bs.ReadInt32
            decomp_len = bs.ReadInt32
            unknown1 = bs.ReadInt32
            offset_to_index_decomp = bs.ReadInt32
            unknown2 = bs.ReadInt32
            zeros = bs.ReadChars(8)
            name = bs.ReadChars(32)
            builddate = bs.ReadChars(32)
            maptype = bs.ReadInt32
            id = bs.ReadInt32
            unknown3 = bs.ReadInt32
            ReDim empty_ints(HEADER_EMPTY_INTS)
            For x = 1 To HEADER_EMPTY_INTS
                empty_ints(x) = bs.ReadInt32
            Next x
            footer = bs.ReadInt32
        End Sub
    End Structure

    '// Index Header Structure
    Public Structure INDEX_HEADER_STRUCT
        Public index_magic As Int32
        Public starting_id As Integer 'The value of the first object identifier in the index.
        Public unknown2 As Integer
        Public tagcount As Integer 'Number of items in the index.
        Public vertex_object_count As Integer
        Public vertex_offset As Integer
        Public indeces_object_count As Integer
        Public indeces_offset As Integer
        Public tagstart As Integer 'Offset of item index.
        Public Sub ReadStruct(ByRef bs As BinaryReader)
            index_magic = bs.ReadInt32
            starting_id = bs.ReadInt32
            unknown2 = bs.ReadInt32
            tagcount = bs.ReadInt32
            vertex_object_count = bs.ReadInt32
            vertex_offset = bs.ReadInt32
            indeces_object_count = bs.ReadInt32
            indeces_offset = bs.ReadInt32
            tagstart = bs.ReadInt32
        End Sub
    End Structure

    '// Index Item Structure
    Public Structure INDEX_ITEM_STRUCT
        Public tagclass As String 'Length is 12
        Public id As Integer
        Public stringoffset As Integer
        Public offset As Integer
        Public zeros As String 'Length is 8
        Public Sub ReadStruct(ByRef bs As BinaryReader)
            Dim x As Integer
            For x = 1 To 12 ' The fucking readchars function doesnt grab the number you tell it
                ' In fact, the number has anthing to do with the parameter passed... =/
                ' It worked fine on the previous functions... POS!!
                tagclass &= Chr(bs.ReadByte)
            Next x
            'MsgBox(Hex(bs.BaseStream.Position))
            id = bs.ReadInt32
            stringoffset = bs.ReadInt32
            offset = bs.ReadInt32
            zeros = bs.ReadChars(8)
        End Sub
    End Structure

    '// Expanded Index Item Structure
    Public Structure INDEX_ITEM_EXPANDED_STRUCT
        Public fileData As INDEX_ITEM_STRUCT
        Public filePath As String
        Public tagClass As String
        Public estimatedMetaSize As Int32
        Public offset As Int32
        Public magic_metadata_offset As Int64
        Public magic_string_offset As Int64
    End Structure

#End Region

    Public Structure TAG_STRUCT
        Public tag As String
        Public name As String
    End Structure

#Region "Public Data Members"

    Public fileNumber As Integer = 0 'The file identifier for opening the map file.
    Public strFilename As String
    Public fileHeader As New FILE_HEADER_STRUCT
    Public indexHeader As New INDEX_HEADER_STRUCT
    Public indexItem() As INDEX_ITEM_EXPANDED_STRUCT
    Public intMagic As Long
    Public fs As FileStream
    Public cTagTypes As New Collection

#End Region

#Region "Functions"

    '/////////////////////////////////////////////////////
    '// Constructor
    '/////////////////////////////////////////////////////
    Public Sub New()
        Dim t As TAG_STRUCT

        'Add all of the known tags
        t.tag = "bitm" : t.name = "Bitmap" : cTagTypes.Add(t, t.tag)
        t.tag = "snd!" : t.name = "Sound" : cTagTypes.Add(t, t.tag)
        t.tag = "lsnd" : t.name = "Looping Sound" : cTagTypes.Add(t, t.tag)
        t.tag = "weap" : t.name = "Weapon" : cTagTypes.Add(t, t.tag)
        t.tag = "vehi" : t.name = "Vehicle" : cTagTypes.Add(t, t.tag)
        t.tag = "mode" : t.name = "Model" : cTagTypes.Add(t, t.tag)
        t.tag = "proj" : t.name = "Projectile" : cTagTypes.Add(t, t.tag)
        t.tag = "lens" : t.name = "Lens" : cTagTypes.Add(t, t.tag)
        t.tag = "scnr" : t.name = "Scenario" : cTagTypes.Add(t, t.tag)
        t.tag = "ssce" : t.name = "Sound Scenery" : cTagTypes.Add(t, t.tag)
        t.tag = "snde" : t.name = "Sound Environment" : cTagTypes.Add(t, t.tag)
    End Sub

    '/////////////////////////////////////////////////////
    '// Open the Halo .map file that is to be processed
    '/////////////////////////////////////////////////////
    Enum eOpenFile As Byte
        Success = &H0
        InvalidFile = &H1
        FileNotFound = &H2
        SharingViolation = &H3
        UnknwonError = &HFF
    End Enum
    Public Function OpenFile(ByVal strFname As String) As eOpenFile
        Try
            fs = New FileStream(strFname, FileMode.Open, FileAccess.ReadWrite)
            Dim br1 As New BinaryReader(fs) 'Open the file
            If br1.ReadChars(4) = "daeh" Then
                OpenFile = eOpenFile.Success 'The header is valid.
                br1.BaseStream.Seek(0, SeekOrigin.Begin)
                Return eOpenFile.Success
            Else
                Return eOpenFile.InvalidFile
            End If
        Catch
            Return eOpenFile.UnknwonError
        End Try
    End Function

    '/////////////////////////////////////////////////////
    '// Close the Map File
    '/////////////////////////////////////////////////////
    Public Function CloseFile()
        Try
            fs.Close()
            Return 0
        Catch
            Return 1
        End Try
    End Function

    '/////////////////////////////////////////////////////
    ' Read in the map's resource index
    '/////////////////////////////////////////////////////
    Enum eReadFile As Byte
        Success = &H0
        Compressed = &H10
        UnknwonError = &HFF
    End Enum
    Public Function ReadFile() As eReadFile
        Dim br1 As New BinaryReader(fs)
        Try
            fileHeader.ReadStruct(br1) 'Read in the file header.
            If br1.BaseStream.Length < fileHeader.decomp_len Then
                Return eReadFile.Compressed
            End If

            br1.BaseStream.Seek(fileHeader.offset_to_index_decomp, SeekOrigin.Begin)
            indexHeader.ReadStruct(br1)

            intMagic = CInt(Unsigned(indexHeader.index_magic) - (fileHeader.offset_to_index_decomp + 36))

            'Read in all of the Index Items
            ReDim indexItem(indexHeader.tagcount)  'Reserve memory for all of the index items.
            For x As Integer = 1 To indexHeader.tagcount
                indexItem(x).offset = br1.BaseStream.Position  'Get the item's offset
                indexItem(x).fileData.ReadStruct(br1)
                indexItem(x).magic_metadata_offset = Unsigned(indexItem(x).fileData.offset) - intMagic
                indexItem(x - 1).estimatedMetaSize = indexItem(x).fileData.offset _
                    - indexItem(x - 1).fileData.offset  'Estimate the size of the previous metadata structure.
                indexItem(x).magic_string_offset = Unsigned(indexItem(x).fileData.stringoffset) - intMagic
            Next

            'Now that we've got an array of all the index item data, calculate the extended data.
            For x As Integer = 1 To indexHeader.tagcount
                indexItem(x).filePath = ReadString(br1, indexItem(x).magic_string_offset)
                indexItem(x).tagClass = getTagClass(indexItem(x).fileData.tagclass, 3)
                'Temporary hacks for known meta sizes can be put here.
                If indexItem(x).tagClass = "snde" Then indexItem(x).estimatedMetaSize = 32
            Next
            Return eReadFile.Success
        Catch
            Return eReadFile.UnknwonError
        End Try
    End Function

    '/////////////////////////////////////////////////////
    '// Return a specified number of bytes from a
    '// specific offset in the current file
    '/////////////////////////////////////////////////////
    Public Function GetBytes(ByVal offset As Long, ByVal size As Integer, ByRef binChunk() As Byte) As Boolean
        Dim br1 As New BinaryReader(fs)
        ReDim binChunk(size) 'Resize the byte array.

        br1.BaseStream.Seek(offset, SeekOrigin.Begin) 'Read in the data
        binChunk = br1.ReadBytes(size)
    End Function

    '/////////////////////////////////////////////////////
    '// Decipher the tag class from the index structure
    '/////////////////////////////////////////////////////
    Public Function getTagClass(ByVal strTag As String, ByVal intTagNum As Short) As String
        'This function will return one of the three tags from the string.
        'Little ENDIAN is assumed.
        strTag = StrReverse(strTag)
        Select Case intTagNum
            Case 1
                getTagClass = strTag.Substring(0, 4)
            Case 2
                getTagClass = strTag.Substring(4, 4)
            Case 3
                getTagClass = strTag.Substring(8, 4)
        End Select
    End Function

    '/////////////////////////////////////////////////////
    '// This function will read the NULL terminated
    '// string from a specified offset
    '/////////////////////////////////////////////////////
    Public Function ReadString(ByRef br As BinaryReader, ByVal offset As Integer) As String
        Dim btchar As Byte
        Dim strFilePath As String
        Try
            br.BaseStream.Seek(offset, SeekOrigin.Begin)

            Do 'Read one character at a time until we reach NULL character.
                btchar = br.ReadByte
                If Not btchar = 0 Then strFilePath &= Chr(btchar)
            Loop Until btchar = 0
            ReadString = strFilePath
        Catch
        End Try
    End Function

    '/////////////////////////////////////////////////////
    '// This function will complete the name of the
    '// tagclass based on its 4 letter abbreviation
    '/////////////////////////////////////////////////////
    Public Function FullTagClassName(ByVal strString As String) As String
        Select Case strString
            Case "obje" ' - Object
                FullTagClassName = "Object"
            Case "unit" ' - Unit 
                FullTagClassName = "Unit"
            Case "bipd" ' - Biped 
                FullTagClassName = "Biped"
            Case "vehi" ' - Vehicle
                FullTagClassName = "Vehicle"
            Case "item" ' - Item 
                FullTagClassName = "Item"
            Case "weap" ' - Weapon
                FullTagClassName = "Weapon"
            Case "eqip" ' - Equipment 
                FullTagClassName = "Equipment"
            Case "garb" ' - Garbage 
                FullTagClassName = "Garbage"
            Case "proj" ' - Projectile 
                FullTagClassName = "Projectile"
            Case "scen" ' - Scenery 
                FullTagClassName = "Scenery"
            Case "ssce" ' - Sound scenery 
                FullTagClassName = "Sound Scenery"
            Case "devi" ' - Device 
                FullTagClassName = "Device"
            Case "mach" ' - Machine 
                FullTagClassName = "Machine"
            Case "ctrl" ' - Control 
                FullTagClassName = "Control"
            Case "lifi" ' - Light fixture 
                FullTagClassName = "Light Fixture"
            Case "plac" ' - Placeholder 
                FullTagClassName = "Placeholder"
            Case "bitm" ' - Bitmap 
                FullTagClassName = "Bitmap"
            Case "shdr" ' - Shader 
                FullTagClassName = "Shader"
            Case "snd!" ' - Sound 
                FullTagClassName = "Sound"
            Case "ustr" ' - Unicode string list 
                FullTagClassName = "Unicode String list"
            Case "scnr" ' - Scenario
                FullTagClassName = "Scenario"
            Case "sbsp" ' - Structure BSP 
                FullTagClassName = "Structure BSP"
            Case "mply" ' - Multiplayer scenario 
                FullTagClassName = "Multiplayer Scenario"
            Case "itmc" ' - Item collection 
                FullTagClassName = "Item Collection"
            Case "hmt " ' - Hud message text 
                FullTagClassName = "HUD MEssage Text"
            Case "trak" ' - Track 
                FullTagClassName = "Track"
            Case "mode" ' - Model
                FullTagClassName = "Model"

                '{'actr', "Actor"},'
                '{'a'ctv', "Actor Variant"},
                '{'a'ntr', "Animation Trigger"},
                '{'sky ', "Sky"},
                '{'lens', "Lensflare"},
                '{'grhi', "Grenade HUD Interface"},
                '{'unhi', "Unit HUD Interface"},
                '{'wphi', "Weapon HUD Interface"},
                '{'coll', "Collision Model"},
                '{'cont', "Contrail"},
                '{'deca', "Decal"},
                '{'effe', "Effect"},
                '{'part', "Particle"},
                '{'pctl', "Particle System"},
                '{'rain', "Weather Particle"},
                '{'matg', "Game Globals"},
                '{'hud#', "HUD Number"},
                '{'hudg', "HUD Globals"},
                '{'jpt!', "Damage"},
                '{'ligh', "Light"},
                '{'glw!', "Glow"},
                '{'mgs2', "Light Volume"},
                '{'lsnd', "Looping Sound"},
                '{'snde', "Sound Environment"},
                '{'phys', "Physics"},
                '{'pphy', "Point Physics"},
                '{'fog ', "Fog"},
                '{'wind', "Wind"},
                '{'senv', "Shader Environment"},
                '{'dobc', "Detail Object Collection"},
                '{'font', "Font"},
                '{'udlg', "Dialog"},
                '{'DeLa', "UI Item"},
                '{'Soul', "UI Item Collection"},
                '{'soso', "Shader Model"},
                '{'sotr', "Shader Transparency"},
                '{'swat', "Shader Water"},
                '{'sgla', "Shader Glass"},
                '{'smet', "Shader Metal"},
                '{'spla', "Shader Plasma"}
            Case Else
                FullTagClassName = "UNKNOWN"
        End Select
    End Function

    '/////////////////////////////////////////////////////
    '// Return the full map name based on its identifier
    '/////////////////////////////////////////////////////
    Public Function MapName() As String
        Select Case fileHeader.name.Trim(Chr(0))
            Case "a10"
                MapName = "Pillar of Autum"
            Case "a30"
                MapName = "Halo"
            Case "a50"
                MapName = "Truth and Reconciliation"
            Case "b30"
                MapName = "Silent Cartographer"
            Case "b40"
                MapName = "Assault on the Control Room"
            Case "c10"
                MapName = "Guilty Spark"
            Case "c20"
                MapName = "Library"
            Case "c40"
                MapName = "Two Betrayals"
            Case "d20"
                MapName = "Keyes"
            Case "d40"
                MapName = "The Maw"
            Case "beavercreek"
                MapName = "Battle Creek"
            Case "bloodgulch"
                MapName = "Blood Gulch"
            Case "boardingaction"
                MapName = "Boarding Action"
            Case "carousel"
                MapName = "Derelict"
            Case "chillout"
                MapName = "Chill Out"
            Case "damnation"
                MapName = "Damnation"
            Case "hangemhigh"
                MapName = "Hang 'em High"
            Case "longest"
                MapName = "Longest"
            Case "prisoner"
                MapName = "Prisoner"
            Case "putput"
                MapName = "Chiron TL34"
            Case "ratrace"
                MapName = "Rat Race"
            Case "sidewinder"
                MapName = "Sidewinder"
            Case "wizard"
                MapName = "Wizard"
            Case "ui"
                MapName = "User Interface"
            Case Else
                MapName = fileHeader.name
        End Select
    End Function

    '/////////////////////////////////////////////////////
    '// Search through the index for a specific resource
    '// id and return its position in the index array.
    '/////////////////////////////////////////////////////
    Public Function LocateByID(ByVal intID As Integer)
        For x As Integer = 1 To indexHeader.tagcount
            If indexItem(x).fileData.id = intID Then Return x
        Next
        Return -1 'Not found
    End Function

#End Region

End Class