'/////////////////////////////////////////////////////////////
'// Halo Map File
'// [bitm] Metadata Structure
'// Reverse-Engineered by Slayer and others
'/////////////////////////////////////////////////////////////
Imports System.IO

Public Class BitmMeta

#Region "Constants"

#Region "Bitm"

    'BITM
    Enum bitmEnum As Integer
        BITM_FORMAT_A8 = &H0
        BITM_FORMAT_Y8 = &H1
        BITM_FORMAT_AY8 = &H2
        BITM_FORMAT_A8Y8 = &H3
        BITM_FORMAT_R5G6B5 = &H6
        BITM_FORMAT_A1R5G5B5 = &H8
        BITM_FORMAT_A4R4G4B4 = &H9
        BITM_FORMAT_X8R8G8B8 = &HA
        BITM_FORMAT_A8R8G8B8 = &HB
        BITM_FORMAT_DXT1 = &HE
        BITM_FORMAT_DXT2AND3 = &HF
        BITM_FORMAT_DXT4AND5 = &H10
        BITM_FORMAT_P8 = &H11

        BITM_TYPE_2D = &H0
        BITM_TYPE_3D = &H1
        BITM_TYPE_CUBEMAP = &H2

        BITM_FLAG_LINEAR = (1 << 4)
    End Enum

#End Region

#Region "DDS"

    'DDS
    'The dwFlags member of the modified DDSURFACEDESC2 structure can be set to one or more of the following values.
    Enum DDSEnum As Integer
        DDSD_CAPS = &H1
        DDSD_HEIGHT = &H2
        DDSD_WIDTH = &H4
        DDSD_PITCH = &H8
        DDSD_PIXELFORMAT = &H1000
        DDSD_MIPMAPCOUNT = &H20000
        DDSD_LINEARSIZE = &H80000
        DDSD_DEPTH = &H800000

        DDPF_ALPHAPIXELS = &H1
        DDPF_FOURCC = &H4
        DDPF_RGB = &H40

        'The dwCaps1 member of the DDSCAPS2 structure can be set to one or more of the following values.
        DDSCAPS_COMPLEX = &H8
        DDSCAPS_TEXTURE = &H1000
        DDSCAPS_MIPMAP = &H400000

        'The dwCaps2 member of the DDSCAPS2 structure can be set to one or more of the following values.
        DDSCAPS2_CUBEMAP = &H200
        DDSCAPS2_CUBEMAP_POSITIVEX = &H400
        DDSCAPS2_CUBEMAP_NEGATIVEX = &H800
        DDSCAPS2_CUBEMAP_POSITIVEY = &H1000
        DDSCAPS2_CUBEMAP_NEGATIVEY = &H2000
        DDSCAPS2_CUBEMAP_POSITIVEZ = &H4000
        DDSCAPS2_CUBEMAP_NEGATIVEZ = &H8000
        DDSCAPS2_VOLUME = &H200000
    End Enum

#End Region

#End Region

#Region "Structures"

#Region "Bitm Metadata Structures"

    Public Structure BITM_HEADER_STRUCT
        Public DXTn As Short
        Public unknown1() As Integer 'has 21 elements (7) == (6)+108(9)
        Public offset_to_first As Integer
        Public unknown23 As Integer  'always 0x0
        Public imageCount As Integer
        Public offset_to_second As Integer
        Public unknown25 As Integer 'always 0x0
        Public Sub readStruct(ByRef br As BinaryReader)
            DXTn = br.ReadInt16
            DXTn = br.ReadInt16
            Dim x As Integer 'loop counter
            Dim unknown1(21)
            For x = 1 To 21
                unknown1(x) = br.ReadInt32
            Next
            offset_to_first = br.ReadInt32
            unknown23 = br.ReadInt32
            imageCount = br.ReadInt32
            offset_to_second = br.ReadInt32
            unknown25 = br.ReadInt32
        End Sub
    End Structure

    Public Structure BITM_FIRST_STRUCT
        Public unknown() As Integer 'has 16 elements - (8) is ALWAYS 0x10000, the rest are zeros
        Public Sub readstruct(ByRef br As BinaryReader)
            Dim x As Integer 'Loop counter
            ReDim unknown(16)
            For x = 1 To 16
                unknown(x) = br.ReadInt32
            Next
        End Sub
    End Structure

    Public Structure BITM_SECOND_STRUCT
        Public id As Integer ' "bitm"
        Public width As Short
        Public height As Short
        Public depth As Short
        Public type As Short
        Public format As Short
        Public flags As Short
        Public reg_point_x As Short
        Public reg_point_y As Short
        Public num_mipmaps As Short 'either mipmaps, or the rest of the cubemaps, etc
        Public pixel_offset As Short
        Public offset As Integer
        Public size As Integer
        Public unknown1 As Integer
        Public unknown2 As Integer  'always 0xFFFFFFFF?
        Public unknown3 As Integer 'always 0x00000000?
        Public unknown4 As Integer 'always 0x024F0040?

        Public Sub readStruct(ByRef br As BinaryReader)
            id = br.ReadInt32
            width = br.ReadInt16
            height = br.ReadInt16
            depth = br.ReadInt16
            type = br.ReadInt16
            format = br.ReadInt16
            flags = br.ReadInt16
            reg_point_x = br.ReadInt16
            reg_point_y = br.ReadInt16
            num_mipmaps = br.ReadInt16
            pixel_offset = br.ReadInt16
            offset = br.ReadInt32
            size = br.ReadInt32
            unknown1 = br.ReadInt32
            unknown2 = br.ReadInt32
            unknown3 = br.ReadInt32
            unknown4 = br.ReadInt32
        End Sub
    End Structure

    Public Structure BITM_ENCAPSULATOR_STRUCT
        Public header As BITM_HEADER_STRUCT
        Public first As BITM_FIRST_STRUCT
        Public second() As BITM_SECOND_STRUCT
    End Structure

#End Region

#Region "DDS Structures"

    Public Structure DDS_HEADER_STRUCTURE
        Public magic As String '"DDS "
        Public ddsd As DDSURFACEDESC2
        Public Sub readStruct(ByRef br As BinaryReader)
            magic = br.ReadChars(4)
            ddsd.readStruct(br)
        End Sub
        Public Sub writeStruct(ByRef bw As BinaryWriter)
            bw.BaseStream.WriteByte(Asc(Mid(magic, 1, 1)))
            bw.BaseStream.WriteByte(Asc(Mid(magic, 2, 1)))
            bw.BaseStream.WriteByte(Asc(Mid(magic, 3, 1)))
            bw.BaseStream.WriteByte(Asc(Mid(magic, 4, 1)))
            ddsd.writeStruct(bw)
        End Sub
        '//////////////////////////////////////////////////////
        '// Generate a DDS Header
        '//////////////////////////////////////////////////////
        Public Sub generate(ByRef b2 As BITM_SECOND_STRUCT)
            magic = "DDS "
            ddsd.generate(b2)
        End Sub
    End Structure

    Public Structure DDSURFACEDESC2
        Public size_of_structure As Integer  '124
        Public flags As Integer
        Public height As Integer
        Public width As Integer
        Public PitchOrLinearSize As Integer
        Public depth As Integer 'Only used for volume textures.
        Public MipMapCount As Integer 'Total Number of MipMaps
        Public Reserved1() As Integer '11 ints long
        Public ddfPixelFormat As DDPIXELFORMAT
        Public ddsCaps As DDSCAPS2
        Public Reserved2 As Integer
        Public Sub readStruct(ByRef br As BinaryReader)
            size_of_structure = br.ReadInt32
            flags = br.ReadInt32
            height = br.ReadInt32
            width = br.ReadInt32
            PitchOrLinearSize = br.ReadInt32
            depth = br.ReadInt32
            MipMapCount = br.ReadInt32
            Dim x As Integer 'Loop counter
            ReDim Reserved1(11)
            For x = 1 To 11
                Reserved1(x) = br.ReadInt32
            Next
            ddfPixelFormat.readStruct(br)
            ddsCaps.readStruct(br)
            Reserved2 = br.ReadInt32
        End Sub
        Public Sub writeStruct(ByRef bw As BinaryWriter)
            bw.Write(size_of_structure)
            bw.Write(flags)
            bw.Write(height)
            bw.Write(width)
            bw.Write(PitchOrLinearSize)
            bw.Write(depth)
            bw.Write(MipMapCount)
            Dim x As Integer 'Loop counter
            For x = 1 To 11
                bw.Write(Reserved1(x))
            Next
            ddfPixelFormat.writeStruct(bw)
            ddsCaps.writeStruct(bw)
            bw.Write(Reserved2)
        End Sub
        '//////////////////////////////////////////////////////
        '// Generate the DDSURFACEDESC2 chunk of the header
        '//////////////////////////////////////////////////////
        Public Sub generate(ByVal b2 As BITM_SECOND_STRUCT)
            'Size of structure. This member must be set to 124.
            size_of_structure = 124

            'Flags to indicate valid fields. Always include DDSD_CAPS, DDSD_PIXELFORMAT,
            'DDSD_WIDTH, DDSD_HEIGHT and either DDSD_PITCH or DDSD_LINEARSIZE.
            flags = flags + DDSEnum.DDSD_CAPS
            flags = flags + DDSEnum.DDSD_PIXELFORMAT
            flags = flags + DDSEnum.DDSD_WIDTH
            flags = flags + DDSEnum.DDSD_HEIGHT
            Select Case b2.format
                Case bitmEnum.BITM_FORMAT_DXT1, bitmEnum.BITM_FORMAT_DXT2AND3, bitmEnum.BITM_FORMAT_DXT4AND5
                    flags = flags + DDSEnum.DDSD_LINEARSIZE
                Case Else
                    flags = flags + DDSEnum.DDSD_PITCH
            End Select

            'Height of the main image in pixels
            height = b2.height

            'Width of the main image in pixels
            width = b2.width

            'For uncompressed formats, this is the number of bytes per scan line (DWORD aligned) for the main
            'image. dwFlags should include DDSD_PITCH in this case. For compressed formats, this is the (total)
            'number of bytes for the main image. dwFlags should be include DDSD_LINEARSIZE in this case.
            If (b2.flags And DDSEnum.DDSD_LINEARSIZE) Then
                PitchOrLinearSize = b2.width * 2 'Assumes 16 bits.
            Else
                PitchOrLinearSize = b2.size  'I don't think this is correct.
            End If

            'For volume textures, this is the depth of the volume.
            'dwFlags should include DDSD_DEPTH in this case.
            depth = 0 'There are no volume textures (that I know of) in Halo.

            'For items with mipmap levels, this is the total number of levels in the mipmap
            'chain of the main image. dwFlags should include DDSD_MIPMAPCOUNT in this case
            MipMapCount = b2.num_mipmaps
            If MipMapCount > 0 Then flags = flags + DDSEnum.DDSD_MIPMAPCOUNT


            'Reserved values - should always equal 0.
            ReDim Reserved1(11)
            Reserved1(1) = 0 : Reserved1(2) = 0 : Reserved1(3) = 0 : Reserved1(4) = 0
            Reserved1(5) = 0 : Reserved1(6) = 0 : Reserved1(7) = 0 : Reserved1(8) = 0
            Reserved1(9) = 0 : Reserved1(10) = 0 : Reserved1(11) = 0

            'A 32-byte value that specifies the pixel format structure.
            ddfPixelFormat.generate(b2)

            'A 16-byte value that specifies the capabilities structure.
            ddsCaps.generate(b2)
        End Sub
    End Structure

    Public Structure DDPIXELFORMAT
        Public size As Integer   '32
        Public Flags As Integer  '4
        Public FourCC As String 'DXT1, DXT2, etc..
        Public RGBBitCount As Integer
        Public RBitMask As Integer
        Public GBitMask As Integer
        Public BBitMask As Integer
        Public RGBAlphaBitMask As Integer
        Public Sub readStruct(ByRef br As BinaryReader)
            size = br.ReadInt32 '32
            Flags = br.ReadInt32 '4
            FourCC = br.ReadChars(4) 'DXT1, DXT2, etc..
            RGBBitCount = br.ReadInt32
            RBitMask = br.ReadInt32
            GBitMask = br.ReadInt32
            BBitMask = br.ReadInt32
            RGBAlphaBitMask = br.ReadInt32
        End Sub
        Public Sub writeStruct(ByRef bw As BinaryWriter)
            bw.Write(size)
            bw.Write(Flags)
            bw.BaseStream.WriteByte(Asc(Mid(FourCC, 1, 1)))
            bw.BaseStream.WriteByte(Asc(Mid(FourCC, 2, 1)))
            bw.BaseStream.WriteByte(Asc(Mid(FourCC, 3, 1)))
            bw.BaseStream.WriteByte(Asc(Mid(FourCC, 4, 1)))
            bw.Write(RGBBitCount)
            bw.Write(RBitMask)
            bw.Write(GBitMask)
            bw.Write(BBitMask)
            bw.Write(RGBAlphaBitMask)
        End Sub
        '//////////////////////////////////////////////////////
        '// Generate the DDPIXELFORMAT chunk of the header
        '//////////////////////////////////////////////////////
        Public Sub generate(ByVal b2 As BITM_SECOND_STRUCT)
            'Size of structure. This member must be set to 32. 
            size = 32

            'Flags to indicate valid fields. Uncompressed formats will usually use DDPF_RGB to indicate
            'an RGB format, while compressed formats will use DDPF_FOURCC with a four-character code. 
            'This is accomplished in the below structure.

            'This is the four-character code for compressed formats. dwFlags should include DDPF_FOURCC in
            'this case. For DXTn compression, this is set to "DXT1", "DXT2", "DXT3", "DXT4", or "DXT5". 
            Flags = 0
            Select Case b2.format
                Case &HE
                    FourCC = "DXT1"
                    Flags = Flags + DDSEnum.DDPF_FOURCC
                Case &HF
                    FourCC = "DXT3"
                    Flags = Flags + DDSEnum.DDPF_FOURCC
                Case &H10
                    FourCC = "DXT4"
                    Flags = Flags + DDSEnum.DDPF_FOURCC
                Case Else
                    FourCC = Chr(0) & Chr(0) & Chr(0) & Chr(0)
                    Flags = Flags + DDSEnum.DDPF_RGB
            End Select

            'For RGB formats, this is the total number of bits in the format. dwFlags should include DDPF_RGB
            'in this case. This value is usually 16, 24, or 32. For A8R8G8B8, this value would be 32. 
            Select Case b2.format
                Case bitmEnum.BITM_FORMAT_R5G6B5
                    RGBBitCount = 16
                Case bitmEnum.BITM_FORMAT_A1R5G5B5
                    RGBBitCount = 16
                Case bitmEnum.BITM_FORMAT_A4R4G4B4
                    RGBBitCount = 16
                Case bitmEnum.BITM_FORMAT_X8R8G8B8
                    RGBBitCount = 32
                Case bitmEnum.BITM_FORMAT_A8R8G8B8
                    RGBBitCount = 32
                Case Else
                    RGBBitCount = 0
            End Select

            'For RGB formats, this contains the masks for the red, green, and blue channels. For A8R8G8B8, these
            'values would be 0x00ff0000, 0x0000ff00, and 0x000000ff respectively. 
            Select Case b2.format
                Case bitmEnum.BITM_FORMAT_R5G6B5
                    RBitMask = &HF800
                    GBitMask = &H7E0
                    BBitMask = &H1F
                    RGBAlphaBitMask = &H401008
                Case bitmEnum.BITM_FORMAT_A1R5G5B5
                    RBitMask = &H7C00
                    GBitMask = &H3E0
                    BBitMask = &H1F
                    RGBAlphaBitMask = &H401008
                Case bitmEnum.BITM_FORMAT_A4R4G4B4
                    RBitMask = &HF000
                    GBitMask = &HF0
                    BBitMask = &HF
                    RGBAlphaBitMask = &H401008
                Case bitmEnum.BITM_FORMAT_X8R8G8B8
                    RBitMask = &HFF000
                    GBitMask = &HFF00
                    BBitMask = &HFF
                    RGBAlphaBitMask = &H401008
                Case bitmEnum.BITM_FORMAT_A8R8G8B8
                    RBitMask = &HFF0000
                    GBitMask = &HFF00
                    BBitMask = &HFF
                    RGBAlphaBitMask = &H401008
                Case Else
                    RBitMask = 0
                    GBitMask = 0
                    BBitMask = 0
                    RGBAlphaBitMask = &H0
            End Select

            'For RGB formats, this contains the mask for the alpha channel, if any. dwFlags should include 
            'DDPF_ALPHAPIXELS in this case. For A8R8G8B8, this value would be 0xff000000. 
            '^^ This was set in the previous block on code.
        End Sub
    End Structure

    Public Structure DDSCAPS2
        Public caps1 As Integer
        Public caps2 As Integer
        Public Reserved() As Integer 'Length of 3 (or maybe 2)
        Public Sub readStruct(ByRef br As BinaryReader)
            caps1 = br.ReadInt32
            caps2 = br.ReadInt32
            ReDim Reserved(2)
            Reserved(1) = br.ReadInt32
            Reserved(2) = br.ReadInt32
        End Sub
        Public Sub writeStruct(ByRef bw As BinaryWriter)
            bw.Write(caps1)
            bw.Write(caps2)
            bw.Write(Reserved(1))
            bw.Write(Reserved(2))
        End Sub
        '//////////////////////////////////////////////////////
        '// Generate the DDSCAPS2 chunk of the header
        '//////////////////////////////////////////////////////
        Public Sub generate(ByRef b2 As BITM_SECOND_STRUCT)
            'DDS files should always include DDSCAPS_TEXTURE. If the file contains mipmaps, DDSCAPS_MIPMAP
            'should be set. For any DDS file with more than one main surface,such as a mipmaps, cubic
            'environment map, or volume texture, DDSCAPS_COMPLEX should also be set. 
            caps1 = caps1 + DDSEnum.DDSCAPS_TEXTURE
            If b2.num_mipmaps > 0 Then caps1 = caps1 + DDSEnum.DDSCAPS_MIPMAP
            caps1 = caps1 + DDSEnum.DDSCAPS_COMPLEX
            'caps1 = &H401008 'DDSCAPS_TEXTURE, DDSCAPS_MIPMAP, DDSCAPS_COMPLEX flags.

            'For cubic environment maps, DDSCAPS2_CUBEMAP should be included as well as one or more faces of
            'the map (DDSCAPS2_CUBEMAP_POSITIVEX, DDSCAPS2_CUBEMAP_NEGATIVEX, DDSCAPS2_CUBEMAP_POSITIVEY,
            ' DDSCAPS2_CUBEMAP_NEGATIVEY, DDSCAPS2_CUBEMAP_POSITIVEZ, DDSCAPS2_CUBEMAP_NEGATIVEZ). For volume
            'textures, DDSCAPS2_VOLUME should be included. 
            '******************************************************************************
            'This is where I stopped - this code needs to be added for cubemaps.
            '******************************************************************************
            caps2 = 0

            'Reserved - all should be set to 0.
            ReDim Reserved(2) : Reserved(1) = 0 : Reserved(2) = 0
        End Sub
    End Structure

#End Region

    '// Memory storage for sound chunks
    Public Structure BITM_BUFFER_STRUCT
        Public bin() As Byte
    End Structure

#End Region

#Region "Public Data Members"

    Public bitm As BITM_ENCAPSULATOR_STRUCT
    Public binBuffer() As BITM_BUFFER_STRUCT

#End Region

#Region "Functions"

    '//////////////////////////////////////////////////////
    '// Decode the specified bitm metadata
    '//////////////////////////////////////////////////////
    Enum eDecodeBitmMetadata As Byte
        Success = &H0
        Failure = &H1
    End Enum
    Public Function DecodeBitmMetadata(ByRef map As HaloMap, ByVal i As Integer) As eDecodeBitmMetadata
        Dim br1 As New BinaryReader(map.fs)
        Dim bChunk() As Byte
        Try
            'Goto the metadata offset.
            br1.BaseStream.Seek(Unsigned(map.indexItem(i).magic_metadata_offset), SeekOrigin.Begin)
            bitm.header.readStruct(br1)
            br1.BaseStream.Seek(Unsigned(bitm.header.offset_to_first) - map.intMagic, SeekOrigin.Begin)
            bitm.first.readstruct(br1)
            'This is the array of structures that hold image specific data.
            ReDim bitm.second(bitm.header.imageCount)
            br1.BaseStream.Seek(Unsigned(bitm.header.offset_to_second) - map.intMagic, SeekOrigin.Begin)
            For y As Integer = 1 To bitm.header.imageCount
                bitm.second(y).readStruct(br1)
            Next
            Return eDecodeBitmMetadata.Success
        Catch
            Return eDecodeBitmMetadata.Failure
        End Try
    End Function

    '//////////////////////////////////////////////////////
    '// Class Constructor
    '//////////////////////////////////////////////////////
    Public Sub New(ByRef map As HaloMap, ByVal i As Integer)
        'Deflash the specified texture set to memory.
        DecodeBitmMetadata(map, i)
        ReDim binBuffer(bitm.header.imageCount)
        For x As Integer = 1 To bitm.header.imageCount
            map.GetBytes(bitm.second(x).offset, bitm.second(x).size, binBuffer(x).bin)
        Next x
    End Sub

    '//////////////////////////////////////////////////////
    '// Extract the specified texture as a DDS
    '//////////////////////////////////////////////////////
    Public Function ExtractDDS(ByRef map As HaloMap, ByVal c As Integer, ByVal strFilename As String)

        Dim bw As BinaryWriter
        Dim bw2 As BinaryWriter
        Dim dds As New DDS_HEADER_STRUCTURE
        Dim bChunk() As Byte
        Dim swiz As Swizzle

        Dim tmp, tmp1, tmp2 As Integer
        For z As Integer = c To c 'bitm(i).header.imageCount
            ReDim bChunk(binBuffer(z).bin.Length)
            swiz = New Swizzle(bitm.second(z).width, bitm.second(z).height, bitm.second(z).depth)
            dds.generate(bitm.second(z))
            Try
                tmp = (bitm.second(z).flags)
                If 1 = 1 Then 'Not (bitm.second(z).flags And bitmEnum.BITM_FLAG_LINEAR) Then
                    bw = New BinaryWriter(New FileStream(strFilename & "." & z & ".dds", FileMode.Create))
                    dds.writeStruct(bw)
                    bw.Write(binBuffer(z).bin)
                    bw.Close()
                Else
                    bw2 = New BinaryWriter(New FileStream(strFilename & ".swizzle." & z & ".dds", FileMode.Create))
                    For y As Integer = 0 To bitm.second(z).height - 1
                        For x As Integer = 0 To bitm.second(z).width - 1
                            tmp1 = ((y * bitm.second(z).width) + x) * 2
                            tmp2 = (swiz.Swizzle(x, y, bitm.second(z).depth)) * 2
                            bChunk(tmp1) = binBuffer(z).bin(tmp2)
                            bChunk(tmp1 + 1) = binBuffer(z).bin(tmp2 + 1)
                        Next
                    Next
                    dds.writeStruct(bw2)
                    bw2.Write(bChunk)
                    bw2.Close()
                End If
            Catch
            End Try
        Next z
    End Function

    '//////////////////////////////////////////////////////
    '// Inject a texture as the specified offset
    '//////////////////////////////////////////////////////
    Public Function InjectDDS(ByRef map As HaloMap, ByVal strFilename As String, _
        ByVal c As Integer, Optional ByVal offset As Integer = 0)
        Dim br As BinaryReader
        Dim bw As BinaryWriter
        Dim dds As New DDS_HEADER_STRUCTURE
        Dim bChunk() As Byte

        If offset = 0 Then offset = bitm.second(c).offset

        'Read in the ADPCM header and ensure that it is valid.
        br = New BinaryReader(New FileStream(strFilename, FileMode.Open, FileAccess.Read)) 'Open the file
        dds.readStruct(br)

        If (Not dds.magic = "DDS ") Then
            InjectDDS = "Invalid Header: " & vbCrLf & _
                "You must choose a valid DDS format texture."
            Exit Function
        End If

        'Check the structure to see if it matches the current sound.
        Dim bCompatible As Boolean = True
        Dim strErrorMsg As String
        If Not dds.ddsd.ddfPixelFormat.FourCC = "DXT" & bitm.header.DXTn + 1 Then
            bCompatible = False
            strErrorMsg &= "DXT format mismatch." & vbCrLf
        End If
        If (Not dds.ddsd.height = bitm.second(c).height) Or (Not dds.ddsd.width = bitm.second(c).width) Then
            bCompatible = False
            strErrorMsg &= "Dimensions are not the same." & vbCrLf
        End If

        If bCompatible Or Not bCompatible Then
            Try
                'Apply the datawriter
                bw = New BinaryWriter(map.fs)

                'Grab a chunk of the file.
                ReDim bChunk(bitm.second(c).size)
                bChunk = br.ReadBytes(bitm.second(c).size)

                'Inject the chunk into the appropriate map offset
                bw.Seek(offset, SeekOrigin.Begin)
                bw.Write(bChunk)

                'Replace the appropriate chunk buffer
                ReDim binBuffer(bitm.second(c).size)
                binBuffer(1).bin = bChunk
                strErrorMsg = "Texture was successfully imported."
            Catch ex As Exception
                strErrorMsg = "An error occured while importing data: " & vbCrLf & _
                        ex.Message
            End Try
        End If
        InjectDDS = strErrorMsg
        br.Close()
    End Function

    Public Function ImageType(ByVal i As Integer) As String
        Select Case i
            Case &H0
                ImageType = "A8"
            Case &H1
                ImageType = "Y8"
            Case &H2
                ImageType = "AY8"
            Case &H3
                ImageType = "A8Y8"
            Case &H6
                ImageType = "R5G6B5"
            Case &H8
                ImageType = "A1R5G5B5"
            Case &H9
                ImageType = "A4R4G4B4"
            Case &HA
                ImageType = "X8R8G8B8"
            Case &HB
                ImageType = "A8R8G8B8"
            Case &HE
                ImageType = "DXT1"
            Case &HF
                ImageType = "DXT2/DXT3"
            Case &H10
                ImageType = "DXT4/DXT5"
            Case &H11
                ImageType = "P8"
            Case Else
                ImageType = "Unknown"
        End Select
    End Function

#End Region

End Class
