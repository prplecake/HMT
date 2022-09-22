Imports System.IO
Imports HMTLib
Imports System.Drawing
Imports System.Runtime.InteropServices

Public Class BitmGUI
    Inherits System.Windows.Forms.UserControl

    Public map As HaloMap.Map
    Public bitm As BitmLib.BitmMeta
    Public b As Bitmap
    Public ptr As New System.IntPtr
    Public fTex As New frmTextureFullsize
    <MarshalAs(UnmanagedType.ByValArray)> Public decodedChunk() As Byte

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef m As HaloMap.Map, ByVal i As Integer)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        SelectBitmap(m, i)
    End Sub

    'UserControl1 overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
        Try
            Marshal.FreeHGlobal(ptr)
        Catch
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents tvImages As System.Windows.Forms.TreeView
    Friend WithEvents lblImageInfo As System.Windows.Forms.Label
    Friend WithEvents btnInjectTexture As System.Windows.Forms.Button
    Friend WithEvents lblBitmInfo As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnSaveTexture As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog As System.Windows.Forms.SaveFileDialog
    Friend WithEvents pb As System.Windows.Forms.PictureBox
    Friend WithEvents lbl As System.Windows.Forms.Label
    Friend WithEvents txtOffset As System.Windows.Forms.TextBox
    Friend WithEvents btnRevert As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnEOF As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.tvImages = New System.Windows.Forms.TreeView
        Me.lblImageInfo = New System.Windows.Forms.Label
        Me.btnInjectTexture = New System.Windows.Forms.Button
        Me.lblBitmInfo = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnSaveTexture = New System.Windows.Forms.Button
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog = New System.Windows.Forms.SaveFileDialog
        Me.pb = New System.Windows.Forms.PictureBox
        Me.lbl = New System.Windows.Forms.Label
        Me.txtOffset = New System.Windows.Forms.TextBox
        Me.btnRevert = New System.Windows.Forms.Button
        Me.btnUpdate = New System.Windows.Forms.Button
        Me.btnEOF = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tvImages
        '
        Me.tvImages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tvImages.ImageIndex = -1
        Me.tvImages.Location = New System.Drawing.Point(8, 109)
        Me.tvImages.Name = "tvImages"
        Me.tvImages.SelectedImageIndex = -1
        Me.tvImages.Size = New System.Drawing.Size(296, 107)
        Me.tvImages.TabIndex = 34
        '
        'lblImageInfo
        '
        Me.lblImageInfo.BackColor = System.Drawing.Color.White
        Me.lblImageInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblImageInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblImageInfo.Location = New System.Drawing.Point(8, 224)
        Me.lblImageInfo.Name = "lblImageInfo"
        Me.lblImageInfo.Size = New System.Drawing.Size(296, 24)
        Me.lblImageInfo.TabIndex = 33
        Me.lblImageInfo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnInjectTexture
        '
        Me.btnInjectTexture.BackColor = System.Drawing.Color.Black
        Me.btnInjectTexture.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnInjectTexture.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnInjectTexture.Location = New System.Drawing.Point(8, 440)
        Me.btnInjectTexture.Name = "btnInjectTexture"
        Me.btnInjectTexture.Size = New System.Drawing.Size(72, 24)
        Me.btnInjectTexture.TabIndex = 32
        Me.btnInjectTexture.Text = "Inject Texture"
        '
        'lblBitmInfo
        '
        Me.lblBitmInfo.BackColor = System.Drawing.Color.White
        Me.lblBitmInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblBitmInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBitmInfo.Location = New System.Drawing.Point(8, 24)
        Me.lblBitmInfo.Name = "lblBitmInfo"
        Me.lblBitmInfo.Size = New System.Drawing.Size(296, 48)
        Me.lblBitmInfo.TabIndex = 29
        Me.lblBitmInfo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Firebrick
        Me.Label6.Location = New System.Drawing.Point(8, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(120, 16)
        Me.Label6.TabIndex = 30
        Me.Label6.Text = "Texture Information"
        '
        'btnSaveTexture
        '
        Me.btnSaveTexture.BackColor = System.Drawing.Color.Black
        Me.btnSaveTexture.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveTexture.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnSaveTexture.Location = New System.Drawing.Point(8, 256)
        Me.btnSaveTexture.Name = "btnSaveTexture"
        Me.btnSaveTexture.Size = New System.Drawing.Size(72, 24)
        Me.btnSaveTexture.TabIndex = 31
        Me.btnSaveTexture.Text = "Save Texture"
        '
        'pb
        '
        Me.pb.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pb.BackColor = System.Drawing.Color.Black
        Me.pb.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pb.Location = New System.Drawing.Point(0, 0)
        Me.pb.Name = "pb"
        Me.pb.Size = New System.Drawing.Size(208, 208)
        Me.pb.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pb.TabIndex = 37
        Me.pb.TabStop = False
        '
        'lbl
        '
        Me.lbl.BackColor = System.Drawing.Color.White
        Me.lbl.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl.Location = New System.Drawing.Point(128, 84)
        Me.lbl.Name = "lbl"
        Me.lbl.Size = New System.Drawing.Size(32, 9)
        Me.lbl.TabIndex = 38
        Me.lbl.Text = "Offset:"
        Me.lbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtOffset
        '
        Me.txtOffset.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtOffset.Location = New System.Drawing.Point(168, 80)
        Me.txtOffset.Name = "txtOffset"
        Me.txtOffset.Size = New System.Drawing.Size(96, 20)
        Me.txtOffset.TabIndex = 39
        Me.txtOffset.Text = ""
        '
        'btnRevert
        '
        Me.btnRevert.BackColor = System.Drawing.Color.Black
        Me.btnRevert.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRevert.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnRevert.Location = New System.Drawing.Point(64, 80)
        Me.btnRevert.Name = "btnRevert"
        Me.btnRevert.Size = New System.Drawing.Size(48, 20)
        Me.btnRevert.TabIndex = 40
        Me.btnRevert.Text = "Revert"
        '
        'btnUpdate
        '
        Me.btnUpdate.BackColor = System.Drawing.Color.Black
        Me.btnUpdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdate.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnUpdate.Location = New System.Drawing.Point(8, 80)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(48, 20)
        Me.btnUpdate.TabIndex = 41
        Me.btnUpdate.Text = "Update"
        '
        'btnEOF
        '
        Me.btnEOF.BackColor = System.Drawing.Color.Black
        Me.btnEOF.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEOF.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnEOF.Location = New System.Drawing.Point(272, 80)
        Me.btnEOF.Name = "btnEOF"
        Me.btnEOF.Size = New System.Drawing.Size(32, 20)
        Me.btnEOF.TabIndex = 42
        Me.btnEOF.Text = "EOF"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.pb)
        Me.Panel1.Location = New System.Drawing.Point(96, 256)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(208, 208)
        Me.Panel1.TabIndex = 43
        '
        'BitmGUI
        '
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.btnEOF)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.btnRevert)
        Me.Controls.Add(Me.txtOffset)
        Me.Controls.Add(Me.lbl)
        Me.Controls.Add(Me.tvImages)
        Me.Controls.Add(Me.lblImageInfo)
        Me.Controls.Add(Me.btnInjectTexture)
        Me.Controls.Add(Me.lblBitmInfo)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.btnSaveTexture)
        Me.Name = "BitmGUI"
        Me.Size = New System.Drawing.Size(320, 480)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Functions"

    Public Sub Unload()
        Marshal.FreeHGlobal(ptr)
        bitm.ReleaseFile()
    End Sub

    '///////////////////////////////////////////////////////////////////
    '// Change the current sound
    '///////////////////////////////////////////////////////////////////
    Public Sub SelectBitmap(ByRef m As HaloMap.Map, ByVal intIndex As Integer)

        map = m

        Dim s As String = map.indexItem(intIndex).filePath

        'If we already have a bitmap open, we need to destroy it so that the bitmaps file is no longer in use
        If Not bitm Is Nothing Then
            bitm.ReleaseFile()
        End If

        If map.pc Then
            Try
                Dim strFilename As String = map.fs.Name
                strFilename = Mid(strFilename, 1, Len(strFilename) - Len(getLevel(strFilename, 0))) & "bitmaps.map"
                bitm = New BitmMeta(map, intIndex, strFilename)
            Catch
                MsgBox("Error accessing bitmaps.map - make sure that it exists in" & vbCrLf & _
                    "the same folder as the map you are trying to open.")
                Exit Sub
            End Try
        Else
            bitm = New BitmMeta(map, intIndex)
        End If


        'Build the string to be displayed in the sound info box.
        lblBitmInfo.Text = "Resource Name: " & getLevel(s, 0) & vbCrLf & vbCrLf & _
            "(" & bitm.bitm.header.imageCount & " images in the set)"

        FillInImageList()
        tvImages.SelectedNode = tvImages.Nodes.Item(0)
    End Sub
    '///////////////////////////////////////////////////////////////////
    '// Update the tracklist
    '///////////////////////////////////////////////////////////////////
    Private Sub FillInImageList()
        tvImages.Nodes.Clear()
        For x As Integer = 1 To bitm.bitm.header.imageCount
            tvImages.Nodes.Add(x & "- " & bitm.bitm.second(x).width & "x" & bitm.bitm.second(x).height)
        Next
    End Sub
    Private Sub bitm_changeImage(ByVal i As Integer)
        'Update the Sub image Information
        Dim bSupported As Boolean = False

        If bitm.bitm.second(i).format = bitm.bitmEnum.BITM_FORMAT_DXT1 Or _
            bitm.bitm.second(i).format = bitm.bitmEnum.BITM_FORMAT_DXT2AND3 Or _
            bitm.bitm.second(i).format = bitm.bitmEnum.BITM_FORMAT_DXT4AND5 Or _
            bitm.bitm.second(i).format = bitm.bitmEnum.BITM_FORMAT_A1R5G5B5 Or _
            bitm.bitm.second(i).format = bitm.bitmEnum.BITM_FORMAT_A4R4G4B4 Or _
            bitm.bitm.second(i).format = bitm.bitmEnum.BITM_FORMAT_A8R8G8B8 Or _
            bitm.bitm.second(i).format = bitm.bitmEnum.BITM_FORMAT_R5G6B5 Or _
            bitm.bitm.second(i).format = bitm.bitmEnum.BITM_FORMAT_X8R8G8B8 Then
            bSupported = True
        Else
            bSupported = False
        End If

        If bSupported Then
            btnInjectTexture.Enabled = True
        Else
            btnInjectTexture.Enabled = False
        End If

        lblImageInfo.Text = _
            IIf(bSupported, "", "Unsupported pixel format - ") & bitm.ImageType(bitm.bitm.second(i).format) & "( " & _
            bitm.bitm.second(i).size & " bytes )"
        txtOffset.Text = Hex(bitm.bitm.second(i).offset)
        'Initialize a new bitmap object based on the decoded data
        Dim l As New ImageLib
        Dim ImageSize As Integer

        ImageSize = CLng(bitm.bitm.second(i).width) * CLng(bitm.bitm.second(i).height) * 4
        ReDim decodedChunk(ImageSize)
        Dim swiz As Swizzle
        Dim width As Integer = bitm.bitm.second(i).width
        Dim height As Integer = bitm.bitm.second(i).height
        Dim bShowPreview As Boolean = True

        Try
            Select Case bitm.bitm.second(i).format
                Case bitm.bitmEnum.BITM_FORMAT_DXT1
                    decodedChunk = l.DecodeDXT1(height, width, bitm.binBuffer(i).bin)
                Case bitm.bitmEnum.BITM_FORMAT_DXT2AND3
                    decodedChunk = l.DecodeDXT23(height, width, bitm.binBuffer(i).bin)
                Case bitm.bitmEnum.BITM_FORMAT_DXT4AND5
                    decodedChunk = l.DecodeDXT45(height, width, bitm.binBuffer(i).bin)
                Case bitm.bitmEnum.BITM_FORMAT_A8R8G8B8, bitm.bitmEnum.BITM_FORMAT_X8R8G8B8, bitm.bitmEnum.BITM_FORMAT_R5G6B5
                    If Not map.pc Then
                        decodedChunk = bitm.Swizzle(bitm.binBuffer(i).bin, width, height, -1, 32, BitmMeta.SwizzleType.DeSwizzle)
                    End If
                    'Case bitm.bitmEnum.BITM_FORMAT_A1R5G5B5
                    '    If Not map.pc Then
                    '    decodedChunk = bitm.Swizzle(bitm.binBuffer(i).bin, width, height, -1, 16, BitmMeta.SwizzleType.DeSwizzle)
                    '    End If
                Case Else
                    ReDim decodedChunk(64 * 64 * 4)
                    ImageSize = (64 * 64 * 4)
                    height = 64
                    width = 64
                    'decodedChunk.Clear(decodedChunk, 0, ImageSize)
                    bShowPreview = False
            End Select
        Catch
            'MsgBox("Error decoding image")
        End Try

        'Reserve a blobk of unmanaged memory for the byte array
        Marshal.FreeHGlobal(ptr)
        ptr = Marshal.AllocHGlobal(ImageSize)

        'Copy the decoded byte array to unmanaged memory
        Dim api As New WINAPI
        api.CopyMemory(ptr, decodedChunk, ImageSize)

        'Create the bitmap object
        b = New System.Drawing.Bitmap(width, height, width * 4, Imaging.PixelFormat.Format32bppArgb, ptr)
        If Not bShowPreview Then
            b = l.Overlay(width, height, b, "No Preview Available", _
                New System.Drawing.Font("Verdana", 6), Color.White, False, False)
        End If
        pb.Image = b 'The picture box displays the bitmap
    End Sub
#End Region

#Region "Event Handlers"
    Private Sub tvImages_SelectedIndexChanged(ByVal sender As System.Object, _
        ByVal e As TreeViewEventArgs) Handles tvImages.AfterSelect
        bitm_changeImage(e.Node.Index + 1)
    End Sub
    Private Sub btnSaveTexture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveTexture.Click
        Dim strFilename As String

        SaveFileDialog.AddExtension = True
        SaveFileDialog.DefaultExt = "*.dds"
        SaveFileDialog.Filter = "DDS Texture (*.dds)|*.dds"
        SaveFileDialog.FileName = getLevel(bitm.fileName, 0) & _
                tvImages.SelectedNode.Index + 1 & ".dds"
        Try
            If SaveFileDialog.ShowDialog() = DialogResult.Cancel Then Exit Sub
        Catch
            SaveFileDialog.FileName = "IllegalFilename" & _
                tvImages.SelectedNode.Index + 1 & ".dds"
            If SaveFileDialog.ShowDialog() = DialogResult.Cancel Then Exit Sub
        End Try
        strFilename = SaveFileDialog.FileName

        bitm.ExtractDDS(map, tvImages.SelectedNode.Index + 1, strFilename)
    End Sub
    Private Sub btnInjectTexture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInjectTexture.Click
        Dim strFilename As String
        Dim fi As FileInfo
        Dim i As Integer = tvImages.SelectedNode.Index + 1

        OpenFileDialog.AddExtension = True
        OpenFileDialog.DefaultExt = "*.dds"
        OpenFileDialog.Filter = "DDS Texture (*.dds)|*.dds"
        OpenFileDialog.FileName = ""
        If OpenFileDialog.ShowDialog() = DialogResult.Cancel Then Exit Sub

        'Open the source file.
        strFilename = OpenFileDialog.FileName
        fi = New FileInfo(strFilename)
        Dim bCheckSize As Boolean = True
        If fi.Length < bitm.bitm.second(i).size + 128 Then
            bCheckSize = False
            Select Case MsgBox("The texture you are trying to import is smaller than the existing" & vbCrLf & _
                            "data.  HMT will automatically update the size speficied in the map file" & vbCrLf & _
                            "to match that of the image being injected.", MsgBoxStyle.OKCancel)
                Case MsgBoxResult.Cancel
                    Exit Sub
                Case MsgBoxResult.OK
                    Try
                        bitm.bitm.second(tvImages.SelectedNode.Index + 1).size = fi.Length
                        Dim bw As New BinaryWriter(map.fs)
                        bw.BaseStream.Seek(Unsigned(bitm.bitm.header.offset_to_second) - map.intMagic + (48 * tvImages.SelectedNode.Index), SeekOrigin.Begin)
                        bitm.bitm.second(tvImages.SelectedNode.Index + 1).writeStruct(bw)
                        bw.Flush()
                    Catch
                        MsgBox("There was an error updating the structure.", MsgBoxStyle.Critical)
                    End Try
            End Select
        End If
        If (bCheckSize) And (fi.Length > bitm.bitm.second(i).size + 128) Then
            Select Case MsgBox("The texture you are trying to import is larger than the existing" & vbCrLf & _
                            "data!  Would you like HMT to update the stored size to match the size" & vbCrLf & _
                            "of the image being injected?  Choosing no will truncate the file so" & vbCrLf & _
                            "that it conforms to the size specified in the map file." & vbCrLf & _
                            "Note: Injecting larger files without truncating can result in map" & vbCrLf & _
                            "corruption, and it is possible that the file may not work after" & vbCrLf & _
                            "truncating.  It is recommended that you inject files of the same size" & vbCrLf & _
                            "unless you have changed their offset to an position in the file that will" & vbCrLf & _
                            "not be affected." & vbCrLf & vbCrLf & _
                            "Update size?", MsgBoxStyle.YesNoCancel)
                Case MsgBoxResult.Cancel
                    Exit Sub
                Case MsgBoxResult.Yes
                    Try
                        bitm.bitm.second(tvImages.SelectedNode.Index + 1).size = fi.Length
                        Dim bw As New BinaryWriter(map.fs)
                        bw.BaseStream.Seek(Unsigned(bitm.bitm.header.offset_to_second) - map.intMagic + (48 * tvImages.SelectedNode.Index), SeekOrigin.Begin)
                        bitm.bitm.second(tvImages.SelectedNode.Index + 1).writeStruct(bw)
                        bw.Flush()
                    Catch
                        MsgBox("There was an error updating the structure.", MsgBoxStyle.Critical)
                    End Try
            End Select
            bCheckSize = False
        End If
        MsgBox(bitm.InjectDDS(map, strFilename, tvImages.SelectedNode.Index + 1))
        bitm_changeImage(tvImages.SelectedNode.Index + 1)
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Try
            bitm.bitm.second(tvImages.SelectedNode.Index + 1).offset = CInt("&H" & txtOffset.Text)
            Dim bw As New BinaryWriter(map.fs)
            bw.BaseStream.Seek(Unsigned(bitm.bitm.header.offset_to_second) - map.intMagic + (48 * tvImages.SelectedNode.Index), SeekOrigin.Begin)
            bitm.bitm.second(tvImages.SelectedNode.Index + 1).writeStruct(bw)
            bw.Flush()
            MsgBox("Update was successful.", MsgBoxStyle.Information)
        Catch
            MsgBox("There was an error updating the structure.", MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnEOF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEOF.Click
        If map.pc Then
            txtOffset.Text = Hex(bitm.fs.Length)
        Else
            txtOffset.Text = Hex(map.fileHeader.decomp_len)
        End If
    End Sub

    Private Sub btnRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRevert.Click
        txtOffset.Text = Hex(bitm.bitm.second(tvImages.SelectedNode.Index + 1).offset)
    End Sub

    Private Sub pb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pb.Click
        Try
            fTex.Close()
        Catch
        End Try
        fTex = New frmTextureFullsize
        If b.Width < pb.Width And b.Height < pb.Height Then
            MsgBox("Full sized texture is smaller than the preview window")
            Exit Sub
        End If
        fTex.pb.Width = b.Width
        fTex.pb.Height = b.Height

        'Size the window according to the image size
        If b.Width > 512 Or b.Height > 512 Then
            fTex.Height = b.Height
            fTex.Width = b.Width
        Else
            fTex.Height = 512
            fTex.Width = 512
        End If

        fTex.pb.Location = New System.Drawing.Point((fTex.Width - fTex.pb.Width) / 2, (fTex.Height - fTex.pb.Height) / 2)
        fTex.pb.Refresh()
        fTex.pb.ResumeLayout()
        fTex.pb.Image = b
        fTex.Show()
    End Sub

#End Region

End Class

'/////////////////////////////////////////////////////////////
'// Halo Map File
'// (bitm) Metadata Structure
'/////////////////////////////////////////////////////////////

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
        Public unknown1() As Integer 'has 22 elements (7) == (6)+108(9)
        Public offset_to_first As Integer
        Public unknown23 As Integer  'always 0x0
        Public imageCount As Integer
        Public offset_to_second As Integer
        Public unknown25 As Integer 'always 0x0
        Public Sub readStruct(ByRef br As BinaryReader)
            Dim x As Integer 'loop counter
            Dim unknown1(22)
            Console.WriteLine(br.BaseStream.Position)
            For x = 1 To 22
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

    Public Structure BITM_SECOND_STRUCT  '48 bytes
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
        Public Sub writeStruct(ByRef bw As BinaryWriter)
            bw.Write(id)
            bw.Write(width)
            bw.Write(height)
            bw.Write(depth)
            bw.Write(type)
            bw.Write(format)
            bw.Write(flags)
            bw.Write(reg_point_x)
            bw.Write(reg_point_y)
            bw.Write(num_mipmaps)
            bw.Write(pixel_offset)
            bw.Write(offset)
            bw.Write(size)
            bw.Write(unknown1)
            bw.Write(unknown2)
            bw.Write(unknown3)
            bw.Write(unknown4)
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
            Dim RGBBitCount As Integer = 0
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
            End Select
            If (b2.flags And DDSEnum.DDSD_PITCH) Then _
                PitchOrLinearSize = b2.width * (RGBBitCount / 8)
            If (b2.flags And DDSEnum.DDSD_LINEARSIZE) Then _
                PitchOrLinearSize = b2.size 'I don't think this is correct.

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
                    Flags = Flags + DDSEnum.DDPF_ALPHAPIXELS
                Case bitmEnum.BITM_FORMAT_A4R4G4B4
                    RGBBitCount = 16
                    Flags = Flags + DDSEnum.DDPF_ALPHAPIXELS
                Case bitmEnum.BITM_FORMAT_X8R8G8B8
                    Flags = Flags + DDSEnum.DDPF_ALPHAPIXELS
                    RGBBitCount = 32
                Case bitmEnum.BITM_FORMAT_A8R8G8B8
                    Flags = Flags + DDSEnum.DDPF_ALPHAPIXELS
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
                    RGBAlphaBitMask = &H0
                Case bitmEnum.BITM_FORMAT_A1R5G5B5
                    RBitMask = &H7C00
                    GBitMask = &H3E0
                    BBitMask = &H1F
                    RGBAlphaBitMask = &H8000
                Case bitmEnum.BITM_FORMAT_A4R4G4B4
                    RBitMask = &HF000
                    GBitMask = &HF0
                    BBitMask = &HF
                    RGBAlphaBitMask = &HF00000
                Case bitmEnum.BITM_FORMAT_X8R8G8B8
                    RBitMask = &HFF000
                    GBitMask = &HFF00
                    BBitMask = &HFF
                    RGBAlphaBitMask = &HFF000000
                Case bitmEnum.BITM_FORMAT_A8R8G8B8
                    RBitMask = &HFF0000
                    GBitMask = &HFF00
                    BBitMask = &HFF
                    RGBAlphaBitMask = &HFF000000
                Case Else
                    RBitMask = 0
                    GBitMask = 0
                    BBitMask = 0
                    RGBAlphaBitMask = &H0
            End Select

            'For RGB formats, this contains the mask for the alpha channel, if any. dwFlags should include 
            'DDPF_ALPHAPIXELS in this case. For A8R8G8B8, this value would be 0xff000000. 
            '^^ This was set in the previous block of code.
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

    '// Memory storage for textures in the set
    Public Structure BITM_BUFFER_STRUCT
        Public bin() As Byte
    End Structure

#End Region

#Region "Public Data Members"

    Public bitm As BITM_ENCAPSULATOR_STRUCT
    Public binBuffer() As BITM_BUFFER_STRUCT
    Public fs As FileStream
    Public fileName As String

#End Region

#Region "Functions"

    '//////////////////////////////////////////////////////
    '// Decode the specified bitm metadata
    '//////////////////////////////////////////////////////
    Enum eDecodeBitmMetadata As Byte
        Success = &H0
        Failure = &H1
    End Enum
    Public Function DecodeBitmMetadata(ByRef map As HaloMap.Map, ByVal i As Integer) As eDecodeBitmMetadata
        Dim br1 As New BinaryReader(map.fs)
        Dim bChunk() As Byte
        Try
            'Goto the metadata offset.
            fileName = map.indexItem(i).filePath
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
    Public Sub New(ByRef map As HaloMap.Map, ByVal i As Integer, Optional ByVal strFilename As String = "")
        'Deflash the specified texture set to memory.
        Dim br As BinaryReader

        DecodeBitmMetadata(map, i)

        ReDim binBuffer(bitm.header.imageCount)

        If map.pc Then
            'PC version - need to get images from the external file.
            'TODO: add some exception handling here.
            fs = New FileStream(strFilename, FileMode.Open, FileAccess.ReadWrite)

            br = New BinaryReader(fs)
            For x As Integer = 1 To bitm.header.imageCount
                ReDim binBuffer(x).bin(bitm.second(x).size)
                br.BaseStream.Seek(bitm.second(x).offset, SeekOrigin.Begin)
                binBuffer(x).bin = br.ReadBytes(bitm.second(x).size)
            Next x
        Else
            'Xbox version - files are stored in the map file.
            For x As Integer = 1 To bitm.header.imageCount
                map.GetBytes(bitm.second(x).offset, bitm.second(x).size, binBuffer(x).bin)
            Next x
        End If
    End Sub

    Public Sub ReleaseFile()
        Try
            If Not fs Is Nothing Then fs.Close()
        Catch
        End Try
    End Sub

    '//////////////////////////////////////////////////////
    '// Extract the specified texture as a DDS
    '//////////////////////////////////////////////////////
    Public Function ExtractDDS(ByRef map As HaloMap.Map, ByVal c As Integer, ByVal strFilename As String)

        Dim bw As BinaryWriter
        Dim bw2 As BinaryWriter
        Dim dds As New DDS_HEADER_STRUCTURE
        Dim bChunk() As Byte
        Dim swiz As Swizzle

        Dim tmp, tmp1, tmp2 As Integer
        Dim width, height As Integer
        Try
            For z As Integer = c To c 'bitm(i).header.imageCount
                width = bitm.second(z).width
                height = bitm.second(z).height
                ReDim bChunk(binBuffer(z).bin.Length)
                bChunk = binBuffer(z).bin

                If Not map.pc Then
                    If bitm.second(z).format = bitmEnum.BITM_FORMAT_A8R8G8B8 Or _
                        bitm.second(z).format = bitmEnum.BITM_FORMAT_X8R8G8B8 Then _
                        bChunk = Swizzle(binBuffer(z).bin, width, height, -1, 32, BitmMeta.SwizzleType.DeSwizzle)
                    If bitm.second(z).format = bitmEnum.BITM_FORMAT_A1R5G5B5 Or _
                        bitm.second(z).format = bitmEnum.BITM_FORMAT_R5G6B5 Or _
                        bitm.second(z).format = bitmEnum.BITM_FORMAT_A4R4G4B4 Then _
                        bChunk = Swizzle(binBuffer(z).bin, width, height, -1, 16, BitmMeta.SwizzleType.DeSwizzle)
                End If

                dds.generate(bitm.second(z))
                Try
                    tmp = (bitm.second(z).flags)
                    bw = New BinaryWriter(New FileStream(strFilename & "." & z & ".dds", FileMode.Create))
                    dds.writeStruct(bw)
                    bw.Write(bChunk)
                    bw.Close()
                Catch
                End Try
            Next z
        Catch
            MsgBox("Error extracting texture.")
        End Try

    End Function

    '//////////////////////////////////////////////////////
    '// Inject a texture as the specified offset
    '//////////////////////////////////////////////////////
    Public Function InjectDDS(ByRef map As HaloMap.Map, ByVal strFilename As String, _
        ByVal c As Integer, Optional ByVal offset As Integer = 0)
        Dim br As New BinaryReader(New FileStream(strFilename, FileMode.Open, FileAccess.Read)) 'Open the file
        Dim bw As BinaryWriter
        Dim dds As New DDS_HEADER_STRUCTURE
        Dim bChunk() As Byte

        If offset = 0 Then offset = bitm.second(c).offset

        If map.pc Then
            bw = New BinaryWriter(fs)
        Else
            bw = New BinaryWriter(map.fs)
        End If

        dds.readStruct(br)

        If (Not dds.magic = "DDS ") Then
            InjectDDS = "Invalid Header: " & vbCrLf & _
                "You must choose a valid DDS format texture."
            Exit Function
        End If

        'Check the structure to see if it matches the current sound.
        Dim bCompatible As Boolean = True
        Dim strErrorMsg As String

        'Make sure images are the same pixel format
        Select Case dds.ddsd.ddfPixelFormat.FourCC
            Case "DXT1"
                If Not bitm.second(c).format = bitmEnum.BITM_FORMAT_DXT1 Then
                    bCompatible = False
                    strErrorMsg &= "Image being injected must be DXT1." & vbCrLf
                End If
            Case "DXT2", "DXT3"
                If (Not bitm.second(c).format = bitmEnum.BITM_FORMAT_DXT2AND3) Then
                    bCompatible = False
                    strErrorMsg &= "Image being injected must be DXT2 or DXT3." & vbCrLf
                End If
            Case "DXT4", "DXT5"
                If (Not bitm.second(c).format = bitmEnum.BITM_FORMAT_DXT4AND5) Then
                    bCompatible = False
                    strErrorMsg &= "Image being injected must be DXT4 or DXT5." & vbCrLf
                End If
            Case Else
        End Select

        Dim width As Integer = bitm.second(c).width
        Dim height As Integer = bitm.second(c).height
        If bCompatible Then
            'Update the bitmap structure to be the same dimensions as the new file.
            bitm.second(c).height = dds.ddsd.height
            bitm.second(c).width = dds.ddsd.width
            Dim bw2 As New BinaryWriter(map.fs)
            bw2.BaseStream.Seek(Unsigned(bitm.header.offset_to_second) - map.intMagic + (48 * (c - 1)), SeekOrigin.Begin)
            bitm.second(c).writeStruct(bw2)
            bw2.Flush()
            Try
                'Grab a chunk of the file.
                ReDim bChunk(bitm.second(c).size)
                bChunk = br.ReadBytes(bitm.second(c).size)

                'Swizzle the chunk if neccessary
                If Not map.pc Then
                    If bitm.second(c).format = bitmEnum.BITM_FORMAT_A8R8G8B8 Or _
                        bitm.second(c).format = bitmEnum.BITM_FORMAT_X8R8G8B8 Then _
                        bChunk = Swizzle(binBuffer(c).bin, width, height, -1, 32, BitmMeta.SwizzleType.Swizzle)
                    If bitm.second(c).format = bitmEnum.BITM_FORMAT_A1R5G5B5 Or _
                        bitm.second(c).format = bitmEnum.BITM_FORMAT_R5G6B5 Or _
                        bitm.second(c).format = bitmEnum.BITM_FORMAT_A4R4G4B4 Then _
                        bChunk = Swizzle(binBuffer(c).bin, width, height, -1, 16, BitmMeta.SwizzleType.Swizzle)
                End If

                'Inject the chunk into the appropriate map offset
                bw.Seek(offset, SeekOrigin.Begin)
                bw.Write(bChunk)

                'Replace the appropriate chunk buffer
                ReDim binBuffer(bitm.second(c).size)
                binBuffer(c).bin = bChunk
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

    '/////////////////////////////////////////
    '// Swizzle a set of Pixels
    '// Donated by Stephen Cakebread
    '/////////////////////////////////////////
    Public Enum SwizzleType As Byte
        Swizzle = &H0
        DeSwizzle = &H1
    End Enum
    Public Function Swizzle(ByVal bin() As Byte, ByVal width As Integer, ByVal height As Integer, _
    ByVal depth As Integer, ByVal BitCount As Integer, ByVal Action As SwizzleType) As Byte()
        Dim swiz As New Swizzle(width, height, depth)
        Dim tmp1, tmp2 As Integer
        Dim bChunk() As Byte

        ReDim bChunk(bin.Length - 1)
        For y As Integer = 0 To height - 1
            For x As Integer = 0 To width - 1
                Select Case Action
                    Case SwizzleType.DeSwizzle
                        tmp1 = ((y * width) + x) * (BitCount / 8)
                        tmp2 = (swiz.Swizzle(x, y, -1)) * (BitCount / 8)
                    Case SwizzleType.Swizzle
                        tmp2 = ((y * width) + x) * (BitCount / 8)
                        tmp1 = (swiz.Swizzle(x, y, -1)) * (BitCount / 8)
                End Select
                For i As Integer = 0 To (BitCount / 8) - 1
                    bChunk(tmp1 + i) = bin(tmp2 + i)
                Next
            Next
        Next
        Return bChunk
    End Function

#End Region

End Class

Public Class Swizzle

    Public m_MaskX As Integer = 0
    Public m_MaskY As Integer = 0
    Public m_MaskZ As Integer = 0
    Public Sub New(ByVal width As Integer, ByVal height As Integer, ByVal depth As Integer)
        Dim bit As Integer = 1
        Dim idx As Integer = 1

        While (bit < width) Or (bit < height) Or (bit < depth)

            If (bit < width) Then
                m_MaskX = m_MaskX Or idx
                idx <<= 1
            End If

            If (bit < height) Then
                m_MaskY = m_MaskY Or idx
                idx <<= 1
            End If

            If (bit < depth) Then
                m_MaskZ = m_MaskZ Or idx
                idx <<= 1
            End If

            bit <<= 1
        End While
    End Sub

    Public Function Swizzle(ByVal Sx As Integer, ByVal Sy As Integer, Optional ByVal Sz As Integer = -1) As Int64
        Swizzle = SwizzleAxis(Sx, m_MaskX) Or SwizzleAxis(Sy, m_MaskY) Or (IIf(Sz <> -1, SwizzleAxis(Sz, m_MaskZ), 0))
    End Function

    Public Function SwizzleAxis(ByVal Value As Integer, ByVal Mask As Integer) As Int64

        Dim Result As Int64
        Dim bit As Integer = 1

        While bit <= Mask
            If (Mask And bit) Then
                Result = Result Or (Value And bit)
            Else
                Value <<= 1
            End If
            bit <<= 1
        End While
        SwizzleAxis = Result
    End Function

End Class

Class ImageLib
    Public Structure RGBA_COLOR_STRUCT
        Public r, g, b, a As Integer
    End Structure
    Public Function short_to_rgba(ByVal color As Integer) As RGBA_COLOR_STRUCT
        Dim rc As RGBA_COLOR_STRUCT
        color = Unsigned(color)
        rc.r = (((color >> 11) And &H1F) * &HFF) / 31
        rc.g = (((color >> 5) And &H3F) * &HFF) / 63
        rc.b = (((color >> 0) And &H1F) * &HFF) / 31
        rc.a = 255
        Return rc
    End Function

    Public Function rgba_to_int(ByVal c As RGBA_COLOR_STRUCT) As Integer
        Return (c.a << 24) Or (c.r << 16) Or (c.g << 8) Or c.b
    End Function

    Public Function GradientColors(ByVal Col1 As RGBA_COLOR_STRUCT, _
        ByVal Col2 As RGBA_COLOR_STRUCT) As RGBA_COLOR_STRUCT
        Dim ret As RGBA_COLOR_STRUCT
        ret.r = ((Col1.r * 2 + Col2.r)) / 3
        ret.g = ((Col1.g * 2 + Col2.g)) / 3
        ret.b = ((Col1.b * 2 + Col2.b)) / 3
        ret.a = 255
        Return ret
    End Function

    Public Function GradientColorsHalf(ByVal Col1 As RGBA_COLOR_STRUCT, _
        ByVal Col2 As RGBA_COLOR_STRUCT) As RGBA_COLOR_STRUCT
        Dim ret As RGBA_COLOR_STRUCT
        ret.r = (Col1.r / 2 + Col2.r / 2)
        ret.g = (Col1.g / 2 + Col2.g / 2)
        ret.b = (Col1.b / 2 + Col2.b / 2)
        ret.a = 255
        Return ret
    End Function

    '//////////////////////////
    '// DecodeDXT1 
    '//////////////////////////
    Public Function DecodeDXT1(ByVal height As Integer, ByVal width As Integer, _
        ByVal SourceData() As Byte) As Byte()

        Dim DestData() As Byte
        Dim Color(4) As RGBA_COLOR_STRUCT
        Dim i As Integer
        Dim dptr As Integer
        Dim CColor As RGBA_COLOR_STRUCT
        Dim CData As Integer
        Dim ChunksPerHLine As Integer = width / 4
        Dim trans As Boolean
        Dim zeroColor As RGBA_COLOR_STRUCT
        Dim c1, c2 As Integer

        ReDim DestData(((width * height) * 4) - 1)

        If ChunksPerHLine = 0 Then ChunksPerHLine += 1

        For i = 0 To (width * height) - 1 Step 16

            c1 = (CInt(SourceData(dptr + 1)) << 8) Or (SourceData(dptr))
            c2 = (CInt(SourceData(dptr + 3)) << 8) Or (SourceData(dptr + 2))

            If c1 > c2 Then
                trans = False
            Else
                trans = True
            End If

            Color(0) = short_to_rgba(c1)
            Color(1) = short_to_rgba(c2)
            If Not trans Then
                Color(2) = GradientColors(Color(0), Color(1))
                Color(3) = GradientColors(Color(1), Color(0))
            Else
                Color(2) = GradientColorsHalf(Color(0), Color(1))
                Color(3) = zeroColor
            End If

            CData = (CInt(SourceData(dptr + 4)) << 0) Or _
                (CInt(SourceData(dptr + 5)) << 8) Or _
                (CInt(SourceData(dptr + 6)) << 16) Or _
                (CInt(SourceData(dptr + 7)) << 24)

            Dim ChunkNum As Integer = i / 16
            Dim XPos As Long = ChunkNum Mod ChunksPerHLine
            Dim YPos As Long = (ChunkNum - XPos) / ChunksPerHLine
            Dim ttmp As Long
            Dim ttmp2 As Long

            Dim sizeh As Integer = IIf(height < 4, height, 4)
            Dim sizew As Integer = IIf(width < 4, width, 4)
            Dim tStr As String
            For x As Integer = 0 To sizeh - 1
                For y As Integer = 0 To sizew - 1
                    CColor = Color(CData And &H3)
                    CData >>= 2
                    ttmp = ((YPos * 4 + x) * width + XPos * 4 + y) * 4
                    ttmp2 = rgba_to_int(CColor)
                    DestData(ttmp) = CColor.b
                    DestData(ttmp + 1) = CColor.g
                    DestData(ttmp + 2) = CColor.r
                    DestData(ttmp + 3) = CColor.a
                Next
            Next
            dptr += 8
        Next
        Return DestData
    End Function
    '//////////////////////////
    '// DecodeDXT2/3
    '//////////////////////////
    Public Function DecodeDXT23(ByVal height As Integer, ByVal width As Integer, _
        ByVal SourceData() As Byte) As Byte()

        Dim DestData() As Byte
        Dim Color(4) As RGBA_COLOR_STRUCT
        Dim i As Integer
        Dim CColor As RGBA_COLOR_STRUCT
        Dim CData As Integer
        Dim ChunksPerHLine As Integer = width / 4
        Dim trans As Boolean
        Dim zeroColor As RGBA_COLOR_STRUCT
        Dim c1, c2 As Integer

        ReDim DestData(((width * height) * 4) - 1)

        If ChunksPerHLine = 0 Then ChunksPerHLine += 1

        For i = 0 To (width * height) - 1 Step 16

            Color(0) = short_to_rgba(CInt(SourceData(i + 8)) Or CInt(SourceData(i + 9)) << 8)
            Color(1) = short_to_rgba(CInt(SourceData(i + 10)) Or CInt(SourceData(i + 11)) << 8)
            Color(2) = GradientColors(Color(0), Color(1))
            Color(3) = GradientColors(Color(1), Color(0))

            CData = (CInt(SourceData(i + 12)) << 0) Or _
                (CInt(SourceData(i + 13)) << 8) Or _
                (CInt(SourceData(i + 14)) << 16) Or _
                (CInt(SourceData(i + 15)) << 24)

            Dim ChunkNum As Integer = i / 16
            Dim XPos As Long = ChunkNum Mod ChunksPerHLine
            Dim YPos As Long = (ChunkNum - XPos) / ChunksPerHLine
            Dim ttmp As Long
            Dim ttmp2 As Long
            Dim alpha As Integer

            Dim sizeh As Integer = IIf(height < 4, height, 4)
            Dim sizew As Integer = IIf(width < 4, width, 4)
            Dim tStr As String
            For x As Integer = 0 To sizeh - 1
                alpha = SourceData(i + (2 * x)) Or CInt(SourceData(i + (2 * x) + 1)) << 8
                For y As Integer = 0 To sizew - 1
                    CColor = Color(CData And &H3)
                    CData >>= 2
                    CColor.a = (alpha And &HF) * 16
                    alpha >>= 4
                    ttmp = ((YPos * 4 + x) * width + XPos * 4 + y) * 4

                    DestData(ttmp) = CColor.b
                    DestData(ttmp + 1) = CColor.g
                    DestData(ttmp + 2) = CColor.r
                    DestData(ttmp + 3) = CColor.a
                Next
            Next
        Next
        Return DestData
    End Function
    '//////////////////////////
    '// DecodeDXT4/5
    '//////////////////////////
    Public Function DecodeDXT45(ByVal height As Integer, ByVal width As Integer, _
        ByVal SourceData() As Byte) As Byte()

        Dim DestData() As Byte
        Dim Color(4) As RGBA_COLOR_STRUCT
        Dim i As Integer
        Dim CColor As RGBA_COLOR_STRUCT
        Dim CData As Integer
        Dim ChunksPerHLine As Integer = width / 4
        Dim trans As Boolean
        Dim zeroColor As RGBA_COLOR_STRUCT
        Dim c1, c2 As Integer

        ReDim DestData(((width * height) * 4) - 1)

        If ChunksPerHLine = 0 Then ChunksPerHLine += 1

        For i = 0 To (width * height) - 1 Step 16

            Color(0) = short_to_rgba(CInt(SourceData(i + 8)) Or CInt(SourceData(i + 9)) << 8)
            Color(1) = short_to_rgba(CInt(SourceData(i + 10)) Or CInt(SourceData(i + 11)) << 8)
            Color(2) = GradientColors(Color(0), Color(1))
            Color(3) = GradientColors(Color(1), Color(0))

            CData = (CInt(SourceData(i + 12)) << 0) Or _
                (CInt(SourceData(i + 13)) << 8) Or _
                (CInt(SourceData(i + 14)) << 16) Or _
                (CInt(SourceData(i + 15)) << 24)

            Dim Alpha(8) As Byte

            Alpha(0) = SourceData(i)
            Alpha(1) = SourceData(i + 1)

            'Do the alphas
            If (Alpha(0) > Alpha(1)) Then
                '// 8-alpha block:  derive the other six alphas.
                '// Bit code 000 = alpha_0, 001 = alpha_1, others are interpolated.
                Alpha(2) = (6 * Alpha(0) + 1 * Alpha(1) + 3) / 7 '// bit code 010
                Alpha(3) = (5 * Alpha(0) + 2 * Alpha(1) + 3) / 7 '// bit code 011
                Alpha(4) = (4 * Alpha(0) + 3 * Alpha(1) + 3) / 7 '// bit code 100
                Alpha(5) = (3 * Alpha(0) + 4 * Alpha(1) + 3) / 7 '// bit code 101
                Alpha(6) = (2 * Alpha(0) + 5 * Alpha(1) + 3) / 7 '// bit code 110
                Alpha(7) = (1 * Alpha(0) + 6 * Alpha(1) + 3) / 7 '// bit code 111
            Else
                '// 6-alpha block.
                '// Bit code 000 = alpha_0, 001 = alpha_1, others are interpolated.
                Alpha(2) = (4 * Alpha(0) + 1 * Alpha(1) + 2) / 5 '// Bit code 010
                Alpha(3) = (3 * Alpha(0) + 2 * Alpha(1) + 2) / 5 '// Bit code 011
                Alpha(4) = (2 * Alpha(0) + 3 * Alpha(1) + 2) / 5 '// Bit code 100
                Alpha(5) = (1 * Alpha(0) + 4 * Alpha(1) + 2) / 5 '// Bit code 101
                Alpha(6) = 0            '// Bit code 110
                Alpha(7) = 255          '// Bit code 111
            End If

            '// Byte	Alpha
            '// 0	Alpha_0
            '// 1	Alpha_1 
            '// 2	(0)(2) (2 LSBs), (0)(1), (0)(0)
            '// 3	(1)(1) (1 LSB), (1)(0), (0)(3), (0)(2) (1 MSB)
            '// 4	(1)(3), (1)(2), (1)(1) (2 MSBs)
            '// 5	(2)(2) (2 LSBs), (2)(1), (2)(0)
            '// 6	(3)(1) (1 LSB), (3)(0), (2)(3), (2)(2) (1 MSB)
            '// 7	(3)(3), (3)(2), (3)(1) (2 MSBs)
            '// (0

            '// Read an int and a short
            Dim tmpdword As Long
            Dim tmpword As Integer
            Dim alphaDat As Long

            tmpword = SourceData(i + 2) Or (CInt(SourceData(i + 3)) << 8)
            tmpdword = SourceData(i + 4) Or (CInt(SourceData(i + 5)) << 8) Or (SourceData(i + 6) << 16) Or (CInt(SourceData(i + 7)) << 24)

            alphaDat = tmpword Or (tmpdword << 16)

            Dim ChunkNum As Integer = i / 16
            Dim XPos As Long = ChunkNum Mod ChunksPerHLine
            Dim YPos As Long = (ChunkNum - XPos) / ChunksPerHLine
            Dim ttmp As Long
            Dim ttmp2 As Long

            Dim sizeh As Integer = IIf(height < 4, height, 4)
            Dim sizew As Integer = IIf(width < 4, width, 4)
            Dim tStr As String
            For x As Integer = 0 To sizeh - 1
                For y As Integer = 0 To sizew - 1
                    CColor = Color(CData And &H3)
                    CData >>= 2
                    CColor.a = Alpha(alphaDat And &H7)
                    alphaDat >>= 3
                    ttmp = ((YPos * 4 + x) * width + XPos * 4 + y) * 4

                    DestData(ttmp) = CColor.b
                    DestData(ttmp + 1) = CColor.g
                    DestData(ttmp + 2) = CColor.r
                    DestData(ttmp + 3) = CColor.a
                Next
            Next
        Next
        Return DestData
    End Function
    Public Function Overlay(ByVal width As Integer, ByVal height As Integer, ByVal img As Image, ByVal OverlayText As String, ByVal OverlayFont As Font, _
        ByVal OverlayColor As Color, ByVal AddAlpha As Boolean, ByVal AddShadow As Boolean) As Bitmap
        If OverlayText > "" Then
            ' create bitmap and graphics used for drawing
            Dim bmp As New Bitmap(img)
            Dim g As Graphics = Graphics.FromImage(bmp)

            Dim alpha As Integer = 255
            If AddAlpha Then
                ' Compute transparency: Longer text should be less transparent or it gets lost.
                alpha = 90 + (OverlayText.Length * 2)
                If alpha >= 255 Then alpha = 255
            End If
            ' Create the brush based on the color and alpha
            Dim b As New SolidBrush(Color.FromArgb(alpha, OverlayColor))

            ' Measure the text to render (unscaled, unwrapped)
            Dim s As SizeF = g.MeasureString(OverlayText, OverlayFont)

            ' Enlarge font to ~80% fill (estimated by AREA)
            Dim zoom As Single = CSng(Math.Sqrt((CDbl(img.Width * img.Height) * 0.8) / CDbl(s.Width * s.Height)))
            Dim sty As FontStyle = OverlayFont.Style
            Dim f As New Font(OverlayFont.FontFamily, CSng(OverlayFont.Size) * zoom, sty)
            Console.WriteLine("Starting Zoom: " & zoom & ", Font Size: " & f.Size & ", Alpha: " & alpha)

            ' Measure using new font size, allow to wrap as needed.
            ' Could rotate the overlay at a 30-45 deg. angle (trig would give correct angle).
            ' Of course, then the area covered would be less than "straight" text.
            ' I'll leave those calculations for someone else....
            Dim strFormat As New StringFormat
            Dim charFit, linesFit As Integer
            strFormat.FormatFlags = StringFormatFlags.NoClip 'Or StringFormatFlags.LineLimit 'Or StringFormatFlags.MeasureTrailingSpaces
            strFormat.Trimming = StringTrimming.Word
            Dim layout As New SizeF(CSng(width) * 0.9!, CSng(height) * 1.5!) ' fit to width, allow height to go over
            Console.WriteLine("Target size: " & layout.Width & ", " & layout.Height)
            s = g.MeasureString(OverlayText, f, layout, strFormat, charFit, linesFit)

            ' Reduce size until it actually fits...
            ' Most text only has to be reduced 1 or 2 times.
            Do Until s.Height <= CSng(height) * 0.9! AndAlso s.Width <= layout.Width
                Console.WriteLine("Reducing font size: area required = " & s.Width & ", " & s.Height)
                zoom = Math.Max(s.Height / (CSng(height) * 0.9!), s.Width / layout.Width)
                zoom = f.Size / zoom
                If zoom > 16 Then zoom = CSng(Math.Floor(zoom)) ' use a whole number to reduce "jaggies"
                If zoom >= f.Size Then zoom -= 1
                f = New Font(OverlayFont.FontFamily, zoom, sty)
                s = g.MeasureString(OverlayText, f, layout, strFormat, charFit, linesFit)
                If zoom <= 1 Then Exit Do ' bail
            Loop
            Console.WriteLine("Final Font Size: " & f.Size & ", area: " & s.Width & ", " & s.Height)

            ' Determine draw area (centered)
            Dim rect As New RectangleF((bmp.Width - s.Width) / 2, (bmp.Height - s.Height) / 2, layout.Width, CSng(width) * 0.9!)

            If AddShadow Then
                ' Add "drop shadow" at half transparency and offset by 1/10 font size
                Dim shadow As New SolidBrush(Color.FromArgb(CInt(alpha / 2), OverlayColor))
                Dim sRect As New RectangleF(rect.X - (f.Size * 0.1!), rect.Y - (f.Size * 0.1!), rect.Width, rect.Height)
                g.DrawString(OverlayText, f, shadow, sRect, strFormat)
            End If

            ' Finally, draw centered text!
            g.DrawString(OverlayText, f, b, rect, strFormat)

            ' clean-up
            g.Dispose()
            b.Dispose()
            f.Dispose()
            Return bmp
        Else
            ' nothing to overlay!
            Return New Bitmap(img)
        End If
    End Function
End Class

