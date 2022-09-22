Imports System.IO
Imports System.Xml
Imports HMTLib

Public Class MapGUI
    Inherits System.Windows.Forms.UserControl

    Public WithEvents map As HaloMap.Map
    Public CurrentID As Integer

    Public Event MetaUpdated(ByVal intID As Integer)

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef m As HaloMap.Map, ByVal intID As Integer)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        map = m
        CurrentID = intID
        PopulateControls(intID)
    End Sub

    'UserControl1 overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents txtFilename As System.Windows.Forms.TextBox
    Friend WithEvents btnSaveMeta As System.Windows.Forms.Button
    Friend WithEvents btnInjectMeta As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnRevert As System.Windows.Forms.Button
    Friend WithEvents cTag1 As System.Windows.Forms.ComboBox
    Friend WithEvents cTag3 As System.Windows.Forms.ComboBox
    Friend WithEvents cTag2 As System.Windows.Forms.ComboBox
    Friend WithEvents lblMetaSize As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog As System.Windows.Forms.SaveFileDialog
    Friend WithEvents txtMetaOffset As System.Windows.Forms.TextBox
    Friend WithEvents btnShowSwap As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents listOffsets As System.Windows.Forms.ListBox
    Friend WithEvents pSwap As System.Windows.Forms.Panel
    Friend WithEvents btnSaveTrack As System.Windows.Forms.Button
    Friend WithEvents pExtract As System.Windows.Forms.Panel
    Friend WithEvents pb As System.Windows.Forms.ProgressBar
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents RecursiveMeta As System.Windows.Forms.CheckBox
    Friend WithEvents FolderBrowserDialog As System.Windows.Forms.FolderBrowserDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblID = New System.Windows.Forms.Label
        Me.txtFilename = New System.Windows.Forms.TextBox
        Me.btnSaveMeta = New System.Windows.Forms.Button
        Me.btnInjectMeta = New System.Windows.Forms.Button
        Me.cTag1 = New System.Windows.Forms.ComboBox
        Me.cTag3 = New System.Windows.Forms.ComboBox
        Me.cTag2 = New System.Windows.Forms.ComboBox
        Me.btnSave = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnRevert = New System.Windows.Forms.Button
        Me.lblMetaSize = New System.Windows.Forms.Label
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog = New System.Windows.Forms.SaveFileDialog
        Me.txtMetaOffset = New System.Windows.Forms.TextBox
        Me.btnShowSwap = New System.Windows.Forms.Button
        Me.pSwap = New System.Windows.Forms.Panel
        Me.btnSaveTrack = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.listOffsets = New System.Windows.Forms.ListBox
        Me.pExtract = New System.Windows.Forms.Panel
        Me.Label3 = New System.Windows.Forms.Label
        Me.pb = New System.Windows.Forms.ProgressBar
        Me.RecursiveMeta = New System.Windows.Forms.CheckBox
        Me.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog
        Me.pSwap.SuspendLayout()
        Me.pExtract.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblID
        '
        Me.lblID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblID.Location = New System.Drawing.Point(6, 8)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(166, 16)
        Me.lblID.TabIndex = 0
        Me.lblID.Text = "ID: "
        '
        'txtFilename
        '
        Me.txtFilename.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFilename.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFilename.Location = New System.Drawing.Point(7, 28)
        Me.txtFilename.Name = "txtFilename"
        Me.txtFilename.Size = New System.Drawing.Size(352, 20)
        Me.txtFilename.TabIndex = 1
        Me.txtFilename.Text = ""
        '
        'btnSaveMeta
        '
        Me.btnSaveMeta.BackColor = System.Drawing.Color.Black
        Me.btnSaveMeta.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveMeta.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnSaveMeta.Location = New System.Drawing.Point(8, 80)
        Me.btnSaveMeta.Name = "btnSaveMeta"
        Me.btnSaveMeta.Size = New System.Drawing.Size(56, 21)
        Me.btnSaveMeta.TabIndex = 33
        Me.btnSaveMeta.Text = "Save Meta"
        '
        'btnInjectMeta
        '
        Me.btnInjectMeta.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnInjectMeta.BackColor = System.Drawing.Color.Black
        Me.btnInjectMeta.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnInjectMeta.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnInjectMeta.Location = New System.Drawing.Point(166, 80)
        Me.btnInjectMeta.Name = "btnInjectMeta"
        Me.btnInjectMeta.Size = New System.Drawing.Size(58, 21)
        Me.btnInjectMeta.TabIndex = 34
        Me.btnInjectMeta.Text = "Inject Meta"
        '
        'cTag1
        '
        Me.cTag1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cTag1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cTag1.ItemHeight = 12
        Me.cTag1.Location = New System.Drawing.Point(228, 4)
        Me.cTag1.Name = "cTag1"
        Me.cTag1.Size = New System.Drawing.Size(41, 20)
        Me.cTag1.TabIndex = 35
        '
        'cTag3
        '
        Me.cTag3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cTag3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cTag3.ItemHeight = 12
        Me.cTag3.Location = New System.Drawing.Point(318, 4)
        Me.cTag3.Name = "cTag3"
        Me.cTag3.Size = New System.Drawing.Size(41, 20)
        Me.cTag3.TabIndex = 37
        '
        'cTag2
        '
        Me.cTag2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cTag2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cTag2.ItemHeight = 12
        Me.cTag2.Location = New System.Drawing.Point(273, 4)
        Me.cTag2.Name = "cTag2"
        Me.cTag2.Size = New System.Drawing.Size(41, 20)
        Me.cTag2.TabIndex = 38
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.BackColor = System.Drawing.Color.Black
        Me.btnSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnSave.Location = New System.Drawing.Point(296, 54)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(58, 21)
        Me.btnSave.TabIndex = 39
        Me.btnSave.Text = "Save"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 57)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 16)
        Me.Label2.TabIndex = 43
        Me.Label2.Text = "Meta"
        '
        'btnRevert
        '
        Me.btnRevert.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRevert.BackColor = System.Drawing.Color.Black
        Me.btnRevert.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRevert.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnRevert.Location = New System.Drawing.Point(232, 54)
        Me.btnRevert.Name = "btnRevert"
        Me.btnRevert.Size = New System.Drawing.Size(58, 21)
        Me.btnRevert.TabIndex = 44
        Me.btnRevert.Text = "Revert"
        '
        'lblMetaSize
        '
        Me.lblMetaSize.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblMetaSize.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMetaSize.Location = New System.Drawing.Point(232, 82)
        Me.lblMetaSize.Name = "lblMetaSize"
        Me.lblMetaSize.Size = New System.Drawing.Size(128, 17)
        Me.lblMetaSize.TabIndex = 45
        Me.lblMetaSize.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtMetaOffset
        '
        Me.txtMetaOffset.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMetaOffset.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMetaOffset.Location = New System.Drawing.Point(96, 54)
        Me.txtMetaOffset.Name = "txtMetaOffset"
        Me.txtMetaOffset.Size = New System.Drawing.Size(128, 20)
        Me.txtMetaOffset.TabIndex = 46
        Me.txtMetaOffset.Text = ""
        '
        'btnShowSwap
        '
        Me.btnShowSwap.BackColor = System.Drawing.Color.Black
        Me.btnShowSwap.Font = New System.Drawing.Font("Verdana", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnShowSwap.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnShowSwap.ImageAlign = System.Drawing.ContentAlignment.BottomRight
        Me.btnShowSwap.Location = New System.Drawing.Point(49, 55)
        Me.btnShowSwap.Name = "btnShowSwap"
        Me.btnShowSwap.Size = New System.Drawing.Size(40, 17)
        Me.btnShowSwap.TabIndex = 47
        Me.btnShowSwap.Text = "Swap"
        '
        'pSwap
        '
        Me.pSwap.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pSwap.Controls.Add(Me.btnSaveTrack)
        Me.pSwap.Controls.Add(Me.Label1)
        Me.pSwap.Controls.Add(Me.listOffsets)
        Me.pSwap.Location = New System.Drawing.Point(0, 0)
        Me.pSwap.Name = "pSwap"
        Me.pSwap.Size = New System.Drawing.Size(360, 104)
        Me.pSwap.TabIndex = 48
        Me.pSwap.Visible = False
        '
        'btnSaveTrack
        '
        Me.btnSaveTrack.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSaveTrack.BackColor = System.Drawing.Color.Black
        Me.btnSaveTrack.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveTrack.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnSaveTrack.Location = New System.Drawing.Point(296, 6)
        Me.btnSaveTrack.Name = "btnSaveTrack"
        Me.btnSaveTrack.Size = New System.Drawing.Size(56, 18)
        Me.btnSaveTrack.TabIndex = 13
        Me.btnSaveTrack.Text = "Cancel"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(288, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Swap offset with:"
        '
        'listOffsets
        '
        Me.listOffsets.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.listOffsets.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.listOffsets.Location = New System.Drawing.Point(8, 27)
        Me.listOffsets.Name = "listOffsets"
        Me.listOffsets.Size = New System.Drawing.Size(344, 67)
        Me.listOffsets.TabIndex = 2
        '
        'pExtract
        '
        Me.pExtract.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pExtract.Controls.Add(Me.Label3)
        Me.pExtract.Controls.Add(Me.pb)
        Me.pExtract.Location = New System.Drawing.Point(0, 0)
        Me.pExtract.Name = "pExtract"
        Me.pExtract.Size = New System.Drawing.Size(360, 104)
        Me.pExtract.TabIndex = 14
        Me.pExtract.Visible = False
        '
        'Label3
        '
        Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.Location = New System.Drawing.Point(16, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(328, 48)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Scanning metadata for reflexive offsets and dependencies.  This process may take " & _
        "several seconds... please be patient..."
        '
        'pb
        '
        Me.pb.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pb.Location = New System.Drawing.Point(8, 72)
        Me.pb.Name = "pb"
        Me.pb.Size = New System.Drawing.Size(344, 24)
        Me.pb.TabIndex = 0
        '
        'RecursiveMeta
        '
        Me.RecursiveMeta.Location = New System.Drawing.Point(72, 80)
        Me.RecursiveMeta.Name = "RecursiveMeta"
        Me.RecursiveMeta.Size = New System.Drawing.Size(80, 24)
        Me.RecursiveMeta.TabIndex = 50
        Me.RecursiveMeta.Text = "Recursive"
        '
        'MapGUI
        '
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.pExtract)
        Me.Controls.Add(Me.pSwap)
        Me.Controls.Add(Me.btnShowSwap)
        Me.Controls.Add(Me.txtMetaOffset)
        Me.Controls.Add(Me.lblMetaSize)
        Me.Controls.Add(Me.btnRevert)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.cTag2)
        Me.Controls.Add(Me.cTag3)
        Me.Controls.Add(Me.cTag1)
        Me.Controls.Add(Me.btnInjectMeta)
        Me.Controls.Add(Me.btnSaveMeta)
        Me.Controls.Add(Me.txtFilename)
        Me.Controls.Add(Me.lblID)
        Me.Controls.Add(Me.RecursiveMeta)
        Me.Name = "MapGUI"
        Me.Size = New System.Drawing.Size(368, 107)
        Me.pSwap.ResumeLayout(False)
        Me.pExtract.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Event Handlers"

    'Update the progress bar when an update event from the map is fired
    Private Sub UpdateProgressBar(ByVal intValue As Integer) Handles map.SaveMeta_Status
        If intValue > pb.Maximum Then intValue = pb.Maximum
        If intValue < pb.Minimum Then intValue = pb.Minimum
        pb.Value = intValue
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        'Write all of the data back to the index, and reload the map file
        'Reselect the current node, if possible after reloading map
        Dim bw As New BinaryWriter(map.fs)
        Dim s As String

        bw.Seek(map.indexItem(CurrentID).offset, SeekOrigin.Begin)

        'Generate the tagclass
        s = reverseString(cTag1.Text)
        s &= reverseString(cTag2.Text)
        s &= reverseString(cTag3.Text)

        Dim bData As Byte
        For x As Integer = 1 To 12
            bData = Asc(s.Chars(x - 1))
            bw.Write(bData) 'Write the tagclass
        Next x

        bw.Seek(8, SeekOrigin.Current)
        bw.Write(Unsigned("&H" & txtMetaOffset.Text) + map.intMagic)

        bw.Seek(Unsigned(map.indexItem(CurrentID).fileData.stringoffset) - map.intMagic, SeekOrigin.Begin)
        For x As Integer = 1 To txtFilename.Text.Length
            bData = Asc(txtFilename.Text.Chars(x - 1))
            bw.Write(bData)
        Next x
        bw.Write(Chr(0))

        bw.Flush()

        RaiseEvent MetaUpdated(CurrentID)

    End Sub

    Private Sub btnShowSwap_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowSwap.Click
        pSwap.Visible = True
    End Sub

    Private Sub btnSaveTrack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveTrack.Click
        pSwap.Visible = False
    End Sub

    Private Sub listOffsets_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles listOffsets.SelectedIndexChanged
        Dim intLengthOfHex As Integer = listOffsets.SelectedItem.IndexOf(")") - 1
        txtMetaOffset.Text = Mid(listOffsets.SelectedItem, 2, intLengthOfHex)

        pSwap.Visible = False
    End Sub

    Private Sub btnRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRevert.Click
        PopulateControls(CurrentID)
    End Sub

    Private Sub btnSaveMeta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveMeta.Click
        Dim strPath, strFilename As String

        Try
            If RecursiveMeta.Checked Then
                FolderBrowserDialog.Description = "Choose the root folder where the files will be extracted." & vbCrLf & _
                    "Note: The files will be placed in the correct folder hierarchy."
                If FolderBrowserDialog.ShowDialog = DialogResult.Cancel Then Exit Sub
                strPath = FolderBrowserDialog.SelectedPath
                map.bFileInUse = True
                strFilename = map.indexItem(CurrentID).filePath & "." & map.indexItem(CurrentID).tagClass & ".meta"

                map.SaveMeta(CurrentID, strFilename, strPath & "\", True)

                map.bFileInUse = False
                MsgBox("Finished Extracting.", MsgBoxStyle.Information)
            Else
                'Get the correct filename.
                SaveFileDialog.AddExtension = True
                SaveFileDialog.DefaultExt = "*.meta"
                SaveFileDialog.Filter = "Binary Metadata (*.meta)|*.meta"
                SaveFileDialog.FileName = getLevel(map.indexItem(CurrentID).filePath, 0) & "." & map.indexItem(CurrentID).tagClass
                If SaveFileDialog.ShowDialog() = DialogResult.Cancel Then Exit Sub
                strFilename = SaveFileDialog.FileName
                map.bFileInUse = True
                strPath = strFilename.Remove(strFilename.Length - getLevel(strFilename, 0).Length, getLevel(strFilename, 0).Length)
                map.SaveMeta(CurrentID, getLevel(strFilename, 0), strPath, False)
                map.bFileInUse = False
            End If
        Catch ex As Exception
            MsgBox("There was an error while extracting: " & ex.Message)
            map.bFileInUse = False
        End Try
    End Sub

#End Region

#Region "Functions"

    Public Sub PopulateControls(ByVal intID As Integer)
        CurrentID = intID
        lblID.Text = "ID: " & Unsigned(map.indexItem(intID).fileData.id) & _
            " @ 0x" & Hex(map.indexItem(intID).offset)
        Dim item As HaloMap.Map.TAG_STRUCT
        For Each item In map.cTagTypes
            cTag1.Items.Add(item.tag)
            cTag2.Items.Add(item.tag)
            cTag3.Items.Add(item.tag)
        Next
        cTag1.SelectedItem = map.getTagClass(map.indexItem(intID).fileData.tagclass, 3)
        cTag2.SelectedItem = map.getTagClass(map.indexItem(intID).fileData.tagclass, 2)
        cTag3.SelectedItem = map.getTagClass(map.indexItem(intID).fileData.tagclass, 1)
        lblMetaSize.Text = "Meta Size: " & IIf(Unsigned(map.indexItem(intID).estimatedMetaSize) <= 5000000, _
            Unsigned(map.indexItem(intID).estimatedMetaSize) & " bytes", "(error)")
        If Unsigned(map.indexItem(intID).estimatedMetaSize <= 5000000) Then
            btnSaveMeta.Enabled = True : btnInjectMeta.Enabled = True
        Else
            btnSaveMeta.Enabled = False : btnInjectMeta.Enabled = False
        End If
        txtFilename.Text = map.indexItem(intID).filePath
        txtFilename.MaxLength = map.indexItem(intID).filePath.Length
        txtMetaOffset.Text = Hex(Unsigned(map.indexItem(intID).magic_metadata_offset))
        'Add all compatible offsets to listOffsets control
        listOffsets.Items.Clear()
        For x As Integer = 1 To map.indexHeader.tagcount
            If map.indexItem(x).tagClass = map.indexItem(intID).tagClass Then
                listOffsets.Items.Add("(" & Hex(map.indexItem(x).magic_metadata_offset) & ") " & map.indexItem(x).filePath)
            End If
        Next
    End Sub

    Private Sub btnInjectMeta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInjectMeta.Click
        Dim strFilename As String

        OpenFileDialog.FileName = ""
        OpenFileDialog.AddExtension = True
        OpenFileDialog.DefaultExt = "*.meta"
        OpenFileDialog.Filter = "Binary Metadata (*.meta)|*.meta"
        If OpenFileDialog.ShowDialog() = DialogResult.Cancel Then Exit Sub
        strFilename = OpenFileDialog.FileName

        map.InjectMeta(CurrentID, strFilename, strHexToDec("0x" & txtMetaOffset.Text))
    End Sub

#End Region

End Class

'///////////////////////////////////////////////////////////
'// Halo Map Tools
'// Map File Base Class
'///////////////////////////////////////////////////////////
Public Class Map

#Region "Structures"

    '// Main File Header Structure
    Public Structure FILE_HEADER_STRUCT
        Const HEADER_EMPTY_INTS As Integer = 484
        Public id As Integer
        Public version As Integer
        Public decomp_len As Integer
        Public zeros As Integer
        Public offset_to_index_decomp As Int32
        Public MetadataSize As Integer
        <VBFixedString(8)> Public zeros2 As String
        <VBFixedString(32)> Public name As String
        <VBFixedString(32)> Public builddate As String
        Public maptype As Integer
        Public unknown3 As Integer
        Public empty_ints() As Integer 'Length will be dimensioned as HEADER_EMPTY_INTS
        Public footer As String
        Public Sub ReadStruct(ByRef bs As BinaryReader)
            Dim x As Integer 'Loop Counter
            id = bs.ReadInt32
            version = bs.ReadInt32
            decomp_len = bs.ReadInt32
            zeros = bs.ReadInt32
            offset_to_index_decomp = bs.ReadInt32
            MetadataSize = bs.ReadInt32
            zeros2 = bs.ReadChars(8)
            name = bs.ReadChars(32)
            builddate = bs.ReadChars(32)
            maptype = bs.ReadInt32
            id = bs.ReadInt32
            unknown3 = bs.ReadInt32
            ReDim empty_ints(HEADER_EMPTY_INTS)
            For x = 1 To HEADER_EMPTY_INTS
                empty_ints(x) = bs.ReadInt32
            Next x
            footer = bs.ReadChars(4)
        End Sub
        Public Sub ReadDemoStruct(ByRef bs As BinaryReader)
            bs.BaseStream.Seek(&H58C, SeekOrigin.Begin)
            name = bs.ReadChars(32)
            bs.BaseStream.Seek(&H5EC, SeekOrigin.Begin)
            offset_to_index_decomp = bs.ReadInt32
        End Sub
        Public Sub WriteStruct(ByRef bw As BinaryWriter)
            bw.Write(id)
            bw.Write(version)
            bw.Write(decomp_len)
            bw.Write(zeros)
            bw.Write(offset_to_index_decomp)
            bw.Write(MetadataSize)
            bw.Write(zeros2.ToCharArray)
            bw.Write(name.ToCharArray)
            bw.Write(builddate.ToCharArray)
            bw.Write(maptype)
            bw.Write(id)
            bw.Write(unknown3)
            ReDim empty_ints(HEADER_EMPTY_INTS)
            For x As Integer = 1 To HEADER_EMPTY_INTS
                bw.Write(empty_ints(x))
            Next x
            bw.Write(footer.ToCharArray)
        End Sub
    End Structure

    Public Structure TAG_STRUCT
        Public tag As String
        Public name As String
        Public intTag As Integer
    End Structure

    '// Index Header Structure
    Public Structure INDEX_HEADER_STRUCT
        Public index_magic As Int32
        Public starting_id As Long 'The value of the first object identifier in the index.
        Public unknown2 As Integer
        Public unknown3 As Integer 'PC Version Only
        Public tagcount As Integer 'Number of items in the index.
        Public vertex_object_count As Integer
        Public vertex_offset As Integer
        Public indeces_object_count As Integer
        Public indeces_offset As Integer
        Public tagstart As Integer 'Offset of item index.
        Public Sub ReadStruct(ByRef bs As BinaryReader, ByVal pc As Boolean)
            index_magic = bs.ReadInt32
            starting_id = bs.ReadInt32
            unknown2 = bs.ReadInt32
            tagcount = bs.ReadInt32
            vertex_object_count = bs.ReadInt32
            vertex_offset = bs.ReadInt32
            indeces_object_count = bs.ReadInt32
            indeces_offset = bs.ReadInt32
            If pc Then unknown3 = bs.ReadInt32
            tagstart = bs.ReadInt32
        End Sub
        Public Sub WriteStruct(ByRef bw As BinaryWriter, ByVal pc As Boolean)
            bw.Write(index_magic)
            'I really, really, really fucking wish thsere were unsigned ints
            Dim uIntStartingID As System.UInt32 = Convert.ToUInt32(Unsigned(starting_id))
            bw.Write(uIntStartingID)
            bw.Write(unknown2)
            bw.Write(tagcount)
            bw.Write(vertex_object_count)
            bw.Write(vertex_offset)
            bw.Write(indeces_object_count)
            bw.Write(indeces_offset)
            If pc Then bw.Write(unknown3)
            bw.Write(tagstart)
        End Sub
    End Structure

    '// Index Item Structure
    Public Structure INDEX_ITEM_STRUCT
        Public tagclass As String 'Length is 12
        Public id As Long
        Public stringoffset As Integer
        Public offset As Integer
        Public zeros As String 'Length is 8
        Public Sub ReadStruct(ByRef bs As BinaryReader)
            Dim x As Integer
            For x = 1 To 12
                tagclass &= Chr(bs.ReadByte)
            Next x
            id = bs.ReadInt32
            stringoffset = bs.ReadInt32
            offset = bs.ReadInt32
            zeros = bs.ReadChars(8)
        End Sub
        Public Sub WriteStruct(ByRef bw As BinaryWriter)
            Dim bData As Byte
            For x As Integer = 1 To 12
                bData = Asc(tagclass.Chars(x - 1))
                bw.Write(bData) 'Write the tagclass
            Next x
            bw.Write(Convert.ToUInt32(id))
            bw.Write(stringoffset)
            bw.Write(offset)
            For x As Integer = 1 To 8
                bData = CByte(0)
                bw.Write(bData)
            Next x
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

#Region "Public Data Members"

    Public fileNumber As Integer = 0 'The file identifier for opening the map file.
    Public strFilename As String
    Public fileHeader As New FILE_HEADER_STRUCT
    Public indexHeader As New INDEX_HEADER_STRUCT
    Public indexItem() As INDEX_ITEM_EXPANDED_STRUCT
    Public intMagic As Long
    Public fs As FileStream
    Public cTagTypes As New Collection
    Public pc As Boolean = False
    Public tagMinimum As Integer
    Public tagMaximum As Integer
    Public bFileInUse As Boolean = False
#End Region

#Region "Functions"

    '/////////////////////////////////////////////////////
    '// Constructor
    '/////////////////////////////////////////////////////
    Public Sub New()
        Dim t As TAG_STRUCT

        'Add all of the known tags
        t.tag = Chr(255) & Chr(255) & Chr(255) & Chr(255) : t.name = "(none)" : cTagTypes.Add(t, t.tag)
        t.tag = "bitm" : t.name = "Bitmap" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "snd!" : t.name = "Sound" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "lsnd" : t.name = "Looping Sound" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "weap" : t.name = "Weapon" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "vehi" : t.name = "Vehicle" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "mode" : t.name = "Model (Xbox)" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "mod2" : t.name = "Model (PC)" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "proj" : t.name = "Projectile" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "lens" : t.name = "Lens" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "scnr" : t.name = "Scenario" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "ssce" : t.name = "Sound Scenery" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "snde" : t.name = "Sound Environment" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "obje" : t.name = "Object" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "unit" : t.name = "Unit" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "bipd" : t.name = "Biped" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "item" : t.name = "Item" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "eqip" : t.name = "Equipment" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "garb" : t.name = "Garbage" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "scen" : t.name = "Scenery" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "devi" : t.name = "Device" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "mach" : t.name = "Machine" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "ctrl" : t.name = "Control" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "lifi" : t.name = "Light Fixture" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "plac" : t.name = "Placeholder" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "shdr" : t.name = "Shader" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "ustr" : t.name = "Unicode String list" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "sbsp" : t.name = "Structure BSP" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "mply" : t.name = "Multiplayer Scenario" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "itmc" : t.name = "Item Collection" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "hmt " : t.name = "HUD Message Text" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "trak" : t.name = "Track" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "actr" : t.name = "Actor" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "actv" : t.name = "Actor Variant" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "antr" : t.name = "Animation Trigger" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "sky " : t.name = "Sky" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "grhi" : t.name = "Grenade HUD Interface" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "unhi" : t.name = "Unit HUD Interface" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "wphi" : t.name = "Weapon HUD Interface" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "coll" : t.name = "Collision Model" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "cont" : t.name = "Contrail" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "deca" : t.name = "Decal" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "effe" : t.name = "Effect" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "part" : t.name = "Particle" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "pctl" : t.name = "Particle System" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "rain" : t.name = "Weather Particle" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "matg" : t.name = "Game Globals" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "hud#" : t.name = "HUD Number" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "hudg" : t.name = "HUD Globals" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "jpt!" : t.name = "Damage" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "ligh" : t.name = "Light" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "glw!" : t.name = "Glow" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "mgs2" : t.name = "Light Volume" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "phys" : t.name = "Physics" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "pphy" : t.name = "Point Physics" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "fog " : t.name = "Fog" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "wind" : t.name = "Wind" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "senv" : t.name = "Shader Environment" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "dobc" : t.name = "Detail Object Collection" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "font" : t.name = "Font" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "udlg" : t.name = "Dialog" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "soso" : t.name = "Shader Model" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "sotr" : t.name = "Shader Transparency" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "swat" : t.name = "Shader Water" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "sgla" : t.name = "Shader Glass" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "smet" : t.name = "Shader Metal" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "spla" : t.name = "Shader Plasma" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "schi" : t.name = "Shader Transparancy Variant" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "scex" : t.name = "Unknown PC Shader" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "tagc" : t.name = "Unknown PC Tag" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "foot" : t.name = "Unknown PC Tag" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "ant!" : t.name = "Antenna" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "str#" : t.name = "Unknown PC Tag" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "colo" : t.name = "Unknown PC Tag" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "flag" : t.name = "Unknown PC Tag" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "metr" : t.name = "Meter" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "devc" : t.name = "PC Device Default" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "DeLa" : t.name = "DeLa" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)
        t.tag = "Soul" : t.name = "Unknown PC Tag" : t.intTag = StringToInt(t.tag) : cTagTypes.Add(t, t.tag)

        'Find the minimum and maximum
        tagMinimum = &HFFFFFFF
        tagMaximum = 0
        For Each t2 As TAG_STRUCT In cTagTypes
            If Not t2.name = "(none)" Then
                If t2.intTag < tagMinimum Then tagMinimum = t2.intTag
                If t2.intTag > tagMaximum Then tagMaximum = t2.intTag
            End If
        Next
    End Sub

    '////////////////////////////////////////////
    '// Convert a string to an integer
    '////////////////////////////////////////////
    Private Function StringToInt(ByVal s As String) As Integer
        Dim i As Integer
        s = reverseString(s)
        If s.Length > 4 Then Return 0
        For x As Integer = 1 To s.Length
            i = i Or (CInt(Asc(Mid(s, x, 1))) << (8 * (x - 1)))
        Next
        Return i
    End Function

    '/////////////////////////////////////////////////////
    '// Open the Halo .map file that is to be processed
    '/////////////////////////////////////////////////////
    Enum eOpenFile As Byte
        Success = &H1
        Xbox = &H2
        PC = &H4
        PCDemo = &H8
        InvalidFile = &H10
        FileNotFound = &H20
        SharingViolation = &H40
        UnknwonError = &HFF
    End Enum
    Public Function OpenFile(ByVal strFname As String) As Byte
        Dim result As Byte

        Try
            strFilename = strFname
            fs = New FileStream(strFname, FileMode.Open, FileAccess.ReadWrite)
            Dim br1 As New BinaryReader(fs) 'Open the file
            fileHeader.ReadStruct(br1)
            Select Case fileHeader.version
                Case 5
                    result = result Or eOpenFile.Success Or eOpenFile.Xbox
                Case 7, 13
                    result = result Or eOpenFile.Success Or eOpenFile.PC
                Case Else
                    'Check and see if this is the demo
                    br1.BaseStream.Seek(&H2D1, SeekOrigin.Begin)
                    If br1.ReadInt32 = 909587760 Then 'This is the build for demo maps
                        result = result Or eOpenFile.Success Or eOpenFile.PCDemo
                    Else
                        result = result Or eOpenFile.InvalidFile
                    End If
            End Select

            br1.BaseStream.Seek(0, SeekOrigin.Begin)
            Return result
        Catch
            Return eOpenFile.UnknwonError Or eOpenFile.InvalidFile
        End Try
    End Function
    Public Function OpenFile(ByRef mapFS As FileStream) As Byte
        Dim result As Byte

        'Try
        strFilename = mapFS.Name
        fs = mapFS
        Dim br1 As New BinaryReader(fs) 'Open the file
        fileHeader.ReadStruct(br1)
        Select Case fileHeader.version
            Case 5
                result = result Or eOpenFile.Success Or eOpenFile.Xbox
            Case 7, 13
                result = result Or eOpenFile.Success Or eOpenFile.PC
            Case Else
                'Check and see if this is the demo
                br1.BaseStream.Seek(&H2D1, SeekOrigin.Begin)
                If br1.ReadInt32 = 909587760 Then 'This is the build for demo maps
                    result = result Or eOpenFile.Success Or eOpenFile.PCDemo
                Else
                    result = result Or eOpenFile.InvalidFile
                End If
        End Select

        br1.BaseStream.Seek(0, SeekOrigin.Begin)
        Return result
        'Catch
        'Return eOpenFile.UnknwonError Or eOpenFile.InvalidFile
        'End Try
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
    '// Read in the map's resource index
    '/////////////////////////////////////////////////////
    Enum eReadFile As Byte
        Success = &H0
        Compressed = &H10
        UnknwonError = &HFF
    End Enum
    Public Structure OffsetLookup
        Public Offset As Integer
        Public Key As Integer
    End Structure
    Public Function ReadFile(ByVal mapType As eOpenFile) As eReadFile
        Dim br1 As New BinaryReader(fs)

        'Try
        br1.BaseStream.Seek(0, SeekOrigin.Begin)
        Select Case mapType
            Case eOpenFile.PCDemo
                fileHeader.ReadDemoStruct(br1) 'Read in the file header.
                pc = True
            Case eOpenFile.PC
                fileHeader.ReadStruct(br1) 'Read in the file header.
                pc = True
            Case eOpenFile.Xbox
                fileHeader.ReadStruct(br1) 'Read in the file header.
                pc = False
        End Select

        'We can't handle compresed files right now so who cares...
        If Unsigned(br1.BaseStream.Length) < Unsigned(fileHeader.decomp_len) Then
            If Not mapType And eOpenFile.PCDemo Then Return eReadFile.Compressed
        End If

        br1.BaseStream.Seek(fileHeader.offset_to_index_decomp, SeekOrigin.Begin)
        indexHeader.ReadStruct(br1, pc)

        intMagic = CLng(Unsigned(indexHeader.index_magic) - (fileHeader.offset_to_index_decomp + _
                IIf(pc, 40, 36)))

        'Read in all of the Index Items
        Dim LookupTable(indexHeader.tagcount) As OffsetLookup
        ReDim indexItem(indexHeader.tagcount)  'Reserve memory for all of the index items.
        For x As Integer = 1 To indexHeader.tagcount
            indexItem(x).offset = br1.BaseStream.Position  'Get the item's offset
            indexItem(x).fileData.ReadStruct(br1)
            indexItem(x).magic_metadata_offset = Unsigned(indexItem(x).fileData.offset) - intMagic
            indexItem(x).magic_string_offset = Unsigned(indexItem(x).fileData.stringoffset) - intMagic
            indexItem(x).tagClass = getTagClass(indexItem(x).fileData.tagclass, 3)
            'SBSP tags have invalid offsets, so they cannot be included in in meta size calculation
            If indexItem(x - 1).tagClass = "sbsp" Then
                indexItem(x - 1).magic_metadata_offset = indexItem(x).magic_metadata_offset
            End If
        Next
        If indexItem(indexHeader.tagcount).tagClass = "sbsp" Then _
            indexItem(indexHeader.tagcount).magic_metadata_offset = 0

        'Sort the LookupTable
        Dim idx, max As Integer
        Dim indexTemp(indexHeader.tagcount) As INDEX_ITEM_EXPANDED_STRUCT
        indexTemp = indexItem.Clone
        For y As Integer = LookupTable.Length - 1 To 0 Step -1
            max = 0 : idx = 0
            For x As Integer = 0 To indexTemp.Length - 1
                If indexTemp(x).offset > max Then
                    max = indexTemp(x).offset
                    idx = x
                End If
            Next
            indexTemp(idx).offset = 0
            LookupTable(y).Offset = max
            LookupTable(y).Key = idx
        Next

        'For debugging purposes
        'Write the sorted LookUp table to a file
        'Dim lutOut As New StreamWriter(New FileStream(fs.Name & ".sortedindex.txt", FileMode.Create))
        'Now that we've got an array of all the index item data, calculate the extended data.
        For x As Integer = 1 To indexHeader.tagcount
            'Find the current item in the lookup table
            For t As Integer = 0 To LookupTable.Length - 1
                Try
                    If LookupTable(t).Key = x Then
                        indexItem(x - 1).estimatedMetaSize = indexItem(x).magic_metadata_offset _
                            - indexItem(LookupTable(t - 1).Key).magic_metadata_offset
                        If indexItem(x - 1).estimatedMetaSize < 0 Then _
                            indexItem(x - 1).estimatedMetaSize = 0
                        'Estimate the size of the previous metadata structure.
                    End If
                Catch
                End Try
            Next
            'lutOut.WriteLine(x - 1 & " : " & indexItem(x - 1).magic_metadata_offset & " : " & indexItem(x - 1).estimatedMetaSize & _
            '" : " & indexItem(x - 1).tagClass)
            indexItem(x).filePath = ReadString(br1, indexItem(x).magic_string_offset)
        Next

        indexItem(indexHeader.tagcount).estimatedMetaSize = Unsigned(fs.Length - Unsigned(indexItem(indexHeader.tagcount).magic_metadata_offset))

        If indexItem(indexHeader.tagcount).tagClass = "sbsp" Then _
            indexItem(indexHeader.tagcount).estimatedMetaSize = 0
        'lutOut.WriteLine(indexHeader.tagcount & " : " & indexItem(indexHeader.tagcount).magic_metadata_offset & " : " & _
        'indexItem(indexHeader.tagcount).estimatedMetaSize & " : " & indexItem(indexHeader.tagcount).tagClass)
        indexItem(indexHeader.tagcount).filePath = ReadString(br1, indexItem(indexHeader.tagcount).magic_string_offset)
        'lutOut.Close()

        'indexItem(indexHeader.tagcount).estimatedMetaSize = Unsigned(fileHeader.decomp_len) - Unsigned(indexItem(indexHeader.tagcount).magic_metadata_offset)
        Return eReadFile.Success
        'Catch
        '   Return eReadFile.UnknwonError
        'End Try
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

    Public Function GetInts(ByVal offset As Long, ByVal number As Long, ByRef intChunk() As Integer) As Boolean
        Dim br1 As New BinaryReader(fs)
        ReDim intChunk(number - 1) 'Resize the byte array.

        br1.BaseStream.Seek(offset, SeekOrigin.Begin) 'Read in the data
        For x As Integer = 0 To number - 1
            intChunk(x) = br1.ReadInt32
        Next
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

        'Console.WriteLine(br.BaseStream.Position)
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
            Case "deathisland"
                MapName = "Death Island"
            Case "dangercanyon"
                MapName = "Danger Canyon"
            Case "gephyrophobia"
                MapName = "Gephyrophobia"
            Case "timberland"
                MapName = "Timberland"
            Case "icefields"
                MapName = "Ice Fields"
            Case "infinity"
                MapName = "Infinity"
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

    Public Function IsTag(ByVal i As Integer) As Boolean
        If i < tagMinimum Or i > tagMaximum Then Return False
        Dim t As TAG_STRUCT
        For Each t In cTagTypes
            If t.intTag = i Then Return True
        Next
        Return False
    End Function

    Public Function IsTag(ByVal strString As String) As Boolean
        Dim t As TAG_STRUCT
        Dim strNullTag As String = Chr(255) & Chr(255) & Chr(255) & Chr(255)
        For Each t In cTagTypes
            If (t.tag.ToUpper = strString.ToUpper) Then
                If t.tag <> strNullTag Then Return True
            End If
        Next
        Return False
    End Function

    Public Function IsID(ByVal ID As Integer) As Boolean
        'Check through all of the index items and see if this number exists as an ident
        If (ID > 0) Then Return False
        If ((ID And &H7FFFFFFF) < &H61740000) Then Return False
        For x As Integer = 1 To indexHeader.tagcount
            If indexItem(x).fileData.id = ID Then Return True
        Next
        Return False
    End Function

    Public Function IsOffset(ByVal intOffset As Long) As Boolean
        intOffset = Unsigned(intOffset)
        Dim unsignedMagic As Integer = Unsigned(intMagic)
        If intOffset > unsignedMagic And intOffset <= (unsignedMagic + fs.Length) Then
            Return True
        Else
            Return False
        End If
    End Function

    '////////////////////////////////////////////////////////////////
    '// Scans a chunk of metadata for reflexive offsets and idents,
    '// then saves the results in an XML file along with the binary
    '// metadata itself, which is stored in a .meta file
    '////////////////////////////////////////////////////////////////
    Public Event SaveMeta_Status(ByVal PercentCompleted As Integer)
    Public Function SaveMeta(ByVal intID As Integer, ByVal strFilename As String, ByVal strPath As String, ByVal Recursive As Boolean, _
            Optional ByVal StartOffset As Long = 0, Optional ByVal Length As Long = 0) As Map.OffsetRange

        'If the specified folder doesn't exist, create it.
        Dim start As DateTime = Now()
        Dim strfilepath As String = strFilename.Remove(strFilename.Length - getLevel(strFilename, 0).Length, getLevel(strFilename, 0).Length)
        If Not Directory.Exists(strPath & strfilepath) Then Directory.CreateDirectory(strPath & strfilepath)

        Dim bw As New BinaryWriter(New FileStream(strPath & strFilename, FileMode.Create))
        Dim txtWriter As New XmlTextWriter((strPath) & strFilename.Remove(strFilename.Length - 5, 5) & ".xml", Nothing)
        txtWriter.Formatting = Formatting.Indented
        Dim bChunk() As Integer

        'Get the raw metadata from the map offset
        If StartOffset = 0 And Length = 0 Then
            GetInts(Unsigned(indexItem(intID).fileData.offset) - intMagic, indexItem(intID).estimatedMetaSize / 4, bChunk)
        Else
            GetInts(StartOffset, Length / 4, bChunk)
        End If

        'Write the map name to the file
        txtWriter.WriteStartDocument()
        txtWriter.WriteComment("Halo Map Tools: Metadata Structure File")
        txtWriter.WriteStartElement("Results")
        txtWriter.WriteElementString("Map", MapName())
        txtWriter.WriteElementString("Tag", indexItem(intID).fileData.tagclass)
        txtWriter.WriteElementString("Filename", indexItem(intID).filePath)

        'Examine the metadata and write the structure file
        Dim tmpInt As Integer
        Dim tmpTag As Integer
        Dim tmpIntMagic As Long
        Dim tmpStr As String
        Dim tmpID As Long
        Dim intL As Integer
        Dim ProcessTags As Boolean = False

        For x As Integer = 0 To bChunk.Length - 1

            'Fire an event every 32 dwords that gives the current percent complete
            If x Mod 32 = 0 Then RaiseEvent SaveMeta_Status(CInt(x / (bChunk.Length / 100)))

            tmpInt = bChunk(x)
            'Don't process it if it's any of these values.
            If (Not (tmpInt = 0)) And (Not (tmpInt = &HFFFFFFFF)) And (Not (tmpInt = &HCACACACA)) Then
                tmpIntMagic = Unsigned(tmpInt) - intMagic

                'See if the word is a tag.  If it is, process it.
                Dim ReflexiveTagCount As Integer = 0
                Dim intTranslation As Integer
                If IsOffset(tmpInt) Then
                    'Is it refelxive?
                    If ((tmpIntMagic) >= indexItem(intID).magic_metadata_offset) And _
                        ((tmpIntMagic) <= (indexItem(intID).magic_metadata_offset + indexItem(intID).estimatedMetaSize)) Then
                        'Could be a reflexive - make sure we have a valid chunkcount before it
                        'If bChunk(x - 1) > 0 And bChunk(x - 1) < 200000 Then
                        'This seems to be a reflexive offset.
                        txtWriter.WriteStartElement("Reflexive")
                        txtWriter.WriteElementString("Location", "0x" & Hex(x * 4))
                        intTranslation = tmpIntMagic - indexItem(intID).magic_metadata_offset
                        txtWriter.WriteElementString("Translation", "0x" & Hex(intTranslation))
                        txtWriter.WriteElementString("ChunkCount", bChunk(x - 1))
                        txtWriter.WriteEndElement()
                        'Else'
                        If bChunk(x - 1) > 200000 Then
                            Console.WriteLine(strFilename & " - 0x" & Hex(x * 4) & " - " & bChunk(x - 1))
                        End If
                    End If
                Else
                    If IsID(bChunk(x)) Then
                        'Check to see if the ID is preceded by a tag.  If it is, then we have a 
                        'Dependency.  If not, then it's a LoneID.
                        tmpTag = IIf(x >= 3, bChunk(x - 3), 0)
                        tmpID = bChunk(x)
                        If IsTag(tmpTag) Then
                            'This is a Dependency
                            Try
                                txtWriter.WriteStartElement("Dependency")
                                txtWriter.WriteElementString("Location", "0x" & Hex(x * 4))
                                txtWriter.WriteElementString("Tagclass", indexItem(LocateByID(tmpID)).fileData.tagclass)
                                txtWriter.WriteElementString("Filename", indexItem(LocateByID(tmpID)).filePath)
                                txtWriter.WriteEndElement()
                            Catch
                                txtWriter.WriteComment("Invalid ID found at 0x" & Hex(x) & " : " & tmpID)
                            End Try
                        Else
                            'We've found a LoneID.. yey!
                            Try
                                txtWriter.WriteStartElement("LoneID")
                                txtWriter.WriteElementString("Location", "0x" & Hex(x * 4))
                                txtWriter.WriteElementString("Tagclass", indexItem(LocateByID(tmpID)).fileData.tagclass)
                                txtWriter.WriteElementString("Filename", indexItem(LocateByID(tmpID)).filePath)
                                txtWriter.WriteEndElement()
                            Catch
                                txtWriter.WriteComment("Invalid ID found at 0x" & Hex(x) & " : " & tmpID)
                            End Try
                        End If
                        Application.DoEvents()
                    End If
                End If
            End If
        Next

        'Write the binary metadata to the .meta file
        For x As Integer = 0 To bChunk.Length - 1
            bw.Write(bChunk(x))
        Next x
        bw.Close()

        'Now that we've written the binary metadata and the XML structure file, test to see if this
        'tag has binary attachments.  If so, call the appropriate Raw meta extractor.
        Dim oRange As New Map.OffsetRange
        Select Case indexItem(intID).tagClass
            Case "mod2"
                oRange = SaveRawMod2(strPath & strfilepath, getLevel(indexItem(intID).filePath, 0) & "." & indexItem(intID).tagClass, _
                    indexItem(intID).magic_metadata_offset, txtWriter)
        End Select

        'Close the XML document
        txtWriter.WriteEndElement()
        txtWriter.WriteEndDocument()
        txtWriter.Close()
        txtWriter = Nothing

        If Recursive Then
            'Open the meta file, and recursively call this function for all of it's dependencies
            Dim xmlD As New XmlDocument
            xmlD.Load((strPath) & strFilename.Remove(strFilename.Length - 5, 5) & ".xml")
            Dim nList As XmlNodeList = xmlD.SelectNodes("/Results/Dependency")
            Dim strTagClass As String

            For Each n As XmlNode In nList
                strTagClass = n.ChildNodes(1).InnerText
                strFilename = n.ChildNodes(2).InnerText
                intID = LocateByID(LocateByTagClassAndFilename(strTagClass, strFilename))
                If Not File.Exists(strPath & strFilename & "." & reverseString(Mid(strTagClass, 1, 4)) & ".meta") Then
                    SaveMeta(intID, strFilename & "." & reverseString(Mid(strTagClass, 1, 4)) & ".meta", strPath, True, _
                        indexItem(intID).magic_metadata_offset, indexItem(intID).estimatedMetaSize)
                End If
            Next
        End If
        Return oRange
    End Function

    '/////////////////////////////////////////////////////////////
    '// Inject metadata into the map at the specified offset,
    '// correcting idents and reflexive offsets based on the
    '// corresponding XML structure file
    '/////////////////////////////////////////////////////////////
    Public Sub InjectMeta(ByVal intID As Long, ByVal strFilename As String, ByVal StartOffset As Integer, _
        Optional ByVal vertexInjectionOffset As Integer = 0, Optional ByVal indexInjectionOffset As Integer = 0)

        Dim br As New BinaryReader(New FileStream(strFilename, FileMode.Open, FileAccess.ReadWrite))
        Dim sr As StreamReader
        Dim bw As New BinaryWriter(fs)
        Dim bChunk() As Byte
        Dim strMapName As String

        ReDim bChunk(br.BaseStream.Length - 1)
        bChunk = br.ReadBytes(br.BaseStream.Length)

        'If a structure file exists with the raw meta file, we need to parse it
        'and make sure all dependencies exist as well as update the reflexives.
        If File.Exists(strFilename.Remove((strFilename.Length - 5), 5) & ".xml") Then
            Dim strTemp, strRootTag As String
            Dim strTag, strFilePath As String
            Dim ID, intLocation As Integer
            Dim intTranslation As Integer
            Dim strID As String
            Dim strTemp2 As String
            Dim intTemp As Integer
            Dim xmlN As XmlNode
            Dim tmpID As Integer
            Dim xmlD As New XmlDocument

            'Load the XML file
            xmlD.Load(strFilename.Remove((strFilename.Length - 5), 5) & ".xml")

            'Select all of the nodes in the document
            '** vvv Hack here
            Dim intOriginalOffset As Integer
            '** ^^^
            For Each xmlN In xmlD.SelectNodes("/Results/*")
                Select Case xmlN.Name
                    Case "Reflexive"
                        intLocation = strHexToDec(xmlN.ChildNodes(0).InnerText)
                        intTranslation = strHexToDec(xmlN.ChildNodes(1).InnerText)
                        strTemp2 = Hex(Unsigned(StartOffset + intTranslation) + intMagic)
                        bChunk(intLocation) = Val("&H" & Mid(strTemp2, 7, 2))
                        bChunk(intLocation + 1) = Val("&H" & Mid(strTemp2, 5, 2))
                        bChunk(intLocation + 2) = Val("&H" & Mid(strTemp2, 3, 2))
                        bChunk(intLocation + 3) = Val("&H" & Mid(strTemp2, 1, 2))
                    Case "Dependency", "LoneID"
                        intLocation = strHexToDec(xmlN.ChildNodes(0).InnerText)
                        strTag = xmlN.ChildNodes(1).InnerText
                        strFilePath = xmlN.ChildNodes(2).InnerText

                        'We need to get the ID here based on the tag and the filename
                        strID = Hex(LocateByTagClassAndFilename(strTag, strFilePath))
                        If strID = "0" Then
                            Console.WriteLine("The dependency ID was not found")
                        End If

                        bChunk(intLocation) = Val("&H" & Mid(strID, 7, 2))
                        bChunk(intLocation + 1) = Val("&H" & Mid(strID, 5, 2))
                        bChunk(intLocation + 2) = Val("&H" & Mid(strID, 3, 2))
                        bChunk(intLocation + 3) = Val("&H" & Mid(strID, 1, 2))
                    Case "Map"
                        strMapName = xmlN.InnerText
                End Select
            Next

            'My temporary hack - look for mod2 corrections from the block
            For Each xmlN In xmlD.SelectNodes("/Results/Asset/*")
                Select Case xmlN.Name
                    Case "Correction" 'Raw pointer offset correction
                        'Temporary hack to finishe the damn LibraryMP mod
                        intLocation = strHexToDec(xmlN.ChildNodes(0).InnerText)
                        intTranslation = strHexToDec(xmlN.ChildNodes(1).InnerText)
                        If strMapName = "Update" Then
                            Select Case xmlN.ParentNode.ChildNodes(0).InnerText
                                Case "VertexBlock"
                                    'Current Vertex Size, before the new model
                                    intTranslation += 7259476 'Original Death Island
                                    intTranslation += 664496 'Wraith Vertex Size
                                Case "IndexBlock"
                                    'Current Index size, before the new model
                                    intTranslation += 359476 'Original Death Island
                                    intTranslation += 31316 'Wraith Index Size
                            End Select
                            intOriginalOffset = CLng(bChunk(intLocation)) + _
                                                (CLng(bChunk(intLocation + 1)) << 8) + _
                                                (CLng(bChunk(intLocation + 2)) << 16) + _
                                                (CLng(bChunk(intLocation + 3)) << 24)
                            '***************** Hack *************************************************
                            '*** intOriginalOffset will be set to 0 to allow for the model data
                            '*** to immediately follow existing data
                            '*************** End Hack ***********************************************
                            intOriginalOffset = 0
                            strTemp2 = Hex(intTranslation + intOriginalOffset)
                            strTemp2 = strTemp2.PadLeft(8, "0")
                            bChunk(intLocation) = Val("&H" & Mid(strTemp2, 7, 2))
                            bChunk(intLocation + 1) = Val("&H" & Mid(strTemp2, 5, 2))
                            bChunk(intLocation + 2) = Val("&H" & Mid(strTemp2, 3, 2))
                            bChunk(intLocation + 3) = Val("&H" & Mid(strTemp2, 1, 2))
                        End If
                End Select
            Next
            xmlD = Nothing
        End If

        'Load the XML file
        Dim strRawName As String = strFilename.Remove((strFilename.Length - 5), 5) & ".raw.xml"
        If File.Exists(strRawName) Then
            Dim xmlD As New XmlDocument
            Dim xmlN As XmlNode
            Dim intLocation, intTranslation As Integer
            Dim strTemp2 As String

            xmlD.Load(strRawName)

            For Each xmlN In xmlD.SelectNodes("/Results/*")
                Select Case xmlN.Name
                    Case "Vertex"
                        intLocation = strHexToDec(xmlN.ChildNodes(0).InnerText)
                        intTranslation = strHexToDec(xmlN.ChildNodes(1).InnerText)
                        strTemp2 = Hex(Unsigned((vertexInjectionOffset - indexHeader.vertex_offset) + intTranslation)).PadLeft(8, "0")
                        bChunk(intLocation) = Val("&H" & Mid(strTemp2, 7, 2))
                        bChunk(intLocation + 1) = Val("&H" & Mid(strTemp2, 5, 2))
                        bChunk(intLocation + 2) = Val("&H" & Mid(strTemp2, 3, 2))
                        bChunk(intLocation + 3) = Val("&H" & Mid(strTemp2, 1, 2))
                    Case "Index"
                        intLocation = strHexToDec(xmlN.ChildNodes(0).InnerText)
                        intTranslation = strHexToDec(xmlN.ChildNodes(1).InnerText)
                        strTemp2 = Hex(Unsigned(((indexInjectionOffset - indexHeader.vertex_offset) - indexHeader.indeces_offset) + intTranslation)).PadLeft(8, "0")
                        bChunk(intLocation) = Val("&H" & Mid(strTemp2, 7, 2))
                        bChunk(intLocation + 1) = Val("&H" & Mid(strTemp2, 5, 2))
                        bChunk(intLocation + 2) = Val("&H" & Mid(strTemp2, 3, 2))
                        bChunk(intLocation + 3) = Val("&H" & Mid(strTemp2, 1, 2))
                End Select
            Next
        End If

        'bw.Seek(Unsigned(indexItem(intID).fileData.offset) - intMagic, SeekOrigin.Begin)
        bw.Seek(StartOffset, SeekOrigin.Begin)
        bw.Write(bChunk, 0, bChunk.Length - 4)
        bw.Flush()
    End Sub

    '////////////////////////////////////////////////
    '// Return an item's position in the map index
    '// based on its filename and tag
    '////////////////////////////////////////////////
    Public Function LocateByTagClassAndFilename(ByVal strDependencyTagClass As String, _
            ByVal strDependencyFilename As String) As Long

        For x As Integer = 1 To indexHeader.tagcount
            If indexItem(x).fileData.tagclass = strDependencyTagClass _
            And indexItem(x).filePath = strDependencyFilename Then
                Return indexItem(x).fileData.id
            End If
        Next
        Return 0
    End Function

#End Region

    '////////////////////////////////////////////
    '// Halo Map Tools                         //
    '// -------------------------------------- //
    '// Mod2 Helper - Incomplete               //
    '// Will process a mod2 tag and allow for  //
    '// Raw Offsets to be placed in the XML    //
    '// structure files.                       //
    '////////////////////////////////////////////

#Region "Structures"

#Region "Ported from Model.h"

Public Structure CHUNK
    Public count As Integer
    Public offset As UInt32
    Public unknown As UInt32
    Public Sub Read(ByRef br As BinaryReader)
        count = br.ReadInt32
        offset = br.ReadUInt32
        unknown = br.ReadUInt32
    End Sub
End Structure

Public Structure STRUCT_MODEL_HEADER
    Public zero1 As UInt32
    Public unknown1a As UInt32
    Public offset1 As UInt32
    Public offset2 As UInt32
    Public offset3 As UInt32
    Public offset4 As UInt32
    Public offset5 As UInt32
    Public unknown2() As Short '6 shorts
    Public unknown1b() As UInt32 '2 dwords
    Public uScale As Single
    Public vScale As Single
    Public unknown3() As UInt32 '29 dwords
    Public AttachmentPoints As CHUNK
    Public Bones As CHUNK
    Public Models As CHUNK
        Public SubModels As CHUNK
    Public Shaders As CHUNK
    Public Sub Read(ByRef br As BinaryReader)
        zero1 = br.ReadUInt32
        unknown1a = br.ReadUInt32
        offset1 = br.ReadUInt32
        offset2 = br.ReadUInt32
        offset3 = br.ReadUInt32
        offset4 = br.ReadUInt32
        offset5 = br.ReadUInt32
            ReDim unknown2(10)
            For x As Integer = 1 To 10
                unknown2(x) = br.ReadInt16
            Next
            uScale = br.ReadSingle
            vScale = br.ReadSingle
            Dim unknown3(29)
            For x As Integer = 1 To 29
                unknown3(x) = br.ReadUInt32
            Next x
            AttachmentPoints.Read(br)
            Bones.Read(br)
            Models.Read(br)
            SubModels.Read(br)
            Shaders.Read(br)
    End Sub
End Structure

Public Structure STRUCT_PC_MODEL_BONE
    Public name() As Char '32 bytes
    Public unknown1 As Short() '4 shorts
    Public unknown2 As Single()  '29 floats
End Structure

Public Structure STRUCT_SHADER_DESCRIPTION
    Public tag As UInt32
    Public namePtr As UInt32
    Public zero As UInt32
    Public Shader_TagID As UInt32
    Public unknown() As UInt32 '4 dwords
End Structure

Public Structure STRUCT_MODEL_CHUNK2
    Public unknown1 As UInt32
    Public unknown2 As UInt32
    Public unknown3() As UInt32 '14 dwords
    Public chunk1 As CHUNK
    Public name As String 'Length is 32 bytes
    Public unknown4() As UInt32 '8 dwords
    Public chunk2 As CHUNK
    Public name2 As String 'Length is 12 bytes
End Structure

    Public Structure STRUCT_MODEL_SUBMODEL_HEADER
        Dim count As Int32
        Dim offset As UInt32
        Dim zeros() As Integer '10
        Public Sub Read(ByRef br As BinaryReader)
            Dim x As Integer
            count = br.ReadInt32()
            offset = br.ReadUInt32()
            ReDim zeros(9)
            For x = 0 To 9
                zeros(x) = br.ReadInt32()
            Next
        End Sub
    End Structure

    Public Structure STRUCT_XBOX_MODEL_VERTEX
        Public unknown() As UInt32 '8 dwords
    End Structure

    Public Structure STRUCT_MODEL_TRIANGLESTRIP_HEADER
        Dim StartingOffset As Integer 'Helper - not a part of the struct
        Dim zero As Integer
        Dim number As Short
        Dim oheckseffeffeffeff As Short
        Dim unknown() As Integer  '6
        Dim zeros() As Integer  '9
        Dim one As Integer
        Dim index_count As Integer
        Dim index_offset As Int32
        Dim unknown1_offset As Int32
        Dim unknown1_count As Integer
        Dim vertex_count As Integer
        Dim unknown2() As Integer  '2
        Dim vertex_offset As Int32
        Dim zeros2() As Integer
        Public Sub readstruct(ByRef br As BinaryReader)
            Dim x As Integer
            ReDim unknown(6)
            ReDim zeros(9)
            ReDim unknown2(2)
            StartingOffset = br.BaseStream.Position
            zero = br.ReadInt16()
            number = br.ReadInt32()
            oheckseffeffeffeff = br.ReadInt16()
            For x = 0 To 5
                unknown(x) = br.ReadInt32()
            Next
            For x = 0 To 8
                zeros(x) = br.ReadInt32()
            Next
            one = br.ReadInt32()
            index_count = br.ReadInt32() + 2
            index_offset = br.ReadInt32()
            unknown1_offset = br.ReadInt32()
            unknown1_count = br.ReadInt32()
            vertex_count = br.ReadInt32()
            unknown2(1) = br.ReadInt32()
            unknown2(2) = br.ReadInt32()
            vertex_offset = br.ReadInt32()
        End Sub
        Public Sub ReadPC(ByRef br As BinaryReader)
            Dim x As Integer
            ReDim unknown(6)
            ReDim zeros(9)
            ReDim unknown2(2)
            ReDim zeros2(7)
            StartingOffset = br.BaseStream.Position
            zero = br.ReadInt32()
            number = br.ReadInt16()
            oheckseffeffeffeff = br.ReadInt16()
            For x = 0 To 5
                unknown(x) = br.ReadInt32()
            Next
            For x = 0 To 8
                zeros(x) = br.ReadInt32()
            Next
            one = br.ReadInt32()
            index_count = br.ReadInt32() + 2
            index_offset = br.ReadInt32()
            unknown1_offset = br.ReadInt32()
            unknown1_count = br.ReadInt32()
            vertex_count = br.ReadInt32()
            unknown2(1) = br.ReadInt32()
            unknown2(2) = br.ReadInt32()
            vertex_offset = br.ReadInt32()
            For x = 0 To 6 ' PC ONLY
                zeros2(x) = br.ReadInt32()
            Next
        End Sub
        Public Sub writestruct(ByRef bw As BinaryWriter)
            Dim x As Integer
            bw.Write(zero)
            bw.Write(number)
            bw.Write(oheckseffeffeffeff)
            For x = 0 To 5
                bw.Write(unknown(x))
            Next
            For x = 0 To 8
                bw.Write(zeros(x))
            Next
            bw.Write(one)
            bw.Write(index_count - 2)
            x = Val("&H" & Hex(index_offset))
            bw.Write(x)
            bw.Write(unknown1_offset)
            bw.Write(unknown1_count)
            bw.Write(vertex_count)
            bw.Write(unknown2(1))
            bw.Write(unknown2(2))
            x = Val("&H" & Hex(vertex_offset))
            bw.Write(x)
        End Sub
        Public Sub writestructpc(ByRef bw As BinaryWriter)
            Dim x As Integer
            bw.Write(zero)
            bw.Write(number)
            bw.Write(oheckseffeffeffeff)
            For x = 0 To 5
                bw.Write(unknown(x))
            Next
            For x = 0 To 8
                bw.Write(zeros(x))
            Next
            bw.Write(one)
            bw.Write(index_count - 2)
            x = Val("&H" & Hex(index_offset))
            bw.Write(x)
            bw.Write(unknown1_offset)
            bw.Write(unknown1_count)
            bw.Write(vertex_count)
            bw.Write(unknown2(1))
            bw.Write(unknown2(2))
            x = Val("&H" & Hex(vertex_offset))
            bw.Write(x)
            For x = 0 To 6 ' PC ONLY
                bw.Write(zeros2(x))
            Next
        End Sub
    End Structure

    Public Structure STRUCT_MODEL_COMPRESSED_VERT
        Public coord() As Single '3 floats
        Public unknown1() As Short '10 shorts
    End Structure

    Public Structure STRUCT_MODEL_UV
        Public u As Single
        Public v As Single
    End Structure

    Public Structure STRUCT_MODEL_SUBMESH_INFO
        Public header As STRUCT_MODEL_SUBMODEL_HEADER
        Public TriangleStrips() As STRUCT_MODEL_TRIANGLESTRIP_HEADER
        Public CmpVert() As STRUCT_MODEL_COMPRESSED_VERT
        Public TexVert() As STRUCT_MODEL_UV
        Public Index() As UInt16
    End Structure

    Public Structure STRUCT_MODEL_VERT_REFERENCE
        Public unknown() As UInt32 '3 dwords
    End Structure

#End Region

    Public Structure OffsetRange
        Public minVOffset As Integer
        Public maxVOffset As Integer
        Public minIOffset As Integer
        Public maxIOffset As Integer
    End Structure

#End Region

#Region "Public Variables"

Public modelHeader As STRUCT_MODEL_HEADER
Public modelBones As STRUCT_PC_MODEL_BONE
Public modelSubmeshInfo As STRUCT_MODEL_SUBMESH_INFO
Public m_IndexOffset As Integer
Public m_VertexOffset As Integer

#End Region


'////////////////////////////////////////////////////////////////////////
'// Read the metadata and save the raw data for the specified model
'////////////////////////////////////////////////////////////////////////
    Public Function SaveRawMod2(ByVal strPath As String, ByVal strFilename As String, _
        ByVal offset As Integer, ByRef txtWriter As XmlTextWriter) As OffsetRange

        'Vertex and Index list magic offsets
        Dim vOffset As Integer = indexHeader.vertex_offset
        Dim iOffset As Integer = indexHeader.indeces_offset + indexHeader.vertex_offset
        Dim strTag As String = Mid(strFilename, strFilename.Length - 3, 4)

        'Binary Vertex/Index files
        Dim bw As New BinaryWriter(New FileStream(strPath & strFilename & ".vertices.bin", FileMode.Create))
        Dim bw2 As New BinaryWriter(New FileStream(strPath & strFilename & ".indices.bin", FileMode.Create))

        txtWriter.WriteComment("Attached Binary File Structure")

        'Create a new reader on the map filestream
        Dim br As New BinaryReader(fs)
        br.BaseStream.Seek(offset, SeekOrigin.Begin)

        'Read the model header
        modelHeader.Read(br)

        Dim SubModel(modelHeader.SubModels.count - 1) As STRUCT_MODEL_SUBMODEL_HEADER

        'Since vb.net doesn't support jagged arrays, we will need to keep track of the max number of
        'triangle strips per sub-model, so that we can dimension the array appropriately.. what a
        'waste of memory... :/
        Dim maxTriangleStrips As Integer = 0
        Dim TriangleStrip(modelHeader.SubModels.count - 1, 0) As STRUCT_MODEL_TRIANGLESTRIP_HEADER

        'Go to the Submesh area of the meta
        Dim sOffset As Integer = Convert.ToInt64(UMath(HMTUtility.OperandEnum.Subtract, modelHeader.SubModels.offset, intMagic))
        sOffset += 36 '36 bytes of garbage data
        Dim metaStartOffset As Integer

        Dim minVertexOffset As Long = &HFFFFFFF
        Dim maxVertexOffset As Long = 0
        Dim minIndexOffset As Long = &HFFFFFFF
        Dim maxIndexOffset As Long = 0

        For x As Integer = 0 To modelHeader.SubModels.count - 1

            br.BaseStream.Seek(sOffset, SeekOrigin.Begin)
            Console.WriteLine(br.BaseStream.Position)
            SubModel(x).Read(br)
            sOffset = br.BaseStream.Position

            If SubModel(x).count > 0 Then

                'Go to the Submesh area of the meta
                br.BaseStream.Seek(Convert.ToInt64( _
                    UMath(HMTUtility.OperandEnum.Subtract, SubModel(x).offset, intMagic)), SeekOrigin.Begin)

                If SubModel(x).count > maxTriangleStrips Then maxTriangleStrips = SubModel(x).count
                ReDim Preserve TriangleStrip(modelHeader.SubModels.count - 1, maxTriangleStrips - 1)

                For y As Integer = 0 To SubModel(x).count - 1
                    TriangleStrip(x, y).ReadPC(br)
                Next

            End If
        Next

        'Get the min and max offsets of the vertex and index blocks
        maxIndexOffset = TriangleStrip(0, SubModel(0).count - 1).index_offset
        maxIndexOffset += (2 * (TriangleStrip(0, SubModel(0).count - 1).index_count))
        minIndexOffset = TriangleStrip(modelHeader.SubModels.count - 1, SubModel(modelHeader.SubModels.count - 1).count - 1).index_offset
        'For some reason, and extra indice is needed, so i'm not subtracting 1 from the index count.

        maxVertexOffset = TriangleStrip(0, SubModel(0).count - 1).vertex_offset
        maxVertexOffset += (68 * (TriangleStrip(0, SubModel(0).count - 1).vertex_count - 1))
        minVertexOffset = TriangleStrip(modelHeader.SubModels.count - 1, SubModel(modelHeader.SubModels.count - 1).count - 1).vertex_offset
        'There are 9 extra vertices for some reason... I'm just leaving the data for now.

        'Write the vertex Asset Block to the XML
        txtWriter.WriteStartElement("Asset")
        txtWriter.WriteElementString("Type", "VertexBlock")
        txtWriter.WriteElementString("Filename", strFilename & ".vertices")

        'We need to loop through the model data again to write the index translations to the XML
        For x As Integer = 0 To modelHeader.SubModels.count - 1
            For y As Integer = 0 To SubModel(x).count - 1
                metaStartOffset = TriangleStrip(x, y).StartingOffset - offset
                Console.WriteLine("Vertex Block Pointer found at " & metaStartOffset + 100)

                txtWriter.WriteStartElement("Correction")
                txtWriter.WriteElementString("Location", "0x" & Hex(metaStartOffset + 100))
                txtWriter.WriteElementString("Translation", "0x" & Hex(TriangleStrip(x, y).vertex_offset - minVertexOffset))
                txtWriter.WriteEndElement()
            Next
        Next
        txtWriter.WriteEndElement()

        'Write the index Asset Block to the XML
        txtWriter.WriteStartElement("Asset")
        txtWriter.WriteElementString("Type", "IndexBlock")
        txtWriter.WriteElementString("Filename", strFilename & ".indices")

        'We need to loop through the model data again to write the index translations to the XML
        For x As Integer = 0 To modelHeader.SubModels.count - 1
            For y As Integer = 0 To SubModel(x).count - 1
                metaStartOffset = TriangleStrip(x, y).StartingOffset - offset
                Console.WriteLine("Index Block Pointer found at " & metaStartOffset + 76)

                txtWriter.WriteStartElement("Correction")
                txtWriter.WriteElementString("Location", "0x" & Hex(metaStartOffset + 76))
                txtWriter.WriteElementString("Translation", "0x" & Hex(TriangleStrip(x, y).index_offset - minIndexOffset))
                txtWriter.WriteEndElement()
                txtWriter.WriteStartElement("Correction")
                txtWriter.WriteElementString("Location", "0x" & Hex(metaStartOffset + 80))
                txtWriter.WriteElementString("Translation", "0x" & Hex(TriangleStrip(x, y).index_offset - minIndexOffset))
                txtWriter.WriteEndElement()
            Next
        Next
        txtWriter.WriteEndElement()

        minVertexOffset += indexHeader.vertex_offset
        maxVertexOffset += indexHeader.vertex_offset
        minIndexOffset += indexHeader.vertex_offset + indexHeader.indeces_offset
        maxIndexOffset += indexHeader.vertex_offset + indexHeader.indeces_offset

        Dim bChunk() As Byte
        GetBytes(minVertexOffset, maxVertexOffset - minVertexOffset, bChunk)
        bw.Write(bChunk)
        GetBytes(minIndexOffset, maxIndexOffset - minIndexOffset, bChunk)
        bw2.Write(bChunk)

        bw.Close()
        bw2.Close()

        Dim oRange As New OffsetRange
        oRange.minVOffset = minVertexOffset
        oRange.maxVOffset = maxVertexOffset
        oRange.minIOffset = minIndexOffset
        oRange.maxIOffset = maxIndexOffset

        Return oRange
    End Function

End Class




