'///////////////////////////////////////////////////
'// Halo Map Tools                                //
'// [snd!] Plugin File                            //
'// (c)2003 MonoxideC                             //
'// --------------------------------------------- //
'// Classes Contained: 2                          //
'// [1] SndGUI  - User interface                  //
'// [2] SndMeta - Data processing code            //
'///////////////////////////////////////////////////

Imports System.IO
Imports HMTLib

Public Class SndGUI
    Inherits Windows.Forms.UserControl

    Public map As HaloMap.Map
    Public snd As SndMeta
    Public api As WINAPI
    Public trackBuffer() As Byte

    Public Const SND_ASYNC As Integer = 1
    Public Const SND_MEMORY As Integer = 4

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef m As HaloMap.Map, ByVal i As Integer)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        'Deflash the specified texture set to memory.
        SelectSound(m, i)
    End Sub

    'UserControl overrides dispose to clean up the component list.
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
    Friend WithEvents pSound As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblTrackInfo As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents tvTrackInfo As System.Windows.Forms.TreeView
    Friend WithEvents lblSndInfo As System.Windows.Forms.Label
    Friend WithEvents lblChunk As System.Windows.Forms.Label
    Friend WithEvents tb1 As System.Windows.Forms.TrackBar
    Friend WithEvents btnSaveTrack As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnLoadChunk As System.Windows.Forms.Button
    Friend WithEvents btnImportTrack As System.Windows.Forms.Button
    Friend WithEvents btnSaveChunk As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog As System.Windows.Forms.SaveFileDialog
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pSound = New System.Windows.Forms.Panel
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblTrackInfo = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.tvTrackInfo = New System.Windows.Forms.TreeView
        Me.lblSndInfo = New System.Windows.Forms.Label
        Me.lblChunk = New System.Windows.Forms.Label
        Me.tb1 = New System.Windows.Forms.TrackBar
        Me.btnSaveTrack = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnLoadChunk = New System.Windows.Forms.Button
        Me.btnImportTrack = New System.Windows.Forms.Button
        Me.btnSaveChunk = New System.Windows.Forms.Button
        Me.SaveFileDialog = New System.Windows.Forms.SaveFileDialog
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.pSound.SuspendLayout()
        CType(Me.tb1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pSound
        '
        Me.pSound.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pSound.BackColor = System.Drawing.Color.White
        Me.pSound.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pSound.Controls.Add(Me.Label3)
        Me.pSound.Controls.Add(Me.lblTrackInfo)
        Me.pSound.Controls.Add(Me.Label5)
        Me.pSound.Controls.Add(Me.tvTrackInfo)
        Me.pSound.Controls.Add(Me.lblSndInfo)
        Me.pSound.Controls.Add(Me.lblChunk)
        Me.pSound.Controls.Add(Me.tb1)
        Me.pSound.Controls.Add(Me.btnSaveTrack)
        Me.pSound.Controls.Add(Me.Label1)
        Me.pSound.Controls.Add(Me.Label2)
        Me.pSound.Controls.Add(Me.btnLoadChunk)
        Me.pSound.Controls.Add(Me.btnImportTrack)
        Me.pSound.Controls.Add(Me.btnSaveChunk)
        Me.pSound.Location = New System.Drawing.Point(0, 0)
        Me.pSound.Name = "pSound"
        Me.pSound.Size = New System.Drawing.Size(320, 440)
        Me.pSound.TabIndex = 28
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Firebrick
        Me.Label3.Location = New System.Drawing.Point(6, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(208, 14)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "Tracks contained in the selected sound structure:"
        '
        'lblTrackInfo
        '
        Me.lblTrackInfo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTrackInfo.BackColor = System.Drawing.Color.White
        Me.lblTrackInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTrackInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTrackInfo.Location = New System.Drawing.Point(6, 238)
        Me.lblTrackInfo.Name = "lblTrackInfo"
        Me.lblTrackInfo.Size = New System.Drawing.Size(304, 27)
        Me.lblTrackInfo.TabIndex = 19
        Me.lblTrackInfo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Firebrick
        Me.Label5.Location = New System.Drawing.Point(6, 222)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(120, 16)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Track Information"
        '
        'tvTrackInfo
        '
        Me.tvTrackInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tvTrackInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tvTrackInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tvTrackInfo.ImageIndex = -1
        Me.tvTrackInfo.Location = New System.Drawing.Point(6, 88)
        Me.tvTrackInfo.Name = "tvTrackInfo"
        Me.tvTrackInfo.SelectedImageIndex = -1
        Me.tvTrackInfo.Size = New System.Drawing.Size(304, 126)
        Me.tvTrackInfo.TabIndex = 22
        Me.tvTrackInfo.Text = "TreeViewEX"
        '
        'lblSndInfo
        '
        Me.lblSndInfo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSndInfo.BackColor = System.Drawing.Color.White
        Me.lblSndInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSndInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSndInfo.Location = New System.Drawing.Point(6, 18)
        Me.lblSndInfo.Name = "lblSndInfo"
        Me.lblSndInfo.Size = New System.Drawing.Size(304, 46)
        Me.lblSndInfo.TabIndex = 6
        Me.lblSndInfo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblChunk
        '
        Me.lblChunk.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblChunk.BackColor = System.Drawing.Color.White
        Me.lblChunk.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblChunk.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblChunk.Location = New System.Drawing.Point(6, 318)
        Me.lblChunk.Name = "lblChunk"
        Me.lblChunk.Size = New System.Drawing.Size(304, 35)
        Me.lblChunk.TabIndex = 8
        Me.lblChunk.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb1
        '
        Me.tb1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tb1.Location = New System.Drawing.Point(6, 390)
        Me.tb1.Maximum = 1
        Me.tb1.Minimum = 1
        Me.tb1.Name = "tb1"
        Me.tb1.Size = New System.Drawing.Size(304, 45)
        Me.tb1.TabIndex = 11
        Me.tb1.Value = 1
        '
        'btnSaveTrack
        '
        Me.btnSaveTrack.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSaveTrack.BackColor = System.Drawing.Color.Black
        Me.btnSaveTrack.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveTrack.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnSaveTrack.Location = New System.Drawing.Point(6, 270)
        Me.btnSaveTrack.Name = "btnSaveTrack"
        Me.btnSaveTrack.Size = New System.Drawing.Size(72, 24)
        Me.btnSaveTrack.TabIndex = 12
        Me.btnSaveTrack.Text = "Save Track"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Firebrick
        Me.Label1.Location = New System.Drawing.Point(6, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 16)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Sound Information"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Firebrick
        Me.Label2.Location = New System.Drawing.Point(6, 302)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(120, 16)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Chunk Information"
        '
        'btnLoadChunk
        '
        Me.btnLoadChunk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLoadChunk.BackColor = System.Drawing.Color.Black
        Me.btnLoadChunk.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLoadChunk.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnLoadChunk.Location = New System.Drawing.Point(238, 358)
        Me.btnLoadChunk.Name = "btnLoadChunk"
        Me.btnLoadChunk.Size = New System.Drawing.Size(72, 24)
        Me.btnLoadChunk.TabIndex = 15
        Me.btnLoadChunk.Text = "Import Chunk"
        '
        'btnImportTrack
        '
        Me.btnImportTrack.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnImportTrack.BackColor = System.Drawing.Color.Black
        Me.btnImportTrack.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnImportTrack.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnImportTrack.Location = New System.Drawing.Point(238, 270)
        Me.btnImportTrack.Name = "btnImportTrack"
        Me.btnImportTrack.Size = New System.Drawing.Size(72, 24)
        Me.btnImportTrack.TabIndex = 16
        Me.btnImportTrack.Text = "Import Track"
        '
        'btnSaveChunk
        '
        Me.btnSaveChunk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSaveChunk.BackColor = System.Drawing.Color.Black
        Me.btnSaveChunk.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveChunk.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnSaveChunk.Location = New System.Drawing.Point(6, 358)
        Me.btnSaveChunk.Name = "btnSaveChunk"
        Me.btnSaveChunk.Size = New System.Drawing.Size(72, 24)
        Me.btnSaveChunk.TabIndex = 17
        Me.btnSaveChunk.Text = "Save Chunk"
        '
        'SndGUI
        '
        Me.Controls.Add(Me.pSound)
        Me.Name = "SndGUI"
        Me.Size = New System.Drawing.Size(320, 440)
        Me.pSound.ResumeLayout(False)
        CType(Me.tb1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub Unload()
        snd.releasefile()
    End Sub

#Region "Event Handlers"
    'The chunk selection TickerBar was scrolled by the user.
    Private Sub tb1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb1.Scroll
        UpdateSndChunkInfo()
    End Sub

    'The user clicked the 'Save Track' button.
    Private Sub btnSaveCurrentFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveTrack.Click
        SaveTrack()
    End Sub

    'Save the current chunk to an ADPCM file
    Private Sub btnSaveChunk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveChunk.Click
        Dim strFilename As String

        'Get the correct filename.
        SaveFileDialog.AddExtension = True
        SaveFileDialog.DefaultExt = "*.wav"
        If snd.Struct.sndHeader.format <> 3 Then
            SaveFileDialog.Filter = "ADPCM Wave (*.wav)|*.wav"
        Else
            SaveFileDialog.Filter = "OGG Media File (*.ogg)|*.ogg"
        End If
        SaveFileDialog.FileName = getLevel(snd.fileName, 0) & _
                "." & "chunk" & tb1.Value & ".wav"
        If SaveFileDialog.ShowDialog() = DialogResult.Cancel Then Exit Sub
        strFilename = SaveFileDialog.FileName

        'Save the chunks.
        Dim strMsg = (snd.SaveChunks(0, tb1.Value, tb1.Value, strFilename))
        If strMsg <> "" Then MsgBox(strMsg)
    End Sub

    '///////////////////////////////////////////////////////////////////
    '// Replace the current chunk with an ADPCM file
    '///////////////////////////////////////////////////////////////////
    Private Sub btnLoadChunk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadChunk.Click
        Dim strFilename As String

        OpenFileDialog.AddExtension = True
        OpenFileDialog.FileName = ""

        OpenFileDialog.AddExtension = True
        OpenFileDialog.DefaultExt = "*.ogg"
        If snd.Struct.sndHeader.format = 3 Then 'Ogg
            OpenFileDialog.Filter = "OGG Media (*.ogg)|*.ogg"
        Else
            OpenFileDialog.Filter = "ADPCM Wave (*.wav)|*.wav"
        End If
        If OpenFileDialog.ShowDialog() = DialogResult.Cancel Then Exit Sub
        strFilename = OpenFileDialog.FileName
        Select Case snd.Struct.sndHeader.format
            Case 3 'Ogg
                MsgBox(snd.ImportSoundFileIntoTrack(0, map.fs, strFilename, SndMeta.SoundType.Ogg, tb1.Value))
            Case Else
                MsgBox(snd.ImportSoundFileIntoTrack(0, map.fs, strFilename, SndMeta.SoundType.ADPCM, tb1.Value))
        End Select
    End Sub

    '///////////////////////////////////////////////////////////////////
    '// Replace the current track with a new file
    '///////////////////////////////////////////////////////////////////
    Private Sub btnImportTrack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImportTrack.Click
        Dim strFilename As String
        Dim fi As FileInfo

        OpenFileDialog.AddExtension = True
        OpenFileDialog.DefaultExt = "*.wav"
        If snd.Struct.sndHeader.format = 3 Then 'Ogg
            OpenFileDialog.Filter = "OGG Media (*.ogg)|*.ogg"
        Else
            OpenFileDialog.Filter = "ADPCM Wave (*.wav)|*.wav"
        End If
        OpenFileDialog.FileName = ""
        If OpenFileDialog.ShowDialog() = DialogResult.Cancel Then Exit Sub

        'Open the source file.
        strFilename = OpenFileDialog.FileName
        fi = New FileInfo(strFilename)
        If fi.Length > snd.Struct.size Then _
            If MsgBox("The sound you are trying to import is larger than the existing" & vbCrLf & _
                "data and will be truncated to fit in the allotted space.  Proceed?", _
                MsgBoxStyle.OKCancel) = MsgBoxResult.Cancel Then Exit Sub

        'Invoke the ImportADPCMFileIntoTrack method of the snd object.
        Select Case snd.Struct.sndHeader.format
            Case 3 'Ogg
                MsgBox(snd.ImportSoundFileIntoTrack(tvTrackInfo.SelectedNode.Index + 1, snd.fs, strFilename, SndMeta.SoundType.Ogg))
            Case Else
                If map.pc Then
                    MsgBox(snd.ImportSoundFileIntoTrack(tvTrackInfo.SelectedNode.Index + 1, snd.fs, strFilename, SndMeta.SoundType.ADPCM))
                Else
                    MsgBox(snd.ImportSoundFileIntoTrack(tvTrackInfo.SelectedNode.Index + 1, map.fs, strFilename, SndMeta.SoundType.ADPCM))
                End If
        End Select
    End Sub
#End Region

#Region "Utility Function"
    '///////////////////////////////////////////////////////////////////
    '// Change the current sound
    '///////////////////////////////////////////////////////////////////
    Public Sub SelectSound(ByRef m As HaloMap.Map, ByVal intID As Integer)

        'Need to fix this so that it doesnt depend on a form.
        'Probably should move status into the Map class.
        'fMain.bReady = False
        'If fMain.bChangingFiles Then Exit Sub

        'If we already have a bitmap open, we need to destroy it so that the bitmaps file is no longer in use
        If Not snd Is Nothing Then
            snd.ReleaseFile()
        End If
        map = m

        'Make sure we have a map loaded.
        If map.indexItem Is Nothing Then Exit Sub

        If map.pc Then
            Dim strFilename As String = map.fs.Name
            strFilename = Mid(strFilename, 1, Len(strFilename) - Len(getLevel(strFilename, 0))) & "sounds.map"
            snd = New SndMeta(map, intID, strFilename)
        Else
            snd = New SndMeta(map, intID)
        End If

        'Build the string to be displayed in the sound info box.
        lblSndInfo.Text = "Sound Name: " & getLevel(snd.fileName, 0) & vbCrLf & vbCrLf & _
            IIf((snd.Struct.sndHeader.channels + 1) = 1, "Mono", "Stereo") & " / " & _
            snd.Struct.sampleRate & "Hz" & IIf(snd.Struct.sndHeader.format = 3, " - Ogg Vorbis", " - ADPCM") & vbCrLf & _
            IIf(snd.Struct.sndHeader.format = 3, "Lengh not calculated (", _
                snd.CalculateADPCMLength(snd.Struct.size, _
                snd.Struct.sampleRate, snd.Struct.sndHeader.channels + 1) & " seconds (") & _
            snd.Struct.size & " bytes) in " & snd.Struct.sndHeader.track_count & " track(s)."

        tb1.Maximum = snd.Struct.sndHeader.chunk_count
        tb1.Value = 1
        UpdateSndTrackList()
        tvTrackInfo.SelectedNode = tvTrackInfo.Nodes.Item(0)
        snd_changeTrack(1)
    End Sub

    '///////////////////////////////////////////////////////////////////
    '// User changed to a different chunk
    '///////////////////////////////////////////////////////////////////
    Private Sub UpdateSndChunkInfo()

        lblChunk.Text = snd.Struct.sndChunk(tb1.Value).name.Trim(Chr(0)) & vbCrLf & _
            "Chunk " & tb1.Value & " of " & tb1.Maximum & vbCrLf & _
            IIf(snd.Struct.sndHeader.format = 3, "Length not calculated (", _
            "Length: " & snd.CalculateADPCMLength(snd.Struct.sndChunk(tb1.Value).size, _
            snd.Struct.sampleRate, snd.Struct.sndHeader.channels + 1) & _
            " seconds (") & snd.Struct.sndChunk(tb1.Value).size & " bytes)"
    End Sub

    '///////////////////////////////////////////////////////////////////
    '// Update the tracklist
    '///////////////////////////////////////////////////////////////////
    Private Sub UpdateSndTrackList()
        Dim x As Integer

        tvTrackInfo.Nodes.Clear()

        For x = 1 To snd.Struct.sndHeader.track_count
            tvTrackInfo.Nodes.Add(snd.Struct.sndChunk(x).name.Trim(Chr(0)))
        Next
    End Sub

    Private Sub snd_changeTrack(ByVal i As Integer)
        'Update the Track Information
        lblTrackInfo.Text = snd.Struct.sndTrack(i).name & vbCrLf & _
                IIf(snd.Struct.sndHeader.format = 3, "Length not calculated (", _
                snd.CalculateADPCMLength(snd.Struct.sndTrack(i).size, _
                snd.Struct.sampleRate, snd.Struct.sndHeader.channels + 1) & " seconds (") & _
                snd.Struct.sndTrack(i).size & " bytes) in " & _
                snd.Struct.sndTrack(i).numChunks & " chunk(s)."
        UpdateSndChunkInfo()
    End Sub
#End Region

#Region "Save/Import Procedures (Most are in their event handlers)"

    '///////////////////////////////////////////////////////////////////
    '// Save the currently selected track
    '///////////////////////////////////////////////////////////////////
    Private Sub SaveTrack(Optional ByVal strFilename As String = "")
        Dim z, i As Integer 'Loop counters
        Dim strTrackName As String
        Dim f1 As Integer
        Dim br As BinaryReader
        Dim bw As BinaryWriter
        Dim bChunk() As Byte
        Dim fi As FileInfo
        Dim mbResult As MsgBoxResult

        'Browse for the file.
        If strFilename = "" Then
            SaveFileDialog.AddExtension = True
            SaveFileDialog.DefaultExt = "*.wav"
            If snd.Struct.sndHeader.format <> 3 Then
                SaveFileDialog.Filter = "ADPCM Wave (*.wav)|*.wav"
            Else
                SaveFileDialog.Filter = "OGG Media File (*.ogg)|*.ogg"
            End If
            SaveFileDialog.FileName = getLevel(snd.fileName, 0) & _
                    "." & snd.Struct.sndTrack(tvTrackInfo.SelectedNode.Index + 1).name
            If SaveFileDialog.ShowDialog = DialogResult.Cancel Then Exit Sub
            strFilename = SaveFileDialog.FileName
        End If
            If strFilename = "" Then Exit Sub 'The user didnt select a file.

            'Call the SaveChunks method of the SndMeta Class
            Dim strMsg = (snd.SaveChunks(tvTrackInfo.SelectedNode.Index + 1, 1, _
                    snd.Struct.sndTrack(tvTrackInfo.SelectedNode.Index + 1).numChunks, strFilename))
            If strMsg <> "" Then MsgBox(strMsg)
    End Sub

#End Region

    'Private Sub btnPlay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlay.Click
    'Dim len As Integer
    'Dim count As Integer
    'Dim totalLength As Integer
    '
    'For x As Integer = 1 To snd.Struct.sndTrack(1).numChunks
    'totalLength += snd.cb(x).bin.Length
    'ReDim Preserve trackBuffer(totalLength)
    'For z As Integer = 1 To snd.cb(x).bin.Length - 1
    'trackBuffer(count) = snd.cb(x).bin(z - 1)
    'count += 1
    'Next
    'Next
    'api.PlaySound(trackBuffer, IntPtr.Zero, SND_ASYNC Or SND_MEMORY)
    'End Sub

    Private Sub tvTrackInfo_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvTrackInfo.AfterSelect
        snd_changeTrack(e.Node.Index + 1)
    End Sub
End Class

Public Class SndMeta

#Region "Structures"

    '///////////////////////////////////
    '// Header structure
    '///////////////////////////////////
    Public Structure SND_HEADER_STRUCT
        Public unknown1() As Integer                        ' Unknown - Length is 27 ints
        'Unknown1(1) and (2) seem to represent some type of collection - so far, the following holds true:
        'if Unknown1(1) is <= 1 then the sound is 22050 .. >= 2 then it is 44100
        Public channels As Short                            ' Number of channels
        Public format As Short                           ' Sound Format - so far 0001 = adpcm and 0003 = ogg
        Public id As Integer                                ' !dns
        Public unknown2() As Integer                        ' Unknown - Length is 10 ints
        Public default_ptr As Integer                       ' Default Pointer
        Public unknown3 As Integer                          ' 0x00000000
        Public default_txt As String                        ' "default" - Length is 32 chars.
        Public unknown4() As String                         ' Unknown - Length is 3 ints
        Public track_count As Integer                       ' Number of tracks in the sequence.
        Public unknown4b() As String                        ' Unknown - Length is 3 ints
        Public chunk_count As Integer                       ' The number of tracks.
        Public track_ptr As Integer                         ' Track Pointer (Needs to be magicked)
        Public unknown6 As Integer                          ' 0x00000000
        Public Sub ReadStruct(ByRef bs As BinaryReader)
            Dim x As Integer 'Loop counter
            ReDim unknown1(27)
            For x = 1 To 27
                unknown1(x) = bs.ReadInt32
            Next
            channels = bs.ReadInt16
            format = bs.ReadInt16
            id = bs.ReadInt32
            ReDim unknown2(10)
            For x = 1 To 10
                unknown2(x) = bs.ReadInt32
            Next
            default_ptr = bs.ReadInt32
            unknown3 = bs.ReadInt32
            default_txt = bs.ReadChars(32)
            ReDim unknown4(3)
            For x = 1 To 3
                unknown4(x) = bs.ReadInt32
            Next
            track_count = bs.ReadInt32
            ReDim unknown4b(3)
            For x = 1 To 3
                unknown4b(x) = bs.ReadInt32
            Next
            chunk_count = bs.ReadInt32
            track_ptr = bs.ReadInt32
            unknown6 = bs.ReadInt32
        End Sub
        Public Sub WriteStruct(ByVal bw As BinaryWriter, ByVal o As Long)
            Dim x As Integer 'Loop counter
            bw.Seek(o, SeekOrigin.Begin)
            ReDim unknown1(27)
            For x = 1 To 27
                bw.Write(unknown1(x))
            Next
            bw.Write(channels)
            bw.Write(format)
            bw.Write(id)
            ReDim unknown2(10)
            For x = 1 To 10
                bw.Write(unknown2(x))
            Next
            bw.Write(default_ptr)
            bw.Write(unknown3)
            bw.Write(default_txt)
            ReDim unknown4(7)
            For x = 1 To 7
                bw.Write(unknown4(x))
            Next
            bw.Write(chunk_count)
            bw.Write(track_ptr)
            bw.Write(unknown6)
        End Sub
    End Structure

    '///////////////////////////////////
    '// Chunk Structure
    '///////////////////////////////////
    Public Structure SND_CHUNK_STRUCT
        Public name As String                                   ' Name of the Sound - Length is 32 chars.
        Public unknown1a As Integer                             ' 0x00000000
        Public unknown1b As Integer                             ' 0x0000803F
        Public unknown2 As Short                                ' 0x0100
        Public chunk_id As Short                                ' Increments with each portion
        ' Last of each track is 0xFFFF - Starts at track_count
        Private unknown3() As Integer                           ' Unknown - Length is 5 ints.
        Public size As Integer                                  ' Size of chunk.
        Public unknown4 As Integer                              ' 0x00000000
        Public offset As Integer                                ' Location of the chunk - not magicked.
        Private unknown5() As Integer                           ' 0x00000000 - Length is 4 ints
        Public size_dealy_a As Integer                          ' Equals previous chunk dealy + current track size.
        Public unknown6() As Integer                            ' Unknown - Length is 4 ints.
        Public size_dealy_b As Integer                          ' Same as size_dealy_a
        Public unknown7 As Integer                             ' 0x00000000
        Public unknown8 As Integer                             ' 0x00000000
        Public Sub ReadStruct(ByRef bs As BinaryReader)
            Dim x As Integer 'Loop counter
            name = bs.ReadChars(32)
            unknown1a = bs.ReadInt32
            unknown1b = bs.ReadInt32
            unknown2 = bs.ReadInt16
            chunk_id = bs.ReadInt16
            ReDim unknown3(5)
            For x = 1 To 5
                unknown3(x) = bs.ReadInt32
            Next
            size = bs.ReadInt32
            unknown4 = bs.ReadInt32
            offset = bs.ReadInt32
            ReDim unknown5(4)
            For x = 1 To 4
                unknown5(x) = bs.ReadInt32
            Next
            size_dealy_a = bs.ReadInt32
            ReDim unknown6(4)
            For x = 1 To 4
                unknown6(x) = bs.ReadInt32
            Next
            size_dealy_b = bs.ReadInt32
            unknown7 = bs.ReadInt32
            unknown8 = bs.ReadInt32
        End Sub
        Public Sub WriteStruct(ByRef bw As BinaryWriter, ByVal o As Integer)
            Dim x As Integer 'Loop counter
            bw.Write(name)
            bw.Write(unknown1a)
            bw.Write(unknown1b)
            bw.Write(unknown2)
            bw.Write(chunk_id)
            ReDim unknown3(5)
            For x = 1 To 5
                bw.Write(unknown3(x))
            Next
            bw.Write(size)
            bw.Write(unknown4)
            bw.Write(offset)
            ReDim unknown5(4)
            For x = 1 To 4
                bw.Write(unknown5(x))
            Next
            bw.Write(size_dealy_a)
            ReDim unknown6(4)
            For x = 1 To 4
                bw.Write(unknown6(x))
            Next
            bw.Write(size_dealy_b)
            bw.Write(unknown7)
            bw.Write(unknown8)
        End Sub
    End Structure

    '///////////////////////////////////
    '// Track Structure
    '///////////////////////////////////
    Public Structure SND_TRACK_STRUCT
        Public name As String
        Public size As Integer
        Public numChunks As Integer
        Public chunk() As Integer 'Chunk index list
    End Structure

    '///////////////////////////////////
    '// Encapsulates other structres for
    '// easier access.
    '///////////////////////////////////
    Public Structure SND_ENCAPSULATOR_STRUCT
        Public sndHeader As SND_HEADER_STRUCT
        Public sndChunk() As SND_CHUNK_STRUCT
        Public sndTrack() As SND_TRACK_STRUCT
        Public size As Integer
        Public sampleRate As Integer
    End Structure

    '///////////////////////////////////
    '// XBox ADPCM File Header Structure
    '///////////////////////////////////
    Public Structure ADPCM_FILE_HEADER
        <VBFixedString(4)> Public RIFFHeader As String ' "RIFF"
        Public intFileSize As Int32
        <VBFixedString(4)> Public WAVEHeader As String
        <VBFixedString(4)> Public fmtHeader As String
        Public intWaveSectionChunkSize As Int32  'The size of the WAVE type format (2 bytes) + mono/stereo flag
        '(2 bytes) + sample rate (4 bytes) + bytes/sec (4 bytes) + block alignment (2 bytes) + 
        'bits/sample (2 bytes) + ExtraParameter Data (4bytes) = 20 bytes.
        'This is usually 16 (or 0x10).
        Public shortWaveTypeFormat As Short 'ADPCM Type
        Public shortNumChannels As Short 'Mono (0x01) or stereo (0x02)
        Public intSampleRate As Int32 'Sample Rate
        Public intByteRate As Int32 ' Number of Channels * SampleRate * bits per sample / 8
        Public shortBlockAlignment As Short 'Block alignment
        Public shortBitsPerSample As Short  'Bits/Sample
        Public shortExtraParamSize As Short 'Specifies the size of the extra parameters (only in non-pcm)
        Public shortExtraParam_BlockSize As Short 'Extra Parameter for ADPCM only.
        <VBFixedString(4)> Public DATAHeader As String
        Public intSizeOfDataChunk As Int32 'Number of bytes of data is included in the data section. 
        Public Sub readStruct(ByRef br As BinaryReader)
            RIFFHeader = br.ReadChars(4)
            intFileSize = br.ReadInt32
            WAVEHeader = br.ReadChars(4)
            fmtHeader = br.ReadChars(4)
            intWaveSectionChunkSize = br.ReadInt32
            shortWaveTypeFormat = br.ReadInt16
            shortNumChannels = br.ReadInt16
            intSampleRate = br.ReadInt32
            intByteRate = br.ReadInt32
            shortBlockAlignment = br.ReadInt16
            shortBitsPerSample = br.ReadInt16
            shortExtraParamSize = br.ReadInt16
            shortExtraParam_BlockSize = br.ReadInt16
            DATAHeader = br.ReadChars(4)
            intSizeOfDataChunk = br.ReadInt32
        End Sub
        Public Sub writeStruct(ByRef bw As BinaryWriter)
            Dim x As Integer 'Loop counter
            For x = 1 To RIFFHeader.Length
                bw.BaseStream.WriteByte(Asc(Mid(RIFFHeader, x, 1)))
            Next x
            bw.Write(intFileSize)
            For x = 1 To WAVEHeader.Length
                bw.BaseStream.WriteByte(Asc(Mid(WAVEHeader, x, 1)))
            Next x
            For x = 1 To fmtHeader.Length
                bw.BaseStream.WriteByte(Asc(Mid(fmtHeader, x, 1)))
            Next x
            bw.Write(intWaveSectionChunkSize)
            bw.Write(shortWaveTypeFormat)
            bw.Write(shortNumChannels)
            bw.Write(intSampleRate)
            bw.Write(intByteRate)
            bw.Write(shortBlockAlignment)
            bw.Write(shortBitsPerSample)
            bw.Write(shortExtraParamSize)
            bw.Write(shortExtraParam_BlockSize)
            For x = 1 To DATAHeader.Length
                bw.BaseStream.WriteByte(Asc(Mid(DATAHeader, x, 1)))
            Next x
            bw.Write(intSizeOfDataChunk)
        End Sub
        Public Sub GenerateStruct(ByVal intDataLength, ByVal intNumChannels, _
                ByVal iSampleRate)
            '///////////////////////////////////////////////////
            '// Generate an ADPCM Header based on key variables
            '///////////////////////////////////////////////////
            RIFFHeader = "RIFF"
            WAVEHeader = "WAVE"
            fmtHeader = "fmt "
            intWaveSectionChunkSize = 20
            shortWaveTypeFormat = 105
            shortBitsPerSample = 4
            shortExtraParamSize = 2
            shortExtraParam_BlockSize = 64
            DATAHeader = "data"
            intSizeOfDataChunk = intDataLength - 4
            intFileSize = intSizeOfDataChunk + 44 '48 byte header -4 bytes
            shortNumChannels = intNumChannels
            intSampleRate = iSampleRate
            intByteRate = shortNumChannels * intSampleRate * 4 / 8
            shortBlockAlignment = 36 * shortNumChannels
        End Sub
    End Structure

    '///////////////////////////////////
    '// Memory storage for sound chunks
    '///////////////////////////////////
    Public Structure CHUNK_BUFFER_STRUCT
        Public bin() As Byte
    End Structure

#End Region

#Region "Public Data Members"

    Public Struct As SND_ENCAPSULATOR_STRUCT
    Public cb() As CHUNK_BUFFER_STRUCT
    Public fileName As String
    Public fs As FileStream

#End Region

#Region "Functions"

    '//////////////////////////////////////////////////////
    '// Class Constructor
    '//////////////////////////////////////////////////////
    Public Sub New(ByRef map As HaloMap.Map, ByVal i As Integer, Optional ByVal strFilename As String = "")
        'Deflash the specified texture set to memory.
        If strFilename <> "" Then fs = New FileStream(strFilename, FileMode.Open, FileAccess.ReadWrite)
        DecodeSndMetadata(map, i)
    End Sub

    '//////////////////////////////////////////////////////////////////////
    '// Loops through item Index and decodes all snd! metadata structures
    '// Data is stored in memory to be consumable by other classes.
    '//////////////////////////////////////////////////////////////////////
    Public Function DecodeSndMetadata(ByRef map As HaloMap.Map, ByVal intID As Integer) As Boolean

        Dim strTemp As String
        Dim br1 As New BinaryReader(map.fs)

        Try
            'Seek to the metadata offset and read the Header.
            br1.BaseStream.Seek(CInt(Unsigned(map.indexItem(intID).magic_metadata_offset)), SeekOrigin.Begin)
            Struct.sndHeader.ReadStruct(br1)
            fileName = map.indexItem(intID).filePath
            'Calculate the sample rate.
            If Struct.sndHeader.unknown1(1) <= 1 Then
                Struct.sampleRate = 22050
            Else
                Struct.sampleRate = 44100
            End If
            ReDim Struct.sndChunk(Struct.sndHeader.chunk_count) 'Reserve memory for all of the chunk metadata
            'Go to the offset of the first chunk.
            map.fs.Seek((Unsigned(Struct.sndHeader.track_ptr) - map.intMagic), SeekOrigin.Begin)

            For y As Integer = 1 To Struct.sndHeader.chunk_count 'Loop through all of the chunks and get the binary data.
                Struct.sndChunk(y).ReadStruct(br1) 'Read in the metadata chunk.
                Struct.size += Struct.sndChunk(y).size 'Calculate the total size of the sound data chunks as we go.
            Next

            'Calculate track data
            ReDim Struct.sndTrack(Struct.sndHeader.track_count)
            For z As Integer = 1 To Struct.sndHeader.track_count
                strTemp = Struct.sndChunk(z).name
                Struct.sndTrack(z).name = strTemp.Trim(Chr(0))
                For y As Integer = z To Struct.sndHeader.chunk_count
                    If Struct.sndChunk(y).name = strTemp Then
                        Struct.sndTrack(z).size += Struct.sndChunk(y).size
                        Struct.sndTrack(z).numChunks += 1
                        ReDim Preserve Struct.sndTrack(z).chunk(Struct.sndTrack(z).numChunks)
                        Struct.sndTrack(z).chunk(Struct.sndTrack(z).numChunks) = y
                    End If
                Next
            Next
            'Deflash all the chunks
            ReDim cb(Struct.sndHeader.chunk_count)

            If map.pc Then
                'PC version - need to get images from the external file.
                'TODO: add some exception handling here.
                Dim br As New BinaryReader(fs)
                For x As Integer = 1 To Struct.sndHeader.chunk_count
                    ReDim cb(x).bin(Struct.sndChunk(x).size)
                    br.BaseStream.Seek(Struct.sndChunk(x).offset, SeekOrigin.Begin)
                    cb(x).bin = br.ReadBytes(Struct.sndChunk(x).size)
                Next x
            Else
                For x As Integer = 1 To Struct.sndHeader.chunk_count
                    map.GetBytes(Struct.sndChunk(x).offset, Struct.sndChunk(x).size, cb(x).bin)
                Next x
            End If
        Catch
            MsgBox("There was an error processing the map file.")
            DecodeSndMetadata = False
        End Try
    End Function

    '//////////////////////////////////////////////////////////////////////
    '// Calculate the length in seconds of an ADPCM encoded sound.
    '//////////////////////////////////////////////////////////////////////
    Public Function CalculateADPCMLength(ByVal intLength As Integer, _
            ByVal intSampleRate As Integer, ByVal intChannels As Integer) As Single
        CalculateADPCMLength = Int(((intLength / ((intSampleRate * 4) / 8)) / intChannels) * 100) / 100
    End Function

    '//////////////////////////////////////////////////////////////////////
    '// Split a source ADPCM file into chunks that can be imported into
    '// the map file via the ImportSoundChunk Method.
    '// s = sound index, t = track index, fs = valid Halo Map filestream
    '// with write access.
    '//////////////////////////////////////////////////////////////////////
    Public Enum SoundType
        Ogg = 1
        ADPCM = 2
    End Enum

    Public Function ImportSoundFileIntoTrack(ByVal t As Integer, _
                ByVal mapFileStream As FileStream, ByVal strFilename As String, ByVal intType As SoundType, _
                Optional ByVal intChunkNumber As Short = 0) As String

        Dim br As BinaryReader
        Dim bw As BinaryWriter
        Dim ad As New SndMeta.ADPCM_FILE_HEADER
        Dim bChunk() As Byte
        Dim bCompatible As Boolean = True
        Dim strErrorMsg As String

        br = New BinaryReader(New FileStream(strFilename, FileMode.Open, FileAccess.Read)) 'Open the file

        If intType = SoundType.ADPCM Then
            'Read in the ADPCM header and ensure that it is valid.
            ad.readStruct(br)
            If (Not ad.DATAHeader = "data") Or (Not ad.shortWaveTypeFormat = 105) Then
                ImportSoundFileIntoTrack = "Invalid Header: " & vbCrLf & _
                    "You must choose a valid ADPCM encoded wav file."
                Exit Function
            End If
            'Check the structure to see if it matches the current sound.
            If Struct.sndHeader.channels + 1 <> ad.shortNumChannels Then
                bCompatible = False
                strErrorMsg &= "Incorrect number of channels." & vbCrLf
            End If
            If Struct.sampleRate <> ad.intSampleRate Then
                bCompatible = False
                strErrorMsg &= "Sample rates do not match." & vbCrLf
            End If
        Else
            If MsgBox("Ogg file validation has not been implemented in this release.  Make" & vbCrLf & _
                   "sure that the file you are injecting is the same format as the original.", MsgBoxStyle.OKCancel) _
                   = MsgBoxResult.Cancel Then
                Return "File import was canceled."
            End If
            bCompatible = True
        End If

        If bCompatible Then
            Try
                'Apply the datawriter
                bw = New BinaryWriter(mapFileStream)

                Dim z As Integer ' Loop counter
                Dim v As Integer ' Looked up chunk number
                Dim intStartingChunk, intEndingChunk As Integer

                'Setup the values for the loop based on whether we are importing a file
                'over an entire track or just a chunk.
                If intChunkNumber = 0 Then
                    intStartingChunk = 1
                    intEndingChunk = Struct.sndTrack(t).numChunks
                Else
                    intStartingChunk = intChunkNumber
                    intEndingChunk = intChunkNumber
                End If

                For z = intStartingChunk To intEndingChunk
                    v = z
                    If t > 0 Then v = Struct.sndTrack(t).chunk(z)
                    ReDim bChunk(Struct.sndChunk(v).size)
                    'Grab a sized chunk of the file being split
                    bChunk = br.ReadBytes(Struct.sndChunk(v).size)

                    'Inject the chunk into the appropriate map offset
                    bw.Seek(Struct.sndChunk(v).offset, SeekOrigin.Begin)
                    bw.Write(bChunk)

                    'Replace the appropriate chunk buffer
                    ReDim cb(v).bin(Struct.sndChunk(v).size)
                    cb(v).bin = bChunk
                Next
                ImportSoundFileIntoTrack = "Track was successfully imported."
            Catch ex As Exception
                ImportSoundFileIntoTrack = "An error occured while importing data: " & vbCrLf & _
                        ex.Message
            End Try
            Try
                br.Close()
            Catch
            End Try
        Else 'Structures were not compatible and the sounds was not processed.
            ImportSoundFileIntoTrack = "Could not import the file for the following reasons: " & strErrorMsg
        End If
    End Function

    '//////////////////////////////////////////////////////////////////////
    '// Save the specified block of chunks to a new ADPCM file
    '//////////////////////////////////////////////////////////////////////
    Public Function SaveChunks(ByVal intTrackNumber As Integer, _
            ByVal chunkFirst As Integer, ByVal chunkLast As Integer, _
            ByVal strFilename As String) As String

        Dim bw As New BinaryWriter(New FileStream(strFilename, FileMode.Create, FileAccess.Write))
        Dim z As Integer 'Loop counter
        Dim ad As New ADPCM_FILE_HEADER
        Dim intLength As Integer

        'We must calculate the total size of the chunks before we can write the header.
        If intTrackNumber > 0 Then
            For z = chunkFirst To chunkLast
                intLength += Struct.sndChunk(Struct.sndTrack(intTrackNumber).chunk(z)).size
            Next
        Else
            intLength += cb(chunkFirst).bin.Length()
        End If

        'Write the ADPCM header
        If Struct.sndHeader.format <> 3 Then
            ad.GenerateStruct(intLength, Struct.sndHeader.channels + 1, _
                    Struct.sampleRate)
            ad.writeStruct(bw)
        End If

        Try
            'Write (sequentially) all of the appropriate chunks to the new file.
            If intTrackNumber > 0 Then
                For z = chunkFirst To chunkLast
                    bw.Write(cb(Struct.sndTrack(intTrackNumber).chunk(z)).bin)
                Next
            Else
                bw.Write(cb(chunkFirst).bin)
            End If
            SaveChunks = ""
        Catch ex As Exception
            SaveChunks = "There was an error saving the file:" & vbCrLf & _
            ex.Message
        End Try
        bw.Close()
    End Function

    Public Sub ReleaseFile()
        Try
            fs.Close()
        Catch
        End Try
    End Sub

#End Region

End Class

