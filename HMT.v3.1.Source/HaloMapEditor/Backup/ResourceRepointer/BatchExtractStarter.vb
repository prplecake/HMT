Imports HMTLib
Imports BitmLib
Imports sndLib
Imports HaloMap
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading

Public Class BatchExtractStarter
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef m As HaloMap.Map)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        map = m

    End Sub

    'Form overrides dispose to clean up the component list.
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
    Friend WithEvents XPander1 As XPanderControl.XPander
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents pb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents XPanderList1 As XPanderControl.XPanderList
    Friend WithEvents cbMetadata As System.Windows.Forms.CheckBox
    Friend WithEvents cbTextures As System.Windows.Forms.CheckBox
    Friend WithEvents cbModels As System.Windows.Forms.CheckBox
    Friend WithEvents cbSounds As System.Windows.Forms.CheckBox
    Friend WithEvents OpenFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents FolderBrowserDialog As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents txtBatchPath As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(BatchExtractStarter))
        Me.XPander1 = New XPanderControl.XPander
        Me.pb1 = New System.Windows.Forms.ProgressBar
        Me.cbMetadata = New System.Windows.Forms.CheckBox
        Me.cbTextures = New System.Windows.Forms.CheckBox
        Me.cbModels = New System.Windows.Forms.CheckBox
        Me.cbSounds = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnBrowse = New System.Windows.Forms.Button
        Me.txtBatchPath = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.XPanderList1 = New XPanderControl.XPanderList
        Me.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog
        Me.OpenFile = New System.Windows.Forms.OpenFileDialog
        Me.XPander1.SuspendLayout()
        Me.XPanderList1.SuspendLayout()
        Me.SuspendLayout()
        '
        'XPander1
        '
        Me.XPander1.Anchored = False
        Me.XPander1.Animated = False
        Me.XPander1.AnimationTime = 100
        Me.XPander1.BackColor = System.Drawing.Color.Transparent
        Me.XPander1.BorderStyle = System.Windows.Forms.Border3DStyle.Flat
        Me.XPander1.CanToggle = False
        Me.XPander1.CaptionFont = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold)
        Me.XPander1.CaptionStyle = XPanderControl.XPander.CaptionStyleEnum.Normal
        Me.XPander1.CaptionText = "Batch Extract"
        Me.XPander1.ChevronStyle = XPanderControl.XPander.ChevronStyleEnum.Image
        Me.XPander1.CollapsedHighlightImage = CType(resources.GetObject("XPander1.CollapsedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander1.CollapsedImage = CType(resources.GetObject("XPander1.CollapsedImage"), System.Drawing.Bitmap)
        Me.XPander1.Controls.Add(Me.pb1)
        Me.XPander1.Controls.Add(Me.cbMetadata)
        Me.XPander1.Controls.Add(Me.cbTextures)
        Me.XPander1.Controls.Add(Me.cbModels)
        Me.XPander1.Controls.Add(Me.cbSounds)
        Me.XPander1.Controls.Add(Me.Label1)
        Me.XPander1.Controls.Add(Me.btnBrowse)
        Me.XPander1.Controls.Add(Me.txtBatchPath)
        Me.XPander1.Controls.Add(Me.Button1)
        Me.XPander1.DockPadding.Top = 25
        Me.XPander1.ExpandedHighlightImage = CType(resources.GetObject("XPander1.ExpandedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander1.ExpandedImage = CType(resources.GetObject("XPander1.ExpandedImage"), System.Drawing.Bitmap)
        Me.XPander1.Location = New System.Drawing.Point(8, 8)
        Me.XPander1.Name = "XPander1"
        Me.XPander1.ShowTooltips = True
        Me.XPander1.Size = New System.Drawing.Size(440, 168)
        Me.XPander1.TabIndex = 1
        Me.XPander1.Tag = 0
        Me.XPander1.TooltipText = Nothing
        '
        'pb1
        '
        Me.pb1.Location = New System.Drawing.Point(16, 132)
        Me.pb1.Name = "pb1"
        Me.pb1.Size = New System.Drawing.Size(408, 18)
        Me.pb1.TabIndex = 41
        '
        'cbMetadata
        '
        Me.cbMetadata.Location = New System.Drawing.Point(136, 104)
        Me.cbMetadata.Name = "cbMetadata"
        Me.cbMetadata.Size = New System.Drawing.Size(120, 24)
        Me.cbMetadata.TabIndex = 39
        Me.cbMetadata.Text = "Extract Metadata"
        '
        'cbTextures
        '
        Me.cbTextures.Location = New System.Drawing.Point(16, 104)
        Me.cbTextures.Name = "cbTextures"
        Me.cbTextures.Size = New System.Drawing.Size(120, 24)
        Me.cbTextures.TabIndex = 38
        Me.cbTextures.Text = "Extract Textures"
        '
        'cbModels
        '
        Me.cbModels.Location = New System.Drawing.Point(136, 80)
        Me.cbModels.Name = "cbModels"
        Me.cbModels.Size = New System.Drawing.Size(120, 24)
        Me.cbModels.TabIndex = 37
        Me.cbModels.Text = "Extract Models"
        '
        'cbSounds
        '
        Me.cbSounds.Location = New System.Drawing.Point(16, 80)
        Me.cbSounds.Name = "cbSounds"
        Me.cbSounds.Size = New System.Drawing.Size(120, 24)
        Me.cbSounds.TabIndex = 36
        Me.cbSounds.Text = "Extract Sounds"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 16)
        Me.Label1.TabIndex = 35
        Me.Label1.Text = "Extract To Folder:"
        '
        'btnBrowse
        '
        Me.btnBrowse.BackColor = System.Drawing.Color.Black
        Me.btnBrowse.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowse.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnBrowse.Location = New System.Drawing.Point(368, 56)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(56, 22)
        Me.btnBrowse.TabIndex = 34
        Me.btnBrowse.Text = "Browse"
        '
        'txtBatchPath
        '
        Me.txtBatchPath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBatchPath.Location = New System.Drawing.Point(16, 56)
        Me.txtBatchPath.Name = "txtBatchPath"
        Me.txtBatchPath.Size = New System.Drawing.Size(344, 20)
        Me.txtBatchPath.TabIndex = 0
        Me.txtBatchPath.Text = ""
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Black
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Button1.Location = New System.Drawing.Point(272, 96)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(152, 22)
        Me.Button1.TabIndex = 34
        Me.Button1.Text = "Start Batch Extraction"
        '
        'XPanderList1
        '
        Me.XPanderList1.AutoScroll = True
        Me.XPanderList1.BackColor = System.Drawing.Color.FromArgb(CType(99, Byte), CType(117, Byte), CType(222, Byte))
        Me.XPanderList1.Controls.Add(Me.XPander1)
        Me.XPanderList1.Location = New System.Drawing.Point(0, 0)
        Me.XPanderList1.Name = "XPanderList1"
        Me.XPanderList1.Size = New System.Drawing.Size(458, 408)
        Me.XPanderList1.TabIndex = 2
        '
        'BatchExtractStarter
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(458, 187)
        Me.Controls.Add(Me.XPanderList1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "BatchExtractStarter"
        Me.Text = "Halo Map Tools"
        Me.XPander1.ResumeLayout(False)
        Me.XPanderList1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public map As HaloMap.Map
    Public bitm As BitmLib.BitmMeta
    Public snd As sndLib.SndMeta
    Public strBatchPath As String
    Public Shared bcancelbatch As Boolean = False
    Public th As New Thread(AddressOf BatchProcess)

    Private Sub Form_Closing(ByVal sender As Object, _
            ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        th.Abort()
    End Sub

    Public Sub InitBatch()
        th.Priority = ThreadPriority.Lowest
        th.Start()
        For Each c As Control In Me.Controls
            c.Enabled = False
        Next
        While th.IsAlive
            Application.DoEvents()
        End While
        For Each c As Control In Me.Controls
            c.Enabled = True
        Next
    End Sub

    Public Sub BatchProcess()
        Dim strPath As String
        Dim strName As String
        Dim strFilename As String
        Dim LastFile As String
        Dim t As DateTime = Now()
        Dim startOffset, endOffset As Integer
        Dim bspSize(0) As Integer

        Dim oRange As Map.OffsetRange

        Dim tWriter As New StreamWriter(New FileStream(strBatchPath & "\offsets.txt", FileMode.Create))
        Dim q As String = Chr(34)
        Dim r As String = q & "," & q
        tWriter.WriteLine(q & "Index" & r & "Filename" & r & "VertexMin" & r & "VertexMax" & r & "IndexMin" & r & "IndexMax" & q)

        Try
            pb1.Minimum = 1
            pb1.Maximum = map.indexHeader.tagcount
            pb1.Visible = True
            Me.Refresh()

            Dim strType As String
            Dim x As Integer
            For x = 1 To map.indexHeader.tagcount
                pb1.Value = x

                'Extract binary sound files          
                If cbSounds.Checked And map.indexItem(x).tagClass = "snd!" Then
                    strPath = ""
                    For z As Integer = 1 To NumberOfLevels(map.indexItem(x).filePath)
                        strPath &= getLevel(map.indexItem(x).filePath, z) & "\"
                    Next
                    strPath = strBatchPath & "\" & strPath
                    strName = getLevel(map.indexItem(x).filePath, 0)

                    If map.pc And (Not snd Is Nothing) Then _
                        snd.ReleaseFile()

                    snd = New SndMeta(map, x, IIf(map.pc, "sounds.map", ""))

                    For y As Integer = 1 To snd.Struct.sndHeader.track_count
                        If Not Directory.Exists(strPath) Then Directory.CreateDirectory(strPath)
                        strFilename = strName & IIf(strName <> snd.Struct.sndTrack(y).name, _
                               "." & snd.Struct.sndTrack(y).name, "")
                        snd.SaveChunks(y, 1, snd.Struct.sndTrack(y).numChunks, _
                            strPath & "\" & strFilename & IIf(snd.Struct.sndHeader.format = 3, ".ogg", ".wav"))
                        LastFile = strFilename
                    Next
                End If
                'Extract binary image files
                If cbTextures.Checked And map.indexItem(x).tagClass = "bitm" Then
                    strPath = ""
                    For z As Integer = 1 To NumberOfLevels(map.indexItem(x).filePath) - 1
                        strPath &= getLevel(map.indexItem(x).filePath, z) & "\"
                    Next
                    strPath = strBatchPath & "\" & strPath
                    strName = getLevel(map.indexItem(x).filePath, 0)

                    If map.pc And (Not bitm Is Nothing) Then _
                        bitm.ReleaseFile()

                    bitm = New BitmMeta(map, x, IIf(map.pc, "bitmaps.map", ""))

                    If Not Directory.Exists(strPath) Then Directory.CreateDirectory(strPath)

                    Select Case bitm.bitm.second(1).format
                        Case bitm.bitmEnum.BITM_FORMAT_DXT1
                            strType = "DXT1 - "
                        Case bitm.bitmEnum.BITM_FORMAT_DXT2AND3
                            strType = "DXT23 - "
                        Case bitm.bitmEnum.BITM_FORMAT_DXT4AND5
                            strType = "DXT45 - "
                        Case bitm.bitmEnum.BITM_FORMAT_R5G6B5
                            strType = "R5G5B5 - "
                        Case bitm.bitmEnum.BITM_FORMAT_A1R5G5B5
                            strType = "A1R5G5B5 - "
                        Case bitm.bitmEnum.BITM_FORMAT_A4R4G4B4
                            strType = "A4R4G4B4 - "
                        Case bitm.bitmEnum.BITM_FORMAT_X8R8G8B8
                            strType = "X8R8G8B8 - "
                        Case bitm.bitmEnum.BITM_FORMAT_A8R8G8B8
                            strType = "A8R8G8B8 - "
                        Case Else
                            strType = ""
                    End Select

                    strFilename = strPath & "\" & strType & strName & ".dds"

                    bitm.ExtractDDS(map, 1, strFilename)
                    LastFile = strFilename
                End If

                'Extract metadata (this is what we're really interested in right now)
                If cbMetadata.Checked Then
                    If map.indexItem(x).tagClass <> "sbsp" Then
                        oRange = map.SaveMeta(x, map.indexItem(x).filePath & "." & map.indexItem(x).tagClass & ".meta", strBatchPath & "\", False)
                    Else
                        'Calculate the BSP offsets
                        'Step One - Find the scnr tag
                        For i As Integer = 1 To map.indexHeader.tagcount
                            If map.indexItem(i).tagClass = "scnr" Then
                                startOffset = map.indexItem(i).magic_metadata_offset
                                endOffset = startOffset + map.indexItem(i).estimatedMetaSize
                                map.GetInts(endOffset - 32, 2, bspSize)
                            End If
                        Next
                        oRange = map.SaveMeta(x, map.indexItem(x).filePath & "." & map.indexItem(x).tagClass & ".meta", strBatchPath & "\", False, _
                            bspSize(0), bspSize(1))
                        Exit For
                    End If
                    LastFile = strBatchPath & "\" & map.indexItem(x).filePath & "." & map.indexItem(x).tagClass & ".meta"
                End If

                If (Not oRange.maxVOffset = 0) Then _
                tWriter.WriteLine(x & "," & q & map.indexItem(x).filePath & q & "," & oRange.minVOffset & "," & oRange.maxVOffset & _
                    "," & oRange.minIOffset & "," & oRange.maxIOffset)

                If bcancelbatch Then
                    bcancelbatch = False
                    MsgBox("Extraction canceled.")
                    Exit Sub
                End If
                Application.DoEvents()
            Next

            'Console.WriteLine(Now.Subtract(t))
            pb1.Value = pb1.Minimum

            'f.Hide()
            MsgBox("Batch extraction complete!" & vbCrLf & _
                "Total time: " & Now.Subtract(t).ToString)
        Catch e As Exception
            'f.Hide()
            MsgBox("There was an error extracting one or more objects." & vbCrLf & e.Message & vbCrLf & LastFile)
        End Try

        tWriter.Close()

        'Finally, release PC resource files if neccessary
        If map.pc Then
            Try
                If Not bitm Is Nothing Then bitm.ReleaseFile()
                If Not snd Is Nothing Then snd.ReleaseFile()
            Catch
            End Try
        End If
    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        FolderBrowserDialog.Description = "Choose the folder where the files will be extracted." & vbCrLf & _
                "Note: The files will be placed in the correct folder hierarchy."
        If FolderBrowserDialog.ShowDialog = DialogResult.Cancel Then Exit Sub
        txtBatchPath.Text = FolderBrowserDialog.SelectedPath
    End Sub

    Private Sub btnExtract(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        strBatchPath = txtBatchPath.Text
        If strBatchPath = "" Then
            Exit Sub
            MsgBox("Choose a destination folder.")
        End If
        InitBatch()
    End Sub
End Class
