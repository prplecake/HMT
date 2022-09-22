'//////////////////////////////////////////////////////////
'// Halo Map Tools
'// Version 3.1
'// Created by MonoxideC
'// Released to the public domain on 2/13/2004
'//////////////////////////////////////////////////////////

Imports System
Imports System.IO
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Text
Imports System.Threading
Imports System.Xml
Imports HaloMap
Imports sndLib
Imports BitmLib
Imports HMTLib
Imports ScnrLib
Imports XMLLib
Imports Microsoft.VisualBasic.Interaction

Public Class frmMain
    Inherits Windows.Forms.Form
    Const CONST_APP_TITLE = "Halo Map Tools V3.1"

#Region "Public Data Members"

    Public Map As New Map
    Public snd As SndMeta
    Public bitm As BitmMeta
    Public scnr As ScnrMeta
    Public GUI As Object 'Generic object that hold tag plugins (late bound)
    Public WithEvents mGUI As HaloMap.MapGUI
    Public currentPath As String
    Public bFileIsLoaded As Boolean = False
    Public bChangingFiles As Boolean = True
    Public bReady As Boolean = True
    Public f As New frmBox
    Public intSCounter As Integer = 0
    Public fAbout As New Form
    Public cTagTypes As Collection
    Public bToggleIndexPanel As Boolean = False
    Private RC As New RecentFile
    Public LastTag As String = ""  'Keeps track of the last tag selected, so that the appropriate action _
    'can be taken when the next tag is selected (for example, if the bitm plugin is already loaded,
    'there is no reason to load it again)
#End Region

#Region "Private Data Members"
#End Region

#Region "Constructor"
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Dim oNode As New System.Windows.Forms.TreeNode
        Try
            oNode.Text = "Halo Map File"
            oNode.Nodes.Add("No file loaded.")
            tv1.Nodes.Add(oNode)
        Catch ex As Exception
            MsgBox("Cannot create initial node:" & ex.ToString)
            End
        End Try

        'Check to see if the Data folder exists under the application.
        'If not, create it.
        If Not Directory.Exists(Application.StartupPath & "\data") Then
            Try
                Directory.CreateDirectory(Application.StartupPath & "\data")
            Catch ex As Exception
                MsgBox("Could not create folder Data folder: " & vbCrLf & _
                    ex.ToString & " " & ex.StackTrace)
                End
            End Try
        End If
        RC.Application_Name = "HaloMapToolsV3"
        RC.Session_ID = 13245
        Dim d As DataTable = RC.GetRecentFilesList
        For Each r As DataRow In d.Rows
            Console.WriteLine(r.Item(0))
        Next
        UpdateRecentFilesOnMenu()
    End Sub
#End Region

#Region "Windows Form Designer generated code"
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
    Friend WithEvents openFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents XPanderList1 As XPanderControl.XPanderList
    Friend WithEvents XPander2 As XPanderControl.XPander
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOpen As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents SaveFileDialog As System.Windows.Forms.SaveFileDialog
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents FolderBrowserDialog As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents mnuTools As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents btnOpenMapFile As System.Windows.Forms.Button
    Friend WithEvents XPander1 As XPanderControl.XPander
    Friend WithEvents pb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents btnCloseFile As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents p As System.Windows.Forms.Panel
    Friend WithEvents pRefPanel As System.Windows.Forms.Panel
    Friend WithEvents pMeta As System.Windows.Forms.Panel
    Friend WithEvents XPander3 As XPanderControl.XPander
    Friend WithEvents XPander4 As XPanderControl.XPander
    Friend WithEvents tv1 As TreeViewEX
    Friend WithEvents XPander5 As XPanderControl.XPander
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnSMagic As System.Windows.Forms.Button
    Friend WithEvents btnMagic As System.Windows.Forms.Button
    Friend WithEvents txtMagic As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnlocateTagByMetaOffset As System.Windows.Forms.Button
    Friend WithEvents btnLocateByID As System.Windows.Forms.Button
    Friend WithEvents btnHexDec As System.Windows.Forms.Button
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExtract As System.Windows.Forms.MenuItem
    Friend WithEvents mnuBatchExtract As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExtractRange As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRebuild As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSpaceInsert As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
        Me.openFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.XPanderList1 = New XPanderControl.XPanderList
        Me.XPander5 = New XPanderControl.XPander
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnLocateByID = New System.Windows.Forms.Button
        Me.btnHexDec = New System.Windows.Forms.Button
        Me.txtID = New System.Windows.Forms.TextBox
        Me.btnlocateTagByMetaOffset = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.btnSMagic = New System.Windows.Forms.Button
        Me.btnMagic = New System.Windows.Forms.Button
        Me.txtMagic = New System.Windows.Forms.TextBox
        Me.XPander4 = New XPanderControl.XPander
        Me.tv1 = New HMTLib.TreeViewEX
        Me.btnCloseFile = New System.Windows.Forms.Button
        Me.btnOpenMapFile = New System.Windows.Forms.Button
        Me.pb1 = New System.Windows.Forms.ProgressBar
        Me.XPander3 = New XPanderControl.XPander
        Me.pMeta = New System.Windows.Forms.Panel
        Me.XPander2 = New XPanderControl.XPander
        Me.p = New System.Windows.Forms.Panel
        Me.pRefPanel = New System.Windows.Forms.Panel
        Me.Label4 = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.mnuFile = New System.Windows.Forms.MenuItem
        Me.mnuOpen = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.mnuExit = New System.Windows.Forms.MenuItem
        Me.mnuTools = New System.Windows.Forms.MenuItem
        Me.mnuExtract = New System.Windows.Forms.MenuItem
        Me.mnuBatchExtract = New System.Windows.Forms.MenuItem
        Me.mnuExtractRange = New System.Windows.Forms.MenuItem
        Me.mnuRebuild = New System.Windows.Forms.MenuItem
        Me.mnuSpaceInsert = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.mnuHelp = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.SaveFileDialog = New System.Windows.Forms.SaveFileDialog
        Me.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog
        Me.XPander1 = New XPanderControl.XPander
        Me.XPanderList1.SuspendLayout()
        Me.XPander5.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.XPander4.SuspendLayout()
        Me.XPander3.SuspendLayout()
        Me.XPander2.SuspendLayout()
        Me.pRefPanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'XPanderList1
        '
        Me.XPanderList1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.XPanderList1.AutoScroll = True
        Me.XPanderList1.BackColor = System.Drawing.Color.DarkBlue
        Me.XPanderList1.BackColorDark = System.Drawing.Color.DarkBlue
        Me.XPanderList1.BackColorLight = System.Drawing.Color.DarkBlue
        Me.XPanderList1.Controls.Add(Me.XPander5)
        Me.XPanderList1.Controls.Add(Me.XPander4)
        Me.XPanderList1.Controls.Add(Me.XPander3)
        Me.XPanderList1.Controls.Add(Me.XPander2)
        Me.XPanderList1.Location = New System.Drawing.Point(0, -23)
        Me.XPanderList1.Name = "XPanderList1"
        Me.XPanderList1.Size = New System.Drawing.Size(745, 567)
        Me.XPanderList1.TabIndex = 6
        '
        'XPander5
        '
        Me.XPander5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.XPander5.Anchored = False
        Me.XPander5.Animated = False
        Me.XPander5.AnimationTime = 100
        Me.XPander5.BackColor = System.Drawing.Color.Transparent
        Me.XPander5.BorderStyle = System.Windows.Forms.Border3DStyle.Flat
        Me.XPander5.CanToggle = True
        Me.XPander5.CaptionFont = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold)
        Me.XPander5.CaptionStyle = XPanderControl.XPander.CaptionStyleEnum.Normal
        Me.XPander5.CaptionText = "Utilities"
        Me.XPander5.ChevronStyle = XPanderControl.XPander.ChevronStyleEnum.Image
        Me.XPander5.CollapsedHighlightImage = CType(resources.GetObject("XPander5.CollapsedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander5.CollapsedImage = CType(resources.GetObject("XPander5.CollapsedImage"), System.Drawing.Bitmap)
        Me.XPander5.Controls.Add(Me.Panel1)
        Me.XPander5.DockPadding.Top = 25
        Me.XPander5.ExpandedHighlightImage = CType(resources.GetObject("XPander5.ExpandedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander5.ExpandedImage = CType(resources.GetObject("XPander5.ExpandedImage"), System.Drawing.Bitmap)
        Me.XPander5.Location = New System.Drawing.Point(8, 184)
        Me.XPander5.Name = "XPander5"
        Me.XPander5.ShowTooltips = True
        Me.XPander5.Size = New System.Drawing.Size(384, 104)
        Me.XPander5.TabIndex = 4
        Me.XPander5.Tag = 0
        Me.XPander5.TooltipText = Nothing
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btnLocateByID)
        Me.Panel1.Controls.Add(Me.btnHexDec)
        Me.Panel1.Controls.Add(Me.txtID)
        Me.Panel1.Controls.Add(Me.btnlocateTagByMetaOffset)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.btnSMagic)
        Me.Panel1.Controls.Add(Me.btnMagic)
        Me.Panel1.Controls.Add(Me.txtMagic)
        Me.Panel1.Location = New System.Drawing.Point(8, 32)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(368, 64)
        Me.Panel1.TabIndex = 0
        '
        'btnLocateByID
        '
        Me.btnLocateByID.BackColor = System.Drawing.Color.Black
        Me.btnLocateByID.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLocateByID.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnLocateByID.Location = New System.Drawing.Point(291, 34)
        Me.btnLocateByID.Name = "btnLocateByID"
        Me.btnLocateByID.Size = New System.Drawing.Size(67, 20)
        Me.btnLocateByID.TabIndex = 46
        Me.btnLocateByID.Text = "Locate ID"
        '
        'btnHexDec
        '
        Me.btnHexDec.BackColor = System.Drawing.Color.Black
        Me.btnHexDec.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHexDec.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnHexDec.Location = New System.Drawing.Point(291, 8)
        Me.btnHexDec.Name = "btnHexDec"
        Me.btnHexDec.Size = New System.Drawing.Size(66, 20)
        Me.btnHexDec.TabIndex = 45
        Me.btnHexDec.Text = "Hex/Dec"
        '
        'txtID
        '
        Me.txtID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtID.Location = New System.Drawing.Point(189, 8)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(97, 20)
        Me.txtID.TabIndex = 44
        Me.txtID.Text = ""
        Me.txtID.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnlocateTagByMetaOffset
        '
        Me.btnlocateTagByMetaOffset.BackColor = System.Drawing.Color.Black
        Me.btnlocateTagByMetaOffset.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnlocateTagByMetaOffset.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnlocateTagByMetaOffset.Location = New System.Drawing.Point(114, 35)
        Me.btnlocateTagByMetaOffset.Name = "btnlocateTagByMetaOffset"
        Me.btnlocateTagByMetaOffset.Size = New System.Drawing.Size(67, 20)
        Me.btnlocateTagByMetaOffset.TabIndex = 43
        Me.btnlocateTagByMetaOffset.Text = "Locate Tag"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Black
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Button1.Location = New System.Drawing.Point(114, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(66, 20)
        Me.Button1.TabIndex = 42
        Me.Button1.Text = "Swap Endian"
        '
        'btnSMagic
        '
        Me.btnSMagic.BackColor = System.Drawing.Color.Black
        Me.btnSMagic.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSMagic.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnSMagic.Location = New System.Drawing.Point(61, 35)
        Me.btnSMagic.Name = "btnSMagic"
        Me.btnSMagic.Size = New System.Drawing.Size(48, 20)
        Me.btnSMagic.TabIndex = 41
        Me.btnSMagic.Text = "- Magic"
        '
        'btnMagic
        '
        Me.btnMagic.BackColor = System.Drawing.Color.Black
        Me.btnMagic.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMagic.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnMagic.Location = New System.Drawing.Point(11, 35)
        Me.btnMagic.Name = "btnMagic"
        Me.btnMagic.Size = New System.Drawing.Size(45, 20)
        Me.btnMagic.TabIndex = 40
        Me.btnMagic.Text = "+ Magic"
        '
        'txtMagic
        '
        Me.txtMagic.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMagic.Location = New System.Drawing.Point(12, 8)
        Me.txtMagic.Name = "txtMagic"
        Me.txtMagic.Size = New System.Drawing.Size(97, 20)
        Me.txtMagic.TabIndex = 39
        Me.txtMagic.Text = ""
        Me.txtMagic.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'XPander4
        '
        Me.XPander4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.XPander4.Anchored = True
        Me.XPander4.Animated = False
        Me.XPander4.AnimationTime = 100
        Me.XPander4.BackColor = System.Drawing.Color.Transparent
        Me.XPander4.BorderStyle = System.Windows.Forms.Border3DStyle.Flat
        Me.XPander4.CanToggle = False
        Me.XPander4.CaptionFont = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold)
        Me.XPander4.CaptionStyle = XPanderControl.XPander.CaptionStyleEnum.Normal
        Me.XPander4.CaptionText = "Resource Browser"
        Me.XPander4.CaptionTextHighlightColor = System.Drawing.Color.FromArgb(CType(33, Byte), CType(93, Byte), CType(198, Byte))
        Me.XPander4.ChevronStyle = XPanderControl.XPander.ChevronStyleEnum.Image
        Me.XPander4.CollapsedHighlightImage = CType(resources.GetObject("XPander4.CollapsedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander4.CollapsedImage = CType(resources.GetObject("XPander4.CollapsedImage"), System.Drawing.Bitmap)
        Me.XPander4.Controls.Add(Me.tv1)
        Me.XPander4.Controls.Add(Me.btnCloseFile)
        Me.XPander4.Controls.Add(Me.btnOpenMapFile)
        Me.XPander4.Controls.Add(Me.pb1)
        Me.XPander4.DockPadding.Top = 25
        Me.XPander4.ExpandedHighlightImage = CType(resources.GetObject("XPander4.ExpandedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander4.ExpandedImage = CType(resources.GetObject("XPander4.ExpandedImage"), System.Drawing.Bitmap)
        Me.XPander4.Location = New System.Drawing.Point(8, 294)
        Me.XPander4.Name = "XPander4"
        Me.XPander4.ShowTooltips = True
        Me.XPander4.Size = New System.Drawing.Size(385, 258)
        Me.XPander4.TabIndex = 3
        Me.XPander4.Tag = 1
        Me.XPander4.TooltipText = Nothing
        '
        'tv1
        '
        Me.tv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tv1.HideSelection = False
        Me.tv1.ImageIndex = -1
        Me.tv1.Location = New System.Drawing.Point(8, 32)
        Me.tv1.Name = "tv1"
        Me.tv1.SelectedImageIndex = -1
        Me.tv1.Size = New System.Drawing.Size(368, 186)
        Me.tv1.TabIndex = 31
        Me.tv1.Text = "TreeViewEX"
        '
        'btnCloseFile
        '
        Me.btnCloseFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCloseFile.BackColor = System.Drawing.Color.Black
        Me.btnCloseFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseFile.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnCloseFile.Location = New System.Drawing.Point(8, 227)
        Me.btnCloseFile.Name = "btnCloseFile"
        Me.btnCloseFile.Size = New System.Drawing.Size(88, 24)
        Me.btnCloseFile.TabIndex = 30
        Me.btnCloseFile.Text = "Close Map File"
        Me.btnCloseFile.Visible = False
        '
        'btnOpenMapFile
        '
        Me.btnOpenMapFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnOpenMapFile.BackColor = System.Drawing.Color.Black
        Me.btnOpenMapFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOpenMapFile.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnOpenMapFile.Location = New System.Drawing.Point(8, 227)
        Me.btnOpenMapFile.Name = "btnOpenMapFile"
        Me.btnOpenMapFile.Size = New System.Drawing.Size(88, 24)
        Me.btnOpenMapFile.TabIndex = 21
        Me.btnOpenMapFile.Text = "Open Map File"
        '
        'pb1
        '
        Me.pb1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pb1.Location = New System.Drawing.Point(104, 232)
        Me.pb1.Name = "pb1"
        Me.pb1.Size = New System.Drawing.Size(273, 16)
        Me.pb1.TabIndex = 28
        Me.pb1.Visible = False
        '
        'XPander3
        '
        Me.XPander3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.XPander3.Anchored = False
        Me.XPander3.Animated = False
        Me.XPander3.AnimationTime = 100
        Me.XPander3.BackColor = System.Drawing.Color.Transparent
        Me.XPander3.BorderStyle = System.Windows.Forms.Border3DStyle.Flat
        Me.XPander3.CanToggle = True
        Me.XPander3.CaptionFont = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold)
        Me.XPander3.CaptionStyle = XPanderControl.XPander.CaptionStyleEnum.Normal
        Me.XPander3.CaptionText = "Tag Informaiton"
        Me.XPander3.CaptionTextHighlightColor = System.Drawing.Color.FromArgb(CType(33, Byte), CType(93, Byte), CType(198, Byte))
        Me.XPander3.ChevronStyle = XPanderControl.XPander.ChevronStyleEnum.Image
        Me.XPander3.CollapsedHighlightImage = CType(resources.GetObject("XPander3.CollapsedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander3.CollapsedImage = CType(resources.GetObject("XPander3.CollapsedImage"), System.Drawing.Bitmap)
        Me.XPander3.Controls.Add(Me.pMeta)
        Me.XPander3.DockPadding.Top = 25
        Me.XPander3.ExpandedHighlightImage = CType(resources.GetObject("XPander3.ExpandedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander3.ExpandedImage = CType(resources.GetObject("XPander3.ExpandedImage"), System.Drawing.Bitmap)
        Me.XPander3.Location = New System.Drawing.Point(8, 31)
        Me.XPander3.Name = "XPander3"
        Me.XPander3.ShowTooltips = True
        Me.XPander3.Size = New System.Drawing.Size(385, 147)
        Me.XPander3.TabIndex = 2
        Me.XPander3.Tag = 2
        Me.XPander3.TooltipText = Nothing
        '
        'pMeta
        '
        Me.pMeta.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pMeta.BackColor = System.Drawing.Color.White
        Me.pMeta.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pMeta.Location = New System.Drawing.Point(8, 32)
        Me.pMeta.Name = "pMeta"
        Me.pMeta.Size = New System.Drawing.Size(369, 107)
        Me.pMeta.TabIndex = 33
        '
        'XPander2
        '
        Me.XPander2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.XPander2.Anchored = False
        Me.XPander2.Animated = False
        Me.XPander2.AnimationTime = 20
        Me.XPander2.BackColor = System.Drawing.Color.Transparent
        Me.XPander2.BorderStyle = System.Windows.Forms.Border3DStyle.Flat
        Me.XPander2.CanToggle = False
        Me.XPander2.CaptionCurveRadius = 15
        Me.XPander2.CaptionFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.XPander2.CaptionStyle = XPanderControl.XPander.CaptionStyleEnum.Normal
        Me.XPander2.CaptionText = "Tag Editor"
        Me.XPander2.CaptionTextHighlightColor = System.Drawing.Color.FromArgb(CType(33, Byte), CType(93, Byte), CType(198, Byte))
        Me.XPander2.ChevronStyle = XPanderControl.XPander.ChevronStyleEnum.ArrowsInCircle
        Me.XPander2.CollapsedHighlightImage = CType(resources.GetObject("XPander2.CollapsedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander2.CollapsedImage = CType(resources.GetObject("XPander2.CollapsedImage"), System.Drawing.Bitmap)
        Me.XPander2.Controls.Add(Me.p)
        Me.XPander2.Controls.Add(Me.pRefPanel)
        Me.XPander2.DockPadding.Top = 25
        Me.XPander2.ExpandedHighlightImage = CType(resources.GetObject("XPander2.ExpandedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander2.ExpandedImage = CType(resources.GetObject("XPander2.ExpandedImage"), System.Drawing.Bitmap)
        Me.XPander2.Location = New System.Drawing.Point(401, 31)
        Me.XPander2.Name = "XPander2"
        Me.XPander2.PaneBottomRightColor = System.Drawing.Color.White
        Me.XPander2.ShowTooltips = True
        Me.XPander2.Size = New System.Drawing.Size(336, 523)
        Me.XPander2.TabIndex = 1
        Me.XPander2.Tag = 3
        Me.XPander2.TooltipText = Nothing
        '
        'p
        '
        Me.p.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.p.BackColor = System.Drawing.Color.White
        Me.p.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.p.Location = New System.Drawing.Point(8, 32)
        Me.p.Name = "p"
        Me.p.Size = New System.Drawing.Size(320, 484)
        Me.p.TabIndex = 31
        Me.p.Visible = False
        '
        'pRefPanel
        '
        Me.pRefPanel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pRefPanel.BackColor = System.Drawing.Color.White
        Me.pRefPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pRefPanel.Controls.Add(Me.Label4)
        Me.pRefPanel.Location = New System.Drawing.Point(8, 32)
        Me.pRefPanel.Name = "pRefPanel"
        Me.pRefPanel.Size = New System.Drawing.Size(320, 484)
        Me.pRefPanel.TabIndex = 29
        '
        'Label4
        '
        Me.Label4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.Location = New System.Drawing.Point(72, 136)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(176, 164)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "(There are no controls to display, because a supported tag has not been selected)" & _
        ""
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuTools, Me.MenuItem2, Me.mnuHelp})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuOpen, Me.MenuItem3, Me.MenuItem4, Me.mnuExit})
        Me.mnuFile.Text = "&File"
        '
        'mnuOpen
        '
        Me.mnuOpen.Index = 0
        Me.mnuOpen.Text = "&Open"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "&Close"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 2
        Me.MenuItem4.Text = "-"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 3
        Me.mnuExit.Text = "E&xit"
        '
        'mnuTools
        '
        Me.mnuTools.Index = 1
        Me.mnuTools.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuExtract, Me.mnuRebuild, Me.mnuSpaceInsert})
        Me.mnuTools.Text = "&Tools"
        '
        'mnuExtract
        '
        Me.mnuExtract.Index = 0
        Me.mnuExtract.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuBatchExtract, Me.mnuExtractRange})
        Me.mnuExtract.Text = "Extract"
        '
        'mnuBatchExtract
        '
        Me.mnuBatchExtract.Index = 0
        Me.mnuBatchExtract.Text = "Batch Extract"
        '
        'mnuExtractRange
        '
        Me.mnuExtractRange.Index = 1
        Me.mnuExtractRange.Text = "Extract Range"
        '
        'mnuRebuild
        '
        Me.mnuRebuild.Index = 1
        Me.mnuRebuild.Text = "Rebuild Map"
        '
        'mnuSpaceInsert
        '
        Me.mnuSpaceInsert.Index = 2
        Me.mnuSpaceInsert.Text = "Insert Blank Space into Meta File"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 2
        Me.MenuItem2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem6, Me.MenuItem8})
        Me.MenuItem2.Text = "Options"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 0
        Me.MenuItem6.Text = "Disable Plugins"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 1
        Me.MenuItem8.Text = "Reverse Tag Names"
        '
        'mnuHelp
        '
        Me.mnuHelp.Index = 3
        Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem5})
        Me.mnuHelp.Text = "&Help"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 0
        Me.MenuItem5.Text = "&About"
        '
        'XPander1
        '
        Me.XPander1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.XPander1.Anchored = False
        Me.XPander1.Animated = True
        Me.XPander1.AnimationTime = 100
        Me.XPander1.BackColor = System.Drawing.Color.Transparent
        Me.XPander1.BorderStyle = System.Windows.Forms.Border3DStyle.Flat
        Me.XPander1.CanToggle = True
        Me.XPander1.CaptionFont = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold)
        Me.XPander1.CaptionLeftColor = System.Drawing.Color.Black
        Me.XPander1.CaptionRightColor = System.Drawing.Color.SlateGray
        Me.XPander1.CaptionStyle = XPanderControl.XPander.CaptionStyleEnum.Normal
        Me.XPander1.CaptionText = "Texture Browser"
        Me.XPander1.CaptionTextColor = System.Drawing.Color.White
        Me.XPander1.CaptionTextHighlightColor = System.Drawing.Color.White
        Me.XPander1.ChevronStyle = XPanderControl.XPander.ChevronStyleEnum.ArrowsInCircle
        Me.XPander1.CollapsedHighlightImage = CType(resources.GetObject("XPander1.CollapsedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander1.CollapsedImage = CType(resources.GetObject("XPander1.CollapsedImage"), System.Drawing.Bitmap)
        Me.XPander1.DockPadding.Top = 25
        Me.XPander1.ExpandedHighlightImage = CType(resources.GetObject("XPander1.ExpandedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander1.ExpandedImage = CType(resources.GetObject("XPander1.ExpandedImage"), System.Drawing.Bitmap)
        Me.XPander1.Location = New System.Drawing.Point(8, 8)
        Me.XPander1.Name = "XPander1"
        Me.XPander1.PaneBottomRightColor = System.Drawing.Color.White
        Me.XPander1.ShowTooltips = True
        Me.XPander1.Size = New System.Drawing.Size(720, 104)
        Me.XPander1.TabIndex = 2
        Me.XPander1.Tag = 0
        Me.XPander1.TooltipText = Nothing
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(745, 541)
        Me.Controls.Add(Me.XPanderList1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.MinimumSize = New System.Drawing.Size(753, 570)
        Me.Name = "frmMain"
        Me.Text = "Halo Map Tools"
        Me.XPanderList1.ResumeLayout(False)
        Me.XPander5.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.XPander4.ResumeLayout(False)
        Me.XPander3.ResumeLayout(False)
        Me.XPander2.ResumeLayout(False)
        Me.pRefPanel.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Event Handlers"

    '//////////////////////////////////////////////////////
    '// Set the state of the expander controls and the
    '// window title when the form loads
    '//////////////////////////////////////////////////////
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        XPander3.Collapse()
        XPander3.CanToggle = False
        XPander5.Collapse()
        XPander5.CanToggle = False
        Me.Text = CONST_APP_TITLE
    End Sub

#Region "Node Events"
    'The User clicked on an item in the tree view, so we need to select the correct object data.
    Private Sub ChangeNode(ByVal sender As System.Object, ByVal e As TreeViewEventArgs) Handles tv1.AfterSelect
        If Not bFileIsLoaded Then
            If tv1.SelectedNode.Index = 0 Then
                intSCounter += 1  'This was used in the egg, which cant be acces ATM
            End If
            Exit Sub
        End If

        Dim intID As Integer
        Dim t, t2 As TreeViewEX.TreeNode
        'assembly controlLib
        'Try
        For Each t In tv1.Nodes
            If t Is e.Node.Parent Then
                For Each t2 In t.Nodes
                    If t2.Index = e.Node.Index Then
                        'We found the node 
                        intID = Val(t2.NodeValue)
                        'Create the metadata GUI
                        If mGUI Is Nothing Then
                            pMeta.Controls.Clear()
                            mGUI = New MapGUI(Map, Map.LocateByID(intID))
                            mGUI.Location = New System.Drawing.Point(0, 0)
                            mGUI.Dock = DockStyle.Fill
                            mGUI.Visible = False
                            pMeta.Controls.Add(mGUI)
                            mGUI.Visible = True
                            pMeta.Visible = True
                        Else
                            mGUI.PopulateControls(Map.LocateByID(intID))
                        End If
                        Cursor = Cursors.WaitCursor
                        If Not MenuItem6.Checked Then
                            Select Case (Map.indexItem(Map.LocateByID(intID)).tagClass)
                                Case "snd!"
                                    If LastTag = "snd!" Then
                                        GUI.SelectSound(Map, Map.LocateByID(intID))
                                    Else
                                        If Not GUI Is Nothing Then GUI.Unload()
                                        GUI = New SndGUI(Map, Map.LocateByID(intID))
                                        LastTag = "snd!"
                                    End If
                                Case "bitm"
                                    If LastTag = "bitm" Then
                                        GUI.SelectBitmap(Map, Map.LocateByID(intID))
                                    Else
                                        If Not GUI Is Nothing Then GUI.Unload()
                                        GUI = New BitmGUI(Map, Map.LocateByID(intID))
                                        LastTag = "bitm"
                                    End If
                                Case "scnr"
                                    If Not GUI Is Nothing Then GUI.Unload()
                                    GUI = New ScnrGUI(Map, Map.LocateByID(intID))
                                    LastTag = "scnr"
                                Case Else
                                    GUI = New XMLGUI(Map, Map.indexItem(Map.LocateByID(intID)).tagClass, Map.LocateByID(intID))
                                    LastTag = Map.indexItem(Map.LocateByID(intID)).tagClass
                                    If Not GUI.TagHandled() Then
                                        LastTag = ""
                                        GUI.Unload()
                                        GUI = Nothing
                                    End If
                            End Select
                        End If
                        If Not GUI Is Nothing Then
                            p.Controls.Clear()
                            GUI.Location = New System.Drawing.Point(0, 0)
                            GUI.Dock = DockStyle.Fill
                            p.Controls.Add(GUI)
                            GUI.Visible = True
                            p.Visible = True
                        Else
                            p.Visible = False
                        End If
                        If MenuItem6.Checked Then p.Visible = False
                        If bToggleIndexPanel Then
                            XPander3.CanToggle = True
                            XPander3.Expand()
                            bToggleIndexPanel = False
                        End If
                        Cursor = Cursors.Arrow
                    End If
                Next
            End If
        Next
        'Catch ex As Exception
        'MsgBox("An error occurred while creating a tag plugin:" & vbCrLf & vbCrLf & _
        '    "Out of Memory") 'ex.ToString)
        Cursor = Cursors.Arrow
        'End Try

    End Sub


#End Region

#Region "Menu Events"
    'User clicked 'open' in the file menu.
    Private Sub mnuOpenFile(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpen.Click
        Cursor = Cursors.WaitCursor
        OpenFile()
        Cursor = Cursors.Arrow
    End Sub

    'The user chose 'Generate ADPCM Header from Raw Data' in the Tools menu.
    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim frm As New frmGenerateHeader
        frm.Show()
    End Sub

    'User chose 'Close' from the menu.
    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        If Not bFileIsLoaded Then Exit Sub
        CloseFile()
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        fAbout = Nothing
        fAbout = New frmAbout
        fAbout.Show()
    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim frm As New adpcmencodeGUI
        frm.Show()
    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Application.Exit()
    End Sub
    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim strFilename As String
        DecompressMap(Map)
    End Sub

    'Disable plugins
    Private Sub MenuItem6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        MenuItem6.Checked = Not MenuItem6.Checked
        If MenuItem6.Checked = True Then
            p.Visible = False
        End If
    End Sub

    'Open the popup form that will initiate the build
    Private Sub mnuRebuild_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRebuild.Click
        CloseFile()
        Dim f As New MapRebuildStarter
        f.Show()
        fMain = Me
        fMain.Hide()
    End Sub

    Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click
        MenuItem8.Checked = Not MenuItem8.Checked
        If bFileIsLoaded Then
            Dim tNode As TreeViewEX.TreeNode = tv1.SelectedNode
            Dim id As Integer = 0
            If Not tNode Is Nothing Then id = Val(tNode.Value)
            PopulateNodes(tv1)
            If Not id = 0 Then SelectNodeByID(Map.LocateByID(id))
        End If
    End Sub

#End Region

#Region "Button Events"

    '///////////////////////////////////////////////////////////
    '// This was an event handler for a Swizzle button
    '// At the moment, it's not referenced
    '///////////////////////////////////////////////////////////
    Private Sub SwizzleStuff(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim dds As New BitmMeta.DDS_HEADER_STRUCTURE
        Dim bChunk(), binBuffer() As Byte
        Dim swiz As Swizzle

        openFileDialog.AddExtension = True
        openFileDialog.DefaultExt = "*.dds"
        openFileDialog.Filter = "DDS Texture (*.dds)|*.dds"
        openFileDialog.FileName = ""
        If openFileDialog.ShowDialog() = DialogResult.Cancel Then Exit Sub

        Dim br As New BinaryReader(New FileStream(openFileDialog.FileName, FileMode.Open))
        Dim bw As New BinaryWriter(New FileStream(openFileDialog.FileName & "swizzle.dds", FileMode.Create))

        dds.readStruct(br)
        dds.writeStruct(bw)

        ReDim binBuffer(br.BaseStream.Length - 108) 'Header is 108 bytes
        binBuffer = br.ReadBytes(br.BaseStream.Length - 108)

        bw.Write(bitm.Swizzle(binBuffer, dds.ddsd.width, dds.ddsd.height, -1, _
            dds.ddsd.ddfPixelFormat.RGBBitCount, BitmMeta.SwizzleType.DeSwizzle))
        br.Close()
        bw.Close()
    End Sub

    '///////////////////////////////////////////////////////////
    '// The +Magic button event handler - Just adds the current
    '// map magic number to the value in txtMagic.Text
    '///////////////////////////////////////////////////////////
    Private Sub btnMagic_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMagic.Click
        Try
            txtMagic.Text = Hex(Unsigned(Val("&H" & txtMagic.Text)) + Map.intMagic)
        Catch
            txtMagic.Text = "Overflow"
        End Try
    End Sub

    '////////////////////////////////////////////////////////////
    '// The -Magic button event handler - Just subtracts the
    '// current map magic number from the value in txtMagic.Text
    '////////////////////////////////////////////////////////////
    Private Sub btnSMagic_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSMagic.Click
        Try
            txtMagic.Text = Hex(Unsigned(Val("&H" & txtMagic.Text)) - Map.intMagic)
        Catch
            txtMagic.Text = "Overflow"
        End Try
    End Sub

    '////////////////////////////////////////////////////////////
    '// Swaps the endian-ness of the value in txtMagic.text
    '////////////////////////////////////////////////////////////
    Private Sub SwapEndian_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        txtMagic.Text = SwapENDIAN(txtMagic.Text)
    End Sub

    '//////////////////////////////////////////////////////////////////
    '// Remove any spaces that the user may have inadvertantly entered
    '//////////////////////////////////////////////////////////////////
    Private Sub txtMagic_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMagic.TextChanged
        Dim tmpStr As String = ""
        For x As Integer = 1 To Len(txtMagic.Text)
            If Mid(txtMagic.Text, 1, 1) <> " " Then tmpStr &= Mid(txtMagic.Text, 1, 1)
        Next
    End Sub

    '///////////////////////////////////////////////////////////////////
    '// Lookup a node by it's tag's metadata offset and select it
    '///////////////////////////////////////////////////////////////////
    Private Sub btnlocateTagByMetaOffset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnlocateTagByMetaOffset.Click
        SelectNodeByMetaOffset(Val("&H" & txtMagic.Text))
    End Sub

    '///////////////////////////////////////////////////////
    '// Convert the ident to either hex or decimal (invert)
    '///////////////////////////////////////////////////////
    Private Sub btnHexDec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHexDec.Click
        Dim bHex As Boolean = False
        If Len(txtID.Text.Trim(" ")) < 9 Then bHex = True
        If bHex Then
            txtID.Text = Val("&H" & txtID.Text)
        Else
            txtID.Text = Hex(txtID.Text)
        End If
    End Sub

    '///////////////////////////////////////////////////
    '// Lookup a node by it's tag's ident and select it
    '///////////////////////////////////////////////////
    Private Sub btnLocateByID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLocateByID.Click
        Dim bHex As Boolean = False
        If Len(txtID.Text.Trim(" ")) < 9 Then bHex = True
        If bHex Then
            txtID.Text = Val("&H" & txtID.Text)
        End If
        SelectNodeByID(Map.LocateByID(txtID.Text))
    End Sub

    '////////////////////////////////////////////////
    '// Handles the Open button's event - open a map
    '////////////////////////////////////////////////
    Private Sub btnOpenMapFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenMapFile.Click
        Cursor = Cursors.WaitCursor
        OpenFile()
        Cursor = Cursors.Arrow
    End Sub

    '//////////////////////////////////////////////////
    '// Handles the Close button's event - close a map
    '//////////////////////////////////////////////////
    Private Sub btnCloseFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloseFile.Click
        CloseFile()
    End Sub

#End Region

    '////////////////////////////////////////////////////////////////
    '// Handles external event from the Map plugin
    '// Reloads the map file when the map index changes
    '////////////////////////////////////////////////////////////////
    Private Sub MetaUpdated(ByVal intID As Integer) Handles mGUI.MetaUpdated
        Cursor = Cursors.WaitCursor
        XPander3.Collapse()
        'The meta was updated, so the file needs to be reloaded and the correct treenode reselected
        Dim strTemp As String = Map.strFilename
        CloseFile()
        OpenFile(strTemp)
        SelectNodeByID(intID)
        Cursor = Cursors.Arrow
    End Sub

#End Region

#Region "Open/Close/Save Proceudres"
    '///////////////////////////////////////////////////////////////////
    '// Open a Halo map file, and read in all of its relevant data
    '///////////////////////////////////////////////////////////////////
    Private Sub OpenFile(Optional ByVal s As String = "")

        If bFileIsLoaded Then
            If Not MsgBox("A map file is currently open.  Do you wish to close it?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Exit Sub
            Else
                CloseFile()
            End If
        End If

        Dim x As Integer 'Loop counter.
        Dim strFilename As String

        If bFileIsLoaded Then
            MsgBox("Close the current map file before opening a new one.")
            Exit Sub
        End If

        If s <> "" Then
            strFilename = s
        Else
            openFileDialog.AddExtension = True
            openFileDialog.DefaultExt = "*.map"
            openFileDialog.Filter = "Halo Map File (*.map)|*.map"

            If openFileDialog.ShowDialog = DialogResult.Cancel Then Exit Sub

            strFilename = openFileDialog.FileName
        End If

        If strFilename = "" Then Exit Sub

        'Create the Map object and open the file.
        Dim result As Map.eOpenFile
        result = Map.OpenFile(strFilename)
        If result And Map.eOpenFile.FileNotFound Then
            MsgBox("File not found.", MsgBoxStyle.Critical)
            Exit Sub
        End If
        If result And Map.eOpenFile.InvalidFile.InvalidFile Then
            MsgBox("Not a valid Halo map file.", MsgBoxStyle.Critical)
            Exit Sub
        End If
        If result And Map.eOpenFile.SharingViolation.SharingViolation Then
            MsgBox("The file you seletced is in use by another process.", MsgBoxStyle.Critical)
            Exit Sub
        End If

        Dim readType As Map.eOpenFile
        If result And Map.eOpenFile.PC Then readType = Map.eOpenFile.PC
        If result And Map.eOpenFile.PCDemo Then readType = Map.eOpenFile.PCDemo
        If result And Map.eOpenFile.Xbox Then readType = Map.eOpenFile.Xbox

        Select Case Map.ReadFile(readType)
            Case Map.eReadFile.Compressed
                If MsgBox("File is compressed - Do you wish to decompress?", _
                    MsgBoxStyle.YesNo) = MsgBoxResult.No Then Exit Sub
                DecompressMap(Map)
                CloseFile()
                MsgBox("File has been decompressed.", MsgBoxStyle.Information)
                Exit Sub
            Case Map.eReadFile.UnknwonError
                CloseFile()
                MsgBox("Error reading map file.", MsgBoxStyle.Critical)
                Exit Sub
        End Select

        RC.SaveItem(strFilename)

        Dim strOutputFile As String 'Path to the file that will be written.

        'Populate the TreeView object.
        tv1.Sorted = True
        bChangingFiles = True
        tv1.Visible = False

        tv1.Nodes.Clear()
        LastTag = ""

        Me.Text = CONST_APP_TITLE & " - " & Map.MapName & IIf(Map.pc, " (PC)", " (Xbox)")

        PopulateNodes(tv1)

        tv1.Visible = True
        tv1.Refresh()
        bFileIsLoaded = True
        bChangingFiles = False

        XPander5.CanToggle = True

        bToggleIndexPanel = True

        btnCloseFile.Visible = True
        btnOpenMapFile.Visible = False

        p.Visible = False

        UpdateRecentFilesOnMenu()

        'Select the first node.
        bFileIsLoaded = True
        bReady = True

        'GenerateOffsetList("C:\OffsetListTest.txt")

    End Sub

    '///////////////////////////////////////////////////////////////////
    '// Close the file
    '///////////////////////////////////////////////////////////////////
    Private Sub CloseFile()
        Dim strfile As String

        If Not bFileIsLoaded Then Exit Sub

        If Map.bFileInUse Then
            MsgBox("The map file cannot be closed until the current operation completes.")
            Exit Sub
        End If

        Try
            For Each strfile In Directory.GetFiles(Application.StartupPath & "\data")
                Kill(strfile)
            Next
        Catch
            MsgBox("There was an error while manipulating temporary files." & vbCrLf & _
                "Perhaps one of them is in use by another process.")
        End Try

        If Not GUI Is Nothing Then GUI.unload()
        GUI = Nothing
        System.GC.Collect()
        p.Visible = False

        XPander3.Collapse()
        XPander3.CanToggle = False
        XPander5.Collapse()
        XPander5.CanToggle = True
        Me.Text = CONST_APP_TITLE
        bFileIsLoaded = False
        bChangingFiles = True
        tv1.Nodes.Clear()
        bChangingFiles = False

        XPander5.CanToggle = False

        Dim n As New TreeNode("Halo Map File")
        tv1.Nodes.Add(n)
        n.Nodes.Add("No file loaded")
        btnOpenMapFile.Visible = True
        btnCloseFile.Visible = False
        btnOpenMapFile.Visible = True
        bFileIsLoaded = False
        Map.CloseFile()
    End Sub

    '////////////////////////////////////////////////
    '// Update the recently opened files on the menu
    '////////////////////////////////////////////////
    Public RecentFiles() As MenuItem
    Public MenuSeperator As New MenuItem
    Private Sub UpdateRecentFilesOnMenu()
        Dim d As DataTable = RC.GetRecentFilesList
        Dim s As String

        MenuSeperator = MenuItem4.CloneMenu

        ReDim Preserve RecentFiles(d.Rows.Count - 1)

        For x As Integer = 0 To d.Rows.Count - 1
            RecentFiles(x) = New MenuItem
            RecentFiles(x).Index = x
            s = d.Rows(x).Item(0)
            RecentFiles(x).Text = s
        Next

        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.Clear()
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuOpen, Me.MenuItem3, Me.MenuItem4})

        'Add the new menu items with events
        For x As Integer = 0 To RecentFiles.Length - 1
            Me.mnuFile.MenuItems.Add(RecentFiles(x))
            AddHandler RecentFiles(x).Click, AddressOf ClickHandler
        Next
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {MenuSeperator, Me.mnuExit})

    End Sub

    'Open the appropriate file
    Private Sub ClickHandler(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Click
        Dim tmpMenuItem As MenuItem = CType(sender, MenuItem)
        If bFileIsLoaded Then CloseFile()
        Dim f As System.io.File
        If Not f.Exists(tmpMenuItem.Text) Then
            UpdateRecentFilesOnMenu()
            MsgBox("File not found - it has been removed from the menu.")
        Else
            OpenFile(tmpMenuItem.Text)
        End If
    End Sub

#End Region

#Region "Utility Functions"

    '////////////////////////////////////////////////////////////
    '// Generate an HME-Style Offset List text file
    '////////////////////////////////////////////////////////////
    Private Sub GenerateOffsetList(ByVal strFilename As String)
        'Create the text writer stream
        Dim t As New StreamWriter(New FileStream(strFilename, FileMode.Create))

        'Write the file header area
        t.WriteLine("OUTPUT FROM EXECUTABLE CREATED ON: " & Now().ToString)
        t.WriteLine("-----------------------------------------------------")
        t.WriteLine("[Header] Offset 0x00000000")
        t.WriteLine("-----------------------------------------------------")
        t.WriteLine("Version: " & Map.fileHeader.version)
        t.WriteLine("Decompressed Length: " & Map.fileHeader.decomp_len & " bytes")
        t.WriteLine("Unknown1: " & Map.fileHeader.zeros2 & " (0x" & Hex(Map.fileHeader.zeros2) & ")")
        t.WriteLine("Metadata Size: " & Map.fileHeader.MetadataSize & " (0x" & Hex(Map.fileHeader.MetadataSize) & ")")
        t.WriteLine("Offset to Index: " & Map.fileHeader.offset_to_index_decomp & _
                " (0x" & Hex(Map.fileHeader.offset_to_index_decomp) & ")")
        t.WriteLine("Unknown3: " & Map.fileHeader.unknown3 & " (0x" & Hex(Map.fileHeader.unknown3) & ")")
        t.WriteLine("Name: " & Map.fileHeader.name & " (" & Map.MapName() & ")")
        t.WriteLine("Build (YY.MM.DD.Build): " & Map.fileHeader.builddate)
        Dim mapType As String
        Select Case Map.fileHeader.maptype
            Case 0
                mapType = "Singleplayer"
            Case 1
                mapType = "Multiplayer"
            Case 2
                mapType = "UI"
        End Select
        t.WriteLine("Maptype: " & mapType)
        t.WriteLine("")
        t.WriteLine("-----------------------------------------------------")
        t.WriteLine("[Header] Offset 0x" & Hex(Map.fileHeader.offset_to_index_decomp))
        t.WriteLine("-----------------------------------------------------")
        t.WriteLine("Index Magic: " & Map.indexHeader.index_magic & " (0x" & Hex(Map.indexHeader.index_magic) & ")")
        t.WriteLine("Starting ID: " & Map.indexHeader.starting_id & " (0x" & Hex(Map.indexHeader.starting_id) & ")")
        t.WriteLine("Unknown2: " & Map.indexHeader.unknown2 & " (0x" & Hex(Map.indexHeader.unknown2) & ")")
        t.WriteLine("Tag Count: " & Map.indexHeader.tagcount & " (0x" & Hex(Map.indexHeader.tagcount) & ")")
        t.WriteLine("Vertex Object Count: " & Map.indexHeader.vertex_object_count & " (0x" & Hex(Map.indexHeader.vertex_object_count) & ")")
        t.WriteLine("Vertex Offset: " & Map.indexHeader.vertex_offset & " (0x" & Hex(Map.indexHeader.vertex_offset) & ")")
        t.WriteLine("Indices Object Count: " & Map.indexHeader.indeces_object_count & " (0x" & Hex(Map.indexHeader.indeces_object_count) & ")")
        t.WriteLine("Indices Offset: " & Map.indexHeader.indeces_offset & " (0x" & Hex(Map.indexHeader.indeces_offset) & ")")
        t.WriteLine("Magic: " & Map.intMagic & " (0x" & Hex(Map.intMagic) & ")")
        t.WriteLine("")
        t.WriteLine("-----------------------------------------------------")
        t.WriteLine("Index Entries")
        t.WriteLine("-----------------------------------------------------")

        For x As Integer = 0 To Map.indexHeader.tagcount - 1
            t.WriteLine("")
            t.WriteLine("[Index Item] " & x & " (0x" & Hex(Map.indexItem(x + 1).offset) & ")")
            t.WriteLine("Tag class 0: " & reverseString(Mid(Map.indexItem(x + 1).fileData.tagclass, 1, 4)))
            t.WriteLine("Tag class 1: " & reverseString(Mid(Map.indexItem(x + 1).fileData.tagclass, 5, 4)))
            t.WriteLine("Tag class 2: " & reverseString(Mid(Map.indexItem(x + 1).fileData.tagclass, 8, 4)))
            t.WriteLine("Identifier: 0x" & Hex(Unsigned(Map.indexItem(x + 1).fileData.id)))
            t.WriteLine("Filename Offset: " & "0x" & Hex(Map.indexItem(x + 1).fileData.stringoffset) & _
                " (" & Map.indexItem(x + 1).fileData.stringoffset & ")")
            t.WriteLine("Raw Offset: " & "0x" & Hex(Map.indexItem(x + 1).fileData.offset) & _
                " (" & Map.indexItem(x + 1).fileData.offset & ")")
            t.WriteLine("Offset: " & "0x" & Hex(Map.indexItem(x + 1).magic_metadata_offset) & _
                " (" & Map.indexItem(x + 1).magic_metadata_offset & ")")
            t.WriteLine("Filename: " & Map.indexItem(x + 1).filePath)
        Next
        t.WriteLine("-----------------------------------------------------")
        t.Close()

    End Sub


#Region "snd!"


#End Region

#Region "scnr"
    '///////////////////////////////////////////////////////////////////
    '// Decode Scenery Data
    '///////////////////////////////////////////////////////////////////
    Private Sub SelectScenery(ByVal intID As Integer)
        Dim intIndex As Integer

        bReady = False
        If bChangingFiles Then Exit Sub

        'Make sure we have a map loaded.
        If Map.indexItem Is Nothing Then Exit Sub

        intIndex = Map.LocateByID(intID)

        Dim s As String = Map.indexItem(intIndex).filePath

        scnr = New ScnrMeta(Map, intIndex)

        'hexBox.Text = scnr.Text

        bReady = True
    End Sub
#End Region

#Region "Other"

    '/////////////////////////////////////////////
    '// Populate nodes with all tags
    '/////////////////////////////////////////////
    Private Function PopulateNodes(ByRef tv As TreeViewEX) As Integer
        Dim iCount As Integer
        Dim m As Map.TAG_STRUCT
        Dim node, newNode As TreeViewEX.TreeNode

        tv.Nodes.Clear()

        For Each m In Map.cTagTypes
            newNode = New TreeViewEX.TreeNode : node = New TreeViewEX.TreeNode
            newNode.NodeKey = m.tag
            newNode.NodeValue = m.name
            iCount = 0
            For x As Integer = 1 To Map.indexHeader.tagcount
                If Map.indexItem(x).tagClass = m.tag Then
                    node = New TreeViewEX.TreeNode
                    node.Text = Map.indexItem(x).filePath
                    node.NodeKey = Map.indexItem(x).fileData.id
                    node.NodeValue = Map.indexItem(x).fileData.id
                    newNode.Nodes.Add(node)
                    iCount += 1
                End If
            Next
            If Not MenuItem8.Checked Then
                newNode.Text = "[" & m.tag & "] " & m.name & " (" & iCount & " items)"
            Else
                newNode.Text = m.name & " [" & m.tag & "] (" & iCount & " items)"
            End If

            If iCount > 0 Then tv.Nodes.Add(newNode)
        Next
    End Function

    Private Sub SelectNodeByID(ByVal intID As Integer)
        'The meta was updated, so the file needs to be reloaded and the correct treenode reselected
        Dim t2, t3 As TreeViewEX.TreeNode
        For Each t2 In tv1.Nodes
            For Each t3 In t2.Nodes
                If Map.LocateByID(t3.Value) = Str(intID) Then
                    t2.Expand()
                    tv1.SelectedNode = t3
                End If
            Next
        Next
    End Sub

    Private Sub SelectNodeByMetaOffset(ByVal intOffset As Integer)
        'The meta was updated, so the file needs to be reloaded and the correct treenode reselected
        Dim bFound As Boolean = False
        Dim intID As Integer

        For x As Integer = 1 To Map.indexHeader.tagcount
            If Map.indexItem(x).magic_metadata_offset = intOffset Then
                bFound = True
                intID = Map.indexItem(x).fileData.id
            End If
        Next
        If Not bFound Then
            MsgBox("Tag not found.  Make sure you have subtracted magic!")
            Exit Sub
        End If
        SelectNodeByID(Map.LocateByID(intID))
    End Sub

    Private Function DecompressMap(ByRef m As Map) As String
        'Dim z As New zLibWrapper.zLib
        'Dim bw As BinaryWriter
        'Dim br As New BinaryReader(Map.fs)
        'Dim i As Integer 'Loop counter
        'Dim iLength As Integer = Map.fileHeader.decomp_len
        'Dim chunkSize As Integer = (Map.fs.Length - Map.fs.Position)
        'Dim bChunk() As Byte
        'Dim mb As MemoryStream
        'Dim test As Zip.ZipConstants
        '
        'Dim fs As FileStream = New FileStream(Map.strFilename & ".decomp.map", FileMode.Create)

        MsgBox("Not yet supported... Why is it in the menu then???  Good question...")

        'Map.GetBytes(&H800, chunkSize, bChunk)

        'MsgBox(Hex(test.LOCSIG))

        'mb = New MemoryStream(bChunk)

        'z.DecompressStream(mb, fs)

    End Function

    '///////////////////////////////////////////////////////////////////
    '// Temporarily in this class (will be moved)
    '///////////////////////////////////////////////////////////////////
    Public Sub SelectMeta(ByVal intID As Integer)
        Dim intIndex As Integer

        bReady = False
        If bChangingFiles Then Exit Sub

        'Make sure we have a map loaded.
        If Map.indexItem Is Nothing Then Exit Sub

        intIndex = Map.LocateByID(intID)

        Dim s As String = Map.indexItem(intIndex).filePath & vbCrLf & _
            "0x" & Hex(Map.indexItem(intIndex).offset) & vbCrLf

        'Populate the Hex Box
        If Map.indexItem(intIndex).estimatedMetaSize > 500000 _
            Or Map.indexItem(intIndex).estimatedMetaSize < 1 Then
            'hexBox.Text = s & "Metadata Error"
        Else
            Dim bChunk() As Byte
            Map.GetBytes(Map.indexItem(intIndex).offset, Map.indexItem(intIndex).estimatedMetaSize, bChunk)
            Dim x As Integer
            For x = 1 To IIf(bChunk.Length <= 256, bChunk.Length, 256)
                If bChunk(x - 1) < 16 Then s &= "0"
                s &= Hex(bChunk(x - 1))
                If x Mod 2 = 0 Then s &= " "
                If x Mod 16 = 0 Then s &= vbCrLf
            Next
            If x < bChunk.Length Then s &= "[truncated at 256 bytes]"
            'hexBox.Text = s
        End If
    End Sub

#End Region

#End Region

    Private Sub mnuBatchExtract_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuBatchExtract.Click
        'Make sure a map file is loaded.
        If Not bFileIsLoaded Then
            MsgBox("A valid Halo map file has not been loaded.")
            Exit Sub
        End If
        Dim f As New BatchExtractStarter(Map)
        f.Show()
    End Sub

    '///////////////////////////////////////////////////
    '// Export a user specified range of metadata
    '///////////////////////////////////////////////////
    Private Sub mnuExtractRange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExtractRange.Click
        If Not bFileIsLoaded Then
            MsgBox("A valid Halo map file has not been loaded.")
            Exit Sub
        End If
        Dim StartOffset As Long = Convert.ToInt64(InputBox("Start offset (hex)"), 16)
        Dim MetaLength As Long = Val(InputBox("Size in bytes"))
        FolderBrowserDialog.Description = "Choose the root folder where the files will be extracted." & vbCrLf & _
            "Note: The files will be placed in the correct folder hierarchy."
        If FolderBrowserDialog.ShowDialog = DialogResult.Cancel Then Exit Sub
        Dim strPath As String = FolderBrowserDialog.SelectedPath
        Dim strFilename As String = "extracted_meta"
        Map.bFileInUse = True

        Map.SaveMeta(1, strFilename, strPath & "\", True, StartOffset, MetaLength)

        Map.bFileInUse = False
        MsgBox("Finished Extracting.", MsgBoxStyle.Information)
    End Sub

    Private Sub mnuSpaceInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSpaceInsert.Click
        Dim f As New SCNRHelper
        f.Show()
    End Sub
End Class

