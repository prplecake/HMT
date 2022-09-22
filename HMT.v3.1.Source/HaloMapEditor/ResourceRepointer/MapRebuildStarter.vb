'/////////////////////////////////////////////////
'// HMT Map Rebuilding class
'// Written by MonoxideC
'// ---------------------------------------------
'// Rebuild a maps index and metadata areas
'// from an extracted fileset.
'/////////////////////////////////////////////////

Imports HaloMap
Imports HMTLib
Imports System.IO
Imports System.Xml

Public Class MapRebuildStarter
    Inherits System.Windows.Forms.Form

    Public WithEvents mb As MapBuilder

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents XPanderList1 As XPanderControl.XPanderList
    Friend WithEvents XPander1 As XPanderControl.XPander
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents OpenFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnRemoveFile As System.Windows.Forms.Button
    Friend WithEvents btnAddFile As System.Windows.Forms.Button
    Friend WithEvents fileList As System.Windows.Forms.ListBox
    Friend WithEvents btn7 As System.Windows.Forms.Button
    Friend WithEvents txt7 As System.Windows.Forms.TextBox
    Friend WithEvents btn6 As System.Windows.Forms.Button
    Friend WithEvents txt6 As System.Windows.Forms.TextBox
    Friend WithEvents btn5 As System.Windows.Forms.Button
    Friend WithEvents txt5 As System.Windows.Forms.TextBox
    Friend WithEvents btn4 As System.Windows.Forms.Button
    Friend WithEvents txt4 As System.Windows.Forms.TextBox
    Friend WithEvents btn3 As System.Windows.Forms.Button
    Friend WithEvents txt3 As System.Windows.Forms.TextBox
    Friend WithEvents btn10 As System.Windows.Forms.Button
    Friend WithEvents txt10 As System.Windows.Forms.TextBox
    Friend WithEvents btn9 As System.Windows.Forms.Button
    Friend WithEvents txt9 As System.Windows.Forms.TextBox
    Friend WithEvents btn2 As System.Windows.Forms.Button
    Friend WithEvents txt2 As System.Windows.Forms.TextBox
    Friend WithEvents btn1 As System.Windows.Forms.Button
    Friend WithEvents txt1 As System.Windows.Forms.TextBox
    Friend WithEvents btn12 As System.Windows.Forms.Button
    Friend WithEvents txt12 As System.Windows.Forms.TextBox
    Friend WithEvents btn11 As System.Windows.Forms.Button
    Friend WithEvents txt11 As System.Windows.Forms.TextBox
    Friend WithEvents btn8 As System.Windows.Forms.Button
    Friend WithEvents txt8 As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents txtOriginalFile As System.Windows.Forms.TextBox
    Friend WithEvents btnBuild As System.Windows.Forms.Button
    Friend WithEvents XPander2 As XPanderControl.XPander
    Friend WithEvents fileList_ContextMenu As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents Log As System.Windows.Forms.ListBox
    Friend WithEvents pb1 As System.Windows.Forms.ProgressBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(MapRebuildStarter))
        Me.XPanderList1 = New XPanderControl.XPanderList
        Me.XPander2 = New XPanderControl.XPander
        Me.Log = New System.Windows.Forms.ListBox
        Me.pb1 = New System.Windows.Forms.ProgressBar
        Me.XPander1 = New XPanderControl.XPander
        Me.fileList = New System.Windows.Forms.ListBox
        Me.fileList_ContextMenu = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.btnRemoveFile = New System.Windows.Forms.Button
        Me.btnAddFile = New System.Windows.Forms.Button
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnBrowse = New System.Windows.Forms.Button
        Me.txtOriginalFile = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.btn7 = New System.Windows.Forms.Button
        Me.txt7 = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.btn6 = New System.Windows.Forms.Button
        Me.txt6 = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.btn5 = New System.Windows.Forms.Button
        Me.txt5 = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.btn4 = New System.Windows.Forms.Button
        Me.txt4 = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.btn3 = New System.Windows.Forms.Button
        Me.txt3 = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.btn10 = New System.Windows.Forms.Button
        Me.txt10 = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.btn9 = New System.Windows.Forms.Button
        Me.txt9 = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.btn2 = New System.Windows.Forms.Button
        Me.txt2 = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.btn1 = New System.Windows.Forms.Button
        Me.txt1 = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.btn12 = New System.Windows.Forms.Button
        Me.txt12 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btn11 = New System.Windows.Forms.Button
        Me.txt11 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btn8 = New System.Windows.Forms.Button
        Me.txt8 = New System.Windows.Forms.TextBox
        Me.btnBuild = New System.Windows.Forms.Button
        Me.OpenFile = New System.Windows.Forms.OpenFileDialog
        Me.XPanderList1.SuspendLayout()
        Me.XPander2.SuspendLayout()
        Me.XPander1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'XPanderList1
        '
        Me.XPanderList1.AutoScroll = True
        Me.XPanderList1.BackColor = System.Drawing.Color.FromArgb(CType(99, Byte), CType(117, Byte), CType(222, Byte))
        Me.XPanderList1.Controls.Add(Me.XPander2)
        Me.XPanderList1.Controls.Add(Me.XPander1)
        Me.XPanderList1.Location = New System.Drawing.Point(0, 0)
        Me.XPanderList1.Name = "XPanderList1"
        Me.XPanderList1.Size = New System.Drawing.Size(614, 684)
        Me.XPanderList1.TabIndex = 0
        '
        'XPander2
        '
        Me.XPander2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.XPander2.Anchored = True
        Me.XPander2.Animated = False
        Me.XPander2.AnimationTime = 100
        Me.XPander2.BackColor = System.Drawing.Color.Transparent
        Me.XPander2.BorderStyle = System.Windows.Forms.Border3DStyle.Flat
        Me.XPander2.CanToggle = False
        Me.XPander2.CaptionFont = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold)
        Me.XPander2.CaptionStyle = XPanderControl.XPander.CaptionStyleEnum.Normal
        Me.XPander2.CaptionText = "Rebuild Log"
        Me.XPander2.ChevronStyle = XPanderControl.XPander.ChevronStyleEnum.Image
        Me.XPander2.CollapsedHighlightImage = CType(resources.GetObject("XPander2.CollapsedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander2.CollapsedImage = CType(resources.GetObject("XPander2.CollapsedImage"), System.Drawing.Bitmap)
        Me.XPander2.Controls.Add(Me.Log)
        Me.XPander2.Controls.Add(Me.pb1)
        Me.XPander2.DockPadding.Top = 25
        Me.XPander2.ExpandedHighlightImage = CType(resources.GetObject("XPander2.ExpandedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander2.ExpandedImage = CType(resources.GetObject("XPander2.ExpandedImage"), System.Drawing.Bitmap)
        Me.XPander2.Location = New System.Drawing.Point(8, 560)
        Me.XPander2.Name = "XPander2"
        Me.XPander2.ShowTooltips = True
        Me.XPander2.Size = New System.Drawing.Size(596, 32)
        Me.XPander2.TabIndex = 1
        Me.XPander2.Tag = 0
        Me.XPander2.TooltipText = Nothing
        '
        'Log
        '
        Me.Log.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Log.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Log.Location = New System.Drawing.Point(8, 54)
        Me.Log.Name = "Log"
        Me.Log.Size = New System.Drawing.Size(576, 0)
        Me.Log.TabIndex = 41
        '
        'pb1
        '
        Me.pb1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pb1.Location = New System.Drawing.Point(8, 32)
        Me.pb1.Maximum = 7
        Me.pb1.Name = "pb1"
        Me.pb1.Size = New System.Drawing.Size(576, 16)
        Me.pb1.TabIndex = 0
        '
        'XPander1
        '
        Me.XPander1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.XPander1.Anchored = False
        Me.XPander1.Animated = False
        Me.XPander1.AnimationTime = 100
        Me.XPander1.BackColor = System.Drawing.Color.Transparent
        Me.XPander1.BorderStyle = System.Windows.Forms.Border3DStyle.Flat
        Me.XPander1.CanToggle = False
        Me.XPander1.CaptionFont = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold)
        Me.XPander1.CaptionStyle = XPanderControl.XPander.CaptionStyleEnum.Normal
        Me.XPander1.CaptionText = "Halo Map Rebuilder"
        Me.XPander1.ChevronStyle = XPanderControl.XPander.ChevronStyleEnum.Image
        Me.XPander1.CollapsedHighlightImage = CType(resources.GetObject("XPander1.CollapsedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander1.CollapsedImage = CType(resources.GetObject("XPander1.CollapsedImage"), System.Drawing.Bitmap)
        Me.XPander1.Controls.Add(Me.fileList)
        Me.XPander1.Controls.Add(Me.btnRemoveFile)
        Me.XPander1.Controls.Add(Me.btnAddFile)
        Me.XPander1.Controls.Add(Me.Label14)
        Me.XPander1.Controls.Add(Me.Label2)
        Me.XPander1.Controls.Add(Me.btnBrowse)
        Me.XPander1.Controls.Add(Me.txtOriginalFile)
        Me.XPander1.Controls.Add(Me.GroupBox1)
        Me.XPander1.Controls.Add(Me.btnBuild)
        Me.XPander1.DockPadding.Top = 25
        Me.XPander1.ExpandedHighlightImage = CType(resources.GetObject("XPander1.ExpandedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander1.ExpandedImage = CType(resources.GetObject("XPander1.ExpandedImage"), System.Drawing.Bitmap)
        Me.XPander1.Location = New System.Drawing.Point(8, 8)
        Me.XPander1.Name = "XPander1"
        Me.XPander1.ShowTooltips = True
        Me.XPander1.Size = New System.Drawing.Size(596, 544)
        Me.XPander1.TabIndex = 0
        Me.XPander1.Tag = 1
        Me.XPander1.TooltipText = Nothing
        '
        'fileList
        '
        Me.fileList.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.fileList.ContextMenu = Me.fileList_ContextMenu
        Me.fileList.Location = New System.Drawing.Point(16, 408)
        Me.fileList.Name = "fileList"
        Me.fileList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.fileList.Size = New System.Drawing.Size(552, 93)
        Me.fileList.TabIndex = 40
        '
        'fileList_ContextMenu
        '
        Me.fileList_ContextMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Add"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "Delete"
        '
        'btnRemoveFile
        '
        Me.btnRemoveFile.BackColor = System.Drawing.Color.Black
        Me.btnRemoveFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRemoveFile.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnRemoveFile.Location = New System.Drawing.Point(96, 512)
        Me.btnRemoveFile.Name = "btnRemoveFile"
        Me.btnRemoveFile.Size = New System.Drawing.Size(72, 22)
        Me.btnRemoveFile.TabIndex = 38
        Me.btnRemoveFile.Text = "Remove"
        '
        'btnAddFile
        '
        Me.btnAddFile.BackColor = System.Drawing.Color.Black
        Me.btnAddFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddFile.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnAddFile.Location = New System.Drawing.Point(16, 512)
        Me.btnAddFile.Name = "btnAddFile"
        Me.btnAddFile.Size = New System.Drawing.Size(72, 21)
        Me.btnAddFile.TabIndex = 37
        Me.btnAddFile.Text = "Add"
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.Color.Black
        Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label14.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.White
        Me.Label14.Location = New System.Drawing.Point(16, 392)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(72, 16)
        Me.Label14.TabIndex = 35
        Me.Label14.Text = "Additional Files"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Black
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(24, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 16)
        Me.Label2.TabIndex = 32
        Me.Label2.Text = "Original Map"
        '
        'btnBrowse
        '
        Me.btnBrowse.BackColor = System.Drawing.Color.Black
        Me.btnBrowse.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowse.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnBrowse.Location = New System.Drawing.Point(400, 56)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(56, 19)
        Me.btnBrowse.TabIndex = 33
        Me.btnBrowse.Text = "Browse"
        '
        'txtOriginalFile
        '
        Me.txtOriginalFile.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtOriginalFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOriginalFile.Location = New System.Drawing.Point(24, 56)
        Me.txtOriginalFile.Name = "txtOriginalFile"
        Me.txtOriginalFile.Size = New System.Drawing.Size(368, 18)
        Me.txtOriginalFile.TabIndex = 31
        Me.txtOriginalFile.Text = ""
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.White
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.btn7)
        Me.GroupBox1.Controls.Add(Me.txt7)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.btn6)
        Me.GroupBox1.Controls.Add(Me.txt6)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.btn5)
        Me.GroupBox1.Controls.Add(Me.txt5)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.btn4)
        Me.GroupBox1.Controls.Add(Me.txt4)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.btn3)
        Me.GroupBox1.Controls.Add(Me.txt3)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.btn10)
        Me.GroupBox1.Controls.Add(Me.txt10)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.btn9)
        Me.GroupBox1.Controls.Add(Me.txt9)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.btn2)
        Me.GroupBox1.Controls.Add(Me.txt2)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.btn1)
        Me.GroupBox1.Controls.Add(Me.txt1)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.btn12)
        Me.GroupBox1.Controls.Add(Me.txt12)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.btn11)
        Me.GroupBox1.Controls.Add(Me.txt11)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.btn8)
        Me.GroupBox1.Controls.Add(Me.txt8)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 88)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(560, 296)
        Me.GroupBox1.TabIndex = 34
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Required Files"
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(8, 156)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(120, 13)
        Me.Label13.TabIndex = 52
        Me.Label13.Text = "[bitm] Lag Icon"
        '
        'btn7
        '
        Me.btn7.BackColor = System.Drawing.Color.Black
        Me.btn7.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn7.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn7.Location = New System.Drawing.Point(488, 156)
        Me.btn7.Name = "btn7"
        Me.btn7.Size = New System.Drawing.Size(56, 19)
        Me.btn7.TabIndex = 53
        Me.btn7.Text = "Browse"
        '
        'txt7
        '
        Me.txt7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt7.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt7.Location = New System.Drawing.Point(144, 156)
        Me.txt7.Name = "txt7"
        Me.txt7.Size = New System.Drawing.Size(336, 18)
        Me.txt7.TabIndex = 51
        Me.txt7.Text = ""
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(8, 134)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(120, 13)
        Me.Label12.TabIndex = 49
        Me.Label12.Text = "[ustr] Multiplayer Map List"
        '
        'btn6
        '
        Me.btn6.BackColor = System.Drawing.Color.Black
        Me.btn6.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn6.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn6.Location = New System.Drawing.Point(488, 134)
        Me.btn6.Name = "btn6"
        Me.btn6.Size = New System.Drawing.Size(56, 19)
        Me.btn6.TabIndex = 50
        Me.btn6.Text = "Browse"
        '
        'txt6
        '
        Me.txt6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt6.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt6.Location = New System.Drawing.Point(144, 134)
        Me.txt6.Name = "txt6"
        Me.txt6.Size = New System.Drawing.Size(336, 18)
        Me.txt6.TabIndex = 48
        Me.txt6.Text = ""
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(8, 112)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(120, 13)
        Me.Label11.TabIndex = 46
        Me.Label11.Text = "[ustr] Loading String"
        '
        'btn5
        '
        Me.btn5.BackColor = System.Drawing.Color.Black
        Me.btn5.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn5.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn5.Location = New System.Drawing.Point(488, 112)
        Me.btn5.Name = "btn5"
        Me.btn5.Size = New System.Drawing.Size(56, 19)
        Me.btn5.TabIndex = 47
        Me.btn5.Text = "Browse"
        '
        'txt5
        '
        Me.txt5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt5.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt5.Location = New System.Drawing.Point(144, 112)
        Me.txt5.Name = "txt5"
        Me.txt5.Size = New System.Drawing.Size(336, 18)
        Me.txt5.TabIndex = 45
        Me.txt5.Text = ""
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(8, 90)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(120, 13)
        Me.Label10.TabIndex = 43
        Me.Label10.Text = "[bitm] Loading Background"
        '
        'btn4
        '
        Me.btn4.BackColor = System.Drawing.Color.Black
        Me.btn4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn4.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn4.Location = New System.Drawing.Point(488, 90)
        Me.btn4.Name = "btn4"
        Me.btn4.Size = New System.Drawing.Size(56, 19)
        Me.btn4.TabIndex = 44
        Me.btn4.Text = "Browse"
        '
        'txt4
        '
        Me.txt4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt4.Location = New System.Drawing.Point(144, 90)
        Me.txt4.Name = "txt4"
        Me.txt4.Size = New System.Drawing.Size(336, 18)
        Me.txt4.TabIndex = 42
        Me.txt4.Text = ""
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(8, 68)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(120, 13)
        Me.Label9.TabIndex = 40
        Me.Label9.Text = "[tagc] All Scenario"
        '
        'btn3
        '
        Me.btn3.BackColor = System.Drawing.Color.Black
        Me.btn3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn3.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn3.Location = New System.Drawing.Point(488, 68)
        Me.btn3.Name = "btn3"
        Me.btn3.Size = New System.Drawing.Size(56, 19)
        Me.btn3.TabIndex = 41
        Me.btn3.Text = "Browse"
        '
        'txt3
        '
        Me.txt3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt3.Location = New System.Drawing.Point(144, 68)
        Me.txt3.Name = "txt3"
        Me.txt3.Size = New System.Drawing.Size(336, 18)
        Me.txt3.TabIndex = 39
        Me.txt3.Text = ""
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(8, 222)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(120, 13)
        Me.Label8.TabIndex = 37
        Me.Label8.Text = "[snd!] UI Back Sound"
        '
        'btn10
        '
        Me.btn10.BackColor = System.Drawing.Color.Black
        Me.btn10.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn10.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn10.Location = New System.Drawing.Point(488, 222)
        Me.btn10.Name = "btn10"
        Me.btn10.Size = New System.Drawing.Size(56, 19)
        Me.btn10.TabIndex = 38
        Me.btn10.Text = "Browse"
        '
        'txt10
        '
        Me.txt10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt10.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt10.Location = New System.Drawing.Point(144, 222)
        Me.txt10.Name = "txt10"
        Me.txt10.Size = New System.Drawing.Size(336, 18)
        Me.txt10.TabIndex = 36
        Me.txt10.Text = ""
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(8, 200)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(120, 13)
        Me.Label7.TabIndex = 34
        Me.Label7.Text = "[snd!] UI Forward Sound"
        '
        'btn9
        '
        Me.btn9.BackColor = System.Drawing.Color.Black
        Me.btn9.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn9.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn9.Location = New System.Drawing.Point(488, 200)
        Me.btn9.Name = "btn9"
        Me.btn9.Size = New System.Drawing.Size(56, 19)
        Me.btn9.TabIndex = 35
        Me.btn9.Text = "Browse"
        '
        'txt9
        '
        Me.txt9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt9.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt9.Location = New System.Drawing.Point(144, 200)
        Me.txt9.Name = "txt9"
        Me.txt9.Size = New System.Drawing.Size(336, 18)
        Me.txt9.TabIndex = 33
        Me.txt9.Text = ""
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(8, 46)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(120, 13)
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "[matg] Game Globals"
        '
        'btn2
        '
        Me.btn2.BackColor = System.Drawing.Color.Black
        Me.btn2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn2.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn2.Location = New System.Drawing.Point(488, 46)
        Me.btn2.Name = "btn2"
        Me.btn2.Size = New System.Drawing.Size(56, 19)
        Me.btn2.TabIndex = 32
        Me.btn2.Text = "Browse"
        '
        'txt2
        '
        Me.txt2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt2.Location = New System.Drawing.Point(144, 46)
        Me.txt2.Name = "txt2"
        Me.txt2.Size = New System.Drawing.Size(336, 18)
        Me.txt2.TabIndex = 30
        Me.txt2.Text = ""
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(8, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(120, 13)
        Me.Label5.TabIndex = 28
        Me.Label5.Text = "[scnr] Scenario"
        '
        'btn1
        '
        Me.btn1.BackColor = System.Drawing.Color.Black
        Me.btn1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn1.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn1.Location = New System.Drawing.Point(488, 24)
        Me.btn1.Name = "btn1"
        Me.btn1.Size = New System.Drawing.Size(56, 19)
        Me.btn1.TabIndex = 29
        Me.btn1.Text = "Browse"
        '
        'txt1
        '
        Me.txt1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt1.Location = New System.Drawing.Point(144, 24)
        Me.txt1.Name = "txt1"
        Me.txt1.Size = New System.Drawing.Size(336, 18)
        Me.txt1.TabIndex = 27
        Me.txt1.Text = ""
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(8, 266)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(120, 13)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "[sbsp] BSP Chunk"
        '
        'btn12
        '
        Me.btn12.BackColor = System.Drawing.Color.Black
        Me.btn12.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn12.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn12.Location = New System.Drawing.Point(488, 266)
        Me.btn12.Name = "btn12"
        Me.btn12.Size = New System.Drawing.Size(56, 19)
        Me.btn12.TabIndex = 26
        Me.btn12.Text = "Browse"
        '
        'txt12
        '
        Me.txt12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt12.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt12.Location = New System.Drawing.Point(144, 266)
        Me.txt12.Name = "txt12"
        Me.txt12.Size = New System.Drawing.Size(336, 18)
        Me.txt12.TabIndex = 24
        Me.txt12.Text = ""
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(8, 244)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 13)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "[tagc] Multiplayer Scenario"
        '
        'btn11
        '
        Me.btn11.BackColor = System.Drawing.Color.Black
        Me.btn11.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn11.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn11.Location = New System.Drawing.Point(488, 244)
        Me.btn11.Name = "btn11"
        Me.btn11.Size = New System.Drawing.Size(56, 19)
        Me.btn11.TabIndex = 23
        Me.btn11.Text = "Browse"
        '
        'txt11
        '
        Me.txt11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt11.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt11.Location = New System.Drawing.Point(144, 244)
        Me.txt11.Name = "txt11"
        Me.txt11.Size = New System.Drawing.Size(336, 18)
        Me.txt11.TabIndex = 21
        Me.txt11.Text = ""
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 178)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 13)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "[snd!] UI Cursor Sound"
        '
        'btn8
        '
        Me.btn8.BackColor = System.Drawing.Color.Black
        Me.btn8.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn8.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn8.Location = New System.Drawing.Point(488, 178)
        Me.btn8.Name = "btn8"
        Me.btn8.Size = New System.Drawing.Size(56, 19)
        Me.btn8.TabIndex = 20
        Me.btn8.Text = "Browse"
        '
        'txt8
        '
        Me.txt8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt8.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt8.Location = New System.Drawing.Point(144, 178)
        Me.txt8.Name = "txt8"
        Me.txt8.Size = New System.Drawing.Size(336, 18)
        Me.txt8.TabIndex = 18
        Me.txt8.Text = ""
        '
        'btnBuild
        '
        Me.btnBuild.BackColor = System.Drawing.Color.Black
        Me.btnBuild.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBuild.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnBuild.Location = New System.Drawing.Point(496, 48)
        Me.btnBuild.Name = "btnBuild"
        Me.btnBuild.Size = New System.Drawing.Size(80, 23)
        Me.btnBuild.TabIndex = 29
        Me.btnBuild.Text = "Build"
        '
        'MapRebuildStarter
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(614, 598)
        Me.Controls.Add(Me.XPanderList1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "MapRebuildStarter"
        Me.Text = "Halo Map Tools"
        Me.XPanderList1.ResumeLayout(False)
        Me.XPander2.ResumeLayout(False)
        Me.XPander1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub UpdatePB(ByVal v As Integer) Handles mb.UpdateProgressBar
        pb1.Value = v
    End Sub

    Private Sub UpdateLog(ByVal s As String) Handles mb.AddToLog
        Log.Items.Add(s)
    End Sub

    Private Sub FormClose(ByVal sender As Object, _
        ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        fMain.Show()
    End Sub

    Private Sub btnAddFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddFile.Click
        OpenFile.DefaultExt = ".meta"
        OpenFile.AddExtension = True
        OpenFile.Multiselect = True
        OpenFile.Filter = "Binary Metadata (*.meta)|*.meta"
        OpenFile.ShowDialog()

        For Each s As String In OpenFile.FileNames
            fileList.Items.Add(s)
        Next
    End Sub

    Private Sub btnRemoveFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveFile.Click
        Dim s() As String
        ReDim s(fileList.SelectedItems.Count)
        Dim x As Integer = 0
        For Each str As String In fileList.SelectedItems
            s(x) = str
            x += 1
        Next
        For x = 0 To s.Length - 1
            fileList.Items.Remove(s(x))
        Next
    End Sub

    Private Sub btn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1.Click, _
                                                                                               btn2.Click, _
                                                                                               btn3.Click, _
                                                                                               btn4.Click, _
                                                                                               btn5.Click, _
                                                                                               btn6.Click, _
                                                                                               btn7.Click, _
                                                                                               btn8.Click, _
                                                                                               btn9.Click, _
                                                                                               btn10.Click, _
                                                                                               btn11.Click, _
                                                                                               btn12.Click
        'Browse for a file and place it into the approprate text box
        OpenFile.DefaultExt = ".meta"
        OpenFile.AddExtension = True
        OpenFile.Multiselect = False
        OpenFile.Filter = "Binary Metadata (*.meta)|*.meta"
        If OpenFile.ShowDialog() = DialogResult.Cancel Then Exit Sub

        Select Case CType(sender, Button).Name
            Case "btn1"
                txt1.Text = OpenFile.FileName
            Case "btn2"
                txt2.Text = OpenFile.FileName
            Case "btn3"
                txt3.Text = OpenFile.FileName
            Case "btn4"
                txt4.Text = OpenFile.FileName
            Case "btn5"
                txt5.Text = OpenFile.FileName
            Case "btn6"
                txt6.Text = OpenFile.FileName
            Case "btn7"
                txt7.Text = OpenFile.FileName
            Case "btn8"
                txt8.Text = OpenFile.FileName
            Case "btn9"
                txt9.Text = OpenFile.FileName
            Case "btn10"
                txt10.Text = OpenFile.FileName
            Case "btn11"
                txt11.Text = OpenFile.FileName
            Case "btn12"
                txt12.Text = OpenFile.FileName
        End Select

    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        'Browse for a file and place it into the approprate text box
        OpenFile.DefaultExt = ".map"
        OpenFile.AddExtension = True
        OpenFile.Multiselect = False
        OpenFile.Filter = "Halo Map File (*.map)|*.map"
        If OpenFile.ShowDialog() = DialogResult.Cancel Then Exit Sub

        txtOriginalFile.Text = OpenFile.FileName

        If MsgBox("Do you wish to automatically fill in the required files" & vbCrLf & _
                "based on the path and name of the selected map file?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

            Dim mapname As String = HMTLib.getLevel(txtOriginalFile.Text, 0) 'The map filename
            mapname = Mid(mapname, 1, mapname.Length - 4)
            Dim strSource As String = Mid(txtOriginalFile.Text, 1, txtOriginalFile.Text.Length - (mapname.Length + 4))

            txt1.Text = strSource & "levels\test\" & mapname & "\" & mapname & ".scnr.meta"
            txt2.Text = strSource & "globals\globals.matg.meta"
            txt3.Text = strSource & "ui\ui_tags_loaded_all_scenario_types.tagc.meta"
            txt4.Text = strSource & "ui\shell\bitmaps\background.bitm.meta"
            txt5.Text = strSource & "ui\shell\strings\loading.ustr.meta"
            txt6.Text = strSource & "ui\shell\main_menu\mp_map_list.ustr.meta"
            txt7.Text = strSource & "ui\shell\bitmaps\trouble_brewing.bitm.meta"
            txt8.Text = strSource & "sound\sfx\ui\cursor.snd!.meta"
            txt9.Text = strSource & "sound\sfx\ui\forward.snd!.meta"
            txt10.Text = strSource & "sound\sfx\ui\back.snd!.meta"
            txt11.Text = strSource & "ui\ui_tags_loaded_multiplayer_scenario_type.tagc.meta"
            txt12.Text = strSource & "levels\test\" & mapname & "\" & mapname & ".sbsp.meta"
        End If

    End Sub

    Private Sub btnBuild_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBuild.Click
        'Call the Buildmap Function
        Dim s() As String
        Dim m() As String
        Dim iCount As Integer = 12
        iCount += fileList.Items.Count

        ReDim s(iCount - 1)
        s(0) = txt1.Text
        s(1) = txt2.Text
        s(2) = txt3.Text
        s(3) = txt4.Text
        s(4) = txt5.Text
        s(5) = txt6.Text
        s(6) = txt7.Text
        s(7) = txt8.Text
        s(8) = txt9.Text
        s(9) = txt10.Text
        s(10) = txt11.Text
        s(11) = txt12.Text

        If iCount > 12 Then
            For x As Integer = 12 To iCount - 1
                s(x) = fileList.Items(x - 12)
            Next
        End If

        ReDim m(0)
        m(0) = "C:\Documents and Settings\Administrator\Desktop\BG3\characters\cyborg\cyborg.mod2.raw.xml"

        XPander1.Collapse()
        pb1.Value = 0
        Log.Items.Clear()
        mb = New MapBuilder
        mb.BuildMap(txtOriginalFile.Text, s) ', m)
        XPander1.Collapse()
        pb1.Value = 0
        Me.Close()
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        OpenFile.DefaultExt = ".meta"
        OpenFile.AddExtension = True
        OpenFile.Multiselect = True
        OpenFile.Filter = "Binary Metadata (*.meta)|*.meta"
        OpenFile.ShowDialog()

        For Each s As String In OpenFile.FileNames
            fileList.Items.Add(s)
        Next
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Dim s() As String
        ReDim s(fileList.SelectedItems.Count)
        Dim x As Integer = 0
        For Each str As String In fileList.SelectedItems
            s(x) = str
            x += 1
        Next
        For x = 0 To s.Length - 1
            fileList.Items.Remove(s(x))
        Next
    End Sub
End Class

'////////////////////////////////////////////////////////////////////////////////
'// This class will rebuild a map's index, filedata, and metadata sections from
'// a set of extracted meta files and meta XML structure files, assigning idents
'// and fixing reflexive offsets in the process.  
'////////////////////////////////////////////////////////////////////////////////
Public Class MapBuilder

    Public Event UpdateProgressBar(ByVal v As Integer)
    Public Event AddToLog(ByVal s As String)

    'Used for storing data about each item that is processed
    Public Structure META_FILE_STRUCT
        Public Filename As String
        Public TagClass As String
        Public MetaOffset As Long
        Public ID As Long
        Public Processed As Boolean
    End Structure

    Public strSource As String 'The source folder where processing will be done
    Public Index As MapIndex
    Public metaFile() As META_FILE_STRUCT
    Public IndiceInjectionPoint As Integer
    Public ModelInjectionPoint As Integer

    'Used in building the TreeBuilt.txt output file
    'This may be converted to XML at some point.. or removed.. who knows...
    Public sw3 As New StreamWriter(New FileStream("C:\BuildTree.txt", FileMode.Create))
    Public rDepth As Integer = -1 'Used in generating TreeBuild.txt file

    'The original map file - we use data from it in generating the new map
    Private tmpMap As New Map

    '////////////////////////////////////////////////////////////////
    '// This is the main function that calls all other function
    '////////////////////////////////////////////////////////////////
    Public Sub BuildMap(ByVal strSourceMap As String, ByVal strBuildFiles() As String, Optional ByVal strModels() As String = Nothing)
        Dim TagOrder() As String                               'Processing order
        Dim strMapfile As String = getLevel(strSourceMap, 0)   'The map filename
        strSource = strSourceMap.Trim(strMapfile.ToCharArray)  'Get the source folder

        Dim rfp As New RecursiveFileProcessor  'Our master .meta file list
        Dim t As DateTime = Now                'Used as a start time for measurement

        'Recursively get a list of all meta files in the path
        rfp.ProcessDirectory(strSource, ".meta")

        'Copy the rfp results to the metaFile array
        Dim FileareaSize As Integer = 0
        ReDim metaFile(rfp.fileNumber - 1)
        For x As Integer = 0 To rfp.fileNumber - 1
            metaFile(x).Filename = rfp.FileList(x + 1).filename
            metaFile(x).TagClass = rfp.FileList(x + 1).tag
            metaFile(x).Processed = False
        Next
        rfp = Nothing 'Destroy this to free memory

        'Open the original map file so that we can examine it's structure
        tmpMap.OpenFile(strSourceMap)
        tmpMap.ReadFile(tmpMap.eOpenFile.PC)

        Dim fi As FileInfo
        Dim sTemp As String

        'Create the new map file
        Dim bw As New BinaryWriter(New FileStream(strSourceMap & ".rebuild.map", FileMode.Create))
        Dim bChunk() As Byte 'Byte array used to store the file data

        'Create the index
        Index = New MapIndex(tmpMap)

        'Find the size of the sbsp
        Dim sbspSize As Integer
        fi = New FileInfo(strBuildFiles(strBuildFiles.Length - 1))
        sbspSize = fi.Length
        If sbspSize Mod 4 > 0 Then sbspSize += (4 - (sbspSize Mod 4))

        'Go through the list of files (in tag order) and add files
        Console.WriteLine("Creating Map Index")
        RaiseEvent AddToLog("Creating map Index")
        RaiseEvent UpdateProgressBar(1)
        Dim processResult As Long = 0
        For tagIndex As Integer = 0 To strBuildFiles.Length - 1
            'Find unprocessed files with matching tags
            For metaIndex As Integer = 0 To metaFile.Length - 1
                If metaFile(metaIndex).Filename = (strBuildFiles(tagIndex)) And _
                    Not metaFile(metaIndex).Processed Then
                    'Process this tag
                    Application.DoEvents()
                    If metaFile(metaIndex).TagClass = "sbsp" Then
                        processResult = ProcessFile(metaFile(metaIndex))
                    Else
                        processResult = ProcessFile(metaFile(metaIndex))
                    End If
                End If
                If processResult = -1 Then
                    bw.Close()
                    tmpMap.CloseFile()
                    Exit Sub
                End If
            Next
        Next
        RaiseEvent UpdateProgressBar(2)
        Application.DoEvents()

        'Create a new Map object for the new file
        Dim NewMap As New HaloMap.Map

        Dim newStartOffset As Integer

        If newStartOffset > 0 Then
            If Not MsgBox("Do you wish to use the modified start offset?" & vbCrLf & _
                "Offset: " & newStartOffset, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then newStartOffset = 0
        End If

        'Write the file header so that the file can be correctly opened
        bw.Seek(0, SeekOrigin.Begin)
        NewMap.fileHeader = tmpMap.fileHeader
        If newStartOffset > 0 Then tmpMap.fileHeader.offset_to_index_decomp = newStartOffset
        NewMap.fileHeader.offset_to_index_decomp = tmpMap.fileHeader.offset_to_index_decomp
        NewMap.fileHeader.WriteStruct(bw)

        bw.Seek(0, SeekOrigin.Begin)
        NewMap.indexHeader = Index.Header
        NewMap.indexItem = Index.Index
        NewMap.OpenFile(bw.BaseStream)
        NewMap.intMagic = tmpMap.intMagic
        If newStartOffset > 0 Then tmpMap.intMagic = (1078198312 - (NewMap.fileHeader.offset_to_index_decomp + 40))
        NewMap.intMagic = tmpMap.intMagic
        'Based on the offset_to_decomp and the metadata size, go the the calculated filedata start offset
        Dim FilenameAreaStartOffset As Integer = tmpMap.fileHeader.offset_to_index_decomp + 8 'PC has an extra 4 bytes
        FilenameAreaStartOffset += (Index.Header.tagcount + 1) * 32
        If FilenameAreaStartOffset Mod 4 > 0 Then FilenameAreaStartOffset += (4 - (FilenameAreaStartOffset Mod 4))
        bw.BaseStream.Seek(FilenameAreaStartOffset, SeekOrigin.Begin)

        'Write all of the filename in the correct order (according the to index)
        'While we are writing them, we need to update the FileName offset in the index
        Console.WriteLine("Writing Filenames")
        RaiseEvent AddToLog("Writing Filenames")
        RaiseEvent UpdateProgressBar(3)

        Dim bBuildFile As Boolean = False
        Dim strTempFile As String
        For x As Integer = 1 To Index.Header.tagcount
            NewMap.indexItem(x).fileData.stringoffset = CLng(bw.BaseStream.Position) + tmpMap.intMagic
            'Search the strBuildFiles array for the filename
            bBuildFile = False
            For z As Integer = 0 To strBuildFiles.Length - 1
                strTempFile = Mid(strBuildFiles(z), strSource.Length + 1, strBuildFiles(z).Length - strSource.Length)
                strTempFile = Mid(strTempFile, 1, strTempFile.Length - 10)
                'Console.WriteLine(Index.Index(x).filePath)
                If strTempFile = Index.Index(x).filePath _
                    Or Mid(Index.Index(x).fileData.tagclass.ToUpper, 1, 4) = "ALED" _
                    Or Mid(Index.Index(x).fileData.tagclass.ToUpper, 1, 4) = " TMH" _
                    Or Mid(Index.Index(x).fileData.tagclass.ToUpper, 1, 4) = "CGAT" _
                    Or Mid(Index.Index(x).fileData.tagclass.ToUpper, 1, 4) = "RTSU" Then
                    bBuildFile = True
                End If
            Next

            'Select Case bBuildFile
            'Case True
            bw.Write(Index.Index(x).filePath.ToCharArray)
            bw.Write(CByte(0)) 'Add the Null terminator
            'Case False
            'bw.Write(CByte(0)) 'Add the Null terminator
            'End Select
        Next
        RaiseEvent UpdateProgressBar(4)
        Application.DoEvents()

        'Go to the next dword boundary and store this location as the start offset
        'for the metadata area
        Dim MetaAreaStartOffset As Integer = bw.BaseStream.Position
        If MetaAreaStartOffset Mod 4 > 0 Then MetaAreaStartOffset += 4 - (MetaAreaStartOffset Mod 4)
        bw.BaseStream.Seek(MetaAreaStartOffset, SeekOrigin.Begin)

        Dim intStreamPosition As Integer = MetaAreaStartOffset

        'We need to read in all of the metadata files in Index order and write them to the metadata area
        'During this process, we need to open their structure XML file again and correct the reflexive
        'offsets and the idents.  This can be accomplished by calling the importMeta method of the map class.
        Console.WriteLine("Writing Metadata")
        RaiseEvent AddToLog("Writing Metadata")
        RaiseEvent UpdateProgressBar(5)
        For x As Integer = 1 To NewMap.indexHeader.tagcount
            'Find the corresponding MetaFile so we know the correct file to open
            For y As Integer = 0 To metaFile.Length - 1
                If metaFile(y).ID = NewMap.indexItem(x).fileData.id Then
                    'We found it, hells yeah
                    Application.DoEvents()
                    If Mid(metaFile(y).TagClass, 1, 4) = "psbs" Then
                        NewMap.InjectMeta(x, strSource & metaFile(y).Filename & "." & _
                            reverseString(Mid(metaFile(y).TagClass, 1, 4)) & ".meta", 0) 'SBSP goes at 0x800 (2048), so change this back at somt point.
                    Else
                        bw.BaseStream.Seek(intStreamPosition, SeekOrigin.Begin)
                        NewMap.indexItem(x).fileData.offset = CLng(bw.BaseStream.Position) + tmpMap.intMagic
                        NewMap.indexItem(x).magic_metadata_offset = bw.BaseStream.Position
                        NewMap.InjectMeta(x, strSource & metaFile(y).Filename & "." & _
                                reverseString(Mid(metaFile(y).TagClass, 1, 4)) & ".meta", bw.BaseStream.Position, _
                                ModelInjectionPoint, IndiceInjectionPoint)
                        intStreamPosition = bw.BaseStream.Position
                    End If
                End If
            Next
        Next

        Dim intFileLength As Integer = bw.BaseStream.Position

        'Now all we need to do is wrte the index to the map, one element at a time
        '** Note: Right now we aren't correcting the Index header and the file header to adjust for
        'differences in Meta Size and File size - but that will be simple, and will be added next
        RaiseEvent AddToLog("Writing Index to File")
        RaiseEvent UpdateProgressBar(6)
        bw.BaseStream.Seek(NewMap.fileHeader.offset_to_index_decomp, SeekOrigin.Begin)
        NewMap.indexHeader.WriteStruct(bw, True)
        For x As Integer = 1 To NewMap.indexHeader.tagcount
            NewMap.indexItem(x).fileData.WriteStruct(bw)
        Next

        'And finally, we need to update the fileHeader and IndexHeader to the new sizes.
        bw.Seek(0, SeekOrigin.Begin)
        NewMap.fileHeader.decomp_len = intFileLength
        NewMap.fileHeader.MetadataSize = intFileLength - NewMap.fileHeader.offset_to_index_decomp
        NewMap.fileHeader.WriteStruct(bw)
        bw.Seek(0, SeekOrigin.Begin)
        bw.Write("daeh".ToCharArray)

        RaiseEvent AddToLog("Done!")
        RaiseEvent UpdateProgressBar(7)
        bw.Close()
        Console.WriteLine("Map compile completed in " & Convert.ToString(Now.Subtract(t)) & _
                            " seconds (" & Index.Header.tagcount & " tags processed" & ")")
        RaiseEvent AddToLog("Map compile completed in " & Convert.ToString(Now.Subtract(t)) & _
                            " seconds (" & Index.Header.tagcount & " tags processed" & ")")
        tmpMap.CloseFile()
        sw3.Close()
        Console.WriteLine("Map Rebuild Complete")
        MsgBox("Map Rebuild Completed Successfully")
    End Sub

    'Send a tagclass and filename, and return the array index where it is located
    Private Function LocateMetaFileByTagClassAndFilename(ByVal strTagClass As String, _
        ByVal strFilename As String) As Integer
        For x As Integer = 0 To metaFile.Length - 1
            If metaFile(x).Filename = strSource & strFilename Then
                If metaFile(x).TagClass = reverseString(strTagClass.Substring(0, 4)) Then Return x
                'Console.WriteLine(metaFile(x).TagClass)
            End If
        Next
        Return -1
    End Function

    Public Structure Dependency
        Public OriginalID As Long
        Public Location As Long
        Public TagClass As String
        Public Filename As String
    End Structure

    Public Structure Reflexive
        Public StartOffset As Integer
        Public EndOffset As Integer
        Public Chunk_Count As Integer
        Public Location As Integer
        Public Processed As Boolean
    End Structure

    Public Structure XMLItemList
        Public Offset As Integer
        Public Processed As Boolean
    End Structure

    '////////////////////////////////////////////////////////
    '// This recursive function will process a tag file and
    '// write the processed data to metaOffset in the meta
    '// area of the map file
    '////////////////////////////////////////////////////////
    Public Function ProcessFile(ByRef mFile As META_FILE_STRUCT) As Long
        Dim br As New BinaryReader(New FileStream(mFile.Filename, FileMode.Open, FileAccess.ReadWrite))
        Dim xmlD As New XmlDocument
        Dim xmlN As XmlNode
        Dim bChunk() As Integer
        Dim intID As Long
        Dim dList() As Dependency
        Dim rList() As Reflexive
        Dim xList() As XMLItemList

        'Read in the binary .meta file
        ReDim bChunk(br.BaseStream.Length / 4)
        For x As Integer = 1 To bChunk.Length - 1
            bChunk(x) = br.ReadInt32
        Next
        br.Close()  'No reason to have this open anymore

        'Open the XML structure file and use its contents to process the meta
        If Not File.Exists(mFile.Filename.Remove((mFile.Filename.Length - 5), 5) & ".xml") Then
            Console.WriteLine("Structure file not found for " & mFile.Filename & " - File skipped.")
            Exit Function
        Else
            'Console.WriteLine("Reading from " & mFile.Filename)
        End If
        xmlD.Load(mFile.Filename.Remove((mFile.Filename.Length - 5), 5) & ".xml")

        'We don't need the map name, but it's in there just in case
        mFile.TagClass = xmlD.SelectSingleNode("/Results/Tag").InnerText
        mFile.Filename = xmlD.SelectSingleNode("/Results/Filename").InnerText

        'Add this item to the map index
        'Increment the meta cursor past the current meta block
        intID = Index.Add(mFile)
        mFile.ID = intID

        Dim strMapName As String = ""
        Dim strTagClass As String = mFile.TagClass
        Dim strFilename As String = mFile.Filename
        Dim strDependencyFilename As String
        Dim strDependencyTagClass As String
        Dim intDependencyID
        Dim StrType As String
        Dim intLocation, intTranslation As Integer
        Dim strID, strTemp2 As String
        Dim xmlNList As XmlNodeList

        'Read in all of the dependencies and store them in a array.
        xmlNList = xmlD.SelectNodes("/Results/Dependency")
        Dim dependencyCount As Integer = 0
        ReDim dList(xmlNList.Count)
        For Each xmlN In xmlNList
            dependencyCount += 1
            dList(dependencyCount).Location = strHexToDec(xmlN.ChildNodes.Item(0).InnerText)
            dList(dependencyCount).TagClass = xmlN.ChildNodes.Item(1).InnerText
            dList(dependencyCount).Filename = xmlN.ChildNodes.Item(2).InnerText
            If Mid(dList(dependencyCount).TagClass, 1, 4) = "psbs" Then
                dependencyCount -= 1
            End If
        Next

        'Read in all of the reflexives so we can determine blocks
        xmlNList = xmlD.SelectNodes("/Results/Reflexive")
        Dim reflexiveCount As Integer = 1
        ReDim rList(xmlNList.Count + 1)
        For Each xmlN In xmlNList
            reflexiveCount += 1
            rList(reflexiveCount).Location = strHexToDec(xmlN.ChildNodes.Item(0).InnerText)
            rList(reflexiveCount).StartOffset = strHexToDec(xmlN.ChildNodes.Item(1).InnerText)
            rList(reflexiveCount - 1).EndOffset = rList(reflexiveCount).StartOffset - 1
            rList(reflexiveCount).Chunk_Count = xmlN.ChildNodes.Item(2).InnerText
            If Mid(dList(dependencyCount).TagClass, 1, 4) = "psbs" Then
                dependencyCount -= 1
            End If
        Next
        rList(reflexiveCount).EndOffset = (bChunk.Length * 4)

        'At this point, we have rList - an array containing startOffset and endOffset for
        'all reflexives in the xml file

        'Go through the list one last time and make a sequential list of all of the nodes in the file.
        Dim XMLNodeCount As Integer = 0
        xmlNList = xmlD.SelectNodes("/Results/*")
        For Each xmlN In xmlNList
            If xmlN.Name = "Dependency" Or xmlN.Name = "Reflexive" Then
                XMLNodeCount += 1
                ReDim Preserve xList(XMLNodeCount)
                xList(XMLNodeCount).Offset = strHexToDec(xmlN.ChildNodes.Item(0).InnerText)
                xList(XMLNodeCount).Processed = False
            End If
        Next

        Dim xChunkCount As Integer
        Dim xTagList As XmlNodeList
        Dim depIndex As Integer
        Dim lngFilenameOffset As Long
        Dim strReflexiveFilename As String
        Dim reflexiveIndex As Integer
        Dim ChunkNumber As Integer = 0
        Dim FirstDependencyInChunkRange As Integer

        xmlNList = xmlD.SelectNodes("/Results/*")
        For Each xmlN In xmlNList
            Select Case xmlN.Name
                Case "Reflexive"
                    intLocation = strHexToDec(xmlN.ChildNodes.Item(0).InnerText)
                    intTranslation = strHexToDec(xmlN.ChildNodes.Item(1).InnerText)
                    Application.DoEvents()
                    xChunkCount = xmlN.ChildNodes.Item(2).InnerText

                    'Locate this reflexive in the rList based on its startOffset
                    For z As Integer = 1 To rList.Length - 1
                        If rList(z).StartOffset = intTranslation Then
                            reflexiveIndex = z
                            Exit For
                        End If
                    Next
                    ProcessReflexive(reflexiveIndex, intLocation, intTranslation, xList, dList, rList)
                Case "Dependency"
                    Application.DoEvents()
                    intLocation = strHexToDec(xmlN.ChildNodes.Item(0).InnerText)
                    strDependencyTagClass = xmlN.ChildNodes.Item(1).InnerText
                    strDependencyFilename = xmlN.ChildNodes.Item(2).InnerText
                    'Find the dependency in the array based on its location
                    For x As Integer = 1 To dependencyCount
                        If dList(x).Location = intLocation Then
                            depIndex = x
                            Exit For
                        End If
                    Next
                    If Not ProcessDependency(dList(depIndex)) Then Return -1
            End Select
        Next
        Application.DoEvents()

        mFile.Processed = True
        Return intID
    End Function

    '////////////////////////////////////////////
    '// Recursive Reflexive processing function
    '////////////////////////////////////////////
    Public Function ProcessReflexive(ByRef ReflexiveIndex As Integer, ByRef ItemOffset As Integer, ByRef intTranslation As Integer, _
        ByRef xList() As XMLItemList, ByRef dList() As Dependency, ByRef rList() As Reflexive)
        Dim chunkNumber As Integer
        If Not rList(ReflexiveIndex).Processed Then
            'Now we need to process dependencies in list order, if they exist
            'Go through the dependency list and add all dependencies who's Locations are
            'between StartOffset and EndOffset
            chunkNumber = 0
            For x As Integer = 1 To xList.Length - 1
                If xList(x).Offset >= rList(ReflexiveIndex).StartOffset And _
                    xList(x).Offset <= rList(ReflexiveIndex).EndOffset Then
                    'Ok - we know the offset of an ordered item in the file that matches.
                    'Now, we need to determine if it's a reflexive or a dependency and act accordingly
                    For z As Integer = 1 To dList.Length - 1
                        If dList(z).Location = xList(x).Offset Then
                            'Yey we found it :)
                            'We don't want to add the sbsp dependency from here - only from the end of the process
                            If Mid(dList(z).TagClass, 1, 4) <> "psbs" Then ProcessDependency(dList(z))
                            Exit For
                        End If
                    Next
                    For z As Integer = ReflexiveIndex + 1 To rList.Length - 1
                        If rList(z).Location = xList(x).Offset Then
                            'Yey we found it :)
                            Application.DoEvents()
                            ProcessReflexive(z, xList(x).Offset, rList(z).StartOffset, xList, dList, rList)
                            rList(ReflexiveIndex).Processed = True
                            Exit For
                        End If
                    Next
                End If
            Next
        End If
    End Function

    '////////////////////////////////////////////
    '// External Dependency processing function
    '////////////////////////////////////////////
    Private Function ProcessDependency(ByRef dep As Dependency) As Boolean
        Dim intIndex As Integer
        Dim strID As String
        Dim dependencyID As Long

        If Hex(dep.OriginalID) = "FFFFFFFF" Then Exit Function 'Blank dependency

        rDepth += 1

        'Act appropriatele depending on if the map index contains the tag
        If Not Index.Contains(dep.TagClass, dep.Filename) Then
            intIndex = LocateMetaFileByTagClassAndFilename(dep.TagClass, _
                dep.Filename & "." & reverseString(dep.TagClass.Substring(0, 4)) & ".meta")
            sw3.WriteLine(Space(rDepth * 4) & "[" & reverseString(Mid(dep.TagClass, 1, 4)) & "] " & dep.Filename)
            If intIndex > -1 Then
                dependencyID = ProcessFile(metaFile(intIndex))
            Else
                MsgBox("[" & reverseString(Mid(dep.TagClass, 1, 4)) & "] " & dep.Filename & vbCrLf & _
                    "A dependency was referenced that does not exist!" & vbCrLf & _
                    "The map file may not function.")
                Return False
            End If
        Else
            dependencyID = Index.LocateByTagClassAndFilename(dep.TagClass, dep.Filename)
        End If
        Application.DoEvents()
        rDepth -= 1
        Return True
    End Function


    '//////////////////////////////////////////////////////
    '// This class encapsulates an element of a map index
    '//////////////////////////////////////////////////////
    Public Class MapIndex

        Private currentID1 As Long = 0
        Private currentID2 As Long = 57716
        Private intMagic As Long
        Public IndexHeaderSize As Integer
        Public Header As New Map.INDEX_HEADER_STRUCT
        Public Index() As Map.INDEX_ITEM_EXPANDED_STRUCT

        '//////////////////////
        '// Class constructor
        '//////////////////////
        Public Sub New(ByRef tmpMap As HaloMap.Map)
            'Calculate the indexHeader size
            IndexHeaderSize = 32 + IIf(tmpMap.pc, 4, 0) '32 bytes - add 4 more bytes if this is a PC map

            'Create the Header objects
            Header = tmpMap.indexHeader
            Header.tagcount = 0
            intMagic = tmpMap.intMagic
        End Sub

        '///////////////////////////////////////////////////
        '// See if the index contains a tag matching the
        '// supplied criteria
        '///////////////////////////////////////////////////
        Public Function Contains(ByVal strDependencyTagClass As String, ByVal strDependencyFilename As String) As Boolean
            For x As Integer = 1 To Header.tagcount
                If Index(x).filePath = strDependencyFilename Then
                    If Index(x).fileData.tagclass = strDependencyTagClass Then
                        Return True
                    End If
                End If
            Next
            Return False
        End Function

        '///////////////////////////////////////////////////
        '// Accept a metafile value, add it to the index,
        '// write its filename to the filearea, and return
        '// its Ident value.
        '///////////////////////////////////////////////////
        Public Function Add(ByRef metafile As META_FILE_STRUCT) As Long
            'Populate the correct values
            Header.tagcount += 1
            ReDim Preserve Index(Header.tagcount)
            Index(Header.tagcount).fileData.id = (currentID2 << 16) Or (currentID1)
            currentID1 += 1 : currentID2 += 1

            Index(Header.tagcount).fileData.tagclass = metafile.TagClass
            Index(Header.tagcount).filePath = metafile.Filename
            metafile.ID = Index(Header.tagcount).fileData.id

            Return Index(Header.tagcount).fileData.id
        End Function

        '////////////////////////////////////////////////
        '// Return an item's position in the map index
        '// based on its filename and tag
        '////////////////////////////////////////////////
        Public Function LocateByTagClassAndFilename(ByVal strDependencyTagClass As String, _
                ByVal strDependencyFilename As String) As Long

            For x As Integer = 1 To Header.tagcount
                If Index(x).fileData.tagclass = strDependencyTagClass _
                And Index(x).filePath = strDependencyFilename Then
                    Return Index(x).fileData.id
                End If
            Next
        End Function

    End Class

End Class
