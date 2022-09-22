Imports System.io

Public Class frmGenerateHeader
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents txtFilename As System.Windows.Forms.TextBox
    Friend WithEvents of1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cChannels As System.Windows.Forms.ComboBox
    Friend WithEvents cSampleRate As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmGenerateHeader))
        Me.XPanderList1 = New XPanderControl.XPanderList
        Me.XPander1 = New XPanderControl.XPander
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.cChannels = New System.Windows.Forms.ComboBox
        Me.cSampleRate = New System.Windows.Forms.ComboBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtFilename = New System.Windows.Forms.TextBox
        Me.of1 = New System.Windows.Forms.OpenFileDialog
        Me.XPanderList1.SuspendLayout()
        Me.XPander1.SuspendLayout()
        Me.SuspendLayout()
        '
        'XPanderList1
        '
        Me.XPanderList1.AutoScroll = True
        Me.XPanderList1.BackColor = System.Drawing.Color.Black
        Me.XPanderList1.BackColorDark = System.Drawing.Color.Black
        Me.XPanderList1.BackColorLight = System.Drawing.Color.White
        Me.XPanderList1.Controls.Add(Me.XPander1)
        Me.XPanderList1.Location = New System.Drawing.Point(0, 0)
        Me.XPanderList1.Name = "XPanderList1"
        Me.XPanderList1.Size = New System.Drawing.Size(304, 172)
        Me.XPanderList1.TabIndex = 0
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
        Me.XPander1.CaptionLeftColor = System.Drawing.Color.Black
        Me.XPander1.CaptionRightColor = System.Drawing.Color.Red
        Me.XPander1.CaptionStyle = XPanderControl.XPander.CaptionStyleEnum.Normal
        Me.XPander1.CaptionText = "Generate an ADPCM Header"
        Me.XPander1.CaptionTextColor = System.Drawing.Color.White
        Me.XPander1.CaptionTextHighlightColor = System.Drawing.Color.White
        Me.XPander1.ChevronStyle = XPanderControl.XPander.ChevronStyleEnum.Image
        Me.XPander1.CollapsedHighlightImage = CType(resources.GetObject("XPander1.CollapsedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander1.CollapsedImage = CType(resources.GetObject("XPander1.CollapsedImage"), System.Drawing.Bitmap)
        Me.XPander1.Controls.Add(Me.Button3)
        Me.XPander1.Controls.Add(Me.Button2)
        Me.XPander1.Controls.Add(Me.cChannels)
        Me.XPander1.Controls.Add(Me.cSampleRate)
        Me.XPander1.Controls.Add(Me.Button1)
        Me.XPander1.Controls.Add(Me.Label1)
        Me.XPander1.Controls.Add(Me.txtFilename)
        Me.XPander1.DockPadding.Top = 25
        Me.XPander1.ExpandedHighlightImage = CType(resources.GetObject("XPander1.ExpandedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander1.ExpandedImage = CType(resources.GetObject("XPander1.ExpandedImage"), System.Drawing.Bitmap)
        Me.XPander1.Location = New System.Drawing.Point(8, 8)
        Me.XPander1.Name = "XPander1"
        Me.XPander1.PaneBottomRightColor = System.Drawing.Color.White
        Me.XPander1.ShowTooltips = True
        Me.XPander1.Size = New System.Drawing.Size(288, 152)
        Me.XPander1.TabIndex = 0
        Me.XPander1.Tag = 0
        Me.XPander1.TooltipText = Nothing
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.Black
        Me.Button3.ForeColor = System.Drawing.Color.White
        Me.Button3.Location = New System.Drawing.Point(192, 120)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(80, 24)
        Me.Button3.TabIndex = 6
        Me.Button3.Text = "Close"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Black
        Me.Button2.ForeColor = System.Drawing.Color.White
        Me.Button2.Location = New System.Drawing.Point(16, 120)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(168, 24)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Generate Header"
        '
        'cChannels
        '
        Me.cChannels.Location = New System.Drawing.Point(128, 88)
        Me.cChannels.Name = "cChannels"
        Me.cChannels.Size = New System.Drawing.Size(56, 21)
        Me.cChannels.TabIndex = 4
        '
        'cSampleRate
        '
        Me.cSampleRate.Location = New System.Drawing.Point(16, 88)
        Me.cSampleRate.Name = "cSampleRate"
        Me.cSampleRate.Size = New System.Drawing.Size(104, 21)
        Me.cSampleRate.TabIndex = 3
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Black
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(192, 64)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(80, 22)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Browse"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(232, 32)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Generate a header for raw PCM Data - You probably don't need this ;)"
        '
        'txtFilename
        '
        Me.txtFilename.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFilename.Location = New System.Drawing.Point(16, 64)
        Me.txtFilename.Name = "txtFilename"
        Me.txtFilename.Size = New System.Drawing.Size(168, 20)
        Me.txtFilename.TabIndex = 0
        Me.txtFilename.Text = ""
        '
        'frmGenerateHeader
        '
        Me.AutoScale = False
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(304, 171)
        Me.Controls.Add(Me.XPanderList1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(312, 205)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(312, 205)
        Me.Name = "frmGenerateHeader"
        Me.ShowInTaskbar = False
        Me.Text = "Halo Sound Tools"
        Me.TopMost = True
        Me.XPanderList1.ResumeLayout(False)
        Me.XPander1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Browse for the file.
        of1.ShowDialog()
        txtFilename.Text = of1.FileName
    End Sub

    '----------------------------------------------------------------------------------
    ' Call the injectHeader method of the snd class, passing the user chosen properties
    '----------------------------------------------------------------------------------
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim intChannels As Integer
        Dim intSampleRate As Integer = Val(cSampleRate.SelectedText)
        Dim strFilename As String = txtFilename.Text

        Select Case cChannels.SelectedText
            Case "Mono"
                intChannels = 1
            Case "Setereo"
                intChannels = 2
        End Select
        If strFilename = "" Then
            MsgBox("No file was selected!")
            Exit Sub
        End If

        'fMain.snd.InjectHeader(intChannels, intSampleRate, strFilename, strFilename & ".wav")

    End Sub

    Private Sub frmGenerateHeader_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Center the form
        Me.CenterToParent()
        cSampleRate.Items.Clear()
        cSampleRate.Items.Add("22050")
        cSampleRate.Items.Add("44100")
        cChannels.Items.Clear()
        cChannels.Items.Add("Mono")
        cChannels.Items.Add("Stereo")
    End Sub
End Class
