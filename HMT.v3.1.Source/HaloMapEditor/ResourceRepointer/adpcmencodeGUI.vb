Imports System.io

Public Class adpcmencodeGUI
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
    Friend WithEvents txtFilename As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents btnConvert As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(adpcmencodeGUI))
        Me.XPanderList1 = New XPanderControl.XPanderList
        Me.XPander1 = New XPanderControl.XPander
        Me.btnConvert = New System.Windows.Forms.Button
        Me.btnBrowse = New System.Windows.Forms.Button
        Me.txtFilename = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.XPanderList1.SuspendLayout()
        Me.XPander1.SuspendLayout()
        Me.SuspendLayout()
        '
        'XPanderList1
        '
        Me.XPanderList1.AutoScroll = True
        Me.XPanderList1.BackColor = System.Drawing.Color.White
        Me.XPanderList1.BackColorDark = System.Drawing.Color.Black
        Me.XPanderList1.BackColorLight = System.Drawing.Color.White
        Me.XPanderList1.Controls.Add(Me.XPander1)
        Me.XPanderList1.Location = New System.Drawing.Point(0, 0)
        Me.XPanderList1.Name = "XPanderList1"
        Me.XPanderList1.Size = New System.Drawing.Size(368, 176)
        Me.XPanderList1.TabIndex = 0
        '
        'XPander1
        '
        Me.XPander1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.XPander1.Anchored = False
        Me.XPander1.Animated = False
        Me.XPander1.AnimationTime = 100
        Me.XPander1.BackColor = System.Drawing.Color.White
        Me.XPander1.BorderStyle = System.Windows.Forms.Border3DStyle.Flat
        Me.XPander1.CanToggle = False
        Me.XPander1.CaptionFont = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold)
        Me.XPander1.CaptionLeftColor = System.Drawing.Color.Black
        Me.XPander1.CaptionRightColor = System.Drawing.Color.Red
        Me.XPander1.CaptionStyle = XPanderControl.XPander.CaptionStyleEnum.Normal
        Me.XPander1.CaptionText = "XBADPCMEncode GUI"
        Me.XPander1.CaptionTextColor = System.Drawing.Color.White
        Me.XPander1.CaptionTextHighlightColor = System.Drawing.Color.White
        Me.XPander1.ChevronStyle = XPanderControl.XPander.ChevronStyleEnum.Image
        Me.XPander1.CollapsedHighlightImage = CType(resources.GetObject("XPander1.CollapsedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander1.CollapsedImage = CType(resources.GetObject("XPander1.CollapsedImage"), System.Drawing.Bitmap)
        Me.XPander1.Controls.Add(Me.btnConvert)
        Me.XPander1.Controls.Add(Me.btnBrowse)
        Me.XPander1.Controls.Add(Me.txtFilename)
        Me.XPander1.Controls.Add(Me.Label1)
        Me.XPander1.DockPadding.Top = 25
        Me.XPander1.ExpandedHighlightImage = CType(resources.GetObject("XPander1.ExpandedHighlightImage"), System.Drawing.Bitmap)
        Me.XPander1.ExpandedImage = CType(resources.GetObject("XPander1.ExpandedImage"), System.Drawing.Bitmap)
        Me.XPander1.ForeColor = System.Drawing.Color.White
        Me.XPander1.Location = New System.Drawing.Point(8, 10)
        Me.XPander1.Name = "XPander1"
        Me.XPander1.PaneBottomRightColor = System.Drawing.Color.White
        Me.XPander1.ShowTooltips = True
        Me.XPander1.Size = New System.Drawing.Size(352, 158)
        Me.XPander1.TabIndex = 0
        Me.XPander1.Tag = 0
        Me.XPander1.TooltipText = Nothing
        '
        'btnConvert
        '
        Me.btnConvert.BackColor = System.Drawing.Color.Black
        Me.btnConvert.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnConvert.Location = New System.Drawing.Point(248, 124)
        Me.btnConvert.Name = "btnConvert"
        Me.btnConvert.Size = New System.Drawing.Size(88, 24)
        Me.btnConvert.TabIndex = 14
        Me.btnConvert.Text = "Convert"
        '
        'btnBrowse
        '
        Me.btnBrowse.BackColor = System.Drawing.Color.Black
        Me.btnBrowse.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnBrowse.Location = New System.Drawing.Point(16, 124)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(88, 24)
        Me.btnBrowse.TabIndex = 13
        Me.btnBrowse.Text = "Browse"
        '
        'txtFilename
        '
        Me.txtFilename.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFilename.Location = New System.Drawing.Point(16, 96)
        Me.txtFilename.Name = "txtFilename"
        Me.txtFilename.Size = New System.Drawing.Size(320, 20)
        Me.txtFilename.TabIndex = 1
        Me.txtFilename.Text = ""
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(16, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(320, 48)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Browse for a source .wav file and click convert.  The output file will be saved i" & _
        "n the same folder as the original file as NameOfOriginalFile.adpcm.wav."
        '
        'adpcmencodeGUI
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(368, 174)
        Me.Controls.Add(Me.XPanderList1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "adpcmencodeGUI"
        Me.Text = "Halo Map Tools"
        Me.XPanderList1.ResumeLayout(False)
        Me.XPander1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        OpenFileDialog.DefaultExt = "*.wav"
        OpenFileDialog.Filter = "Wave Files (*.wav)|*.wav"
        OpenFileDialog.ShowDialog()
        txtFilename.Text = OpenFileDialog.FileName
    End Sub

    Private Sub adpcmencodeGUI_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim file As String
        Dim bEncoderFound As Boolean = False

        'Make sure that the encoder program exists.  If not, close the window.
        If Directory.Exists(Application.StartupPath & "\encoder") Then
            For Each file In Directory.GetFiles(Application.StartupPath & "\encoder")
                If file.EndsWith("xbadpcmencode.exe") Then
                    bEncoderFound = True
                End If
            Next
        End If
        If Not bEncoderFound Then
            MsgBox("The encoder was not found in the correct path.")
            Me.Close()
        End If
    End Sub

    Private Sub btnConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConvert.Click
        If txtFilename.Text = "" Then
            MsgBox("You must first choose a file to encode.")
            Exit Sub
        End If
        Try
            Dim strCommandLine As String = Application.StartupPath & "\encoder\xbadpcmencode.exe " & _
              Chr(34) & txtFilename.Text & Chr(34) & " " & _
              Chr(34) & txtFilename.Text & ".adpcm.wav" & Chr(34)
            Shell(strCommandLine)
        Catch
            MsgBox("There was an error executing the commandline.")
        End Try
    End Sub
End Class
