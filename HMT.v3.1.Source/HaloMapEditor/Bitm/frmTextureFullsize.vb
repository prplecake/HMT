'//////////////////////////////////////////////
'// This form will show the full sized texture
'// from the thumbnail
'//////////////////////////////////////////////
Public Class frmTextureFullsize
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
    Friend WithEvents pb As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pb = New System.Windows.Forms.PictureBox
        Me.SuspendLayout()
        '
        'pb
        '
        Me.pb.BackColor = System.Drawing.Color.Black
        Me.pb.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pb.Location = New System.Drawing.Point(0, 0)
        Me.pb.Name = "pb"
        Me.pb.Size = New System.Drawing.Size(512, 512)
        Me.pb.TabIndex = 0
        Me.pb.TabStop = False
        '
        'frmTextureFullsize
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(512, 512)
        Me.Controls.Add(Me.pb)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmTextureFullsize"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Full Size Texture"
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Event Handlers"

    Private Sub pb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pb.Click
        Me.Hide()
    End Sub

    Private Sub frmTextureFullsize_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Hide()
    End Sub

#End Region

End Class
