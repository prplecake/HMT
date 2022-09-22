Imports System.Xml
Imports HMTLib

Public Class SCNRHelper
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtFilename As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtNumberOfBytes As System.Windows.Forms.TextBox
    Friend WithEvents txtOffset As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtFilename = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.txtNumberOfBytes = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtOffset = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(376, 48)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "This utility will inject an arbitrarily sized binary chunk of 0x00 bytes into the" & _
        " specified location of a meta file.  The XML file will be updated to reflect off" & _
        "set and translation changes."
        '
        'txtFilename
        '
        Me.txtFilename.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFilename.Location = New System.Drawing.Point(8, 62)
        Me.txtFilename.Name = "txtFilename"
        Me.txtFilename.Size = New System.Drawing.Size(296, 20)
        Me.txtFilename.TabIndex = 1
        Me.txtFilename.Text = ""
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(310, 61)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Browse"
        '
        'txtNumberOfBytes
        '
        Me.txtNumberOfBytes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNumberOfBytes.Location = New System.Drawing.Point(8, 88)
        Me.txtNumberOfBytes.Name = "txtNumberOfBytes"
        Me.txtNumberOfBytes.Size = New System.Drawing.Size(104, 20)
        Me.txtNumberOfBytes.TabIndex = 3
        Me.txtNumberOfBytes.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(112, 91)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(131, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Number of Bytes to Inject"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(248, 88)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(136, 48)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Inject Blank Space"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(112, 112)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(131, 16)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Offset (hex)"
        '
        'txtOffset
        '
        Me.txtOffset.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtOffset.Location = New System.Drawing.Point(8, 112)
        Me.txtOffset.Name = "txtOffset"
        Me.txtOffset.Size = New System.Drawing.Size(104, 20)
        Me.txtOffset.TabIndex = 6
        Me.txtOffset.Text = ""
        '
        'SCNRHelper
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(392, 142)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtOffset)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtNumberOfBytes)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtFilename)
        Me.Controls.Add(Me.Label1)
        Me.Name = "SCNRHelper"
        Me.Text = "Meta Blank Space Injector"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog.AddExtension = True
        OpenFileDialog.DefaultExt = "*.meta"
        OpenFileDialog.Filter = "Halo Meta File (*.meta)|*.meta"
        OpenFileDialog.FileName = ""
        If OpenFileDialog.ShowDialog() = DialogResult.Cancel Then Exit Sub

        Dim strFilename As String = OpenFileDialog.FileName
        Dim strXMLFilename As String = Mid(strFilename, 1, strFilename.Length - 5) & ".xml"

        If Not System.IO.File.Exists(strXMLFilename) Then
            MsgBox("XML structure file not found!")
            Exit Sub
        End If

        txtFilename.Text = strFilename

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim strFilename As String = txtFilename.Text
        If strFilename.Length < 1 Then
            MsgBox("You must select a file!")
            Exit Sub
        End If
        Dim strXMLFilename As String = Mid(strFilename, 1, strFilename.Length - 5) & ".xml"
        If Not System.IO.File.Exists(strFilename) Then
            MsgBox("Metadata file not found!")
            Exit Sub
        End If
        If Not System.IO.File.Exists(strXMLFilename) Then
            MsgBox("XML structure file not found!")
            Exit Sub
        End If
        If Val(txtNumberOfBytes.Text & "") < 1 Then
            MsgBox("You must specify the number of bytes to inject!")
            Exit Sub
        End If
        If Val("&H" & txtOffset.Text) < 0 Then
            MsgBox("You must specify and offset to inject at!")
            Exit Sub
        End If

        'Open the .meta file, and inject the correct number of bytes
        Dim intOffset As Integer = "&H" & txtOffset.Text
        Dim intNumberOfBytes As Integer = Val(txtNumberOfBytes.Text)
        Dim br As System.IO.BinaryReader
        Dim bw As System.IO.BinaryWriter
        Dim fs As System.IO.FileStream
        Dim bChunk() As Byte

        fs = New System.IO.FileStream(strFilename, IO.FileMode.Open)

        br = New System.IO.BinaryReader(fs)
        bw = New System.IO.BinaryWriter(fs)

        ReDim bChunk((br.BaseStream.Length - intOffset) - 1)
        br.BaseStream.Seek(intOffset, IO.SeekOrigin.Begin)
        bChunk = br.ReadBytes(bChunk.Length)
        br.BaseStream.Seek(intOffset, IO.SeekOrigin.Begin)
        For x As Integer = 1 To intNumberOfBytes
            bw.Write(Convert.ToByte(0))
        Next
        bw.Write(bChunk)
        bw.Close()
        br.Close()

        'Now we need to update the XML file
        Dim xmlD As New XmlDocument
        Dim xmlN As XmlNode
        Dim intLocation, intTranslation As Integer

        xmlD.Load(strXMLFilename)
        For Each xmlN In xmlD.SelectNodes("/Results/*")
            Select Case xmlN.Name
                Case "Reflexive"
                    intLocation = strHexToDec(xmlN.ChildNodes(0).InnerText)
                    intTranslation = strHexToDec(xmlN.ChildNodes(1).InnerText)
                    If intLocation >= Val("&H" & txtOffset.Text) Then intLocation += intNumberOfBytes
                    If intTranslation >= Val("&H" & txtOffset.Text) Then intTranslation += intNumberOfBytes
                    xmlN.ChildNodes(0).InnerText = "0x" & Hex(intLocation)
                    xmlN.ChildNodes(1).InnerText = "0x" & Hex(intTranslation)
                Case "Dependency", "LoneID"
                    intLocation = strHexToDec(xmlN.ChildNodes(0).InnerText)
                    If intLocation >= Val("&H" & txtOffset.Text) Then intLocation += intNumberOfBytes
                    xmlN.ChildNodes(0).InnerText = "0x" & Hex(intLocation)
            End Select
        Next
        xmlD.Save(strXMLFilename)
        xmlD = Nothing
        MsgBox("Done!")
    End Sub
End Class
