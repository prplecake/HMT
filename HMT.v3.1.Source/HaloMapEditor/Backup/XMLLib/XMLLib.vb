'////////////////////////////////////////////////////
'// Halo Map Tools
'// XMLLib: XML Plugin Based Meta Editing System
'// Written by: tjc2k4
'////////////////////////////////////////////////////
Imports HMTLib
Imports System.xml
Imports System.IO

Public Class XMLGUI
    Inherits System.Windows.Forms.UserControl
    Dim plugin As XMLMain
    Dim userDidIt As Boolean

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef map As HaloMap.Map, ByVal tag As String, ByVal index As Integer)
        MyBase.New()
        map = map
        Dim b As New BitMask(9)
        b.BitOn(0)
        InitializeComponent()
        plugin = New XMLMain(map, tag)
        If Not plugin.TagHandled() Then Exit Sub
        plugin.ReadTag(map.indexItem(index).magic_metadata_offset)
        CreateGUI()
    End Sub
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub
    Private components As System.ComponentModel.IContainer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        '
        'XMLGUI
        '
        Me.BackColor = System.Drawing.Color.White
        Me.Name = "XMLGUI"
        Me.Size = New System.Drawing.Size(320, 440)

    End Sub

#End Region

    Public Sub Unload()

    End Sub

    Public Function TagHandled() As Boolean
        Return plugin.TagHandled()
    End Function

    Public Sub CreateGUI()
        Dim i, j, k As Integer
        Dim x, y, x2, y2, x3, y3, tabIndex As Integer
        x = 2
        y = 0
        tabIndex = 0
        For i = 0 To plugin.struct.Length - 1
            Dim structGroupBox As New Windows.Forms.GroupBox
            structGroupBox.Text = plugin.struct(i).name
            structGroupBox.Name = plugin.struct(i).name & "GroupBox"
            If plugin.struct(i).name <> "Main" Then
                Dim structComboBox As New Windows.Forms.ComboBox
                Dim structLabel As New Windows.Forms.Label
                structLabel.Text = plugin.struct(i).name
                structLabel.Location = New System.Drawing.Point(2, y)
                structLabel.Width = plugin.struct(i).name.Length * 7
                structComboBox.Name = plugin.struct(i).name & "ComboBox"
                structComboBox.Location = New System.Drawing.Point(structLabel.Width + 10, y)
                AddHandler structComboBox.SelectedIndexChanged, AddressOf somethingChanged
                AddHandler structComboBox.Click, AddressOf SetUserDidIt
                Me.Controls.Add(structLabel)
                Me.Controls.Add(structComboBox)
                For j = 0 To plugin.struct(0).values.Length - 1
                    If plugin.struct(i).name = plugin.struct(0).values(j).name And plugin.struct(0).values(j).type = "reflexive" Then
                        For k = 0 To plugin.struct(0).values(j).count - 1
                            structComboBox.Items.Add(k)
                        Next
                    End If
                Next
                y = y + 25
            End If
            structGroupBox.Location = New System.Drawing.Point(2, y)
            x2 = 10
            y2 = 15
            For j = 0 To plugin.struct(i).values.Length - 1
                Select Case plugin.struct(i).values(j).type
                    Case "float", "string32", "integer", "short"
                        Dim txtBox As New System.Windows.Forms.TextBox
                        Dim tagLabel As New System.Windows.Forms.Label
                        txtBox.Name = plugin.struct(i).values(j).name & "TextBox"
                        txtBox.Tag = plugin.struct(i).values(j).offset
                        If Not plugin.struct(i).values(j).data Is Nothing Then txtBox.Text = plugin.struct(i).values(j).data
                        tagLabel.Name = plugin.struct(i).values(j).name & "Label"
                        tagLabel.Text = plugin.struct(i).values(j).name
                        tagLabel.TabIndex = tabIndex
                        txtBox.TabIndex = tabIndex + 1
                        tagLabel.AutoSize = True
                        tagLabel.Location = New System.Drawing.Point(x2, y2)
                        x2 = Me.Width / 2
                        txtBox.Size = New System.Drawing.Size(75, 20)
                        txtBox.Location = New System.Drawing.Point(x2, y2)
                        AddHandler txtBox.TextChanged, AddressOf somethingChanged
                        structGroupBox.Controls.Add(tagLabel)
                        structGroupBox.Controls.Add(txtBox)
                        txtBox.BringToFront()
                        y2 = y2 + 20
                        x2 = 10
                        tabIndex = tabIndex + 2
                    Case "id32", "id16"
                        Dim tagPanel As New System.Windows.Forms.GroupBox
                        tagPanel.Tag = plugin.struct(i).values(j).offset
                        tagPanel.Name = plugin.struct(i).values(j).name.Replace(" ", "") & "Group"
                        tagPanel.Text = plugin.struct(i).values(j).name
                        tagPanel.Width = Me.Width - 20
                        tagPanel.Height = 20 + (plugin.struct(i).values(j).offset_options.Length \ 2 + _
                                           plugin.struct(i).values(j).offset_options.Length Mod 2) * 16
                        tagPanel.Location = New System.Drawing.Point(x2, y2)
                        tagPanel.TabIndex = tabIndex
                        tabIndex = tabIndex + 1
                        x3 = 10
                        y3 = 15
                        For k = 0 To plugin.struct(i).values(j).offset_options.Length - 1
                            Dim radioButton As New System.Windows.Forms.RadioButton
                            radioButton.Name = plugin.struct(i).values(j).offset_options(k).op_name.Replace(" ", "") & "RadioButton"
                            radioButton.Text = plugin.struct(i).values(j).offset_options(k).op_name
                            radioButton.Tag = plugin.struct(i).values(j).offset_options(k).op_value
                            If plugin.struct(i).values(j).offset_options(k).op_value = plugin.struct(i).values(j).data Then
                                radioButton.Checked = True
                            Else
                                radioButton.Checked = False
                            End If
                            radioButton.Location = New System.Drawing.Point(x3, y3)
                            radioButton.Size = New System.Drawing.Size(radioButton.Name.Length * 5 + 15, 15)
                            radioButton.TabIndex = tabIndex + k
                            AddHandler radioButton.CheckedChanged, AddressOf somethingChanged
                            tagPanel.Controls.Add(radioButton)
                            radioButton.BringToFront()
                            If k Mod 2 = 1 Then
                                x3 = 10
                                y3 = y3 + 15
                            Else
                                x3 = Me.Width / 2
                            End If
                        Next
                        structGroupBox.Controls.Add(tagPanel)
                        y2 = y2 + tagPanel.Height
                        x2 = 10
                        tabIndex = tabIndex + plugin.struct(i).values(j).offset_options.Length
                    Case "bitmask32"
                        Dim tagPanel2 As New System.Windows.Forms.GroupBox
                        tagPanel2.Tag = plugin.struct(i).values(j).offset
                        tagPanel2.Name = plugin.struct(i).values(j).name.Replace(" ", "") & "Group"
                        tagPanel2.Text = plugin.struct(i).values(j).name
                        tagPanel2.Width = Me.Width - 20
                        tagPanel2.Height = 20 + (plugin.struct(i).values(j).offset_options.Length \ 2 + _
                                           plugin.struct(i).values(j).offset_options.Length Mod 2) * 16
                        tagPanel2.TabIndex = tabIndex
                        tagPanel2.Location = New System.Drawing.Point(x2, y2)
                        x3 = 10
                        y3 = 15
                        tabIndex = tabIndex + 1
                        For k = 0 To plugin.struct(i).values(j).offset_options.Length - 1
                            Dim checkBox As New System.Windows.Forms.CheckBox
                            checkBox.Name = plugin.struct(i).values(j).offset_options(k).op_name.Replace(" ", "") & "CheckBox"
                            checkBox.Text = plugin.struct(i).values(j).offset_options(k).op_name
                            checkBox.Tag = plugin.struct(i).values(j).offset_options(k).op_value
                            checkBox.Location = New System.Drawing.Point(x3, y3)
                            checkBox.Size = New System.Drawing.Size(checkBox.Name.Length * 5 + 15, 15)
                            checkBox.TabIndex = tabIndex + k
                            If Not plugin.struct(i).values(j).data Is Nothing Then
                                If plugin.struct(i).values(j).data.BitOn(plugin.struct(i).values(j).offset_options(k).op_value) Then
                                    checkBox.Checked = True
                                End If
                            End If
                            AddHandler checkBox.CheckedChanged, AddressOf somethingChanged
                            tagPanel2.Controls.Add(checkBox)
                            checkBox.BringToFront()
                            If k Mod 2 = 1 Then
                                x3 = 10
                                y3 = y3 + 15
                            Else
                                x3 = Me.Width / 2
                            End If
                        Next
                        structGroupBox.Controls.Add(tagPanel2)
                        y2 = y2 + tagPanel2.Height
                        tabIndex = tabIndex + plugin.struct(i).values(j).offset_options.Length
                        x2 = 10
                End Select
            Next
            structGroupBox.Size = New System.Drawing.Size(Me.Width - 4, y2 + 5)
            y = y + structGroupBox.Height
            Me.Controls.Add(structGroupBox)
        Next
    End Sub
    Private Sub SetUserDidIt(ByVal sender As Object, ByVal e As EventArgs)
        userDidIt = True
    End Sub
    Private Sub somethingChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim i, j, k As Integer
        Try
            If TypeOf (sender) Is ComboBox Then
                If userDidIt Then
                    Dim combo As ComboBox
                    Dim comboName As String
                    Dim index As Integer
                    combo = sender
                    comboName = combo.Name
                    index = combo.SelectedIndex
                    plugin.ReadReflexive(Mid(combo.Name, 1, combo.Name.IndexOf("ComboBox")), index)
                    Me.Controls.Clear()
                    CreateGUI()
                    For i = 0 To Me.Controls.Count - 1
                        If Me.Controls.Item(i).Name = comboName Then
                            combo = Me.Controls.Item(i)
                            userDidIt = False
                            combo.SelectedIndex = index
                            Exit Sub
                        End If
                    Next
                End If
            ElseIf TypeOf (sender) Is TextBox Then
                Dim txt As TextBox
                txt = sender
                plugin.SetValue(Mid(txt.Name, 1, txt.Name.IndexOf("TextBox")), txt.Text)
                If txt.Parent.Name = "MainGroupBox" Then
                    plugin.WriteTag()
                Else
                    Dim c As ComboBox
                    For i = 0 To Me.Controls.Count - 1
                        If Mid(txt.Parent.Name, 1, txt.Parent.Name.IndexOf("GroupBox")) & "ComboBox" = Me.Controls.Item(i).Name Then
                            c = Me.Controls.Item(i)
                            plugin.WriteReflexive(Mid(txt.Parent.Name, 1, txt.Parent.Name.IndexOf("GroupBox")), c.SelectedIndex)
                            Exit Sub
                        End If
                    Next
                End If
            ElseIf TypeOf (sender) Is CheckBox Then
                Dim check As CheckBox
                check = sender
                check.Checked = plugin.SetOption(Mid(check.Parent.Name, 1, check.Parent.Name.IndexOf("Group")), check.Tag)
                If check.Parent.Parent.Name = "MainGroupBox" Then
                    plugin.WriteTag()
                Else
                    Dim c As ComboBox
                    For i = 0 To Me.Controls.Count - 1
                        If Mid(check.Parent.Parent.Name, 1, check.Parent.Parent.Name.IndexOf("GroupBox")) & "ComboBox" = Me.Controls.Item(i).Name Then
                            c = Me.Controls.Item(i)
                            plugin.WriteReflexive(Mid(check.Parent.Parent.Name, 1, check.Parent.Parent.Name.IndexOf("GroupBox")), c.SelectedIndex)
                            Exit Sub
                        End If
                    Next
                End If
            ElseIf TypeOf (sender) Is RadioButton Then
                Dim radio As RadioButton
                radio = sender
                If radio.Checked Then
                    radio.Checked = plugin.SetValue(Mid(radio.Parent.Name, 1, radio.Parent.Name.IndexOf("Group")), radio.Tag)
                End If
                If radio.Parent.Parent.Name = "MainGroupBox" Then
                    plugin.WriteTag()
                Else
                    Dim c As ComboBox
                    For i = 0 To Me.Controls.Count - 1
                        If Mid(radio.Parent.Parent.Name, 1, radio.Parent.Parent.Name.IndexOf("GroupBox")) & "ComboBox" = Me.Controls.Item(i).Name Then
                            c = Me.Controls.Item(i)
                            plugin.WriteReflexive(Mid(radio.Parent.Parent.Name, 1, radio.Parent.Parent.Name.IndexOf("GroupBox")), c.SelectedIndex)
                            Exit Sub
                        End If
                    Next
                End If
            Else
                Exit Sub
            End If
        Catch
            MessageBox.Show("Error while trying to update and write structure.", "HMT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

End Class

Public Class XMLMain
    Dim handled As Boolean
    Structure TAG_STRUCT
        Dim name As String
        Dim size As Integer
        Dim values() As VALUE_STRUCT
    End Structure
    Structure VALUE_STRUCT
        Dim name As String
        Dim offset As Integer
        Dim type As String
        Dim data As Object
        Dim count As Integer
        Dim offset_options() As offset_options
    End Structure
    Structure OFFSET_OPTIONS
        Dim op_value As Integer
        Dim op_name As String
    End Structure
    Public struct() As TAG_STRUCT
    Dim map As HaloMap.Map
    Dim tag As String
    Dim offset As Integer
    Public Sub New(ByRef map As HaloMap.Map, ByVal tag As String)
        Me.map = map
        Me.tag = tag
        LoadPlugin(tag)
    End Sub
    Public Sub WriteTag()
        Dim bw As New BinaryWriter(map.fs)
        Dim i, j, k As Integer
        For i = 0 To struct(0).values.Length - 1 ' read the main struct
            bw.BaseStream.Seek(offset + struct(0).values(i).offset, SeekOrigin.Begin)
            Select Case struct(0).values(i).type
                Case "float"
                    If TypeOf (struct(0).values(i).data) Is String Then If struct(0).values(i).data = "" Then struct(0).values(i).data = 0
                    Dim s As Single
                    s = struct(0).values(i).data
                    bw.Write(s)
                Case "string32"
                    Dim s As String
                    s = struct(0).values(i).data
                    bw.Write(s)
                Case "id32", "integer"
                    If TypeOf (struct(0).values(i).data) Is String Then If struct(0).values(i).data = "" Then struct(0).values(i).data = 0
                    Dim s As Integer
                    s = struct(0).values(i).data
                    bw.Write(s)
                Case "short", "id16"
                    If TypeOf (struct(0).values(i).data) Is String Then If struct(0).values(i).data = "" Then struct(0).values(i).data = 0
                    Dim s As Short
                    s = struct(0).values(i).data
                    bw.Write(s)
                Case "reflexive"
                    If TypeOf (struct(0).values(i).data) Is String Then If struct(0).values(i).data = "" Then struct(0).values(i).data = 0
                    Dim s, o As Integer
                    s = struct(0).values(i).count
                    o = struct(0).values(i).data
                    bw.Write(s)
                    bw.Write(o)
                Case "bitmask32"
                    Dim s As Integer
                    s = struct(0).values(i).data.ToInt()
                    bw.Write(s)
            End Select
        Next
    End Sub
    Public Sub WriteReflexive(ByVal type As String, ByVal num As Integer)
        'first, find offset for struct
        Dim i, j, k As Integer
        Dim bw As New BinaryWriter(map.fs)
        Dim refOffset, structNum
        For i = 0 To struct(0).values.Length - 1
            If struct(0).values(i).name = type Then
                refOffset = struct(0).values(i).data
            End If
        Next
        For i = 0 To struct.Length - 1
            If struct(i).name = type Then
                structNum = i
            End If
        Next
        refOffset = Unsigned(refOffset) - map.intMagic + (num * struct(structNum).size)
        For j = 0 To struct(structNum).values.Length - 1
            bw.BaseStream.Seek(refOffset + struct(structNum).values(j).offset, SeekOrigin.Begin)
            Select Case struct(structNum).values(j).type
                Case "float"
                    Dim s As Single
                    s = struct(structNum).values(j).data
                    bw.Write(s)
                Case "string32"
                    Dim b As Byte = 0
                    'struct(structNum).values(j).data = ""
                    For k = 0 To 31
                        If k >= struct(structNum).values(j).data.Length OrElse struct(structNum).values(j).data.Chars(k) Is Nothing Then
                            bw.Write(b)
                        Else
                            bw.Write(struct(structNum).values(j).data.Chars(k))
                        End If
                    Next
                Case "id32", "integer"
                    Dim s As Integer
                    s = struct(structNum).values(j).data
                    bw.Write(s)
                Case "short", "id16"
                    Dim s As Short
                    s = struct(structNum).values(j).data
                    bw.Write(s)
                Case "bitmask32"
                    Dim s As Integer
                    s = struct(structNum).values(j).data.ToInt()
                    bw.Write(s)
            End Select
        Next
    End Sub
    Public Sub ReadTag(ByVal offset As Integer)
        Dim br As New BinaryReader(map.fs)
        Me.offset = offset
        Dim i, j, k As Integer
        For i = 0 To struct(0).values.Length - 1 ' read the main struct
            br.BaseStream.Seek(offset + struct(0).values(i).offset, SeekOrigin.Begin)
            Select Case struct(0).values(i).type
                Case "float"
                    struct(0).values(i).data = br.ReadSingle()
                Case "string32"
                    struct(0).values(i).data = br.ReadString()
                    struct(0).values(i).data.PadRight(32)
                Case "id32"
                    struct(0).values(i).data = br.ReadInt32()
                Case "integer"
                    struct(0).values(i).data = br.ReadInt32()
                Case "short", "id16"
                    struct(0).values(i).data = br.ReadInt16()
                Case "reflexive"
                    struct(0).values(i).count = br.ReadInt32()
                    struct(0).values(i).data = br.ReadInt32()
                Case "bitmask32"
                    struct(0).values(i).data = New BitMask(br.ReadInt32())
            End Select
        Next
    End Sub
    Public Sub ReadReflexive(ByVal type As String, ByVal num As Integer)
        'first, find offset for struct
        Dim i, j, k As Integer
        Dim br As New BinaryReader(map.fs)
        Dim refOffset, structNum
        For i = 0 To struct(0).values.Length - 1
            If struct(0).values(i).name = type Then
                refOffset = struct(0).values(i).data
            End If
        Next
        For i = 0 To struct.Length - 1
            If struct(i).name = type Then
                structNum = i
            End If
        Next
        refOffset = Unsigned(refOffset) - map.intMagic + (num * struct(structNum).size)
        For j = 0 To struct(structNum).values.Length - 1
            br.BaseStream.Seek(refOffset + struct(structNum).values(j).offset, SeekOrigin.Begin)
            Select Case struct(structNum).values(j).type
                Case "float"
                    struct(structNum).values(j).data = br.ReadSingle()
                Case "string32"
                    struct(structNum).values(j).data = ""
                    k = 0
                    While k < 31 And Not br.PeekChar() = (10) And Not br.PeekChar = (13)
                        struct(structNum).values(j).data = struct(structNum).values(j).data & Chr(br.ReadByte)
                        k += 1
                    End While
                Case "id32", "integer"
                    struct(structNum).values(j).data = br.ReadInt32()
                Case "short", "id16"
                    struct(structNum).values(j).data = br.ReadInt16()
                Case "bitmask32"
                    struct(structNum).values(j).data = New BitMask(br.ReadInt32())
            End Select
        Next
    End Sub

    Public Sub LoadPlugin(ByVal tag As String)
        Dim xmlD As New XmlDocument
        Dim xmlNL As XmlNodeList
        Dim i, j, k, l, m As Integer
        If System.IO.File.Exists(Application.StartupPath & "\plugins\" & tag & ".xml") Then
            xmlD.Load(Application.StartupPath & "\plugins\" & tag & ".xml")
        Else
            handled = False
            Return
        End If
        handled = True

        Dim tmpNL As XmlNodeList = xmlD.GetElementsByTagName("struct")
        ReDim struct(tmpNL.Count - 1)

        For i = 0 To tmpNL.Count - 1 ' for each structure
            Dim xmlStructNL(tmpNL.Count) As XmlNodeList
            xmlStructNL(i) = tmpNL.ItemOf(i).ChildNodes
            Me.struct(i).name = xmlStructNL(i).ItemOf(0).InnerText
            Me.struct(i).size = Val(xmlStructNL(i).ItemOf(1).InnerText)
            ReDim Me.struct(i).values(xmlStructNL(i).Count - 3)

            For j = 0 To xmlStructNL(i).Count - 3 ' for each 'value' node
                Dim xmlValueNL As XmlNodeList
                xmlValueNL = xmlStructNL(i).ItemOf(j + 2).ChildNodes
                Me.struct(i).values(j).type = xmlValueNL.Item(0).InnerText
                Me.struct(i).values(j).offset = Val("&H" & Mid(xmlValueNL.ItemOf(1).InnerText, 3))
                Me.struct(i).values(j).name = xmlValueNL.ItemOf(2).InnerText
                If xmlValueNL.Count - 3 > 0 Then
                    ReDim Me.struct(i).values(j).offset_options(xmlValueNL.Count - 4)
                    For k = 3 To xmlValueNL.Count - 1
                        Select Case xmlValueNL.ItemOf(k).Name
                            Case "option"
                                Me.struct(i).values(j).offset_options(k - 3).op_value = Val(xmlValueNL.ItemOf(k).ChildNodes.ItemOf(0).InnerText)
                                Me.struct(i).values(j).offset_options(k - 3).op_name = xmlValueNL.ItemOf(k).ChildNodes.ItemOf(1).InnerText
                            Case "bitmask"
                                Me.struct(i).values(j).offset_options(k - 3).op_value = Val(xmlValueNL.ItemOf(k).ChildNodes.ItemOf(0).InnerText)
                                Me.struct(i).values(j).offset_options(k - 3).op_name = xmlValueNL.ItemOf(k).ChildNodes.ItemOf(1).InnerText
                        End Select
                    Next
                End If
            Next
        Next
    End Sub
    Public Function TagHandled() As Boolean
        Return handled
    End Function
    Public Function SetValue(ByVal valName As String, ByVal val As Object) As Boolean
        Dim i, j, k As Integer
        For i = 0 To struct.Length - 1
            For j = 0 To struct(i).values.Length - 1
                If struct(i).values(j).name.Replace(" ", "") = valName.Replace(" ", "") Then
                    struct(i).values(j).data = val
                    Return True
                End If
            Next
        Next
        Return False
    End Function
    Public Function SetOption(ByVal valName As String, ByVal val As Object) As Boolean
        Dim i, j, k As Integer
        For i = 0 To struct.Length - 1
            For j = 0 To struct(i).values.Length - 1
                If struct(i).values(j).name.Replace(" ", "") = valName.Replace(" ", "") Then
                    If struct(i).values(j).data Is Nothing Then Return False
                    If struct(i).values(j).type = "bitmask32" Then
                        struct(i).values(j).data.ToggleBit(val)
                        Return struct(i).values(j).data.BitOn(val)
                    End If
                End If
            Next
        Next
    End Function

End Class

Public Class BitMask
    Dim mask As Int32
    Public Sub New(ByVal num As Integer)
        mask = num
    End Sub
    Public Function BitOn(ByVal num As Integer) As Boolean
        Return IIf(mask And (1 << (32 - num)), True, False)
    End Function
    Public Sub SetBit(ByVal num As Integer)
        mask = mask Or (1 << (32 - num))
    End Sub
    Public Sub ToggleBit(ByVal num As Integer)
        If mask And (1 << (32 - num)) Then
            mask = mask And Not (1 << (32 - num))
        Else
            mask = mask Or (1 << (32 - num))
        End If
    End Sub

    Public Function ToInt() As Integer
        Return mask
    End Function

End Class
