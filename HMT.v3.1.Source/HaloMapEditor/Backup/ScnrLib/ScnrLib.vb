'///////////////////////////////////////////////////////////
'// Halo Map Tools
'// [scen] Metadata Structure
'// Reverse-Engineered by dmauro
'///////////////////////////////////////////////////////////

Imports System.IO
Imports HaloMap
Imports HMTLib

Public Class ScnrGUI
    Inherits System.Windows.Forms.UserControl

#Region "Structures"

    Public Structure OBJ_EXTENSION_STRUCT
        Public Offset As Integer
        Public ID As Integer
        Public Coord As ScnrMeta.COORD_STRUCT
    End Structure

    Public Structure ObjectSet_struct
        Public RefInfo As ScnrMeta.SCNR_CHUNK
        Public Ref() As ScnrMeta.REF_STRUCT
        Public ObjInfo As ScnrMeta.SCNR_CHUNK
        Public Obj() As OBJ_EXTENSION_STRUCT
    End Structure

#End Region

#Region "Public Variables"

    Public map As HaloMap.Map
    Public intID As Integer
    Public scnr As ScnrLib.ScnrMeta
    Public OL As ObjectSet_struct
    Public UpdateKnob As Boolean = True

#End Region

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef m As HaloMap.Map, ByVal i As Integer)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        intID = i
        map = m
        scnr = New ScnrMeta(map, intID)
        FillInData()
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
    Friend WithEvents tb1 As System.Windows.Forms.TrackBar
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents listGroups As System.Windows.Forms.ListBox
    Friend WithEvents comboType As System.Windows.Forms.ComboBox
    Friend WithEvents txtRotation As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtX As System.Windows.Forms.TextBox
    Friend WithEvents txtY As System.Windows.Forms.TextBox
    Friend WithEvents txtZ As System.Windows.Forms.TextBox
    Friend WithEvents Knob As KnobControl.KnobControl
    Friend WithEvents lblNumber As System.Windows.Forms.Label
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnRevert As System.Windows.Forms.Button
    Friend WithEvents btnSaveOutput As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.listGroups = New System.Windows.Forms.ListBox
        Me.tb1 = New System.Windows.Forms.TrackBar
        Me.lblNumber = New System.Windows.Forms.Label
        Me.comboType = New System.Windows.Forms.ComboBox
        Me.txtX = New System.Windows.Forms.TextBox
        Me.txtY = New System.Windows.Forms.TextBox
        Me.txtZ = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtRotation = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.lblID = New System.Windows.Forms.Label
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnRevert = New System.Windows.Forms.Button
        Me.btnSaveOutput = New System.Windows.Forms.Button
        Me.Knob = New KnobControl.KnobControl
        CType(Me.tb1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'listGroups
        '
        Me.listGroups.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.listGroups.Location = New System.Drawing.Point(8, 8)
        Me.listGroups.Name = "listGroups"
        Me.listGroups.Size = New System.Drawing.Size(112, 106)
        Me.listGroups.TabIndex = 0
        '
        'tb1
        '
        Me.tb1.AutoSize = False
        Me.tb1.Location = New System.Drawing.Point(132, 32)
        Me.tb1.Name = "tb1"
        Me.tb1.Size = New System.Drawing.Size(178, 16)
        Me.tb1.TabIndex = 1
        '
        'lblNumber
        '
        Me.lblNumber.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblNumber.Location = New System.Drawing.Point(132, 56)
        Me.lblNumber.Name = "lblNumber"
        Me.lblNumber.Size = New System.Drawing.Size(178, 16)
        Me.lblNumber.TabIndex = 2
        '
        'comboType
        '
        Me.comboType.Location = New System.Drawing.Point(8, 176)
        Me.comboType.Name = "comboType"
        Me.comboType.Size = New System.Drawing.Size(304, 21)
        Me.comboType.TabIndex = 3
        '
        'txtX
        '
        Me.txtX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtX.Location = New System.Drawing.Point(8, 224)
        Me.txtX.Name = "txtX"
        Me.txtX.Size = New System.Drawing.Size(80, 20)
        Me.txtX.TabIndex = 4
        Me.txtX.Text = ""
        Me.txtX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtY
        '
        Me.txtY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtY.Location = New System.Drawing.Point(116, 224)
        Me.txtY.Name = "txtY"
        Me.txtY.Size = New System.Drawing.Size(80, 20)
        Me.txtY.TabIndex = 5
        Me.txtY.Text = ""
        Me.txtY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtZ
        '
        Me.txtZ.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtZ.Location = New System.Drawing.Point(224, 224)
        Me.txtZ.Name = "txtZ"
        Me.txtZ.Size = New System.Drawing.Size(80, 20)
        Me.txtZ.TabIndex = 6
        Me.txtZ.Text = ""
        Me.txtZ.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(40, 208)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(16, 16)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "X"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(152, 208)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(16, 16)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Y"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(256, 208)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(16, 16)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Z"
        '
        'txtRotation
        '
        Me.txtRotation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRotation.Location = New System.Drawing.Point(88, 272)
        Me.txtRotation.Name = "txtRotation"
        Me.txtRotation.Size = New System.Drawing.Size(144, 20)
        Me.txtRotation.TabIndex = 10
        Me.txtRotation.Text = ""
        Me.txtRotation.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(95, 256)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(128, 16)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Rotation (Radians)"
        '
        'lblID
        '
        Me.lblID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblID.Location = New System.Drawing.Point(8, 152)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(304, 20)
        Me.lblID.TabIndex = 13
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.Color.Black
        Me.btnSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnSave.Location = New System.Drawing.Point(16, 320)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(72, 24)
        Me.btnSave.TabIndex = 33
        Me.btnSave.Text = "Save"
        '
        'btnRevert
        '
        Me.btnRevert.BackColor = System.Drawing.Color.Black
        Me.btnRevert.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRevert.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnRevert.Location = New System.Drawing.Point(232, 320)
        Me.btnRevert.Name = "btnRevert"
        Me.btnRevert.Size = New System.Drawing.Size(72, 24)
        Me.btnRevert.TabIndex = 34
        Me.btnRevert.Text = "Revert"
        '
        'btnSaveOutput
        '
        Me.btnSaveOutput.BackColor = System.Drawing.Color.Black
        Me.btnSaveOutput.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveOutput.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnSaveOutput.Location = New System.Drawing.Point(232, 368)
        Me.btnSaveOutput.Name = "btnSaveOutput"
        Me.btnSaveOutput.Size = New System.Drawing.Size(72, 24)
        Me.btnSaveOutput.TabIndex = 35
        Me.btnSaveOutput.Text = "Text Output"
        Me.btnSaveOutput.Visible = False
        '
        'Knob
        '
        Me.Knob.ImeMode = System.Windows.Forms.ImeMode.On
        Me.Knob.KnobColor = System.Drawing.Color.Black
        Me.Knob.LargeChange = 5
        Me.Knob.Location = New System.Drawing.Point(128, 312)
        Me.Knob.Maximum = 25
        Me.Knob.Minimum = 0
        Me.Knob.Name = "Knob"
        Me.Knob.ShowLargeScale = True
        Me.Knob.ShowSmallScale = False
        Me.Knob.Size = New System.Drawing.Size(72, 64)
        Me.Knob.SmallChange = 1
        Me.Knob.TabIndex = 36
        Me.Knob.Value = 0
        '
        'ScnrGUI
        '
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.btnSaveOutput)
        Me.Controls.Add(Me.btnRevert)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.lblID)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtRotation)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtZ)
        Me.Controls.Add(Me.txtY)
        Me.Controls.Add(Me.txtX)
        Me.Controls.Add(Me.comboType)
        Me.Controls.Add(Me.lblNumber)
        Me.Controls.Add(Me.tb1)
        Me.Controls.Add(Me.listGroups)
        Me.Controls.Add(Me.Knob)
        Me.Name = "ScnrGUI"
        Me.Size = New System.Drawing.Size(320, 408)
        CType(Me.tb1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Event Handlers"

    Private Sub Knob_Change(ByVal ValueChangedEventHandler As System.Object) Handles Knob.ValueChanged
        Dim ratio As Single = Knob.Maximum / 62
        Dim sng As Single
        If Not UpdateKnob Then
            UpdateKnob = True
            Exit Sub
        End If
        sng = ((Knob.Value / ratio) - 31) / 10
        txtRotation.Text = sng
    End Sub

    Private Sub listGroups_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles listGroups.SelectedIndexChanged
        Select Case listGroups.SelectedItem
            Case "Weapons"
                OL.RefInfo = scnr.Header.Weapon_Ref : ReDim OL.Ref(scnr.Header.Weapon_Ref.Count)
                For x As Integer = 1 To scnr.Header.Weapon_Ref.Count : OL.Ref(x) = scnr.Weapon_ref(x) : Next
                OL.ObjInfo = scnr.Header.Weapon : ReDim OL.Obj(scnr.Header.Weapon.Count)
                For x As Integer = 1 To scnr.Header.Weapon.Count
                    OL.Obj(x).coord = scnr.Weapon(x).coord : OL.Obj(x).ID = scnr.Weapon(x).numid
                    OL.Obj(x).Offset = scnr.Weapon(x).mapOffset
                Next

            Case "Vehicles"
                OL.RefInfo = scnr.Header.Vehicle_Ref : ReDim OL.Ref(scnr.Header.Vehicle_Ref.Count)
                For x As Integer = 1 To scnr.Header.Vehicle_Ref.Count : OL.Ref(x) = scnr.Vehicle_ref(x) : Next
                OL.ObjInfo = scnr.Header.Vehicle : ReDim OL.Obj(scnr.Header.Vehicle.Count)
                For x As Integer = 1 To scnr.Header.Vehicle.Count
                    OL.Obj(x).coord = scnr.Vehicle(x).coord : OL.Obj(x).ID = scnr.Vehicle(x).numid
                    OL.Obj(x).Offset = scnr.Vehicle(x).mapOffset
                Next

            Case "Scenery"
                OL.RefInfo = scnr.Header.Scenery_Ref : ReDim OL.Ref(scnr.Header.Scenery_Ref.Count)
                For x As Integer = 1 To scnr.Header.Scenery_Ref.Count : OL.Ref(x) = scnr.Scenery_ref(x) : Next
                OL.ObjInfo = scnr.Header.Scenery : ReDim OL.Obj(scnr.Header.Scenery.Count)
                For x As Integer = 1 To scnr.Header.Scenery.Count
                    OL.Obj(x).coord = scnr.Scenery(x).coord : OL.Obj(x).ID = scnr.Scenery(x).numid
                    OL.Obj(x).Offset = scnr.Scenery(x).mapOffset
                Next

            Case "Equipment"
                OL.RefInfo = scnr.Header.equipment_Ref : ReDim OL.Ref(scnr.Header.equipment_Ref.Count)
                For x As Integer = 1 To scnr.Header.equipment_Ref.Count : OL.Ref(x) = scnr.Equipment_ref(x) : Next
                OL.ObjInfo = scnr.Header.equipment : ReDim OL.Obj(scnr.Header.equipment.Count)
                For x As Integer = 1 To scnr.Header.equipment.Count
                    OL.Obj(x).coord = scnr.Equipment(x).coord : OL.Obj(x).ID = scnr.Equipment(x).numid
                    OL.Obj(x).Offset = scnr.Equipment(x).mapOffset
                Next

            Case "Player Spawns"
                OL.RefInfo.Count = 0
                OL.ObjInfo = scnr.Header.PlayerSpawn : ReDim OL.Obj(scnr.Header.PlayerSpawn.Count)
                For x As Integer = 1 To scnr.Header.PlayerSpawn.Count
                    OL.Obj(x).coord = scnr.PlayerSpawn(x).coord
                    OL.Obj(x).ID = 99 : OL.Obj(x).Offset = scnr.PlayerSpawn(x).mapOffset
                Next

                'Case "Item Collections"
                '    OL.RefInfo = scnr.Header.equipment_Ref : ReDim OL.Ref(scnr.Header.equipment_Ref.Count)
                '    For x As Integer = 1 To scnr.Header.equipment.Count : OL.Ref(x) = scnr.equipment_ref(x) : Next
                '    OL.ObjInfo = scnr.Header.equipment : ReDim OL.Obj(scnr.Header.equipment.Count)
                '    For x As Integer = 1 To scnr.Header.equipment.Count : OL.Obj(x) = scnr.equipment(x).coord : Next

        End Select
        tb1.Minimum = 1 : tb1.Maximum = OL.ObjInfo.Count
        tb1.Value = 1 : lblNumber.Text = "Object Number 1 of " & OL.ObjInfo.Count
        comboType.Enabled = True
        comboType.Items.Clear() : comboType.Update()
        If OL.RefInfo.Count = 0 Then
            lblID.Text = "N/A"
            comboType.Enabled = False
        Else
            For x As Integer = 1 To OL.RefInfo.Count
                comboType.Items.Add(x - 1 & ": " & OL.Ref(x).strName)
            Next
        End If
        SelectItem(1)
    End Sub

    Private Sub tb1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb1.Scroll
        lblNumber.Text = "Object Number " & tb1.Value & " of " & OL.ObjInfo.Count
        SelectItem(tb1.Value)
    End Sub
    Public Sub SelectItem(ByVal i As Integer)
        If comboType.Enabled Then comboType.SelectedIndex = OL.Obj(i).ID
        txtX.Text = OL.Obj(i).Coord.x : txtY.Text = OL.Obj(i).Coord.y
        txtZ.Text = OL.Obj(i).Coord.z : txtRotation.Text = OL.Obj(i).Coord.rot
        UpdateKnob = False
        Dim ratio As Single = Knob.Maximum / 62
        Knob.Value = Int(((OL.Obj(i).Coord.rot * 10) + 31) * ratio)
    End Sub
    Private Sub comboType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles comboType.SelectedIndexChanged
        lblID.Text = "ID: " & Unsigned(OL.Ref(comboType.SelectedIndex + 1).identifier) & _
            " @ " & "0x" & Hex(Unsigned(OL.Ref(comboType.SelectedIndex + 1).mapOffset))
    End Sub

    Private Sub btnRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRevert.Click
        SelectItem(tb1.Value)
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim bw As New BinaryWriter(map.fs)
        Dim x As Integer = tb1.Value
        Try
            OL.Obj(x).Coord.x = txtX.Text : OL.Obj(x).Coord.y = txtY.Text
            OL.Obj(x).Coord.z = txtZ.Text : OL.Obj(x).Coord.rot = txtRotation.Text
            OL.Obj(x).ID = comboType.SelectedIndex
            bw.Seek(OL.Obj(x).Offset, SeekOrigin.Begin)
            If OL.Obj(x).ID <> 99 Then
                Write(OL.Obj(x).ID)
                bw.Seek(6, SeekOrigin.Current)
            End If
            bw.Write(OL.Obj(x).Coord.x)
            bw.Write(OL.Obj(x).Coord.y)
            bw.Write(OL.Obj(x).Coord.z)
            If OL.Obj(x).Coord.rot <> 99 Then bw.Write(OL.Obj(x).Coord.rot)
            bw.Flush()
            MsgBox("Changes were saved.")
        Catch
            MsgBox("There was an error saving the changes.")
        End Try
    End Sub

    Private Sub btnSaveOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveOutput.Click
        Dim w As New StreamWriter(New FileStream("c:\output.txt", FileMode.Create))

        'Write all of the values to the file
        w.WriteLine(listGroups.SelectedItem & "----------------------------------------------------")
        For x As Integer = 1 To tb1.Maximum
            w.WriteLine("Offset: " & Hex(OL.Obj(x).Offset))
            w.WriteLine("ID: " & OL.Obj(x).ID & " - X: " & OL.Obj(x).Coord.x & ", Y: " & OL.Obj(x).Coord.y & _
                ", Z: " & OL.Obj(x).Coord.z & "R: " & OL.Obj(x).Coord.rot & vbCrLf & "------------------------")
        Next
        w.Flush()
        w.Close()
    End Sub

#End Region

#Region "Functions"

    Public Sub Unload()
        'I can add in predestruction code here, if it becomes neccessary
    End Sub

    Public Sub FillInData()
        scnr.DecodeScnrMetadata(map, intID)
        If scnr.Header.Weapon.Count > 0 Then listGroups.Items.Add("Weapons")
        If scnr.Header.Vehicle.Count > 0 Then listGroups.Items.Add("Vehicles")
        If scnr.Header.Scenery.Count > 0 Then listGroups.Items.Add("Scenery")
        If scnr.Header.equipment.Count > 0 Then listGroups.Items.Add("Equipment")
        If scnr.Header.PlayerSpawn.Count > 0 Then listGroups.Items.Add("Player Spawns")
        'If scnr.Header.ItemColl.Count > 0 Then listGroups.Items.Add("Item Collections")
        listGroups.SelectedItem = listGroups.Items.Item(0)
    End Sub

#End Region

End Class

'//////////////////////////////////////
'// [scnr] Scenario Metadata Class
'// Implements all decoding functions
'//////////////////////////////////////
Public Class ScnrMeta

#Region "Structures"

    '///////////////////////////////////
    '// Header structure
    '///////////////////////////////////
    Public Structure SCNR_CHUNK
        Public Count As Integer
        Public Offset As Integer
        Public Unknown As Integer
        Public Function ReadStruct(ByRef fs As FileStream) As SCNR_CHUNK
            Dim br As BinaryReader = New BinaryReader(fs)
            ReadStruct.Count = br.ReadInt32
            ReadStruct.Offset = br.ReadInt32
            ReadStruct.Unknown = br.ReadInt32
        End Function
    End Structure

    Public Structure SCNR_HEADER_STRUCT
        <VBFixedString(272)> Public unneeded1 As String
        Public offsetindex As Integer 'Unsigned
        <VBFixedString(8)> Public unneeded2 As String
        Public offsetendofindex As Integer 'Unsigned
        <VBFixedString(228)> Public zero1 As String

        Public PlayerSpawnsMaybe As SCNR_CHUNK
        Public Scenery As SCNR_CHUNK
        Public Scenery_Ref As SCNR_CHUNK
        Public Biped As SCNR_CHUNK
        Public Biped_Ref As SCNR_CHUNK
        Public Vehicle As SCNR_CHUNK
        Public Vehicle_Ref As SCNR_CHUNK
        Public equipment As SCNR_CHUNK
        Public equipment_Ref As SCNR_CHUNK
        Public Weapon As SCNR_CHUNK
        Public Weapon_Ref As SCNR_CHUNK
        Public LocationNames As SCNR_CHUNK
        Public Machine As SCNR_CHUNK
        Public Machine_Ref As SCNR_CHUNK
        Public Control As SCNR_CHUNK
        Public Control_Ref As SCNR_CHUNK
        Public LightFixture As SCNR_CHUNK
        Public LightFixture_Ref As SCNR_CHUNK
        Public SoundScenery As SCNR_CHUNK
        Public SoundScenery_Ref As SCNR_CHUNK

        <VBFixedArray(7)> Public Unknown1() As SCNR_CHUNK

        Public PlayerStartingProfile As SCNR_CHUNK
        Public PlayerSpawn As SCNR_CHUNK
        Public Labels As SCNR_CHUNK
        Public Animations As SCNR_CHUNK
        Public ItemColl As SCNR_CHUNK
        Public ItemColl_Ref As SCNR_CHUNK

        <VBFixedArray(1)> Public Unknown2() As SCNR_CHUNK

        Public Placement1 As SCNR_CHUNK
        Public Placement2 As SCNR_CHUNK
        Public Decal_Ref As SCNR_CHUNK
        Public DetailObjectCollection As SCNR_CHUNK

        <VBFixedString(84)> Public zero3 As String

        Public Actv As SCNR_CHUNK
        Public Squad As SCNR_CHUNK
        Public MoreSquad As SCNR_CHUNK
        Public Cutscene1 As SCNR_CHUNK
        Public Cutscene2 As SCNR_CHUNK

        <VBFixedString(12)> Public zero4 As String

        Public Cutscene3 As SCNR_CHUNK
        <VBFixedString(336)> Public unneeded3 As String 'stuff i dont wanna mess with
        Public Sub ReadStruct(ByRef fs As FileStream)
            Dim br As New BinaryReader(fs)
            For x As Integer = 1 To 272 : unneeded1 &= br.ReadByte : Next
            offsetindex = br.ReadInt32
            For x As Integer = 1 To 8 : unneeded2 &= br.ReadByte : Next
            offsetendofindex = br.ReadInt32
            For x As Integer = 1 To 228 : zero1 &= br.ReadByte : Next

            PlayerSpawnsMaybe = PlayerSpawnsMaybe.ReadStruct(fs)
            Scenery = Scenery.ReadStruct(fs)
            Scenery_Ref = Scenery_Ref.ReadStruct(fs)
            Biped = Biped.ReadStruct(fs)
            Biped_Ref = Biped_Ref.ReadStruct(fs)
            Vehicle = Vehicle.ReadStruct(fs)
            Vehicle_Ref = Vehicle_Ref.ReadStruct(fs)
            equipment = equipment.ReadStruct(fs)
            equipment_Ref = equipment_Ref.ReadStruct(fs)
            Weapon = Weapon.ReadStruct(fs)
            Weapon_Ref = Weapon_Ref.ReadStruct(fs)
            LocationNames = LocationNames.ReadStruct(fs)
            Machine = Machine.ReadStruct(fs)
            Machine_Ref = Machine_Ref.ReadStruct(fs)
            Control = Control.ReadStruct(fs)
            Control_Ref = Control_Ref.ReadStruct(fs)
            LightFixture = LightFixture.ReadStruct(fs)
            LightFixture_Ref = LightFixture_Ref.ReadStruct(fs)
            SoundScenery = SoundScenery.ReadStruct(fs)
            SoundScenery_Ref = SoundScenery_Ref.ReadStruct(fs)

            ReDim Unknown1(7)
            For x As Integer = 1 To 7 : Unknown1(x) = Unknown1(x).ReadStruct(fs) : Next

            PlayerStartingProfile = PlayerStartingProfile.ReadStruct(fs)
            PlayerSpawn = PlayerSpawn.ReadStruct(fs)
            Labels = Labels.ReadStruct(fs)
            Animations = Animations.ReadStruct(fs)
            ItemColl = ItemColl.ReadStruct(fs)
            ItemColl_Ref = ItemColl_Ref.ReadStruct(fs)


            ReDim Unknown2(1)
            For x As Integer = 1 To 1 : Unknown1(x) = Unknown2(x).ReadStruct(fs) : Next

            Placement1 = Placement1.ReadStruct(fs)
            Placement2 = Placement2.ReadStruct(fs)
            Decal_Ref = Decal_Ref.ReadStruct(fs)
            DetailObjectCollection = DetailObjectCollection.ReadStruct(fs)

            For x As Integer = 1 To 84 : zero3 &= br.ReadByte : Next

            Actv = Actv.ReadStruct(fs)
            Squad = Squad.ReadStruct(fs)
            MoreSquad = MoreSquad.ReadStruct(fs)
            Cutscene1 = Cutscene1.ReadStruct(fs)
            Cutscene2 = Cutscene2.ReadStruct(fs)

            For x As Integer = 1 To 12 : zero4 &= br.ReadByte : Next

            Cutscene3 = Cutscene2.ReadStruct(fs)

            For x As Integer = 1 To 340 : unneeded3 &= br.ReadByte : Next
        End Sub
    End Structure

    Public Structure ITEM_COLL_SPAWN
        Public x As Single
        Public y As Single
        Public z As Single
        <VBFixedArray(2)> Public Unknwon1() As Integer 'Unsigned 
        <VBFixedString(4)> Public Tag() As String
        <VBFixedArray(31)> Public Unknwon2() As Integer 'Unsigned 
    End Structure

    Public Structure COORD_STRUCT
        Public x As Single
        Public y As Single
        Public z As Single
        Public rot As Single
        Public Sub ReadStruct(ByRef fs As FileStream)
            Dim br As New BinaryReader(fs)
            x = br.ReadSingle
            y = br.ReadSingle
            z = br.ReadSingle
            rot = br.ReadSingle
        End Sub
    End Structure

    Public Structure Weapon_STRUCT
        Public mapOffset As Integer
        Public numid As Short
        Public flag As Short
        Public unknown1 As Integer
        Public Coord As COORD_STRUCT
        <VBFixedString(68)> Public unknown2 As String
        Public Sub ReadStruct(ByRef fs As FileStream)
            Dim br As New BinaryReader(fs)
            mapOffset = br.BaseStream.Position
            numid = br.ReadInt16
            flag = br.ReadInt16
            unknown1 = br.ReadInt32
            Coord.ReadStruct(fs)
            For x As Integer = 1 To 68 : unknown2 &= br.ReadByte : Next
        End Sub
    End Structure
    Public Structure Scenery_STRUCT
        Public mapOffset As Integer
        Public numid As Short
        Public flag As Short
        Public unknown1 As Integer
        Public Coord As COORD_STRUCT
        <VBFixedString(48)> Public unknown2 As String
        Public Sub ReadStruct(ByRef fs As FileStream)
            Dim br As New BinaryReader(fs)
            mapOffset = br.BaseStream.Position
            numid = br.ReadInt16
            flag = br.ReadInt16
            unknown1 = br.ReadInt32
            Coord.ReadStruct(fs)
            For x As Integer = 1 To 48 : unknown2 &= br.ReadByte : Next
        End Sub
    End Structure
    Public Structure Vehicle_STRUCT
        Public mapOffset As Integer
        Public numid As Short
        Public flag As Short
        Public unknown1 As Integer
        Public Coord As COORD_STRUCT
        <VBFixedString(96)> Public unknown2 As String
        Public Sub ReadStruct(ByRef fs As FileStream)
            Dim br As New BinaryReader(fs)
            mapOffset = br.BaseStream.Position
            numid = br.ReadInt16
            flag = br.ReadInt16
            unknown1 = br.ReadInt32
            Coord.ReadStruct(fs)
            For x As Integer = 1 To 96 : unknown2 &= br.ReadByte : Next
        End Sub
    End Structure
    Public Structure Equipment_STRUCT
        Public mapOffset As Integer
        Public numid As Short
        Public flag As Short
        Public unknown1 As Integer
        Public Coord As COORD_STRUCT
        <VBFixedString(16)> Public unknown2 As String
        Public Sub ReadStruct(ByRef fs As FileStream)
            Dim br As New BinaryReader(fs)
            mapOffset = br.BaseStream.Position
            numid = br.ReadInt16
            flag = br.ReadInt16
            unknown1 = br.ReadInt32
            Coord.ReadStruct(fs)
            For x As Integer = 1 To 16 : unknown2 &= br.ReadByte : Next
        End Sub
    End Structure
    Public Structure PLAYER_SPAWN_STRUCT
        Public mapOffset As Integer
        Public Coord As COORD_STRUCT
        <VBFixedString(9)> Public unknown2 As String
        Public Sub ReadStruct(ByRef fs As FileStream)
            Dim br As New BinaryReader(fs)
            mapOffset = br.BaseStream.Position
            Coord.ReadStruct(fs)
            For x As Integer = 1 To 9 : unknown2 &= br.ReadInt32 : Next
        End Sub
    End Structure
    Public Structure ITEMCOLL_STRUCT
        Public mapOffset As Integer
        Public Coord As COORD_STRUCT
        Public Unknown1() As Integer
        Public tag As String
        Public Unknown2() As Integer
        Public Sub ReadStruct(ByRef fs As FileStream)
            Dim br As New BinaryReader(fs)
            mapOffset = br.BaseStream.Position
            Coord.ReadStruct(fs)
            Coord.rot = 99 'There is no rotation, so make sure we leave it out
            fs.Seek(mapOffset + 12, SeekOrigin.Begin)
            ReDim Unknown1(2)
            Unknown1(0) = br.ReadInt32 : Unknown1(1) = br.ReadInt32
            For x As Integer = 1 To 4 : tag &= br.ReadChar : Next
            ReDim Unknown2(31)
            For x As Integer = 0 To 30 : Unknown2(x) = br.ReadInt32 : Next
        End Sub
    End Structure

    Public Structure REF_STRUCT
        Public mapOffset As Integer
        <VBFixedString(4)> Public Tagname As String
        Public rawFilenameOffset As Integer 'Unsigned, magic
        Public zero1 As Integer
        Public identifier As Integer
        <VBFixedString(32)> Public zero2 As String
        Public strName As String 'Not in binary structure - generated
        Public Sub readStruct(ByRef fs As FileStream)
            Dim br As New BinaryReader(fs)
            mapOffset = br.BaseStream.Position
            For x As Integer = 1 To 4 : Tagname &= br.ReadByte : Next
            rawFilenameOffset = br.ReadInt32
            zero1 = br.ReadInt32
            identifier = br.ReadInt32
            For x As Integer = 1 To 32 : zero2 &= br.ReadByte : Next
        End Sub
    End Structure
#End Region

#Region "Public Structures"

    Public Header As SCNR_HEADER_STRUCT
    Public Weapon() As Weapon_STRUCT : Public Weapon_ref() As REF_STRUCT
    Public Scenery() As Scenery_STRUCT : Public Scenery_ref() As REF_STRUCT
    Public Vehicle() As Vehicle_STRUCT : Public Vehicle_ref() As REF_STRUCT
    Public Equipment() As Equipment_STRUCT : Public Equipment_ref() As REF_STRUCT
    Public ItemColl() As ITEMCOLL_STRUCT : Public ItemColl_Ref() As REF_STRUCT
    Public PlayerSpawn() As PLAYER_SPAWN_STRUCT
    Public Text As String

#End Region

    '//////////////////////////////////////////////////////
    '// Class Constructor
    '//////////////////////////////////////////////////////
    Public Sub New(ByRef map As Map, ByVal intID As Integer)
        DecodeScnrMetadata(map, intID)
    End Sub

#Region "Functions"

    '//////////////////////////////////////////////////////////////////////
    '// Decode all of the known scenario data
    '//////////////////////////////////////////////////////////////////////
    Public Function DecodeScnrMetadata(ByRef map As Map, ByVal intID As Integer) As Boolean

        Dim strTemp As String
        Dim br1 As New BinaryReader(map.fs)
        Dim s As String

        'Seek to the metadata offset and read the Header.
        br1.BaseStream.Seek(CInt(Unsigned(map.indexItem(intID).magic_metadata_offset)), SeekOrigin.Begin)
        Header.ReadStruct(map.fs)

        'Read known structures
        If Header.Vehicle_Ref.Count > 0 Then
            ReDim Vehicle_ref(Header.Vehicle_Ref.Count)
            br1.BaseStream.Seek(CInt(Unsigned(Header.Vehicle_Ref.Offset) - map.intMagic), SeekOrigin.Begin)
            For x As Integer = 1 To Header.Vehicle_Ref.Count : Vehicle_ref(x).ReadStruct(map.fs) : Next
            For x As Integer = 1 To Header.Vehicle_Ref.Count
                Vehicle_ref(x).strName = map.ReadString(br1, _
                    Unsigned(Vehicle_ref(x).rawFilenameOffset) - map.intMagic)
            Next
            If Header.Vehicle.Count > 0 Then
                ReDim Vehicle(Header.Vehicle.Count)
                br1.BaseStream.Seek(CInt(Unsigned(Header.Vehicle.Offset) - map.intMagic), SeekOrigin.Begin)
                For x As Integer = 1 To Header.Vehicle.Count : Vehicle(x).ReadStruct(map.fs) : Next
            End If
        End If

        If Header.Weapon_Ref.Count > 0 Then
            ReDim Weapon_ref(Header.Weapon_Ref.Count)
            br1.BaseStream.Seek(CInt(Unsigned(Header.Weapon_Ref.Offset) - map.intMagic), SeekOrigin.Begin)
            For x As Integer = 1 To Header.Weapon_Ref.Count : Weapon_ref(x).ReadStruct(map.fs) : Next
            For x As Integer = 1 To Header.Weapon_Ref.Count
                Weapon_ref(x).strName = map.ReadString(br1, _
                    Unsigned(Weapon_ref(x).rawFilenameOffset) - map.intMagic)
            Next
            If Header.Weapon.Count > 0 Then
                ReDim Weapon(Header.Weapon.Count)
                br1.BaseStream.Seek(CInt(Unsigned(Header.Weapon.Offset) - map.intMagic), SeekOrigin.Begin)
                For x As Integer = 1 To Header.Weapon.Count : Weapon(x).ReadStruct(map.fs) : Next
            End If
        End If

        If Header.Scenery_Ref.Count > 0 Then
            ReDim Scenery_ref(Header.Scenery_Ref.Count)
            br1.BaseStream.Seek(CInt(Unsigned(Header.Scenery_Ref.Offset) - map.intMagic), SeekOrigin.Begin)
            For x As Integer = 1 To Header.Scenery_Ref.Count : Scenery_ref(x).ReadStruct(map.fs) : Next
            For x As Integer = 1 To Header.Scenery_Ref.Count
                Scenery_ref(x).strName = map.ReadString(br1, _
                    Unsigned(Scenery_ref(x).rawFilenameOffset) - map.intMagic)
            Next
            If Header.Scenery.Count > 0 Then
                ReDim Scenery(Header.Scenery.Count)
                br1.BaseStream.Seek(CInt(Unsigned(Header.Scenery.Offset) - map.intMagic), SeekOrigin.Begin)
                For x As Integer = 1 To Header.Scenery.Count : Scenery(x).ReadStruct(map.fs) : Next
            End If
        End If

        If Header.equipment_Ref.Count > 0 Then
            ReDim Equipment_ref(Header.equipment_Ref.Count)
            br1.BaseStream.Seek(CInt(Unsigned(Header.equipment_Ref.Offset) - map.intMagic), SeekOrigin.Begin)
            For x As Integer = 1 To Header.equipment_Ref.Count : Equipment_ref(x).ReadStruct(map.fs) : Next
            For x As Integer = 1 To Header.equipment_Ref.Count
                Equipment_ref(x).strName = map.ReadString(br1, _
                    Unsigned(Equipment_ref(x).rawFilenameOffset) - map.intMagic)
            Next
            If Header.equipment.Count > 0 Then
                ReDim Equipment(Header.equipment.Count)
                br1.BaseStream.Seek(CInt(Unsigned(Header.equipment.Offset) - map.intMagic), SeekOrigin.Begin)
                For x As Integer = 1 To Header.equipment.Count : Equipment(x).ReadStruct(map.fs) : Next
            End If
        End If

        'If Header.ItemColl_Ref.Count > 0 Then
        'ReDim ItemColl_Ref(Header.ItemColl_Ref.Count)
        'br1.BaseStream.Seek(CInt(Unsigned(Header.ItemColl_Ref.Offset) - map.intMagic), SeekOrigin.Begin)
        'For x As Integer = 1 To Header.ItemColl_Ref.Count : ItemColl_Ref(x).ReadStruct(map.fs) : Next
        'For x As Integer = 1 To Header.ItemColl_Ref.Count
        'ItemColl_Ref(x).strName = map.ReadString(br1, _
        '    Unsigned(ItemColl_Ref(x).rawFilenameOffset) - map.intMagic)
        'Next
        'If Header.ItemColl.Count > 0 Then
        'ReDim ItemColl(Header.ItemColl.Count)
        'br1.BaseStream.Seek(CInt(Unsigned(Header.ItemColl.Offset) - map.intMagic), SeekOrigin.Begin)
        'For x As Integer = 1 To Header.ItemColl.Count : ItemColl(x).ReadStruct(map.fs) : Next
        'End If
        'End If

        If Header.PlayerSpawn.Count > 0 Then
            ReDim PlayerSpawn(Header.PlayerSpawn.Count)
            br1.BaseStream.Seek(CInt(Unsigned(Header.PlayerSpawn.Offset) - map.intMagic), SeekOrigin.Begin)
            For x As Integer = 1 To Header.PlayerSpawn.Count : PlayerSpawn(x).ReadStruct(map.fs) : Next
        End If

    End Function

#End Region

End Class

