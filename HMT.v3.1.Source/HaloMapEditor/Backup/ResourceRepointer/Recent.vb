Public Class RecentFile
    Private _DS As New DataSet
    Private _AppName As String
    Private _SessionID As String = ""

    Public Property Session_ID() As String
        Get
            Return _SessionID
        End Get
        Set(ByVal Value As String)
            _SessionID = Value
        End Set
    End Property

    Public Property Application_Name() As String
        Get
            Return _AppName
        End Get
        Set(ByVal Value As String)
            _AppName = Value
        End Set
    End Property

    Public ReadOnly Property GetRecentFilesList() As DataTable
        Get
            If _AppName.Trim <> "" Then
                If _SessionID.Trim <> "" Then
                    Return LoadList()
                Else
                    MsgBox("Please provide a Session ID", MsgBoxStyle.Critical, "Error")
                End If
            Else
                MsgBox("Please provide Application Name", MsgBoxStyle.Critical, "Error")
            End If
        End Get
    End Property

    Private Function LoadList() As DataTable 'this will get the list
        Try
            _DS.Tables.Clear()
            Dim DT As New DataTable(_AppName)
            DT.Columns.Add(New DataColumn("FileName"))
            DT.Columns.Add(New DataColumn("LastModified"))
            _DS.Tables.Add(DT)
            Dim Str As String = GetSetting(_AppName, "Recent", Me._SessionID, "")
            If Str.Trim <> "" Then
                Dim ST As New System.IO.StringReader(Str)
                Dim DS As New DataSet
                DS.ReadXml(ST)
                _DS.Merge(DS, False, MissingSchemaAction.Ignore)
            End If
            ClearUnFoundItems()
            TrimEntries(7)
            Return DT
        Catch ex As Exception
            MsgBox(ex.Message & " " & ex.StackTrace)
        End Try
    End Function

    Public Sub SaveItem(ByVal FileName As String) 'this will update or insert an item
        Try
            Dim Found As Boolean
            Dim X As Integer
            For X = 0 To _DS.Tables(0).Rows.Count - 1
                If _DS.Tables(0).Rows(X).Item("FileName").ToString.Trim.ToLower = FileName.Trim.ToLower Then
                    _DS.Tables(0).Rows(X).Item("LastModified") = Now.ToString
                    Found = True
                    Exit For
                End If
            Next
            If Found = False Then
                Dim DR As DataRow = _DS.Tables(0).NewRow
                DR.Item("FileName") = FileName
                DR.Item("LastModified") = Now.ToString
                _DS.Tables(0).Rows.Add(DR)
            End If
            Dim ST As New System.IO.StringWriter
            _DS.WriteXml(ST)
            Dim STR As String = ST.ToString
            SaveSetting(_AppName, "Recent", Me._SessionID, STR)
            TrimEntries(7)
        Catch ex As Exception
            MsgBox(ex.Message & " " & ex.StackTrace)
        End Try
    End Sub

    Private Sub ClearUnFoundItems() 'this will clear the list of items that don't exist anymore
        Try
            Dim X As Integer
            For X = 0 To _DS.Tables(0).Rows.Count - 1
                Dim FL As System.IO.File
                If FL.Exists(_DS.Tables(0).Rows(X).Item("FileName")) = False Then
                    _DS.Tables(0).Rows(X).Delete()
                End If
            Next
            _DS.Tables(0).AcceptChanges()
            Dim ST As New System.IO.StringWriter
            _DS.WriteXml(ST)
            Dim STR As String = ST.ToString
            SaveSetting(_AppName, "Recent", Me._SessionID, STR)
        Catch ex As Exception
            MsgBox(ex.Message & " " & ex.StackTrace)
        End Try
    End Sub

    Public Sub ClearItems() 'This will clear all entries
        Try
            For x As Integer = 0 To _DS.Tables(0).Rows.Count - 1
                _DS.Tables(0).Rows(x).Delete()
            Next
        Catch ex As Exception
            MsgBox(ex.Message & " " & ex.StackTrace)
        End Try
    End Sub

    Private Sub TrimEntries(ByVal MaxRows As Integer)
        'We want to remove all extries except for the latest (MaxRows)
        Dim EndRow As Integer = _DS.Tables(0).Rows.Count - MaxRows
        For x As Integer = 0 To EndRow - 1
            _DS.Tables(0).Rows(x).Delete()
        Next
        _DS.Tables(0).AcceptChanges()
        Dim ST As New System.IO.StringWriter
        _DS.WriteXml(ST)
        Dim STR As String = ST.ToString
        SaveSetting(_AppName, "Recent", Me._SessionID, STR)
    End Sub
End Class
