Public Class TreeViewEX
    Inherits System.Windows.Forms.TreeView

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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
        Me.Text = "TreeViewEX"
    End Sub

#End Region

    Public Class TreeNode
        Inherits System.Windows.Forms.TreeNode
        Implements IDictionaryEnumerator

        Private nodeEntry As New DictionaryEntry
        Private enumerator As IEnumerator

        Public Sub New()
            enumerator = MyBase.Nodes.GetEnumerator()
        End Sub

        Public Property NodeKey() As String
            Get
                Return nodeEntry.Key.ToString()
            End Get

            Set(ByVal Value As String)
                nodeEntry.Key = Value
            End Set
        End Property

        Public Property NodeValue() As Object
            Get
                Return nodeEntry.Value
            End Get

            Set(ByVal Value As Object)
                nodeEntry.Value = Value
            End Set
        End Property

        Public Overridable Overloads ReadOnly Property Entry() As DictionaryEntry _
            Implements IDictionaryEnumerator.Entry

            Get
                Return nodeEntry
            End Get
        End Property

        Public Overridable Overloads Function MoveNext() As Boolean _
            Implements IDictionaryEnumerator.MoveNext

            Dim Success As Boolean

            Success = enumerator.MoveNext()
            Return Success
        End Function

        Public Overridable Overloads ReadOnly Property Current() As Object _
            Implements IEnumerator.Current

            Get
                Return enumerator.Current
            End Get
        End Property

        Public Overridable Overloads ReadOnly Property Key() As Object _
            Implements IDictionaryEnumerator.Key

            Get
                Return nodeEntry.Key
            End Get
        End Property

        Public Overridable Overloads ReadOnly Property Value() As Object _
            Implements IDictionaryEnumerator.Value

            Get
                Return nodeEntry.Value
            End Get
        End Property

        Public Overridable Overloads Sub Reset() _
            Implements IEnumerator.Reset

            enumerator.Reset()
        End Sub
    End Class
End Class