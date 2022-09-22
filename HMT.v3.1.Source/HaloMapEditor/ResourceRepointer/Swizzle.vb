'/////////////////////////////////////////
'// Swizzle a set of Pixels
'// Donated by Stephen Cakebread
'/////////////////////////////////////////

Public Class Swizzle

    Public m_MaskX As Integer = 0
    Public m_MaskY As Integer = 0
    Public m_MaskZ As Integer = 0
    Public Sub New(ByVal width As Integer, ByVal height As Integer, ByVal depth As Integer)
        Dim bit As Integer = 1
        Dim idx As Integer = 1

        While (bit < width) Or (bit < height) Or (bit < depth)

            If (bit < width) Then
                m_MaskX = m_MaskX Or idx
                idx <<= 1
            End If

            If (bit < height) Then
                m_MaskY = m_MaskY Or idx
                idx <<= 1
            End If

            If (bit < depth) Then
                m_MaskZ = m_MaskZ Or idx
                idx <<= 1
            End If

            bit <<= 1
        End While
    End Sub

    Public Function Swizzle(ByVal Sx As Integer, ByVal Sy As Integer, Optional ByVal Sz As Integer = -1) As Int64
        Swizzle = SwizzleAxis(Sx, m_MaskX) Or SwizzleAxis(Sy, m_MaskY) Or (IIf(Sz <> -1, SwizzleAxis(Sz, m_MaskZ), 0))
    End Function

    Public Function SwizzleAxis(ByVal Value As Integer, ByVal Mask As Integer) As Int64

        Dim Result As Int64
        Dim bit As Integer = 1

        While bit <= Mask
            If (Mask And bit) Then
                Result = Result Or (Value And bit)
            Else
                Value <<= 1
            End If
            bit <<= 1
        End While
        SwizzleAxis = Result
    End Function

End Class
