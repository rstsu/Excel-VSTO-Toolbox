Imports System.Xml.Linq
Public Module TextHelpers
    Public Function TextBlock(value As XElement) As String
        Return value.Value.Trim().Replace(vbLf, vbCrLf)
    End Function

End Module