Imports System.Linq
Partial Public Module DemoCatalog
    Private ReadOnly _all As List(Of DemoDefinition) = CreateAll()
    Public Function GetAll() As List(Of DemoDefinition)
        Return _all
    End Function
    Public Function GetByCategory(category As DemoCategory) As List(Of DemoDefinition)
        Return _all.
            Where(Function(x) x.Category = category).
            OrderBy(Function(x) x.Title).
            ToList()
    End Function
    Private Function CreateAll() As List(Of DemoDefinition)
        Dim result As New List(Of DemoDefinition)
        result.AddRange(GetRegexExamples())
        result.AddRange(GetPowerQueryExamples())
        result.AddRange(GetVbaExamples())
        result.AddRange(GetFormulaExamples())
        Return result
    End Function
End Module