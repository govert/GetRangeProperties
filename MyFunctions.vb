Imports ExcelDna.Integration

Public Module MyFunctions
    <ExcelFunction(Description:="My first .NET function")>
    Public Function SayHello(ByVal name As String) As String
        SayHello = "Hello " + name
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function GetFormula(<ExcelArgument(AllowReference:=True)> input As Object) As String
        Dim formula As String
        If TypeOf input Is ExcelReference Then
            formula = XlCall.Excel(XlCall.xlfGetFormula, input)
            Return "Formula: " & formula
        Else
            Return "<Not a reference>"
        End If
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function GetBold(<ExcelArgument(AllowReference:=True)> input As Object) As String
        Dim isBold As Boolean
        If TypeOf input Is ExcelReference Then
            isBold = XlCall.Excel(XlCall.xlfGetCell, 20, input)
            Return "Is Bold: " & isBold
        Else
            Return "<Not a reference>"
        End If
    End Function
End Module
