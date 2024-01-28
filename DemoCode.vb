Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Windows.Media.Media3D
Imports Microsoft.Office.Interop.Excel

<HideModuleName>
Module MyGlobals
    ' Public ReadOnly xl As Application = Globals.ThisAddIn.Application
    Public ReadOnly xl As Application = ExcelDna.Integration.ExcelDnaUtil.Application
End Module

Module DemoCode

    Const wkbName As String = "perfDemo.xlsb"

    Dim wkb As Workbook

    Sub StartDemo()
        LoadWorkbook()
        If wkb Is Nothing Then Return

        Dim msg$ = String.Empty
        Dim addrArray = {"a1:z500", "a1:z5000", "a1:z50000"}

        'For demo I've skipped all the xl.Screenupdating/Events etc stuff.
        xl.Cursor = XlMousePointer.xlWait

        'First let's run the VBA for a few sizes
        For Each addr In addrArray
            xl.StatusBar = $"VBA Bold for {addr}..."
            Dim bold = VbaBold(addr)
            Debug.Print(bold)
            xl.StatusBar = $"VBA Formulas for {addr}..."
            Dim fmls = VbaFmls(addr)
            Debug.Print(fmls)
            msg &= $"{bold}{vbLf}{fmls}{vbLf}"
        Next
        'Now repeat for NET
        For Each addr In addrArray
            xl.StatusBar = $"NET Bold for {addr}..." & If(addr = addrArray(2), " Patience! This will take 1 minute or so!", "")
            Dim bold = NetBold(addr)
            Debug.Print(bold)
            xl.StatusBar = $"NET Formulas for {addr}..." & If(addr = addrArray(2), " Patience! This will take 1 minute or so!", "")
            Dim fmls = NetFmls(addr)
            Debug.Print(fmls)
            msg &= $"{bold}{vbLf}{fmls}{vbLf}"
        Next

        xl.StatusBar = Nothing
        xl.Cursor = XlMousePointer.xlDefault

        MsgBox(msg)


    End Sub

    Sub LoadWorkbook()
        If wkb Is Nothing Then
            For Each wb In xl.Workbooks.Cast(Of Workbook)
                If wb.Name.Equals(wkbName, StringComparison.OrdinalIgnoreCase) Then
                    wkb = wb
                    Return
                End If
            Next
            If wkb Is Nothing Then
                Dim path As String

                Dim assy = Reflection.Assembly.GetExecutingAssembly
                Dim aPath = IO.Path.Combine(assy.Location, wkbName)
                Dim bPath = IO.Path.Combine(IO.Path.GetDirectoryName(New Uri(assy.CodeBase).LocalPath), wkbName)

                If IO.File.Exists(aPath) Then
                    path = aPath
                ElseIf IO.File.Exists(bPath) Then
                    path = bPath
                Else
                    MsgBox("Problem locating the content workbook")
                    Return
                End If
                wkb = xl.Workbooks.Open(path)
                ' System.Windows.Forms.Application.DoEvents()

            End If
        End If
    End Sub

#Region "Timed wrappers"
    Function VbaBold(addr As String) As String
        Dim rng = DirectCast(wkb.Worksheets(1), Worksheet).Range(addr)
        Dim sw = Stopwatch.StartNew
        Dim cnt = CountBoldVba(rng)
        Dim ela = sw.ElapsedMilliseconds
        sw.Stop()
        Return $"BOLD VBA:{addr,16}{vbTab}Found:{cnt}{vbTab}Time:{ela}ms."
    End Function
    Function VbaFmls(addr As String) As String
        Dim rng = DirectCast(wkb.Worksheets(1), Worksheet).Range(addr)
        Dim sw = Stopwatch.StartNew
        Dim ret = GetHasFormulaVba(rng)
        Dim ela = sw.ElapsedMilliseconds
        sw.Stop()

        Dim cnt As Integer
        For r = ret.GetLowerBound(0) To ret.GetUpperBound(0)
            For c = ret.GetLowerBound(1) To ret.GetUpperBound(1)
                If ret(r, c) Then cnt += 1
            Next
        Next
        Return $"FMLS VBA:{addr,16}{vbTab}Found:{cnt}{vbTab}Time:{ela}ms."
    End Function
    Function NetBold(addr As String) As String
        Dim rng = DirectCast(wkb.Worksheets(1), Worksheet).Range(addr)
        Dim sw = Stopwatch.StartNew
        Dim cnt = CountBold(rng)
        Dim ela = sw.ElapsedMilliseconds
        sw.Stop()
        Return $"BOLD NET:{addr,16}{vbTab}Found:{cnt}{vbTab}Time:{ela}ms."
    End Function

    Function NetFmls(addr As String) As String
        Dim rng = DirectCast(wkb.Worksheets(1), Worksheet).Range(addr)
        Dim sw = Stopwatch.StartNew
        Dim ret = GetHasFormula(rng)
        Dim ela = sw.ElapsedMilliseconds
        sw.Stop()

        Dim cnt As Integer
        For r = ret.GetLowerBound(0) To ret.GetUpperBound(0)
            For c = ret.GetLowerBound(1) To ret.GetUpperBound(1)
                If ret(r, c) Then cnt += 1
            Next
        Next
        Return $"FMLS NET:{addr,16}{vbTab}Found:{cnt}{vbTab}Time:{ela}ms."
    End Function
#End Region


#Region "VBA Functions"
    Function GetHasFormulaVba(rng As Range) As Boolean(,)
        Dim obj = xl.Run($"{wkb.Name}!GetHasFormula", rng)
        Return DirectCast(obj, Boolean(,))
    End Function
    Function CountBoldVba(rng As Range) As Integer
        Dim ret = DirectCast(xl.Run($"{wkb.Name}!CountBold", rng), Integer)
        Return ret
    End Function
#End Region

#Region "NET functions"
    Function GetHasFormula(rng As Range) As Boolean(,)
        Dim col As Range
        Dim cel As Range
        Dim rows As Integer = rng.Rows.Count
        Dim cols As Integer = rng.Columns.Count
        Dim ret(rows, cols) As Boolean
        For c = 1 To cols
            col = rng.Columns(c)
            If col.HasFormula.Equals(True) Then
                For r = 1 To rows
                    ret(r, c) = True
                Next
            ElseIf IsDBNull(col.HasFormula) Then
                For r = 1 To rows
                    cel = col.Rows(r)
                    ret(r, c) = cel.HasFormula.Equals(True)
                    ' Marshal.ReleaseComObject(cel)
                Next
            End If
            ' Marshal.ReleaseComObject(col)
        Next
        Return ret

    End Function

    Function CountBold(rng As Range) As Integer
        Dim cnt = 0

        Dim rows = rng.Rows.Count
        Dim cols = rng.Columns.Count

        For r = 1 To rows
            For c = 1 To cols
                Dim cel = DirectCast(rng(r, c), Range)
                Dim fnt = cel.Font
                If fnt.Bold.Equals(True) Then cnt += 1
                ' Marshal.ReleaseComObject(fnt)
                ' Marshal.ReleaseComObject(cel)
            Next
        Next
        Return cnt

    End Function
#End Region



End Module
