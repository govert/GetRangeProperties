Imports System.ComponentModel
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Media
Imports ExcelDna.Integration
Imports ExcelDna.Integration.CustomUI
Imports GetRangeProperties.My.Resources

<ComVisible(True)>
Public Class MyRibbon
    Inherits ExcelRibbon

    Private xlRibbon As IRibbonUI

    Public Overrides Function GetCustomUI(RibbonID As String) As String
        Return RibbonResources.Ribbon
    End Function

    'Public Overrides Function LoadImage(imageId As String) As Object
    '    ' This will return the image resource with the name specified in the image='xxxx' tag
    '    Return RibbonResources.ResourceManager.GetObject(imageId)
    'End Function

#Region "Ribbon Callbacks"
    'Create callback methods here.
    'For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
    'Details on callbacks: https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa722523(v=office.12)

    Public Sub OnLoad(ribbonUI As IRibbonUI)
        xlRibbon = ribbonUI
    End Sub
    Public Sub OnAction(ctl As IRibbonControl)
        Select Case ctl.Id
            Case "btnPerf"
                ' We need these QueueAsMacro calls to support the C API
                ExcelAsyncUtil.QueueAsMacro(Sub() DemoCode.LoadWorkbook())
                ExcelAsyncUtil.QueueAsMacro(Sub() DemoCode.StartDemo())
            Case Else
        MsgBox("You clicked " + ctl.Id)
        End Select
    End Sub

    Public Function GetLabel(ctl As IRibbonControl) As Object
        Select Case ctl.Id
            Case "grpPerf" : Return "Perf Demo"
            Case "btnPerf" : Return "Run Test"
            Case Else : Return $"label_{ctl.Id}"
        End Select
    End Function
    Public Function GetImage(ctl As IRibbonControl) As Object
        Select Case ctl.Id
            Case "btnPerf" : Return CreateImage()
            Case Else : Return Nothing
        End Select
    End Function



#Region "Helpers"

    Private Shared Function GetResourceText(resourceName As String) As String
        Dim assy = Assembly.GetExecutingAssembly()
        Dim name = assy.GetManifestResourceNames.Where(Function(e) resourceName.Equals(e, StringComparison.OrdinalIgnoreCase)).FirstOrDefault
        Using resourceReader = New IO.StreamReader(assy.GetManifestResourceStream(name))
            Return resourceReader?.ReadToEnd()
        End Using
        Return Nothing
    End Function

    Private Shared Function CreateImage() As System.Drawing.Bitmap
        'Note: Path/Pen defined @ 300x300
        Dim brs = New SolidColorBrush(ColorConverter.ConvertFromString("#d90000"))
        Dim geo = New PathGeometry(TypeDescriptor.GetConverter(GetType(PathFigureCollection)).ConvertFrom("M0 100 h300 m0 100 h-300 M100 0 v300 m100 0 v-300"))
        Dim pen = New Pen(brs, 30)

        Dim rect = New Windows.Rect(0, 0, 48, 48)
        Dim scale = New ScaleTransform(rect.Width / 300, rect.Width / 300)

        Dim drwVisual As New DrawingVisual With {.Transform = scale}
        Using dc = drwVisual.RenderOpen
            dc.DrawGeometry(Nothing, pen, geo)
        End Using

        Dim bmpSource As New Imaging.RenderTargetBitmap(rect.Width, rect.Height, 96.0#, 96.0#, PixelFormats.Pbgra32)
        bmpSource.Render(drwVisual)

        Using ms As New IO.MemoryStream
            With New Imaging.PngBitmapEncoder
                .Frames.Add(Imaging.BitmapFrame.Create(bmpSource))
                .Save(ms)
            End With
            Return New System.Drawing.Bitmap(ms)
        End Using
    End Function
#End Region


#End Region
End Class
