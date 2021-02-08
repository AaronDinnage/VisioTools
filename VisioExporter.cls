VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Execute()
    Application.Settings.SetRasterExportResolution visRasterUseCustomResolution, 200#, 200#, visRasterPixelsPerInch
    Application.Settings.SetRasterExportSize visRasterFitToSourceSize
    Application.Settings.RasterExportDataFormat = visRasterNonInterlace
    Application.Settings.RasterExportColorFormat = visRaster24Bit
    Application.Settings.RasterExportRotation = visRasterNoRotation
    Application.Settings.RasterExportFlip = visRasterNoFlip
    Application.Settings.RasterExportBackgroundColor = 16777215
    Application.Settings.RasterExportTransparencyColor = 16777215
    Application.Settings.RasterExportUseTransparencyColor = False
    Application.Settings.SVGExportFormat = visSVGExcludeVisioElements

    Dim fileType As String
    fileType = "vsdx"

    Dim directory As String
    directory = ThisDocument.Path
    
    Dim filename As String
    filename = Dir(directory & "*." & fileType)
    
    Dim files As New Collection
    Do While filename <> ""
        filename = Left(filename, Len(filename) - Len(fileType) - 1)
        files.Add (filename)
        filename = Dir
    Loop
    
    For Each file In files
        filename = directory & "" & file & "." & fileType

        Dim doc As Document
        Set doc = Documents.Add(filename)

        Dim page As page
        For Each page In doc.Pages
            If page.Background = 0 Then
                Dim output As String
                output = directory & "" & page.Title

                doc.ExportAsFixedFormat visFixedFormatPDF, output & ".pdf", visDocExIntentPrint, visPrintFromTo, page.Index, page.Index

                page.Export (output & ".png")

                page.Export (output & ".svg")
            End If
        Next

        doc.Close
    Next
End Sub

