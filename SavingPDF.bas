Attribute VB_Name = "SavingPDF"
Option Explicit

Sub SaveWorkbookAsPDF()

    'ThisWorkbook.ExportAsFixedFormat xlTypePDF, "C:\!Excel VBA\Wise Owl\PDFExample.pdf"
    ThisWorkbook.ExportAsFixedFormat xlTypePDF, Environ("UserProfile") & "\MyDocuments\PDFExample.pdf"
    
End Sub

Sub SaveOtherItemsAsPDFs()

    Sheet1.ExportAsFixedFormat xlTypePDF, Environ("UserProfile") & "\MyDocuments\Sheet1.pdf"
    Chart3.ExportAsFixedFormat xlTypePDF, Environ("UserProfile") & "\MyDocuments\Chart3.pdf"
    Sheet1.Range("A1").CurrentRegion.ExportAsFixedFormat xlTypePDF, Environ("UserProfile") & "\MyDocuments\RangeOfCells.pdf"
    
End Sub

Sub ExportEachChartAsPDF()

    Dim ch As Chart
    
    For Each ch In Charts
        ch.ExportAsFixedFormat xlTypePDF, Environ("UserProfile") & "\MyDocuments\" & ch.Name & ".pdf"
    Next ch
    
End Sub

































