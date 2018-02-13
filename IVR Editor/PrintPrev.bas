Attribute VB_Name = "mdlPrintPrev"
Global zoomindex As Integer
Sub GetZoom(zoomlabel As Integer)
'Set up the print previews zoom

        Select Case zoomlabel
            Case 0
                frmPrintPreview.fpSpreadPreview1.PageViewType = 2
                frmPrintPreview.fpSpreadPreview1.PageViewPercentage = 200
            
            Case 1
                frmPrintPreview.fpSpreadPreview1.PageViewType = 2
                frmPrintPreview.fpSpreadPreview1.PageViewPercentage = 150

            Case 2
                frmPrintPreview.fpSpreadPreview1.PageViewType = 2
                frmPrintPreview.fpSpreadPreview1.PageViewPercentage = 100

            Case 3
                frmPrintPreview.fpSpreadPreview1.PageViewType = 2
                frmPrintPreview.fpSpreadPreview1.PageViewPercentage = 75

            Case 4
                frmPrintPreview.fpSpreadPreview1.PageViewType = 2
                frmPrintPreview.fpSpreadPreview1.PageViewPercentage = 50

            Case 5
                frmPrintPreview.fpSpreadPreview1.PageViewType = 2
                frmPrintPreview.fpSpreadPreview1.PageViewPercentage = 25

            Case 6
                frmPrintPreview.fpSpreadPreview1.PageViewType = 2
                frmPrintPreview.fpSpreadPreview1.PageViewPercentage = 10

            Case 7
                frmPrintPreview.fpSpreadPreview1.PageViewType = 3
                
            Case 8
                frmPrintPreview.fpSpreadPreview1.PageViewType = 4
                
            Case 9
                frmPrintPreview.fpSpreadPreview1.PageViewType = 0
                
            Case 10
                frmPrintPreview.fpSpreadPreview1.PageViewType = 5
                frmPrintPreview.fpSpreadPreview1.PageMultiCntH = 2
                frmPrintPreview.fpSpreadPreview1.PageMultiCntV = 1
                
            Case 11
                frmPrintPreview.fpSpreadPreview1.PageViewType = 5
                frmPrintPreview.fpSpreadPreview1.PageMultiCntH = 3
                frmPrintPreview.fpSpreadPreview1.PageMultiCntV = 1
                
            Case 12
                frmPrintPreview.fpSpreadPreview1.PageViewType = 5
                frmPrintPreview.fpSpreadPreview1.PageMultiCntH = 2
                frmPrintPreview.fpSpreadPreview1.PageMultiCntV = 2
                
            Case 13
                frmPrintPreview.fpSpreadPreview1.PageViewType = 5
                frmPrintPreview.fpSpreadPreview1.PageMultiCntH = 3
                frmPrintPreview.fpSpreadPreview1.PageMultiCntV = 2

        End Select
      
End Sub

