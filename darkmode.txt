Sub ToggleDarkLightMode()
    Dim ws As Worksheet
    Dim darkModeBackColor As Long
    Dim darkModeFontColor As Long
    Dim lightModeBackColor As Long
    Dim lightModeFontColor As Long
    Dim originalSheet As Worksheet
    Dim isDarkMode As Boolean
    Set originalSheet = ActiveSheet
    darkModeBackColor = RGB(64, 64, 64)
    darkModeFontColor = RGB(243, 243, 243)
    lightModeBackColor = RGB(256, 256, 256)
    lightModeFontColor = RGB(0, 0, 0)
    Application.ScreenUpdating = False
    isDarkMode = (ThisWorkbook.Sheets(1).Cells(1, 1).Interior.Color = darkModeBackColor)
    For Each ws In ThisWorkbook.Sheets
        ws.Activate
        If isDarkMode Then
            ws.Cells.Interior.ColorIndex = 0
            ws.Cells.Font.Color = lightModeFontColor
            ChangeButtons ws, lightModeBackColor, lightModeFontColor, "Dark Mode"
        Else
            ws.Cells.Interior.Color = darkModeBackColor
            ws.Cells.Font.Color = darkModeFontColor
            ChangeButtons ws, darkModeBackColor, darkModeFontColor, "Light Mode"
        End If
    Next ws
    originalSheet.Activate
    Application.ScreenUpdating = True
End Sub
Sub ChangeButtons(ws As Worksheet, bgColor As Long, fontColor As Long, newText As String)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.TextFrame2.HasText Then
            shp.Fill.ForeColor.RGB = bgColor
            shp.TextFrame.Characters.Font.Color = fontColor
            If shp.Name = "ToggleButton" Then
                shp.OLEFormat.Object.Caption = newText
            End If
        End If
    Next shp
End Sub
