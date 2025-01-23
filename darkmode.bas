'This is the general macro.
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

'This is to be put in the code portion of a UserForm to make the UserForms match the darkmode of the sheet.
Private Sub UserForm_Initialize()
    ApplyDarkMode
End Sub
Private Sub UserForm_Activate()
    ApplyDarkMode
End Sub
Private Sub ApplyDarkMode()
    Dim ctrl As Control
    If ThisWorkbook.Sheets(1).Range("A1").Interior.Color = RGB(64, 64, 64) Then
        Me.BackColor = RGB(64, 64, 64)
        For Each ctrl In Me.Controls
            ctrl.ForeColor = RGB(255, 255, 255)
            ctrl.BackColor = RGB(64, 64, 64)
            If TypeOf ctrl Is TextBox Or TypeOf ctrl Is ComboBox Or TypeOf ctrl Is ListBox Then
                ctrl.BackColor = RGB(50, 50, 50)
            ElseIf TypeOf Control Is Label Or TypeOf ctrl Is CommandButton Or TypeOf ctrl Is Frame Then
                ctrl.BackColor = RGB(64, 64, 64)
            End If
        Next ctrl
    Else
        Me.BackColor = RGB(245, 245, 245)
        For Each ctrl In Me.Controls
            ctrl.ForeColor = RGB(0, 0, 0)
            ctrl.BackColor = RGB(245, 245, 245)
            If TypeOf ctrl Is TextBox Or TypeOf ctrl Is ComboBox Then
                ctrl.BackColor = RGB(245, 245, 245)
            ElseIf TypeOf ctrl Is Label Or TypeOf ctrl Is CommandButton Then
                ctrl.BackColor = RGB(245, 245, 245)
            End If
        Next ctrl
    End If
End Sub
