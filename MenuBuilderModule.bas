Attribute VB_Name = "MenuBuilderModule"
Option Explicit
'
'   The MenuBuilder routine will check for CustomProperty("VisibileInProductionMode" = True)
'   Set this custom property on each sheet in order for it to be included in the menu
'
'
'   Adjust the spacing and color of the menu using the menuFrame Enum below:
'
'
Private Enum menuFrame
    Left = 10
    Top = 10
    Width = 1200
    Height = 40
    bgColor = rgbDarkBlue
    btnSpacing = 10
    btnWdith = 80
    btnHeight = 20
    picPadding = 100
    btnActiveColor = rgbDarkSlateGray
    btnInactiveColor = rgbLightSlateGray
    btnFontSize = 8
End Enum


Public Sub MenuBuilder(book As Workbook)

    Dim ws As Worksheet
    Dim ch As Chart
    Dim iws As Worksheet
    Dim ich As Chart
    Dim sheetList As Collection: Set sheetList = New Collection
    Dim chartList As Collection: Set chartList = New Collection
    Dim frame As Shape
    Dim btn As Shape
    Dim i As Integer: i = 0
    Dim p As Object

    For Each ws In book.Worksheets
        If GetCustomProperty(ws, "VisibleInProductionMode") Then sheetList.Add ws
    Next

    For Each ws In sheetList
        Set frame = ws.Shapes.AddShape(msoShapeRoundedRectangle, menuFrame.Left, menuFrame.Top, menuFrame.Width, menuFrame.Height)
        frame.Name = "menuFrame"
        frame.Line.Visible = msoFalse
        frame.Fill.ForeColor.RGB = menuFrame.bgColor
        frame.Fill.OneColorGradient msoGradientDiagonalDown, 1, 1
        frame.Line.Weight = 0
        frame.ControlFormat.PrintObject = False
        frame.Placement = xlFreeFloating
        Set p = ws.Pictures.Insert("\\sercobpo.com.au\corpdata\Contracts\ATO\Resources\images\serco_transparent.png")
        p.Name = "menuLogo"
        p.Left = menuFrame.Left + menuFrame.btnSpacing
        p.Top = menuFrame.Top + menuFrame.btnSpacing / 2
        p.Width = 60
        p.Height = 30
        p.Placement = xlFreeFloating

        For Each iws In sheetList
            Set btn = ws.Shapes.AddShape(msoShapeRound2SameRectangle, menuFrame.Left + menuFrame.btnSpacing + (menuFrame.btnSpacing * i) + (menuFrame.btnWdith * i) + menuFrame.picPadding, _
            menuFrame.Top + menuFrame.btnSpacing, menuFrame.btnWdith, menuFrame.btnHeight)
            btn.Name = "menuButton_" & iws.Name
            btn.Placement = xlFreeFloating
            btn.ControlFormat.PrintObject = False
            btn.TextFrame.Characters.Text = iws.Name
            btn.TextFrame.Characters.Font.Size = menuFrame.btnFontSize
            btn.TextFrame.HorizontalAlignment = xlHAlignCenter
            btn.TextFrame.VerticalAlignment = xlVAlignCenter
            If ws.Name = iws.Name Then
                btn.Fill.ForeColor.RGB = menuFrame.btnActiveColor
                btn.Line.Weight = 0
                'btn.TextFrame.Characters.Font.Bold = activeButton.Bold
            Else
                btn.Fill.ForeColor.RGB = menuFrame.btnInactiveColor
                btn.Line.Weight = 0
                'btn.TextFrame.Characters.Font.Bold = inactiveButton.Bold
                ws.Hyperlinks.Add btn, "", "'" & iws.Name & "'!A1", iws.Name, iws.Name
            End If
            i = i + 1
        Next

        Set btn = ws.Shapes.AddShape(msoShapeRound2SameRectangle, menuFrame.Left + menuFrame.btnSpacing + (menuFrame.btnSpacing * i) + (menuFrame.btnWdith * i) + menuFrame.picPadding, _
        menuFrame.Top + menuFrame.btnSpacing, menuFrame.btnWdith, menuFrame.btnHeight)
        btn.Name = "menuButton_Refresh"
        btn.Placement = xlFreeFloating
        btn.ControlFormat.PrintObject = False
        btn.TextFrame.Characters.Text = "Refresh"
        btn.TextFrame.HorizontalAlignment = xlHAlignCenter
        btn.TextFrame.VerticalAlignment = xlVAlignCenter
        btn.Fill.ForeColor.RGB = menuFrame.btnActiveColor
        btn.Line.Weight = 0
        'btn.TextFrame.Characters.Font.Bold = activeButton.Bold

        i = 0
    Next

End Sub


Private Function GetCustomProperty(ws As Worksheet, property As String) As Variant
    Dim cp As CustomProperty
    For Each cp In ws.CustomProperties
        If cp.Name = property Then
            GetCustomProperty = cp.Value
            Exit Function
        End If
    Next
End Function


Sub RemoveMenus(workbook As workbook)
    Dim ws As Worksheet
    Dim sh As Shape
    For Each ws In workbook.Worksheets
        For Each sh In ws.Shapes
            'If Strings.Left(sh.Name, 4) = "menu" Then Debug.Print "yes"
            'Debug.Print sh.Name
            'Debug.Print Strings.Left(sh.Name, 4)
            If Strings.Left(sh.Name, 4) = "menu" Then sh.Delete
        Next
    Next
End Sub


