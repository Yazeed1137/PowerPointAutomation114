Attribute VB_Name = "ClipboardAutomaton"

Option Explicit

Private Const DEFAULT_FONT_NAME As String = "DIN Next LT Arabic"
Private Const DEFAULT_FONT_SIZE As Single = 22
Private Const DEFAULT_FONT_COLOR As Long = 5460819 ' #535353
Private Const MAX_CHARS_PER_SLIDE As Long = 190
Private Const PASTE_BUTTON_NAME As String = "btnPasteFromClipboard"
Private Const PASTE_BUTTON_MACRO As String = "PasteClipboardToSlides"

Private Type SlideTextFormat
    FontName As String
    FontSize As Single
    FontColor As Long
    Bold As MsoTriState
    Italic As MsoTriState
    Alignment As PpParagraphAlignment
End Type

' ========================================================
' Entry point:
' - Reads plain text from clipboard
' - Splits text into chunks (max 190 chars each)
' - Places each chunk on a slide
' - Applies slide font/style defaults while preserving discovered style
Sub PasteClipboardToSlides()
    Dim activePres As Presentation
    Dim baseSlide As Slide
    Dim clipboardText As String
    Dim textChunks As Collection
    Dim chunkText As Variant
    Dim currentSlide As Slide
    Dim sourceFormat As SlideTextFormat
    Dim chunkIndex As Long

    Set activePres = Application.ActivePresentation
    If activePres Is Nothing Then Exit Sub

    If activePres.Slides.Count = 0 Then
        MsgBox "Open or create a presentation with at least one slide.", vbExclamation
        Exit Sub
    End If

    clipboardText = GetClipboardText()
    clipboardText = NormalizeClipboardText(clipboardText)

    If Len(clipboardText) = 0 Then
        MsgBox "Clipboard is empty (or does not contain text).", vbExclamation
        Exit Sub
    End If

    Set baseSlide = GetCurrentSlide(activePres)
    If baseSlide Is Nothing Then Set baseSlide = activePres.Slides(1)
    sourceFormat = GetSlideTextFormat(baseSlide)

    Set textChunks = SplitByCharacterLimit(clipboardText, MAX_CHARS_PER_SLIDE)
    chunkIndex = 0

    For Each chunkText In textChunks
        chunkIndex = chunkIndex + 1

        If chunkIndex = 1 Then
            Set currentSlide = baseSlide
        Else
            Set currentSlide = DuplicateSlideAfter(activePres, currentSlide)
        End If

        WriteChunkToSlide currentSlide, CStr(chunkText), sourceFormat
    Next chunkText
End Sub


' ========================================================
' Adds a clickable button on the current slide that runs PasteClipboardToSlides
Sub AddPasteFromClipboardButton()
    Dim activePres As Presentation
    Dim targetSlide As Slide
    Dim btn As Shape

    Set activePres = Application.ActivePresentation
    If activePres Is Nothing Then Exit Sub

    Set targetSlide = GetCurrentSlide(activePres)
    If targetSlide Is Nothing Then Set targetSlide = activePres.Slides(1)

    RemovePasteFromClipboardButton targetSlide

    Set btn = targetSlide.Shapes.AddShape(msoShapeRoundedRectangle, _
        Left:=20, _
        Top:=20, _
        Width:=170, _
        Height:=36)

    btn.Name = PASTE_BUTTON_NAME
    btn.TextFrame.TextRange.Text = "Paste from Clipboard"
    btn.TextFrame.TextRange.Font.Name = "Calibri"
    btn.TextFrame.TextRange.Font.Size = 12
    btn.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter

    btn.Fill.ForeColor.RGB = RGB(83, 83, 83)
    btn.Line.ForeColor.RGB = RGB(83, 83, 83)
    btn.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)

    btn.ActionSettings(ppMouseClick).Action = ppActionRunMacro
    btn.ActionSettings(ppMouseClick).Run = PASTE_BUTTON_MACRO
End Sub

Private Sub RemovePasteFromClipboardButton(ByVal targetSlide As Slide)
    Dim shp As Shape

    For Each shp In targetSlide.Shapes
        If shp.Name = PASTE_BUTTON_NAME Then
            shp.Delete
            Exit Sub
        End If
    Next shp
End Sub


Private Function GetCurrentSlide(ByVal activePres As Presentation) As Slide
    On Error GoTo NoCurrentSlide

    If Not ActiveWindow Is Nothing Then
        Select Case ActiveWindow.ViewType
            Case ppViewNormal, ppViewSlide, ppViewNotesPage
                Set GetCurrentSlide = activePres.Slides(ActiveWindow.View.Slide.SlideIndex)
                Exit Function
            Case ppViewSlideShow
                Set GetCurrentSlide = activePres.Slides(SlideShowWindows(1).View.Slide.SlideIndex)
                Exit Function
        End Select
    End If

NoCurrentSlide:
    Set GetCurrentSlide = Nothing
End Function

Private Function GetClipboardText() As String
    Dim dataObj As Object

    On Error GoTo ClipboardFallback
    Set dataObj = CreateObject("MSForms.DataObject")
    dataObj.GetFromClipboard
    GetClipboardText = dataObj.GetText
    Exit Function

ClipboardFallback:
    GetClipboardText = ""
End Function

Private Function NormalizeClipboardText(ByVal inputText As String) As String
    Dim normalized As String

    normalized = Replace(inputText, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)

    Do While InStr(normalized, vbLf & vbLf & vbLf) > 0
        normalized = Replace(normalized, vbLf & vbLf & vbLf, vbLf & vbLf)
    Loop

    NormalizeClipboardText = Trim(normalized)
End Function

Private Function SplitByCharacterLimit(ByVal sourceText As String, ByVal maxChars As Long) As Collection
    Dim parts As New Collection
    Dim cursor As Long
    Dim breakPos As Long
    Dim piece As String
    Dim startIdx As Long

    If maxChars < 1 Then maxChars = 1

    sourceText = Trim(sourceText)
    If Len(sourceText) = 0 Then
        parts.Add ""
        Set SplitByCharacterLimit = parts
        Exit Function
    End If

    startIdx = 1
    Do While startIdx <= Len(sourceText)
        If Len(sourceText) - startIdx + 1 <= maxChars Then
            piece = Mid$(sourceText, startIdx)
            parts.Add Trim(piece)
            Exit Do
        End If

        breakPos = InStrRev(sourceText, " ", startIdx + maxChars - 1)
        If breakPos < startIdx Then
            breakPos = startIdx + maxChars - 1
        End If

        piece = Mid$(sourceText, startIdx, breakPos - startIdx + 1)
        parts.Add Trim(piece)
        startIdx = breakPos + 1

        Do While startIdx <= Len(sourceText) And Mid$(sourceText, startIdx, 1) = " "
            startIdx = startIdx + 1
        Loop
    Loop

    Set SplitByCharacterLimit = parts
End Function

Private Function GetSlideTextFormat(ByVal sourceSlide As Slide) As SlideTextFormat
    Dim fmt As SlideTextFormat
    Dim shp As Shape

    fmt.FontName = DEFAULT_FONT_NAME
    fmt.FontSize = DEFAULT_FONT_SIZE
    fmt.FontColor = DEFAULT_FONT_COLOR
    fmt.Bold = msoFalse
    fmt.Italic = msoFalse
    fmt.Alignment = ppAlignJustifyLow

    For Each shp In sourceSlide.Shapes
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                With shp.TextFrame.TextRange
                    If Len(.Font.Name) > 0 Then fmt.FontName = .Font.Name
                    If .Font.Size > 0 Then fmt.FontSize = .Font.Size
                    If .Font.Color.RGB <> 0 Then fmt.FontColor = .Font.Color.RGB
                    fmt.Bold = .Font.Bold
                    fmt.Italic = .Font.Italic
                    fmt.Alignment = .ParagraphFormat.Alignment
                End With
                Exit For
            End If
        End If
    Next shp

    If fmt.Alignment <> ppAlignJustifyLow Then fmt.Alignment = ppAlignJustifyLow

    GetSlideTextFormat = fmt
End Function

Private Function DuplicateSlideAfter(ByVal activePres As Presentation, ByVal sourceSlide As Slide) As Slide
    Dim newSlide As Slide

    Set newSlide = sourceSlide.Duplicate()(1)
    newSlide.MoveTo sourceSlide.SlideIndex + 1

    Set DuplicateSlideAfter = activePres.Slides(sourceSlide.SlideIndex + 1)
End Function

Private Sub WriteChunkToSlide(ByVal targetSlide As Slide, ByVal chunkText As String, ByVal fmt As SlideTextFormat)
    Dim textShape As Shape

    Set textShape = GetOrCreatePrimaryTextShape(targetSlide)
    ClearSlideTextBoxes targetSlide

    textShape.TextFrame.TextRange.Text = chunkText
    ApplyFormatting textShape, fmt
End Sub

Private Sub ApplyFormatting(ByVal shp As Shape, ByVal fmt As SlideTextFormat)
    With shp.TextFrame.TextRange
        .Font.Name = fmt.FontName
        .Font.Size = fmt.FontSize
        .Font.Color.RGB = fmt.FontColor
        .Font.Bold = fmt.Bold
        .Font.Italic = fmt.Italic
        .ParagraphFormat.Alignment = ppAlignJustifyLow
    End With
End Sub

Private Function GetOrCreatePrimaryTextShape(ByVal targetSlide As Slide) As Shape
    Dim shp As Shape

    For Each shp In targetSlide.Shapes
        If shp.Type = msoTextBox Then
            Set GetOrCreatePrimaryTextShape = shp
            Exit Function
        End If
    Next shp

    Set GetOrCreatePrimaryTextShape = targetSlide.Shapes.AddTextbox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=40, _
        Top:=60, _
        Width:=targetSlide.Parent.PageSetup.slideWidth - 80, _
        Height:=targetSlide.Parent.PageSetup.slideHeight - 120)
End Function

Private Sub ClearSlideTextBoxes(ByVal targetSlide As Slide)
    Dim shp As Shape

    For Each shp In targetSlide.Shapes
        If shp.Type = msoTextBox Then
            shp.TextFrame.TextRange.Text = ""
        End If
    Next shp
End Sub
