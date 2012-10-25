Attribute VB_Name = "basFormSize"
Option Explicit

' FormLayout()
' Resize a form and all its controls based on screen resolution.
' All controls are resized in TabIndex order. Advanced positioning lets you
' position a control based on the previous (in TabIndex order) control's
' newly resized coordinates. Set any control's Tag property to one of the
' following (without the quotes) to engage advanced positioning.
' "Fixed"
'    The control will not be resized, but it will be repositioned.
' "Multiline"
'    Some controls (Label, OptionButton, TextBox) are autosized
'    based on the font width and/or height instead of the overall ratio,
'    causing problems for multi-line controls that use word wrap. Use
'    this to skip the font-based resizing and just use generic resizing.
' "Left"
'    If the previous control is an OptionButton or CheckBox, align with
'    its caption left. This lets you align "subitems".
' "Right"
'    Position this control to the right of the previous control on the
'    same line. If the control is a TextBox and the previous control is a
'    Label, this setting will line them up vertically as well.
' "Label"
'    If the previous control is a Label, this setting will move the label down
'    slightly to align the text bottom.
' "Etched"
'    Signifies an Etched label, which simulates the etched effect on
'    disabled controls like checkboxes. To use this effect, set the main
'    label's background to transparent, add a new label with the same
'    caption, set its ForeColor to white, set its ZOrder behind the main
'    label (send to back), and set its Tag property to Etched. Be sure
'    the etched label is immediately after the main label in TabOrder.
' Parameters:
' pfrm
'    The form to be resized.
' psngFontSize can be either:
'     - The new font size for the form and all its controls. Font sizes can
'       be multiples of 0.25.
'     - A negative value (eg: -46) to calculate the new font size based on how
'       many lines of text can fit on a full screen in the current resolution.
'       46 lines (-46) results in nice compact text. 41 lines (-41) results
'       in a nice looking larger font. This is a great setting to expose to
'       the user in some way. (But don't show them negatives!)
'
' Sample usage:
'
' Private Sub Form_Load()
'     FormLayout Me, 2
' End Sub
'
Public Sub FormLayout(pfrm As Form, Optional ByVal psngFontSize As Single = -46)

    Const Spacer = " "

    Const OptionSpacer = "o"

    Const CaptionLeftChar = " "

    Const CaptionLeftOffset = 14

    Dim sngOriginalFontSize As Single

    Dim enScaleMode         As ScaleModeConstants

    Dim lngScaleLeft        As Long

    Dim lngScaleTop         As Long

    Dim lngScaleWidth       As Long

    Dim lngScaleHeight      As Long

    Dim lngLeft             As Long

    Dim lngTop              As Long

    Dim lngWidth            As Long

    Dim lngHeight           As Long

    Dim sngX                As Single

    Dim sngY                As Single

    Dim lngOldWidth         As Long

    Dim lngOldHeight        As Long

    Dim lngExtraWidth       As Long

    Dim lngExtraHeight      As Long

    Dim lngScreenHeight     As Long

    Dim lngLines            As Long

    Dim lngTabIndex()       As Long

    Dim i                   As Long

    Dim iMax                As Long

    Dim ctl                 As Control

    Dim ctlPrev             As Control

    Dim ctlEtched           As Control

    Dim lngSSTab            As Long

    Dim strCaption          As String

    Dim strCaptionPrev      As String

    Dim lngComboHeight      As Long

    With pfrm
        ' Remember starting font size (do nothing if it doesn't end up changing)
        sngOriginalFontSize = .FontSize
        ' Save scale settings so we can set ScaleMode to Twips, greatly simplifying things
        enScaleMode = .ScaleMode

        If enScaleMode = vbUser Then
            lngScaleLeft = .ScaleLeft
            lngScaleTop = .ScaleTop
            lngScaleWidth = .ScaleWidth
            lngScaleHeight = .ScaleHeight

        End If

        .ScaleMode = vbTwips
        ' Get existing dimensions
        lngOldWidth = .TextWidth(Spacer)
        lngOldHeight = .TextHeight(Spacer)
        lngExtraWidth = .Width - .ScaleWidth
        lngExtraHeight = .Height - .ScaleHeight

        ' Calculate new font size based on lines of text per screen
        If psngFontSize < 0 Then
            lngLines = Abs(psngFontSize)
            lngScreenHeight = Screen.Height
            .FontSize = 6 ' Arbitrary minimum size
            ' Grow font until the screen holds < requested number of lines (one size too big)
            ' (psngFontSize retains the previous value, so it'll end up with the correct size)
            Do
                psngFontSize = .FontSize
                ' Font sizes have erratic increments; identify the next size up
                .FontSize = .FontSize + 0.25

                If psngFontSize = .FontSize Then .FontSize = .FontSize + 0.5
                If psngFontSize = .FontSize Then .FontSize = .FontSize + 0.75
                If psngFontSize = .FontSize Then .FontSize = .FontSize + 1
                If psngFontSize = .FontSize Then Exit Do
            Loop Until lngScreenHeight \ .TextHeight(Spacer) < lngLines

        End If

        ' Commit to the new font size
        .FontSize = psngFontSize

    End With

    ' Do nothing if font size hasn't changed
    If pfrm.FontSize <> sngOriginalFontSize Then

        With pfrm
            ' Calculate ratios based on change in font size
            sngX = .TextWidth(Spacer) / lngOldWidth
            sngY = .TextHeight(Spacer) / lngOldHeight
            ' Resize form
            lngWidth = .ScaleWidth * sngX + lngExtraWidth
            lngHeight = .ScaleHeight * sngY + lngExtraHeight

            ' Center form if it was already centered, otherwise don't move it
            If .Left <> (Screen.Width - .Width) \ 2 Then lngLeft = .Left Else lngLeft = (Screen.Width - lngWidth) \ 2
            If .Top <> (Screen.Height - .Height) \ 2 Then lngTop = .Top Else lngTop = (Screen.Height - lngHeight) \ 2
            .Move lngLeft, lngTop, lngWidth, lngHeight
            ' Identify TabIndex order
            iMax = .Controls.Count - 1

        End With

        If iMax >= 0 Then
            ReDim lngTabIndex(iMax)

            For Each ctl In pfrm.Controls

                ' Resize lines & shapes now because they don't have a TabIndex
                With ctl

                    Select Case TypeName(ctl)

                        Case "Line"

                            ' Identify left offset (used for controls on an inactive SSTab tab)
                            If TypeName(.Container) = "SSTab" And .X1 < -1500 Then lngSSTab = 75000 Else lngSSTab = 0
                            .X1 = (.X1 + lngSSTab) * sngX - lngSSTab
                            .X2 = (.X2 + lngSSTab) * sngX - lngSSTab
                            .Y1 = .Y1 * sngY
                            .Y2 = .Y2 * sngY
                            iMax = iMax - 1

                        Case "Shape", "Image"

                            ' Identify left offset (used for controls on an inactive SSTab tab)
                            If TypeName(.Container) = "SSTab" And .Left < -1500 Then lngSSTab = 75000 Else lngSSTab = 0
                            .Move (.Left + lngSSTab) * sngX - lngSSTab, .Top * sngY, .Width * sngX, .Height * sngY
                            iMax = iMax - 1

                        Case Else

                            On Error Resume Next

                            lngTabIndex(.TabIndex) = i

                            If Err.Number <> 0 Then iMax = iMax - 1

                            On Error GoTo 0

                            ' Identify ComboBox height
                            If TypeOf ctl Is ComboBox And lngComboHeight = 0 Then
                                .FontSize = pfrm.FontSize
                                lngComboHeight = ctl.Height

                            End If

                    End Select

                End With

                i = i + 1
            Next

            ' Identify standard textbox height now to speed up loop
            If lngComboHeight = 0 Then lngComboHeight = pfrm.TextHeight(Spacer) + 4 * Screen.TwipsPerPixelY

            ' Iterate controls in TabIndex order
            For i = 0 To iMax
                Set ctl = pfrm.Controls(lngTabIndex(i))

                Select Case TypeName(ctl)

                    Case "OptionButton", "CheckBox": strCaption = Replace(Replace(Replace(ctl.Caption, "&&", "~"), "&", ""), "~", "&")

                End Select

                If i <> 0 Then
                    Set ctlPrev = pfrm.Controls(lngTabIndex(i - 1))

                    Select Case TypeName(ctlPrev)

                        Case "OptionButton", "CheckBox": strCaptionPrev = Replace(Replace(Replace(ctlPrev.Caption, "&&", "~"), "&", ""), "~", "&")

                    End Select

                    If ctl.Tag = "Etched" Then Set ctlEtched = pfrm.Controls(lngTabIndex(i - 1))

                End If

                With ctl

                    ' Identify left offset (used for controls on an inactive SSTab tab)
                    If TypeName(.Container) = "SSTab" And .Left < -1500 Then lngSSTab = 75000 Else lngSSTab = 0
                    ' Identify current dimensions
                    lngLeft = .Left + lngSSTab
                    lngTop = .Top
                    lngWidth = .Width
                    lngHeight = .Height

                    ' LEFT
                    Select Case .Tag

                        Case "Left"

                            Select Case TypeName(ctlPrev)

                                Case "OptionButton", "CheckBox": lngLeft = ctlPrev.Left + CaptionLeftOffset * Screen.TwipsPerPixelX + pfrm.TextWidth(CaptionLeftChar)

                                Case Else: lngLeft = ctlPrev.Left

                            End Select

                        Case "Right"

                            Select Case TypeName(ctlPrev)

                                Case "OptionButton", "CheckBox": lngLeft = ctlPrev.Left + (CaptionLeftOffset + 1) * Screen.TwipsPerPixelX + pfrm.TextWidth(strCaptionPrev) + pfrm.TextWidth(OptionSpacer)

                                Case Else: lngLeft = ctlPrev.Left + ctlPrev.Width + pfrm.TextWidth(Spacer)

                            End Select

                        Case "Etched": lngLeft = ctlPrev.Left + Screen.TwipsPerPixelX

                        Case Else: lngLeft = lngLeft * sngX

                    End Select

                    ' TOP
                    Select Case .Tag

                        Case "Etched": lngTop = ctlPrev.Top + Screen.TwipsPerPixelY

                        Case Else: lngTop = lngTop * sngY

                    End Select

                    ' WIDTH
                    Select Case .Tag

                        Case "Fixed"

                        Case "MultiLine": lngWidth = lngWidth * sngX

                        Case "Etched": lngWidth = ctlPrev.Width

                        Case Else

                            Select Case TypeName(ctl)

                                Case "OptionButton", "CheckBox": lngWidth = CaptionLeftOffset * Screen.TwipsPerPixelX + 2 * pfrm.TextWidth(CaptionLeftChar) + pfrm.TextWidth(strCaption)

                                Case "TextBox": If .MaxLength <> 0 Then lngWidth = pfrm.TextWidth("8") * (.MaxLength + 1) Else lngWidth = lngWidth * sngX

                                Case Else: lngWidth = lngWidth * sngX

                            End Select

                    End Select

                    ' HEIGHT
                    Select Case .Tag

                        Case "Fixed"

                        Case "MultiLine": lngHeight = lngHeight * sngY

                        Case "Etched": lngHeight = ctlPrev.Height

                        Case Else

                            Select Case TypeName(ctl)

                                Case "OptionButton", "CheckBox": lngHeight = lngComboHeight

                                Case "ListBox"
                                    lngLines = ctl.Height \ lngOldHeight
                                    lngHeight = pfrm.TextHeight(Spacer) * lngLines + ctl.Height - (lngLines * lngOldHeight)

                                Case "TextBox": lngHeight = lngComboHeight

                                Case Else: lngHeight = lngHeight * sngY

                            End Select

                    End Select

                    ' Apply new formatting
                    On Error Resume Next

                    .Font.Size = pfrm.FontSize

                    On Error GoTo 0

                    Select Case TypeName(ctl)

                        Case "Label"

                            Select Case .Tag

                                Case "MultiLine", "Etched"

                                Case Else
                                    .AutoSize = True
                                    lngHeight = .Height

                                    Select Case .Alignment

                                        Case vbRightJustify: If lngWidth < .Width Then lngLeft = lngLeft - (.Width - lngWidth)

                                        Case vbCenter: If lngWidth < .Width Then lngLeft = lngLeft - (.Width - lngWidth) \ 2

                                    End Select

                                    lngWidth = .Width

                            End Select

                            .Move lngLeft, lngTop, lngWidth, lngHeight

                        Case "ComboBox"
                            lngComboHeight = .Height
                            .Move lngLeft, lngTop, lngWidth

                        Case Else
                            .Move lngLeft, lngTop, lngWidth, lngHeight

                    End Select

                    ' Check for vertical align
                    If i <> 0 Then
                        If TypeOf ctlPrev Is Label Then

                            Select Case .Tag

                                Case "Label", "Right"

                                    ' If previous control is an Etched label, move both labels
                                    If ctlPrev.Tag = "Etched" Then
                                        ctlEtched.Top = ctl.Top + 3 * Screen.TwipsPerPixelY
                                        ctlPrev.Top = ctlEtched.Top + Screen.TwipsPerPixelY
                                    Else
                                        ctlPrev.Top = ctl.Top + 3 * Screen.TwipsPerPixelY

                                    End If

                            End Select

                        End If

                    End If

                End With

                Set ctl = Nothing
            Next
            Set ctlPrev = Nothing
            Set ctlEtched = Nothing

        End If

    End If

    ' Reset ScaleMode to original settings
    With pfrm

        If enScaleMode = vbUser Then
            .ScaleLeft = lngScaleLeft
            .ScaleTop = lngScaleTop
            .ScaleWidth = lngScaleWidth
            .ScaleHeight = lngScaleHeight
        Else
            .ScaleMode = enScaleMode

        End If

    End With

End Sub



Public Function GetOptionCaptionLeft(popt As OptionButton) As Long

    Const Pixels = 15

    Const Char = " "

    GetOptionCaptionLeft = popt.Left + Pixels * Screen.TwipsPerPixelX + popt.Parent.TextWidth(Char)

End Function

Public Function GetOptionHeight(popt As OptionButton) As Long

    Const Pixels = 4

    Const Char = "Q"

    With popt.Parent
        GetOptionHeight = Pixels * Screen.TwipsPerPixelX + .TextHeight(Char)

    End With

End Function

Public Function GetOptionRight(popt As OptionButton) As Long

    Const Pixels = 15

    Const Char = " "

    Dim strCaption As String

    strCaption = popt.Caption & Char
    strCaption = Replace(strCaption, "&&", "~")
    strCaption = Replace(strCaption, "&", "")
    strCaption = Replace(strCaption, "~", "&")

    With popt.Parent
        GetOptionRight = popt.Left + Pixels * Screen.TwipsPerPixelX + .TextWidth(strCaption)

    End With

End Function

Public Function GetOptionWidth(popt As OptionButton) As Long

    Const Pixels = 14

    Const Char = "  "

    Dim strCaption As String

    strCaption = popt.Caption & Char
    strCaption = Replace(strCaption, "&&", "~")
    strCaption = Replace(strCaption, "&", "")
    strCaption = Replace(strCaption, "~", "&")

    With popt.Parent
        GetOptionWidth = Pixels * Screen.TwipsPerPixelX + .TextWidth(strCaption)

    End With

End Function

Public Function OptionButtonCaptionLeft(popt As OptionButton) As Long

    Const Pixels = 14

    Const Char = " "

    OptionButtonCaptionLeft = popt.Left + (Pixels * Screen.TwipsPerPixelX) + popt.Parent.TextWidth(Char)

End Function

Public Function StripAcceleratorKeys(pstrCaption As String) As String

    Dim strReturn As String

    strReturn = pstrCaption
    strReturn = Replace(strReturn, "&&", "~")
    strReturn = Replace(strReturn, "&", "")
    strReturn = Replace(strReturn, "~", "&")
    StripAcceleratorKeys = strReturn

End Function
