Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon

    ' Global variable for storing the animations
    Public animations() As AnimationBasicInfo
    Public animationsLength As Integer


    Private Sub Ribbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        animationsLength = -1
    End Sub

    Private Sub btn_CopyAnimations_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_CopyAnimations.Click
        ' Prompt if no object is selected
        If Globals.ThisAddIn.Application.ActiveWindow.Selection.Type Like PpSelectionType.ppSelectionNone Then
            MsgBox("Please select an object first.")
            Exit Sub
        End If

        ' Get selected object
        Dim selectedObject = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange(1)

        ' Clear the animations
        animations = Nothing
        animationsLength = -1

        ' Get the animations
        Dim currentSlideIndex As Integer
        currentSlideIndex = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex
        For Each animation As Effect In Globals.ThisAddIn.Application.ActivePresentation.Slides(currentSlideIndex).TimeLine.MainSequence
            If animation.Shape Is selectedObject Then
                Dim animationBasicInfo As AnimationBasicInfo
                animationBasicInfo = New AnimationBasicInfo(animation.EffectType)

                animationsLength += 1
                ReDim Preserve animations(animationsLength)
                animations(UBound(animations)) = animationBasicInfo
            End If
        Next

        ''''''''''''''''''''''
        '' Update UI
        ''''''''''''''''''''''
        ' Update buttons
        Globals.Ribbons.Ribbon.btn_PasteAnimations.Enabled = animationsLength > -1
        Globals.Ribbons.Ribbon.btn_OverwriteAnimations.Enabled = animationsLength > -1
        ' Update copied animations count label
        Globals.Ribbons.Ribbon.Label_CopiedAnimationsCount.Label = animationsLength + 1

    End Sub

    Private Sub btn_PasteAnimations_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_PasteAnimations.Click
        ' Prompt if no animations are copied
        If animations Is Nothing Then
            MsgBox("Please copy animations first.")
            Exit Sub
        End If

        ' Prompt if no object is selected
        If Globals.ThisAddIn.Application.ActiveWindow.Selection.Type Like PpSelectionType.ppSelectionNone Then
            MsgBox("Please select an object first.")
            Exit Sub
        End If

        ' Get selected object
        Dim selectedObject = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange(1)

        ' Paste the animations
        Dim currentSlideIndex As Integer
        currentSlideIndex = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex
        For Each animation In animations
            Globals.ThisAddIn.Application.ActivePresentation.
                Slides(currentSlideIndex).TimeLine.MainSequence.AddEffect(Shape:=selectedObject,
                                                                          effectId:=animation.EffectType
                                                                          )
            'trigger:=animation.Effect.TriggerShape
            'Level:=animation.Effect.Level,
        Next

    End Sub

    Private Sub btn_OverwriteAnimations_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_OverwriteAnimations.Click
        ' Prompt if no animations are copied
        If animations Is Nothing Then
            MsgBox("Please copy animations first.")
            Exit Sub
        End If

        ' Prompt if no object is selected
        If Globals.ThisAddIn.Application.ActiveWindow.Selection.Type Like PpSelectionType.ppSelectionNone Then
            MsgBox("Please select an object first.")
            Exit Sub
        End If

        ' Get selected object
        Dim selectedObject = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange(1)

        ' Delete selected object's animations
        Dim currentSlideIndex As Integer
        currentSlideIndex = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex

        For _i = Globals.ThisAddIn.Application.ActivePresentation.Slides(currentSlideIndex).TimeLine.MainSequence.Count To 1 Step -1
            If Globals.ThisAddIn.Application.ActivePresentation.Slides(currentSlideIndex).TimeLine.MainSequence(_i).Shape Is selectedObject Then
                Globals.ThisAddIn.Application.ActivePresentation.Slides(currentSlideIndex).TimeLine.MainSequence(_i).Delete()
            End If
        Next

        'Paste the animations
        For Each animation In animations
            Globals.ThisAddIn.Application.ActivePresentation.
                Slides(currentSlideIndex).TimeLine.MainSequence.AddEffect(Shape:=selectedObject,
                                                                          effectId:=animation.EffectType
                                                                          )
            'trigger:=animation.Effect.TriggerShape
            'Level:=animation.Effect.Level,
        Next


    End Sub
End Class
