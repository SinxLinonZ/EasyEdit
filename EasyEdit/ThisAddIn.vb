Imports System.Windows.Forms
Imports Microsoft.Office.Interop.PowerPoint

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub Application_WindowSelectionChange(Sel As Selection) Handles Application.WindowSelectionChange

        ' Nothing selected
        If Sel.Type Like PpSelectionType.ppSelectionNone Then
            Globals.Ribbons.Ribbon.btn_CopyAnimations.Enabled = False
            Globals.Ribbons.Ribbon.btn_PasteAnimations.Enabled = False
            Globals.Ribbons.Ribbon.btn_OverwriteAnimations.Enabled = False

            ' Slide selected
        ElseIf Sel.Type Like PpSelectionType.ppSelectionSlides Then
            Globals.Ribbons.Ribbon.btn_CopyAnimations.Enabled = False
            Globals.Ribbons.Ribbon.btn_PasteAnimations.Enabled = False
            Globals.Ribbons.Ribbon.btn_OverwriteAnimations.Enabled = False

            ' Text selected
        ElseIf Sel.Type Like PpSelectionType.ppSelectionText Then
            Globals.Ribbons.Ribbon.btn_CopyAnimations.Enabled = True
            Globals.Ribbons.Ribbon.btn_PasteAnimations.Enabled = Globals.Ribbons.Ribbon.animationsLength > -1
            Globals.Ribbons.Ribbon.btn_OverwriteAnimations.Enabled = Globals.Ribbons.Ribbon.animationsLength > -1

            ' Shape selected
        ElseIf Sel.Type Like PpSelectionType.ppSelectionShapes Then
            Globals.Ribbons.Ribbon.btn_CopyAnimations.Enabled = True
            Globals.Ribbons.Ribbon.btn_PasteAnimations.Enabled = Globals.Ribbons.Ribbon.animationsLength > -1
            Globals.Ribbons.Ribbon.btn_OverwriteAnimations.Enabled = Globals.Ribbons.Ribbon.animationsLength > -1

        End If


    End Sub
End Class
