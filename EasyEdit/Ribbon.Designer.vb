Partial Class Ribbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.tab_EasyEdit = Me.Factory.CreateRibbonTab
        Me.Group_AnimationTools = Me.Factory.CreateRibbonGroup
        Me.btn_CopyAnimations = Me.Factory.CreateRibbonButton
        Me.btn_PasteAnimations = Me.Factory.CreateRibbonButton
        Me.tab_EasyEdit.SuspendLayout()
        Me.Group_AnimationTools.SuspendLayout()
        Me.SuspendLayout()
        '
        'tab_EasyEdit
        '
        Me.tab_EasyEdit.Groups.Add(Me.Group_AnimationTools)
        Me.tab_EasyEdit.Label = "Easy Edit"
        Me.tab_EasyEdit.Name = "tab_EasyEdit"
        '
        'Group_AnimationTools
        '
        Me.Group_AnimationTools.Items.Add(Me.btn_CopyAnimations)
        Me.Group_AnimationTools.Items.Add(Me.btn_PasteAnimations)
        Me.Group_AnimationTools.Label = "Animation Tools"
        Me.Group_AnimationTools.Name = "Group_AnimationTools"
        '
        'btn_CopyAnimations
        '
        Me.btn_CopyAnimations.Label = "Copy Animations"
        Me.btn_CopyAnimations.Name = "btn_CopyAnimations"
        Me.btn_CopyAnimations.OfficeImageId = "AnimationGallery"
        Me.btn_CopyAnimations.ShowImage = True
        '
        'btn_PasteAnimations
        '
        Me.btn_PasteAnimations.Label = "Paste Animations"
        Me.btn_PasteAnimations.Name = "btn_PasteAnimations"
        Me.btn_PasteAnimations.OfficeImageId = "Paste"
        Me.btn_PasteAnimations.ShowImage = True
        '
        'Ribbon
        '
        Me.Name = "Ribbon"
        Me.RibbonType = "Microsoft.PowerPoint.Presentation"
        Me.Tabs.Add(Me.tab_EasyEdit)
        Me.tab_EasyEdit.ResumeLayout(False)
        Me.tab_EasyEdit.PerformLayout()
        Me.Group_AnimationTools.ResumeLayout(False)
        Me.Group_AnimationTools.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents tab_EasyEdit As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group_AnimationTools As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btn_CopyAnimations As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btn_PasteAnimations As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon() As Ribbon
        Get
            Return Me.GetRibbon(Of Ribbon)()
        End Get
    End Property
End Class
