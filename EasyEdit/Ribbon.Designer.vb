Partial Class Ribbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.tab_EasyEdit = Me.Factory.CreateRibbonTab
        Me.Group_AnimationTools = Me.Factory.CreateRibbonGroup
        Me.btn_CopyAnimations = Me.Factory.CreateRibbonButton
        Me.btn_PasteAnimations = Me.Factory.CreateRibbonButton
        Me.Box1 = Me.Factory.CreateRibbonBox
        Me.Box2 = Me.Factory.CreateRibbonBox
        Me.Box3 = Me.Factory.CreateRibbonBox
        Me.btn_OverwriteAnimations = Me.Factory.CreateRibbonButton
        Me.Group_AnimationClipboardStatus = Me.Factory.CreateRibbonGroup
        Me.Box4 = Me.Factory.CreateRibbonBox
        Me.Label_CopiedAnimations = Me.Factory.CreateRibbonLabel
        Me.Label_CopiedAnimationsCount = Me.Factory.CreateRibbonLabel
        Me.Box7 = Me.Factory.CreateRibbonBox
        Me.tab_EasyEdit.SuspendLayout()
        Me.Group_AnimationTools.SuspendLayout()
        Me.Box1.SuspendLayout()
        Me.Box2.SuspendLayout()
        Me.Group_AnimationClipboardStatus.SuspendLayout()
        Me.Box4.SuspendLayout()
        Me.Box7.SuspendLayout()
        Me.SuspendLayout()
        '
        'tab_EasyEdit
        '
        Me.tab_EasyEdit.Groups.Add(Me.Group_AnimationClipboardStatus)
        Me.tab_EasyEdit.Groups.Add(Me.Group_AnimationTools)
        Me.tab_EasyEdit.Label = "Easy Edit"
        Me.tab_EasyEdit.Name = "tab_EasyEdit"
        '
        'Group_AnimationTools
        '
        Me.Group_AnimationTools.Items.Add(Me.Box1)
        Me.Group_AnimationTools.Label = "Animation Tools"
        Me.Group_AnimationTools.Name = "Group_AnimationTools"
        '
        'btn_CopyAnimations
        '
        Me.btn_CopyAnimations.Label = "Copy Animations"
        Me.btn_CopyAnimations.Name = "btn_CopyAnimations"
        Me.btn_CopyAnimations.OfficeImageId = "AnimationGallery"
        Me.btn_CopyAnimations.ScreenTip = "Copy selected item's animations."
        Me.btn_CopyAnimations.ShowImage = True
        '
        'btn_PasteAnimations
        '
        Me.btn_PasteAnimations.Label = "Paste Animations"
        Me.btn_PasteAnimations.Name = "btn_PasteAnimations"
        Me.btn_PasteAnimations.OfficeImageId = "Paste"
        Me.btn_PasteAnimations.ScreenTip = "Paste copied animations to selected item."
        Me.btn_PasteAnimations.ShowImage = True
        '
        'Box1
        '
        Me.Box1.Items.Add(Me.Box2)
        Me.Box1.Items.Add(Me.Box3)
        Me.Box1.Name = "Box1"
        '
        'Box2
        '
        Me.Box2.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical
        Me.Box2.Items.Add(Me.btn_CopyAnimations)
        Me.Box2.Items.Add(Me.btn_PasteAnimations)
        Me.Box2.Items.Add(Me.btn_OverwriteAnimations)
        Me.Box2.Name = "Box2"
        '
        'Box3
        '
        Me.Box3.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical
        Me.Box3.Name = "Box3"
        '
        'btn_OverwriteAnimations
        '
        Me.btn_OverwriteAnimations.Label = "Overwrite"
        Me.btn_OverwriteAnimations.Name = "btn_OverwriteAnimations"
        Me.btn_OverwriteAnimations.OfficeImageId = "FormatPainter"
        Me.btn_OverwriteAnimations.ScreenTip = "Delete selected item's animations, then paste."
        Me.btn_OverwriteAnimations.ShowImage = True
        '
        'Group_AnimationClipboardStatus
        '
        Me.Group_AnimationClipboardStatus.Items.Add(Me.Box4)
        Me.Group_AnimationClipboardStatus.Label = "Animation Clipboard Status"
        Me.Group_AnimationClipboardStatus.Name = "Group_AnimationClipboardStatus"
        '
        'Box4
        '
        Me.Box4.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical
        Me.Box4.Items.Add(Me.Box7)
        Me.Box4.Name = "Box4"
        '
        'Label_CopiedAnimations
        '
        Me.Label_CopiedAnimations.Label = "Copied animations: "
        Me.Label_CopiedAnimations.Name = "Label_CopiedAnimations"
        '
        'Label_CopiedAnimationsCount
        '
        Me.Label_CopiedAnimationsCount.Label = "0"
        Me.Label_CopiedAnimationsCount.Name = "Label_CopiedAnimationsCount"
        '
        'Box7
        '
        Me.Box7.Items.Add(Me.Label_CopiedAnimations)
        Me.Box7.Items.Add(Me.Label_CopiedAnimationsCount)
        Me.Box7.Name = "Box7"
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
        Me.Box1.ResumeLayout(False)
        Me.Box1.PerformLayout()
        Me.Box2.ResumeLayout(False)
        Me.Box2.PerformLayout()
        Me.Group_AnimationClipboardStatus.ResumeLayout(False)
        Me.Group_AnimationClipboardStatus.PerformLayout()
        Me.Box4.ResumeLayout(False)
        Me.Box4.PerformLayout()
        Me.Box7.ResumeLayout(False)
        Me.Box7.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents tab_EasyEdit As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group_AnimationTools As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btn_CopyAnimations As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btn_PasteAnimations As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Box1 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents Box2 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents Box3 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents btn_OverwriteAnimations As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group_AnimationClipboardStatus As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Box4 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents Label_CopiedAnimations As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Label_CopiedAnimationsCount As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Box7 As Microsoft.Office.Tools.Ribbon.RibbonBox
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon() As Ribbon
        Get
            Return Me.GetRibbon(Of Ribbon)()
        End Get
    End Property
End Class
