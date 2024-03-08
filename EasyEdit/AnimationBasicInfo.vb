Imports Microsoft.Office.Interop.PowerPoint

Public Class AnimationBasicInfo

    Public EffectType As MsoAnimEffect

    Public Sub New(ByVal _EffectType As MsoAnimEffect)
        EffectType = _EffectType
    End Sub

End Class
