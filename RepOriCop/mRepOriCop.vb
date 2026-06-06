Module mRepOriCop
    Friend Sub SPuebleBarraEstado(acolLabels As Collection)
        acolLabels(1).Content = "Carpeta: " & clsOrionCop.strNombreCarpetaActual
        acolLabels(1).ToolTip = acolLabels(1).Content
        acolLabels(2).Content = "Copropiedad: " & ClsOrionCop.StrNombreCentroUtilActual
        acolLabels(2).ToolTip = acolLabels(2).Content
        acolLabels(3).Content = "Usuario Actual: " & gstrIdUsuario
        acolLabels(3).ToolTip = acolLabels(3).Content
        acolLabels(4).Content = "Periodo Actual: "
        If Not IsNothing(GobjParametros.objAnoActual) Then
            acolLabels(4).Content &= GobjParametros.objAnoActual.strNombrePeriodoActual
        End If
        acolLabels(4).ToolTip = acolLabels(4).Content
    End Sub
End Module
