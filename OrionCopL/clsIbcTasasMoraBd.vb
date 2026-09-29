''' <summary>
''' Acceso a las tasas de mora del centro de utilidad (tabla OriTasasMora) para la herramienta IBC.
''' </summary>
Friend Class ClsIbcTasasMoraBd
    Implements IIbcTasasMora

    ''' <summary>Tasa anual vigente a la fecha (FdblTasaMoraFecha consulta la fecha - 1 día).</summary>
    Friend Function FdblTasaVigente(adtmFecha As Date) As Double Implements IIbcTasasMora.FdblTasaVigente
        Return GobjParametros.FdblTasaMoraFecha(adtmFecha)
    End Function

    ''' <summary>FechaDesde de la última fila de OriTasasMora (GCDTMFECHANULA si no hay).</summary>
    Friend Function FdtmFechaDesdeUltima() As Date Implements IIbcTasasMora.FdtmFechaDesdeUltima
        Dim lobjTasa As New ClsTasaMora(EnuModoInstanciaObjDef.enuNavegable)
        Return lobjTasa.FdtmFechaDesdeUltima
    End Function

    ''' <summary>
    ''' Agrega una fila a OriTasasMora con la tasa anual (fracción) desde la fecha indicada y cierra la
    ''' anterior. Luego relee la BD y lanza error si lo guardado no coincide (evita un 0 silencioso).
    ''' </summary>
    Friend Sub SRegistreTasa(adtmFechaDesde As Date, adblTasaAnual As Double) Implements IIbcTasasMora.SRegistreTasa
        Dim lobjTasa As New ClsTasaMora(EnuModoInstanciaObjDef.enuNavegable)
        lobjTasa.SRegistreTasaAnual(adtmFechaDesde, adblTasaAnual)
        SVerifiqueRegistro(adtmFechaDesde.Date, adblTasaAnual)
    End Sub

    Private Shared Sub SVerifiqueRegistro(adtmFechaDesde As Date, adblTasaAnual As Double)
        ' FdblTasaMoraFecha(f) busca la fila vigente en f - 1: se consulta con FechaDesde + 1
        Dim ldblLeida = GobjParametros.FdblTasaMoraFecha(adtmFechaDesde.AddDays(1))
        If Not ClsIbcCalculo.FblnTasaCoincide(adblTasaAnual, ldblLeida) Then
            Throw New ErrorInesperadoPanLException("La tasa de mora guardada (" & ldblLeida.ToString &
                    ") no coincide con la calculada (" & adblTasaAnual.ToString & ")!")
        End If
        ' La fila anterior debe quedar cerrada en FechaDesde - 1. FdtbTasasMora solo reemplaza (para
        ' mostrar) la FechaHasta de la ÚLTIMA fila con hoy; la penúltima trae el valor real de la BD.
        Dim ldtbTasas = GobjParametros.FdtbTasasMora()
        If ldtbTasas.Rows.Count > 1 Then
            Dim ldrwAnterior As DataRow = ldtbTasas.Rows(ldtbTasas.Rows.Count - 2)
            Dim ldtmHastaAnterior As Date = ClsPanorama.FobjValorCampo(
                    ldrwAnterior(ClsFechaHastaTasaMoraDtm.SstrNombreCampoBd), EnuTipoValor.enuDate)
            If ldtmHastaAnterior <> adtmFechaDesde.AddDays(-1) Then
                Throw New ErrorInesperadoPanLException("La tasa de mora anterior no quedó cerrada en " &
                        Format(adtmFechaDesde.AddDays(-1), GCSTRFMTFECHASIMPLE) & "!")
            End If
        End If
    End Sub
End Class
