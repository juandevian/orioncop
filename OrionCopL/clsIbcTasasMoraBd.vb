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
    ''' <exception cref="ErrorTasaMoraGuardadaException">La fila nueva SÍ se guardó pero quedó inconsistente
    ''' (no se cerró la anterior o la relectura no coincide): no reintentar, requiere corrección manual.</exception>
    ''' <remarks>Cualquier otra excepción proviene de las validaciones previas: no se escribió nada.</remarks>
    Friend Sub SRegistreTasa(adtmFechaDesde As Date, adblTasaAnual As Double) Implements IIbcTasasMora.SRegistreTasa
        Dim lobjTasa As New ClsTasaMora(EnuModoInstanciaObjDef.enuNavegable)
        lobjTasa.SRegistreTasaAnual(adtmFechaDesde, adblTasaAnual)
        ' Desde aquí la fila nueva ya está en la BD: toda falla es "guardada pero inconsistente"
        Dim lstrDetalle = String.Empty
        Dim lobjInterna As Exception = Nothing
        Try
            lstrDetalle = FstrInconsistencia(adtmFechaDesde.Date, adblTasaAnual)
        Catch ex As Exception
            lstrDetalle = "no se pudo verificar lo guardado (" & ex.Message & ")"
            lobjInterna = ex
        End Try
        If Not String.IsNullOrEmpty(lstrDetalle) Then
            Throw New ErrorTasaMoraGuardadaException(lobjTasa.EntOrdinalCreado, adtmFechaDesde.Date,
                    lstrDetalle, lobjInterna)
        End If
    End Sub

    ''' <summary>Relee lo guardado; devuelve la inconsistencia encontrada o vacío si todo coincide.</summary>
    Private Shared Function FstrInconsistencia(adtmFechaDesde As Date, adblTasaAnual As Double) As String
        ' FdblTasaMoraFecha(f) busca la fila vigente en f - 1: se consulta con FechaDesde + 1
        Dim ldblLeida = GobjParametros.FdblTasaMoraFecha(adtmFechaDesde.AddDays(1))
        If Not ClsIbcCalculo.FblnTasaCoincide(adblTasaAnual, ldblLeida) Then
            Return "la tasa leída (" & ldblLeida.ToString & ") no coincide con la calculada (" &
                    adblTasaAnual.ToString & ")"
        End If
        ' La fila anterior debe quedar cerrada en FechaDesde - 1. FdtbTasasMora solo reemplaza (para
        ' mostrar) la FechaHasta de la ÚLTIMA fila con hoy; la penúltima trae el valor real de la BD.
        Dim ldtbTasas = GobjParametros.FdtbTasasMora()
        If ldtbTasas.Rows.Count > 1 Then
            Dim ldrwAnterior As DataRow = ldtbTasas.Rows(ldtbTasas.Rows.Count - 2)
            Dim ldtmHastaAnterior As Date = ClsPanorama.FobjValorCampo(
                    ldrwAnterior(ClsFechaHastaTasaMoraDtm.SstrNombreCampoBd), EnuTipoValor.enuDate)
            If ldtmHastaAnterior <> adtmFechaDesde.AddDays(-1) Then
                Return "la tasa anterior no quedó cerrada en " &
                        Format(adtmFechaDesde.AddDays(-1), GCSTRFMTFECHASIMPLE)
            End If
        End If
        Return String.Empty
    End Function
End Class
