''' <summary>
''' Orquesta la sincronización del IBC con OriTasasMora antes de causar intereses de mora.
''' </summary>
Friend Class ClsIbcSincroniza
#Region "Definiciones"
    Private Const MCENTCERTIFICADOSACONSULTAR As Integer = 12
    Private Const MCSTRFMTFECHA As String = "yyyy-MM-dd"
    Private ReadOnly MobjProveedor As IIbcProveedor
    Private ReadOnly MobjAlmacen As IIbcAlmacen
    Private ReadOnly MobjTasas As IIbcTasasMora
#End Region

#Region "Constructores"
    Friend Sub New(aobjProveedor As IIbcProveedor, aobjAlmacen As IIbcAlmacen, aobjTasas As IIbcTasasMora)
        MobjProveedor = aobjProveedor
        MobjAlmacen = aobjAlmacen
        MobjTasas = aobjTasas
    End Sub
#End Region

#Region "Procedimientos y funciones"
    ''' <summary>
    ''' Deja OriTasasMora con la tasa que corresponde a la fecha de causación según el modo del centro de
    ''' utilidad. Devuelve True si quedó al día (o si el modo es None); False con mensaje si no se pudo.
    ''' </summary>
    ''' <param name="adtmFecha">Fecha de causación. La causación lee la tasa vigente el día anterior
    ''' (FdblTasaMoraFecha consulta la fecha - 1), por eso el certificado es el vigente en adtmFecha - 1.</param>
    ''' <param name="adtmHoy">Fecha de hoy (inyectable): la FechaDesde de una fila nueva no puede superarla.</param>
    ''' <remarks>Solo oculta las fallas esperadas (API caída, validación previa al guardar, fila guardada pero
    ''' inconsistente); cualquier otra excepción se propaga.</remarks>
    Friend Function FblnSincronice(adtmFecha As Date, aenuModo As EnuModoInteres, adblTasaFijaDeseada As Double,
            adblFactor As Double, ablnForzarConsulta As Boolean, adtmHoy As Date, ByRef astrMens As String) As Boolean
        astrMens = String.Empty
        If aenuModo = EnuModoInteres.None Then Return True
        If aenuModo <> EnuModoInteres.EnuFijo AndAlso aenuModo <> EnuModoInteres.EnuVariable Then
            astrMens = "El modo de interés de mora parametrizado (" & CByte(aenuModo).ToString & ") no es válido."
            Return False
        End If
        If aenuModo = EnuModoInteres.EnuVariable AndAlso Not ClsIbcCalculo.FblnFactorValido(adblFactor) Then
            astrMens = "El factor de interés debe ser mayor que 0 y máximo " &
                    ClsIbcCalculo.CDBLFACTORMAXIMO.ToString & "."
            Return False
        End If
        ' Día cuya tasa usa la causación de adtmFecha
        Dim ldtmFechaTasa = adtmFecha.Date.AddDays(-1)
        Dim lstcCert As StcIbcCertificado = Nothing
        Dim lblnTieneCert = ClsIbcCalculo.FblnCertificadoVigente(MobjAlmacen.FlstCertificados(), ldtmFechaTasa,
                lstcCert)
        If ablnForzarConsulta OrElse Not lblnTieneCert Then
            Try
                Dim llstNuevos = MobjProveedor.FlstConsulteCertificados(MCENTCERTIFICADOSACONSULTAR)
                MobjAlmacen.SGuarde(llstNuevos)
                lblnTieneCert = ClsIbcCalculo.FblnCertificadoVigente(MobjAlmacen.FlstCertificados(),
                        ldtmFechaTasa, lstcCert)
            Catch ex As ErrorIbcApiException
                If Not lblnTieneCert Then
                    astrMens = "No se pudo consultar el IBC y no hay certificado local vigente al " &
                            Format(ldtmFechaTasa, MCSTRFMTFECHA) & ": " & ex.Message
                    Return False
                End If
            End Try
        End If
        If Not lblnTieneCert Then
            astrMens = "No existe un certificado de IBC vigente para el " & Format(ldtmFechaTasa, MCSTRFMTFECHA) &
                    " (tasa que usa la causación de intereses del " & Format(adtmFecha, MCSTRFMTFECHA) & ")."
            Return False
        End If
        Dim ldblTasa As Double
        If aenuModo = EnuModoInteres.EnuFijo Then
            ldblTasa = ClsIbcCalculo.FdblTasaFijaEfectiva(adblTasaFijaDeseada, lstcCert.DblIbc)
        Else
            ldblTasa = ClsIbcCalculo.FdblTasaVariable(lstcCert.DblIbc, adblFactor)
        End If
        If ldblTasa <= 0 Then
            astrMens = "La tasa de interés calculada es cero; revise la parametrización."
            Return False
        End If
        ' FdblTasaVigente(adtmFecha) ya consulta la tasa del día anterior
        If ClsIbcCalculo.FblnTasaCoincide(ldblTasa, MobjTasas.FdblTasaVigente(adtmFecha)) Then Return True
        Dim ldtmDesdeUltima = MobjTasas.FdtmFechaDesdeUltima()
        Dim ldtmDesde = If(lstcCert.DtmFechaDesde > ldtmDesdeUltima, lstcCert.DtmFechaDesde,
                ldtmDesdeUltima.AddDays(1))
        If ldtmDesde > adtmHoy Then
            astrMens = "No se puede registrar una tasa nueva con fecha " & Format(ldtmDesde, MCSTRFMTFECHA) &
                    " porque es posterior a hoy. Ya hay una tasa registrada desde esa fecha o una posterior: si no es correcta, elimínela en Tasas de Mora, o cambie el interés a 'Sin parametrizar' y gestione la tasa manualmente."
            Return False
        End If
        If ldtmDesde > ldtmFechaTasa Then
            astrMens = "La nueva tasa de mora tendría fecha desde " & Format(ldtmDesde, MCSTRFMTFECHA) &
                    " y no regiría el " & Format(ldtmFechaTasa, MCSTRFMTFECHA) &
                    ", día cuya tasa usa la causación de intereses del " & Format(adtmFecha, MCSTRFMTFECHA) &
                    ". Ya hay una tasa registrada desde esa fecha o una posterior: si no es correcta, elimínela en Tasas de Mora, o cambie el interés a 'Sin parametrizar' y gestione la tasa manualmente."
            Return False
        End If
        Try
            MobjTasas.SRegistreTasa(ldtmDesde, ldblTasa)
        Catch ex As ErrorTasaMoraGuardadaException
            ' La fila SÍ quedó guardada: el mensaje ya indica corregir a mano y no reintentar
            astrMens = ex.Message
            Return False
        Catch ex As ErrorInesperadoPanLException
            ' Validación previa al guardar: no se escribió nada
            astrMens = "No se registró la nueva Tasa de Mora (" & Format(ldtmDesde, MCSTRFMTFECHA) & "): " &
                    ex.Message
            Return False
        End Try
        Return True
    End Function
#End Region
End Class
