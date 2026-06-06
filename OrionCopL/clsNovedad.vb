Friend Class ClsNovedad
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriNovedades"
    Private MobjMiFactura As ClsFactura = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As clsCBObjetoPan, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        henuTipoObjeto = enuModoInstanciaObjDef.enuDeColeccion
        HblnEsSuprimible = False
        hblnEsAnulable = False
        '
        drwRegistroActual = adrwObjeto
        DtbTablaColeccion = drwRegistroActual.Table
        henuTipoPermiso = EnuPermisosDef.enuHeredado
    End Sub
#End Region
#Region "Propiedades"
#Region "Propiedades indentificadoras"
    Protected Overrides ReadOnly Property HstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Protected Overrides ReadOnly Property HenuIdClase As EnuIdClasesPanDef
        Get
            Return EnuIdClasesPanDef.enuNovedad
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Novedad"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjAliasCont_NovStr As New ClsAliasCont_NovStr(Me)
    Friend ReadOnly Property ObjBaseDec As New ClsBaseDec(Me)
    Friend ReadOnly Property ObjEsPrefactura_NovBln As New ClsEsPreFacturaBln(Me)
    Friend ReadOnly Property ObjFactorDbl As New ClsFactorDbl(Me)
    Friend ReadOnly Property ObjFechaNovedadDtm As New ClsFechaNovedadDtm(Me)
    Friend ReadOnly Property ObjIdAno_NovShr As New ClsIdAno_NovShr(Me)
    Friend ReadOnly Property ObjIdCarpeta_NovShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_NovShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdCuentaCr_NovStr As New ClsIdCuentaCr_NovStr(Me)
    Friend ReadOnly Property ObjIdCuentaDb_NovStr As New ClsIdCuentaDb_NovStr(Me)
    Friend ReadOnly Property ObjIdDocOrigenEnt As New ClsIdDocOrigenEnt(Me)
    Friend ReadOnly Property ObjIdFactura_NovEnt As New ClsIdFactura_NovEnt(Me)
    Friend ReadOnly Property ObjIdItemFact_NovShr As New ClsIdItemFacturaShr(Me)
    Friend ReadOnly Property ObjIdItemDocOrigen_NovShr As New ClsIdItemDocOrigen_NovShr(Me)
    Friend ReadOnly Property ObjIdNovedadShr As New ClsIdNovedadShr(Me)
    Friend ReadOnly Property ObjIdPredioAgrupador_NovStr As New ClsIdPredioAgrupador_NovStr(Me)
    Friend ReadOnly Property ObjIdServicio_NovShr As New ClsIdServicio_NovShr(Me)
    Friend ReadOnly Property ObjIdTercero_NovDbl As New ClsIdTercero_NovDbl(Me)
    Friend ReadOnly Property ObjIdTerceroCtaCr_NovDbl As New ClsIdTerceroCtaCr_NovDbl(Me)
    Friend ReadOnly Property ObjIdTipoDocOrigenByt As New ClsIdTipoDocOrigenByt(Me)
    Friend ReadOnly Property ObjIdTipoNovedadByt As New ClsIdTipoNovedadByt(Me)
    Friend ReadOnly Property ObjPrefijoDocOrigen_NovStr As New ClsPrefijoDocOrigen_NovStr(Me)
    Friend ReadOnly Property ObjPrefijoFact_NovStr As New ClsPrefijoFact_NovStr(Me)
    Friend ReadOnly Property ObjValor_NovDec As New ClsValor_NovDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAnuladoBln)
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjAliasCont_NovStr)
                HcolPropiedades.Add(ObjBaseDec)
                HcolPropiedades.Add(ObjEsPrefactura_NovBln)
                HcolPropiedades.Add(ObjFactorDbl)
                HcolPropiedades.Add(ObjFechaNovedadDtm)
                HcolPropiedades.Add(ObjIdAno_NovShr)
                HcolPropiedades.Add(ObjIdCarpeta_NovShr)
                HcolPropiedades.Add(ObjIdCentroUtil_NovShr)
                HcolPropiedades.Add(ObjIdCuentaCr_NovStr)
                HcolPropiedades.Add(ObjIdCuentaDb_NovStr)
                HcolPropiedades.Add(ObjIdDocOrigenEnt)
                HcolPropiedades.Add(ObjIdFactura_NovEnt)
                HcolPropiedades.Add(ObjIdItemFact_NovShr)
                HcolPropiedades.Add(ObjIdItemDocOrigen_NovShr)
                HcolPropiedades.Add(ObjIdNovedadShr)
                HcolPropiedades.Add(ObjIdPredioAgrupador_NovStr)
                HcolPropiedades.Add(ObjIdServicio_NovShr)
                HcolPropiedades.Add(ObjIdTercero_NovDbl)
                HcolPropiedades.Add(ObjIdTerceroCtaCr_NovDbl)
                HcolPropiedades.Add(ObjIdTipoDocOrigenByt)
                HcolPropiedades.Add(ObjIdTipoNovedadByt)
                HcolPropiedades.Add(ObjPrefijoDocOrigen_NovStr)
                HcolPropiedades.Add(ObjPrefijoFact_NovStr)
                HcolPropiedades.Add(ObjValor_NovDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend ReadOnly Property ObjMIFactura As ClsFactura
        Get
            If MobjMiFactura Is Nothing Then
                Dim lstrPref As String = ObjPrefijoFact_NovStr.ObjValorPro
                Dim lentIdFact As Integer = ObjIdFactura_NovEnt.ObjValorPro
                Dim lobjValorLlave As Object = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFact}
                MobjMiFactura = New ClsFactura()
                MobjMiFactura.SAbra(lobjValorLlave)
            End If
            Return MobjMiFactura
        End Get
    End Property

#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        If enuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            sNumereObj()
            objFechaCreacionDtm.objValorPro = Date.Now
        ElseIf enuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
            ObjValor_NovDec.SValide()
        End If
        MyBase.sActualice(ablnExigeRequeridos)
    End Sub
    Protected Overrides Function SAnuleEnObj() As Boolean
        Dim lblnAnulo = FblnEsAnulable()
        If lblnAnulo Then
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                ObjAnuladoBln.ObjValorPro = True
                ObjValor_NovDec.ObjValorPro = 0
                SActualice(True)
            Else
                Throw New ErrorInesperadoPanLException("Anulando. Estado inesperado del objeto Novedad")
            End If
        End If
        Return True
    End Function
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdNovedadShr.ToString
        End Get
    End Property
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        MobjMiFactura = Nothing
    End Sub
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If enuEstadoActualizacion = enuEstadoObjetoDef.enuCreando Then
            Dim lshrIdNovedad As Short
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND (" &
                    ClsPrefijoFact_NovStr.SstrNombreCampoBd & " = '" & ObjPrefijoFact_NovStr.ObjValorPro &
                    "' OR " & ClsPrefijoFact_NovStr.SstrNombreCampoBd & " = '***')" &
                    " AND " & ClsIdFactura_NovEnt.SstrNombreCampoBd & " = " &
                    ObjIdFactura_NovEnt.ObjValorPro
            lshrIdNovedad = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                    ClsIdNovedadShr.SstrNombreCampoBd, ObjIdNovedadShr.EnuTipoValor,
                    lstrFiltro) + 1
            ObjIdNovedadShr.ObjValorPro = lshrIdNovedad
        End If
    End Sub
    Friend Sub SModifiqueADefinitiva(astrPrefijo As String)
        EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        ObjEsPrefactura_NovBln.ObjValorPro = False
        ObjPrefijoFact_NovStr.ObjValorPro = astrPrefijo
        ObjPrefijoDocOrigen_NovStr.ObjValorPro = astrPrefijo
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsAliasCont_NovStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "AliasCont"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Alias Contable"
        HshrLongitud = 50
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1,
                ShrLongitud, BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsBaseDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Base"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Base"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Decimal.MaxValue,
                BlnEsRequerido, EnuTipoValor.enuDecimal)
        If HblnEsValido Then
            Dim lobjPadre As ClsNovedad = ObjPadre
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsFactorDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Factor"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Factor"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, 1, BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            Dim lobjPadre As ClsNovedad = ObjPadre
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsFechaNovedadDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaNovedad"
    Private ReadOnly MobjPadre As ClsNovedad = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaNovedad"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin = DateSerial(1990, 1, 1)
        Dim ldtmFechaMax = Date.Now
        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
        If HblnEsValido Then
            Dim lobjPadre As ClsNovedad = ObjPadre
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HobjValorNew = DateSerial(HobjValorNew.Year, HobjValorNew.Month, HobjValorNew.Day)
                If Not ClsOrionCop.BlnProcesoEspecial Then
                    Dim lobjFac As ClsFactura = MobjPadre.ObjMIFactura
                    Dim lenuModoCM As EnuModoCausaMora = lobjFac.FenuModoCausaMora
                    Dim ldtmFecIniPer =
                            GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
                    ldtmFechaMin = If(lenuModoCM = EnuModoCausaMora.EnuUltimoDia AndAlso
                            GblnCausandoFM, ldtmFecIniPer.AddDays(-1), ldtmFecIniPer)
                    ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                    If GblnCausandoFM Then
                        ldtmFechaMax = ldtmFechaMax.AddDays(1)
                    End If
                End If
                HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin,
                        ldtmFechaMax, BlnEsRequerido)
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, GCSTRFMTFECHASIMPLE)
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class
Friend Class ClsIdAno_NovShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAno"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Id Año Novedad"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Year(Date.MaxValue),
                BlnEsRequerido, EnuTipoValor)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdCuentaCr_NovStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCuentaCr"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdCuentaCr_Nov"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            Dim lobjPadre As ClsNovedad = ObjPadre
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
            Else
                If Not ClsOrionCop.BlnProcesoEspecial Then
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
                End If
            End If
        End If
    End Sub
    Friend ReadOnly Property StrNombreCuenta As String
        Get
            Dim lstrNombreCuenta = String.Empty
            If HblnEsValido Then
                lstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            End If
            Return lstrNombreCuenta
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdCuentaDb_NovStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCuentaDb"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdCuentaDb_Nov"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            Dim lobjPadre As ClsNovedad = ObjPadre
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
            Else
                If Not ClsOrionCop.BlnProcesoEspecial Then
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
                End If
            End If
        End If
    End Sub
    Friend ReadOnly Property StrNombreCuenta As String
        Get
            Dim lstrNombreCuenta = String.Empty
            If HblnEsValido Then
                lstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorPro.ToString)
            End If
            Return lstrNombreCuenta
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdDocOrigenEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdDocOrigen"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdDocumentoOrigen"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido,
                EnuTipoValor.enuInteger)
        If HblnEsValido Then
            Dim lobjPadre As ClsNovedad = ObjPadre
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando AndAlso
                    Not ClsOrionCop.BlnFacturando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdFactura_NovEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdFactura"
    Private ReadOnly MobjPadre As ClsNovedad = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdFactura"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando AndAlso
                Not ClsOrionCop.BlnFacturando Then
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido, EnuTipoValor)
            If Not (BlnLeyendoOrigen OrElse ClsOrionCop.BlnFacturando) Then
                If HblnEsValido Then
                    Dim lstrPrefijo = MobjPadre.ObjPrefijoFact_NovStr.ObjValorPro
                    Dim lobjLlavePrincipal() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefijo, HobjValorNew,
                            MobjPadre.ObjIdItemFact_NovShr.ObjValorPro,
                            MobjPadre.ObjIdNovedadShr.ObjValorPro}
                    HblnEsValido = MobjPadre.FblnExisteLlave(lobjLlavePrincipal)
                End If
            End If
        Else
            HblnEsValido = True
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
    Private Sub ClsIdFactura_NovEnt_EvnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso HblnEsValido Then
            MobjPadre.ObjIdPredioAgrupador_NovStr.SValide()
        End If
    End Sub
End Class
Friend Class ClsIdItemDocOrigen_NovShr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsNovedad = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdItemDocOrigen"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdItemDocOrigen"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lentVlrMin = 1
        If MobjPadre.ObjIdTipoDocOrigenByt.ObjValorPro = EnuTipoDocOri.EnuNotaRevCr Then
            lentVlrMin = 0
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, lentVlrMin, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdNovedadShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdNovedad"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdNovedad"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 5
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdPredioAgrupador_NovStr
    'Herencia
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredioAgrupador"
    Private ReadOnly MobjPadre As ClsNovedad = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdPredioAgrupador_Factura"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud,
                BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                If Not GblnActualizandoApp Then
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
                End If
            Else
                HblnEsValido = False
                If MobjPadre.ObjPrefijoFact_NovStr.BlnEsValido AndAlso
                        MobjPadre.ObjIdFactura_NovEnt.BlnEsValido Then
                    Dim lstrPre = MobjPadre.ObjPrefijoFact_NovStr.ObjValorPro
                    Dim lentIdFac = MobjPadre.ObjIdFactura_NovEnt.ObjValorPro
                    Dim lobjFac = New ClsFactura()
                    Dim lobjVlrLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPre, lentIdFac}
                    lobjFac.SAbra(lobjVlrLlave)
                    If lobjFac.BlnExiste Then
                        HblnEsValido = (lobjFac.ObjIdPredioAgrupador_FacStr.ObjValorPro = HobjValorNew)
                        If Not HblnEsValido Then
                            Throw New ErrorInesperadoPanLException("Inconsistencia en Predio Agrupador")
                        End If
                    Else
                        If TypeOf (MobjPadre.ObjPadre) Is ClsItemFactura Then
                            Dim lobjItemFac As ClsItemFactura = MobjPadre.ObjPadre
                            HblnEsValido = (lobjItemFac.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando)
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsIdServicio_NovShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdServicio"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdServicio"
        HenuTipoValor = EnuTipoValor.enuUShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, 999,
                BlnEsRequerido, EnuTipoValor)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdTercero_NovDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCliente"
    Private ReadOnly MobjPadre As ClsNovedad = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTerceroCliente"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC,
                GCDBLMAXTERC, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                Dim lobjLlaveCliente() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
                lobjCliente.SAbra(lobjLlaveCliente)
                HblnEsValido = lobjCliente.BlnExiste
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
''' <summary>
''' Id. del Tercero para la cuenta Cr del Servicio. 
''' </summary>
''' <remarks>Normalmente es el Id. del cliente pero cuando es 
''' llevado a una cuenta de dinero recibido para terceros (Pasivo) este tercero es un proveedor</remarks>
Friend Class ClsIdTerceroCtaCr_NovDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCuentaCr"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdTerceroCuentaCr"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC,
                GCDBLMAXTERC, BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdTipoDocOrigenByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoDocOrigen"
    Private ReadOnly MobjPadre As ClsNovedad = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTipoDocOrigen"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoDocOri.EnuFactura,
                EnuTipoDocOri.EnuNotaRevCr, BlnEsRequerido)
        If HblnEsValido Then
            Dim lobjPadre As ClsNovedad = ObjPadre
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Private Sub ClsIdTipoDocOrigenByt_evnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        MobjPadre.ObjIdItemDocOrigen_NovShr.SValide()
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdTipoNovedadByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoNovedad"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdTipoNovedad"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoNov.EnuDbCap,
                EnuTipoNov.EnuRDbIvaInt, BlnEsRequerido)
        If Not HblnEsValido Then
            HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoNov.EnuRDbCap,
                    EnuTipoNov.EnuRCrRetCre, BlnEsRequerido)
        End If
        If Not HblnEsValido Then
            HblnEsValido = HobjValorNew = EnuTipoNov.EnuCrIvaGas
        End If
        If HblnEsValido Then
            Dim lobjPadre As ClsNovedad = ObjPadre
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsPrefijoFact_NovStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoFactura"
    Private ReadOnly MobjPadre As ClsNovedad = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "PrefijoFactura_Nov"
        HshrLongitud = 5
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then
            HobjValorNew = String.Empty
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
    Private Sub ClsPrefijoFact_NovStr_EvnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lstrPreFac = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuFacturaVenta)
            If lstrPreFac = HobjValorNew Then
                MobjPadre.ObjIdPredioAgrupador_NovStr.SValide()
            End If
        End If
    End Sub
End Class
Friend Class ClsPrefijoDocOrigen_NovStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoDocOrigen"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PrefijoDocOrigen"
        HshrLongitud = 5
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then
            HobjValorNew = String.Empty
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsValor_NovDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"
    Private ReadOnly MobjPadre As ClsNovedad = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Valor_Novedad"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0.01, Decimal.MaxValue,
                BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                If ClsOrionCop.BlnProcesoEspecial Then
                    HblnEsValido = (HobjValorNew >= 0)
                Else
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
                End If
            End If
        Else
            If HobjValorNew = 0 Then
                HblnEsValido = MobjPadre.ObjAnuladoBln.ObjValorPro
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
#End Region