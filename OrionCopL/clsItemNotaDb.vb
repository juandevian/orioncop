Friend Class ClsItemNotaDb
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriItemsNotaDb"
    ' Variables de modulo
    Private ReadOnly MobjPadre As ClsNotaDb = Nothing
    Private mobjFactura As ClsFactura = Nothing
    Private MobjItemFac As ClsItemFactura = Nothing
    Private McolNovedades As Collection = Nothing
    Private MdtbNovedades As DataTable = Nothing
    '
    Private MdecValorIva As Decimal = 0D
    Private MdecValorAntesIva As Decimal = 0D
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As clsNotaDb, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        MobjPadre = aobjPadre
        henuTipoObjeto = enuModoInstanciaObjDef.enuDeColeccion
        HblnEsSuprimible = False
        '
        drwRegistroActual = adrwObjeto
        DtbTablaColeccion = DrwRegistroActual.Table
        HenuTipoPermiso = EnuPermisosDef.enuHeredado
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
            Return EnuIdClasesPanDef.enuItemNotaDb
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Item Nota Db."
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjBaseMoraDec As New ClsBaseMoraDec(Me)
    Friend ReadOnly Property ObjDiasMoraEnt As New ClsDiasMoraEnt(Me)
    Friend ReadOnly Property ObjFechaCausoMora_Dtm As New ClsFechaCausoMora_Dtm(Me)
    Friend ReadOnly Property ObjIdCarpeta_ItemNotaDbShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_ItemNotaDbShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdFactura_ItemNotaDbEnt As New ClsIdFactura_ItemNotaDbEnt(Me)
    Friend ReadOnly Property ObjIdItemFac_ItemNotaDbShr As New ClsIdItemFac_ItemNotaDbShr(Me)
    Friend ReadOnly Property ObjIdItemNotaDbShr As New ClsIdItemNotaDbShr(Me)
    Friend ReadOnly Property ObjIdNotaDb_ItemNotaDbEnt As New ClsIdNotaDb_ItemNotaDbEnt(Me)
    Friend ReadOnly Property ObjPrefijo_ItemNotaDbStr As New ClsPrefijo_NotaDbStr(Me)
    Friend ReadOnly Property ObjPrefijoFact_ItemNotaDbStr As New ClsPrefijoFact_ItemNotaDbStr(Me)
    Friend ReadOnly Property ObjTarifaIva_ItemNotaDbDbl As New ClsTarifaIva_ItemNotaDbDbl(Me)
    Friend ReadOnly Property ObjTasaMora_ItemNotaDbDbl As New ClsTasaMora_ItemNotaDbDbl(Me)
    Friend ReadOnly Property ObjValor_ItemNotaDbDec As New ClsValor_ItemNotaDbDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAnuladoBln)
                HcolPropiedades.Add(ObjBaseMoraDec)
                HcolPropiedades.Add(ObjDiasMoraEnt)
                HcolPropiedades.Add(ObjFechaCausoMora_Dtm)
                HcolPropiedades.Add(ObjIdCarpeta_ItemNotaDbShr)
                HcolPropiedades.Add(ObjIdCentroUtil_ItemNotaDbShr)
                HcolPropiedades.Add(ObjIdFactura_ItemNotaDbEnt)
                HcolPropiedades.Add(ObjIdItemFac_ItemNotaDbShr)
                HcolPropiedades.Add(ObjIdItemNotaDbShr)
                HcolPropiedades.Add(ObjIdNotaDb_ItemNotaDbEnt)
                HcolPropiedades.Add(ObjPrefijo_ItemNotaDbStr)
                HcolPropiedades.Add(ObjPrefijoFact_ItemNotaDbStr)
                HcolPropiedades.Add(ObjTarifaIva_ItemNotaDbDbl)
                HcolPropiedades.Add(ObjTasaMora_ItemNotaDbDbl)
                HcolPropiedades.Add(ObjValor_ItemNotaDbDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend ReadOnly Property ObjMiPadre As ClsNotaDb
        Get
            Return MobjPadre
        End Get
    End Property
    Friend ReadOnly Property ObjFactura As ClsFactura
        Get
            If ObjPrefijoFact_ItemNotaDbStr.BlnEsValido AndAlso ObjIdFactura_ItemNotaDbEnt.BlnEsValido Then
                If IsNothing(mobjFactura) Then
                    Dim lstrPrefFac As String = ObjPrefijoFact_ItemNotaDbStr.ObjValorPro
                    Dim lentIdFac As Integer = ObjIdFactura_ItemNotaDbEnt.ObjValorPro
                    Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFac, lentIdFac}
                    mobjFactura = New ClsFactura()
                    mobjFactura.SAbra(lobjValorLlave)
                End If
            Else
                mobjFactura = Nothing
            End If
            Return mobjFactura
        End Get
    End Property
    Friend ReadOnly Property DecValorAntesIva As Decimal
        Get
            Dim ldecValorAntes As Decimal
            Dim ldecValor As Decimal = ObjValor_ItemNotaDbDec.ObjValorPro
            Dim ldblTasaIva As Double = ObjTarifaIva_ItemNotaDbDbl.ObjValorPro
            If ldblTasaIva > 0 Then
                ldecValorAntes = ldecValor / (1 + ldblTasaIva)
                MdecValorAntesIva = Math.Round(ldecValorAntes, 0)
            Else
                MdecValorAntesIva = ldecValor
            End If
            MdecValorIva = ldecValor - MdecValorAntesIva
            Return MdecValorAntesIva
        End Get
    End Property
    Friend ReadOnly Property DecValorBaseIva As Decimal
        Get
            Dim ldecVlrBaseIva = 0D
            If DecValorIva > 0 Then
                ldecVlrBaseIva = DecValorAntesIva
            End If
            Return ldecVlrBaseIva
        End Get
    End Property
    Friend ReadOnly Property DecValorIva As Decimal
        Get
            MdecValorAntesIva = DecValorAntesIva
            Return MdecValorIva
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.sVacie()
        mdtbNovedades = Nothing
        McolNovedades = Nothing
        MobjItemFac = Nothing
        mobjFactura = Nothing
    End Sub
    Public Overrides Function FblnEsAnulable() As Boolean
        Return MobjPadre.FblnEsAnulable
    End Function
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        GobjPanDat.SControleProcesoObj(True)
        Try
            ClsPanorama.SActualiceCol(McolNovedades)
            MyBase.SActualice(ablnExigeRequeridos)
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            GobjPanDat.SControleProcesoObj(False)
        End Try
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Dim lstrPrefijo As String = ObjPrefijo_ItemNotaDbStr.ObjValorPro
            Dim lstrIdObjeto = String.Empty
            If Not String.IsNullOrEmpty(lstrPrefijo) Then
                lstrIdObjeto = lstrPrefijo & "-"
            End If
            lstrIdObjeto &= ObjIdItemNotaDbShr.ToString
            Return lstrIdObjeto
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    '
#End Region
#Region "Manejo Novedad"
    Friend ReadOnly Property ColNovedades
        Get
            If IsNothing(McolNovedades) Then
                If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    Dim lstrFiltro As String = ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" &
                            ObjPrefijo_ItemNotaDbStr.ObjValorPro & "' AND " &
                            ClsIdDocOrigenEnt.SstrNombreCampoBd &
                            " = " & ObjIdNotaDb_ItemNotaDbEnt.ObjValorPro & " AND " &
                            ClsIdItemDocOrigen_NovShr.SstrNombreCampoBd & " = " &
                            ObjIdItemNotaDbShr.ObjValorPro
                    SCargueDtbNovedades()
                    Dim ldrwNovedades = MdtbNovedades.Select(lstrFiltro)
                    For Each ldrwNov As DataRow In MdtbNovedades.Rows
                        Dim lobjNovedad = New ClsNovedad(Me, ldrwNov)
                        lobjNovedad.SLeaValores(True)
                        McolNovedades.Add(lobjNovedad, lobjNovedad.ObjIdNovedadShr.ToString())
                    Next
                End If
            End If
            Return McolNovedades
        End Get
    End Property
    Friend Sub SGenereNovedades(astrPrefijoNotaDb As String, aentIdNotaDb As Integer,
            astrIdPredioAgr As String)
        Dim ldecValorAntesIva = 0D, ldecValorIva = 0D
        Dim ldecValorBase As Decimal = ObjBaseMoraDec.ObjValorPro
        Dim lstrCodigoCtaIvaGen = String.Empty
        Dim lstrCodigoCtaMoraCr = FstrIdCtaMoraCr(lstrCodigoCtaIvaGen)
        McolNovedades = New Collection
        If ObjTarifaIva_ItemNotaDbDbl.ObjValorPro > 0 Then
            ldecValorIva = ObjValor_ItemNotaDbDec.ObjValorPro -
                    Math.Round(ObjValor_ItemNotaDbDec.ObjValorPro /
                    (1 + ObjTarifaIva_ItemNotaDbDbl.ObjValorPro))
            ldecValorBase = Math.Round(ObjBaseMoraDec.ObjValorPro /
                    (1 + ObjTarifaIva_ItemNotaDbDbl.ObjValorPro))
        End If
        ldecValorAntesIva = ObjValor_ItemNotaDbDec.ObjValorPro - ldecValorIva
        ' Novedad: Débito=CxC Intereses Mora, Crédito=Ing Intereses Mora 
        Dim lobjNovedad = FobjNuevaNovedad(astrIdPredioAgr)
        Dim ldtmFechaNovedad As Date = MobjPadre.ObjFecha_NotaDbDtm.ObjValorPro
        With lobjNovedad
            .ObjBaseDec.ObjValorPro = ldecValorBase
            .ObjFactorDbl.ObjValorPro = ObjTasaMora_ItemNotaDbDbl.ObjValorPro
            .ObjIdCuentaDb_NovStr.ObjValorPro = GobjParametros.ObjIdCtaIntMoraDbStr.ObjValorPro
            .ObjIdCuentaCr_NovStr.ObjValorPro = lstrCodigoCtaMoraCr
            .ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuDbInt
            .ObjValor_NovDec.ObjValorPro = ldecValorAntesIva
            .ObjIdDocOrigenEnt.ObjValorPro = aentIdNotaDb
            .ObjPrefijoDocOrigen_NovStr.ObjValorPro = astrPrefijoNotaDb
            .ObjIdPredioAgrupador_NovStr.ObjValorPro = astrIdPredioAgr
            .ObjIdAno_NovShr.ObjValorPro = 0
            .ObjIdServicio_NovShr.ObjValorPro = GCSHRIDMORA
            .ObjFechaNovedadDtm.ObjValorPro = ldtmFechaNovedad
        End With
        McolNovedades.Add(lobjNovedad)
        ' Novedad: Débito=CxC Intereses Mora, Crédito=Iva Generado 
        If ldecValorIva > 0 Then
            lobjNovedad = FobjNuevaNovedad(astrIdPredioAgr)
            With lobjNovedad
                .ObjBaseDec.ObjValorPro = ldecValorAntesIva
                .ObjFactorDbl.ObjValorPro = ObjTarifaIva_ItemNotaDbDbl.ObjValorPro
                .ObjIdCuentaDb_NovStr.ObjValorPro = GobjParametros.ObjIdCtaIntMoraDbStr.ObjValorPro
                .ObjIdCuentaCr_NovStr.ObjValorPro = lstrCodigoCtaIvaGen
                .ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuDbIvaInt
                .ObjValor_NovDec.ObjValorPro = ldecValorIva
                .ObjIdDocOrigenEnt.ObjValorPro = aentIdNotaDb
                .ObjPrefijoDocOrigen_NovStr.ObjValorPro = astrPrefijoNotaDb
                .ObjIdPredioAgrupador_NovStr.ObjValorPro = astrIdPredioAgr
                .ObjIdAno_NovShr.ObjValorPro = 0
                .ObjIdServicio_NovShr.ObjValorPro = GCSHRIDMORA
                .ObjFechaNovedadDtm.ObjValorPro = ldtmFechaNovedad
            End With
            McolNovedades.Add(lobjNovedad)
        End If
    End Sub
    Private Function FobjNuevaNovedad(astrIdPreAgr As String) As ClsNovedad
        Dim lobjNovedad As ClsNovedad = Nothing
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lstrIdPreAgr As String = ObjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
            SCargueDtbNovedades()
            If Not IsNothing(MdtbNovedades) Then
                Dim ldrwNovedad As DataRow = MdtbNovedades.NewRow
                lobjNovedad = New ClsNovedad(Me, ldrwNovedad) With {
                    .EnuPermisosObj = EnuPermisosDef.enuCrear
                }
                With lobjNovedad
                    .SCreeObj(Nothing)
                    .ObjAnuladoBln.ObjValorPro = False
                    .ObjEsPrefactura_NovBln.ObjValorPro = False
                    .ObjFechaCreacionDtm.ObjValorPro = Date.Now
                    .ObjIdCarpeta_NovShr.ObjValorPro = GshrIdCarpeta
                    .ObjIdCentroUtil_NovShr.ObjValorPro = GshrIdCentroUtil
                    .ObjIdPredioAgrupador_NovStr.ObjValorPro = astrIdPreAgr
                    .ObjPrefijoFact_NovStr.ObjValorPro = ObjPrefijoFact_ItemNotaDbStr.ObjValorPro
                    .ObjIdFactura_NovEnt.ObjValorPro = ObjIdFactura_ItemNotaDbEnt.ObjValorPro
                    .ObjIdItemFact_NovShr.ObjValorPro = ObjIdItemFac_ItemNotaDbShr.ObjValorPro
                    .ObjIdItemDocOrigen_NovShr.ObjValorPro = ObjIdItemNotaDbShr.ObjValorPro
                    .ObjIdTercero_NovDbl.ObjValorPro = MobjPadre.ObjIdCliente_NotaDbDbl.ObjValorPro
                    .ObjAliasCont_NovStr.ObjValorPro = MobjPadre.FstrAliasCon()
                    .ObjIdTipoDocOrigenByt.ObjValorPro = EnuTipoDocOri.EnuNotaDb
                    .ObjPrefijoDocOrigen_NovStr.ObjValorPro = ObjPrefijo_ItemNotaDbStr.ObjValorPro
                    .ObjIdTerceroCtaCr_NovDbl.ObjValorPro = 0
                End With
            End If
        End If
        Return lobjNovedad
    End Function
    Private Function FstrIdCtaMoraCr(ByRef astrCodigoCtaIvaGen As String) As String
        Dim lobjFactura As New ClsFactura()
        Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil,
                ObjPrefijoFact_ItemNotaDbStr.ObjValorPro, ObjIdFactura_ItemNotaDbEnt.ObjValorPro}
        lobjFactura.SAbra(lobjValorLlave)
        Dim lobjItemFac As ClsItemFactura = lobjFactura.ColItemsFactura(ObjIdItemFac_ItemNotaDbShr.ToString)
        Dim lstrIdCtaMoraCr As String = lobjItemFac.FstrIdCtaMora
        Dim lstrIdCtaIvaGen As String = lobjItemFac.ObjServicio.ObjCodigoCuentaIvaStr.ObjValorPro
        astrCodigoCtaIvaGen = lstrIdCtaIvaGen
        Return lstrIdCtaMoraCr
    End Function
    Private Sub SCargueDtbNovedades()
        If IsNothing(MdtbNovedades) Then
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsPrefijoFact_NovStr.SstrNombreCampoBd &
            " = '" & ObjPrefijoFact_ItemNotaDbStr.ObjValorPro & "' AND " & ClsIdFactura_NovEnt.SstrNombreCampoBd &
            " = " & ObjIdFactura_ItemNotaDbEnt.ObjValorPro
            Dim lstrIndice(,) As String = {{ClsIdNovedadShr.SstrNombreCampoBd, "ASC"}}
            MdtbNovedades = ClsPanorama.FdtbDataTable(ClsNovedad.SstrNombreTabla, {"*"}, lstrIndice, lstrFiltro)
        End If
    End Sub
#End Region
#Region "Datos EFactura"
    Friend ReadOnly Property StrPrice As String
        Get
            Return Format(DecValorAntesIva, "#0.00")
        End Get
    End Property
    Friend ReadOnly Property StrLinExtAmo As String
        Get
            Return Format(DecValorAntesIva, "#0.00")
        End Get
    End Property
    Friend ReadOnly Property StrLinTotTax As String
        Get
            Return Format(DecValorIva, "#0.00")
        End Get
    End Property
    Friend ReadOnly Property StrLinTot As String
        Get
            Return Format(ObjValor_ItemNotaDbDec.ObjValorPro, "#0.00")
        End Get
    End Property
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsBaseMoraDec
    Inherits clsCBPropiedad
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
            Dim lobjPadre As ClsItemNotaDb = ObjPadre
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
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
Friend Class ClsDiasMoraEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DiasMora"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Días de Mora"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Integer.MaxValue, BlnEsRequerido)
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
Friend Class ClsFechaCausoMora_Dtm
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsItemNotaDb = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaCausoMora"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha Causo Mora"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin As Date =
                GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim ldtmFechaMax As Date = If(GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro,
                Date.Today, GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo.AddDays(1))
        HblnEsValido = MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando OrElse
                ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
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
Friend Class ClsIdFactura_ItemNotaDbEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdFactura"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdFactura_ItemNotaDb"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido)
        If HblnEsValido Then
            Dim lobjPadre As ClsItemNotaDb = ObjPadre
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
Friend Class ClsIdItemFac_ItemNotaDbShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdItemFactura"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdItemFact_ItemNotaDb"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue, BlnEsRequerido)
        Dim lobjPadre As ClsItemNotaDb = ObjPadre
        If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
            HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
Friend Class ClsIdItemNotaDbShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdItemNotaDb"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdItemNotaDb"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 4
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HobjValorPro = True
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
Friend Class ClsIdNotaDb_ItemNotaDbEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdNotaDb"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdNotaDb_ItemNotaDb"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsItemNotaDb = ObjPadre
        Dim lobjNotaDb As ClsNotaDb = lobjPadre.ObjPadre
        If lobjNotaDb.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HblnEsValido = True
        Else
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido)
            If HblnEsValido Then
                HblnEsValido = (HobjValorNew = lobjNotaDb.ObjIdNotaDbEnt.ObjValorPro)
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
Friend Class ClsPrefijoFact_ItemNotaDbStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoFactura"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PrefijoFactura"
        HshrLongitud = 5
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then
            HobjValorNew = String.Empty
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            Dim lobjPadre As ClsItemNotaDb = ObjPadre
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
Friend Class ClsTasaMora_ItemNotaDbDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "TasaMora"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TasaMora"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, 1, HblnEsRequerido,
                EnuTipoValor.enuDouble)
        If HblnEsValido Then
            HobjValorNew = Math.Round(HobjValorNew, 6)
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
            Return Format(HobjValorPro, "p")
        End If
    End Function
End Class
Friend Class ClsTarifaIva_ItemNotaDbDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "TarifaIva"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TarifaDelIva_ItemNotaDb"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, 0.5, BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
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
            Return Format(HobjValorPro, "p")
        End If
    End Function
End Class
Friend Class ClsValor_ItemNotaDbDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "ValorItem_NotaDb"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0.01, Decimal.MaxValue,
                BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            Dim lobjPadre As ClsItemNotaDb = ObjPadre
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
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
#End Region