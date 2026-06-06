Friend Class ClsItemNotaCr
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriItemsNotaCr"
    ' Variables de modulo
    Private ReadOnly MobjPadre As clsNotaCr = Nothing
    Private McolNovedades As Collection = Nothing
    Private MobjFactura As clsFactura = Nothing
    Private MobjItemFac As clsItemFactura = Nothing
#End Region

#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwItemNotaCr">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As clsNotaCr, adrwItemNotaCr As DataRow)
        HobjPadre = aobjPadre
        MobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsSuprimible = False
        '
        DrwRegistroActual = adrwItemNotaCr
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
            Return EnuIdClasesPanDef.enuItemNotaCr
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Item Nota Cr."
        End Get
    End Property
#End Region

#Region "Propiedades Prop"
    Friend ReadOnly Property ObjBaseDscto_NotaCrDec As New ClsBaseDscto_NotaCrDec(Me)
    Friend ReadOnly Property ObjEsReversionIvaBln As New ClsEsReversionIvaBln(Me)
    Friend ReadOnly Property ObjIdCarpeta_ItemNotaCrShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_ItemNotaCrShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdFactura_ItemNotaCrEnt As New ClsIdFactura_ItemNotaCrEnt(Me)
    Friend ReadOnly Property ObjIdItemFac_ItemNotaCrShr As New ClsIdItemFac_ItemNotaCrShr(Me)
    Friend ReadOnly Property ObjIdItemNotaCrShr As New ClsIdItemNotaCrShr(Me)
    Friend ReadOnly Property ObjIdNotaCr_ItemNotaCrEnt As New ClsIdNotaCr_ItemNotaCrEnt(Me)
    Friend ReadOnly Property ObjIdTipoDscto_ItemNotaCrByt As New ClsIdTipoDscto_ItemNotaCrByt(Me)
    Friend ReadOnly Property ObjPrefijo_ItemNotaCrStr As New ClsPrefijo_NotaCrStr(Me)
    Friend ReadOnly Property ObjPrefijoFact_ItemNotaCrStr As New ClsPrefijoFact_ItemNotaCrStr(Me)
    Friend ReadOnly Property ObjTasaDscto_ItemNotaCrDbl As New ClsTasaDscto_ItemNotaCrDbl(Me)
    Friend ReadOnly Property ObjValor_ItemNotaCrDec As New ClsValor_ItemNotaCrDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjEsReversionIvaBln)
                HcolPropiedades.Add(ObjIdCarpeta_ItemNotaCrShr)
                HcolPropiedades.Add(ObjIdCentroUtil_ItemNotaCrShr)
                HcolPropiedades.Add(ObjPrefijo_ItemNotaCrStr)
                HcolPropiedades.Add(ObjIdNotaCr_ItemNotaCrEnt)
                HcolPropiedades.Add(ObjIdItemNotaCrShr)
                HcolPropiedades.Add(ObjPrefijoFact_ItemNotaCrStr)
                HcolPropiedades.Add(ObjIdFactura_ItemNotaCrEnt)
                HcolPropiedades.Add(ObjIdItemFac_ItemNotaCrShr)
                HcolPropiedades.Add(ObjIdTipoDscto_ItemNotaCrByt)
                HcolPropiedades.Add(ObjBaseDscto_NotaCrDec)
                HcolPropiedades.Add(ObjTasaDscto_ItemNotaCrDbl)
                HcolPropiedades.Add(ObjValor_ItemNotaCrDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region

#Region "Otras propiedades"
    Friend ReadOnly Property StrNroNotaCr As String
        Get
            Dim lstrNroNotaCr = ClsPanorama.FstrNumeroDcto(ObjPrefijo_ItemNotaCrStr.ObjValorPro,
                    ObjIdItemNotaCrShr.ObjValorPro)
            Return lstrNroNotaCr
        End Get
    End Property
    Friend ReadOnly Property StrNroFactura_ItemNotaCr As String
        Get
            Dim lstrNroFact_ItemNotaCr = ClsPanorama.FstrNumeroDcto(
                    ObjPrefijoFact_ItemNotaCrStr.ObjValorPro,
                    ObjIdFactura_ItemNotaCrEnt.ObjValorPro)
            Return lstrNroFact_ItemNotaCr
        End Get
    End Property
    Friend ReadOnly Property ObjFactura As ClsFactura
        Get
            If ObjPrefijoFact_ItemNotaCrStr.BlnEsValido AndAlso ObjIdFactura_ItemNotaCrEnt.BlnEsValido Then
                Dim lstrPrefFac As String = ObjPrefijoFact_ItemNotaCrStr.ObjValorPro
                Dim lentIdFac As Integer = ObjIdFactura_ItemNotaCrEnt.ObjValorPro
                Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFac, lentIdFac}
                MobjFactura = New ClsFactura()
                MobjFactura.SAbra(lobjValorLlave)
            Else
                MobjFactura = Nothing
            End If
            Return MobjFactura
        End Get
    End Property
    Friend ReadOnly Property ObjItemFac As ClsItemFactura
        Get
            If ObjIdItemFac_ItemNotaCrShr.BlnEsValido Then
                If IsNothing(MobjItemFac) Then
                    If Not IsNothing(ObjFactura) Then
                        Dim lshrIdItemFac As Short = ObjIdItemFac_ItemNotaCrShr.ObjValorPro
                        MobjItemFac = ObjFactura.ColItemsFactura(lshrIdItemFac.ToString)
                    End If
                End If
            Else
                MobjItemFac = Nothing
            End If
            Return MobjItemFac
        End Get
    End Property
    Friend ReadOnly Property DblTasaDsctoServicio(aenuTipodscto As EnuTipoDescuento) As Double
        Get
            Dim ldblTasadscto = 0.0
            If aenuTipodscto > EnuTipoDescuento.EnuReteFuente Then
                If Not IsNothing(ObjItemFac) Then
                    Select Case aenuTipodscto
                        Case EnuTipoDescuento.EnuReteFuente
                            ldblTasadscto = ObjItemFac.ObjServicio.ObjTarifaRetFteDbl.ObjValorPro
                        Case EnuTipoDescuento.EnuReteIca
                            ldblTasadscto = ObjItemFac.ObjServicio.ObjTarifaRetIcaDbl.ObjValorPro
                        Case EnuTipoDescuento.EnuReteIva
                            ldblTasadscto = GobjParametros.ObjTarifaReteIvaDbl.ObjValorPro
                    End Select
                End If
            End If
            Return ldblTasadscto
        End Get
    End Property
    Friend ReadOnly Property StrItemFactura As String
        Get
            Dim lstrItemFra = String.Empty
            If ObjIdItemFac_ItemNotaCrShr.BlnEsValido Then
                lstrItemFra = ObjItemFac.ObjIdItemFacturaShr.ToString & "-" &
                        ObjItemFac.ObjDetalle_ItemFactStr.ObjValorPro
            End If
            Return lstrItemFra
        End Get
    End Property
    Friend ReadOnly Property BlnEsDscto As Boolean
        Get
            Dim lenuTipoDscto As EnuTipoDescuento = ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
            Dim lblnEsDscto As Boolean = (lenuTipoDscto = EnuTipoDescuento.EnuDsctoCapital OrElse
                    lenuTipoDscto = EnuTipoDescuento.EnuDsctoIntMora OrElse
                    lenuTipoDscto = EnuTipoDescuento.EnuDsctoPP)
            Return lblnEsDscto
        End Get
    End Property
    Friend ReadOnly Property BlnEsRetencion As Boolean
        Get
            Dim lenuTipoDscto As EnuTipoDescuento = ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
            Dim lblnEsRet As Boolean = Not (lenuTipoDscto = EnuTipoDescuento.EnuDsctoCapital OrElse
                    lenuTipoDscto = EnuTipoDescuento.EnuDsctoIntMora OrElse
                    lenuTipoDscto = EnuTipoDescuento.EnuDsctoPP OrElse
                    lenuTipoDscto = EnuTipoDescuento.EnuCancelaIva)
            Return lblnEsRet
        End Get
    End Property
    Friend ReadOnly Property DecBaseIva As Decimal
        Get
            Dim ldecBaseIva = 0D
            If MobjPadre.ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuAnulaFac Then
                ldecBaseIva = Math.Round(ObjValor_ItemNotaCrDec.ObjValorPro /
                        (1 + ObjItemFac.ObjTarifaIva_ItemFactDbl.ObjValorPro), 0)
            End If
            Return ldecBaseIva
        End Get
    End Property
    Friend ReadOnly Property DecValorIva As Decimal
        Get
            Dim ldecValorIva = 0D
            If MobjPadre.ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuAnulaFac Then
                ldecValorIva = ObjValor_ItemNotaCrDec.ObjValorPro - DecBaseIva
            End If
            Return ldecValorIva
        End Get
    End Property
#End Region
#End Region

#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        MobjFactura = Nothing
        MobjItemFac = Nothing
        McolNovedades = Nothing
    End Sub
    Protected Overrides Sub SInicialiceObj()
        ObjEsReversionIvaBln.ObjValorPro = False
        ObjIdCarpeta_ItemNotaCrShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtil_ItemNotaCrShr.ObjValorPro = GshrIdCentroUtil
        ObjIdFactura_ItemNotaCrEnt.ObjValorPro = 0
        ObjIdNotaCr_ItemNotaCrEnt.ObjValorPro = 0
        ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro = EnuTipoDescuento.None
        ObjPrefijoFact_ItemNotaCrStr.ObjValorPro = String.Empty
        ObjValor_ItemNotaCrDec.ObjValorPro = 0
        Dim lstrPrefijo = GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.enuNotaCr)
        If IsNothing(lstrPrefijo) Then lstrPrefijo = String.Empty
        ObjPrefijo_ItemNotaCrStr.ObjValorPro = lstrPrefijo
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Dim lstrPrefijo As String = ObjPrefijo_ItemNotaCrStr.ObjValorPro
            Dim lstrIdObjeto = String.Empty
            If lstrPrefijo <> "" Then
                lstrIdObjeto = lstrPrefijo & "-"
            End If
            lstrIdObjeto &= ObjIdItemNotaCrShr.ToString
            Return lstrIdObjeto
        End Get
    End Property
#End Region

#Region "Novedades"
    Friend ReadOnly Property ColNovedades As Collection
        Get
            If IsNothing(McolNovedades) Then
                McolNovedades = New Collection
                Dim ldtbNovedades As DataTable = MobjPadre.DtbNovedades()
                If Not IsNothing(ldtbNovedades) AndAlso ldtbNovedades.Rows.Count > 0 Then
                    Dim lstrfiltro = ClsIdItemDocOrigen_NovShr.SstrNombreCampoBd & " = " &
                            ObjIdItemNotaCrShr.ObjValorPro
                    Dim ldrwNovedades() As DataRow = ldtbNovedades.Select(lstrfiltro)
                    For Each ldrwNovedad As DataRow In ldrwNovedades
                        Dim lobjNovedad As New ClsNovedad(Me, ldrwNovedad)
                        lobjNovedad.SLeaValores(True)
                        McolNovedades.Add(lobjNovedad, lobjNovedad.ObjIdNovedadShr.ToString)
                    Next
                End If
            End If
            Return McolNovedades
        End Get
    End Property
#End Region

#Region "Procedimientos del objeto"
    Friend Function FblnEsValidaFechaNotaCr(adtmFechaNotaCr As Date,
            aobjFactura As ClsFactura) As Boolean
        Dim lblnEsValida = False
        If ObjPrefijoFact_ItemNotaCrStr.BlnEsValido AndAlso
                ObjIdFactura_ItemNotaCrEnt.BlnEsValido Then
            If MobjPadre.ObjIdPredioAgrupador_NotaCrStr.BlnEsValido Then
                lblnEsValida = (adtmFechaNotaCr >= aobjFactura.ObjFechaFacturaDtm.ObjValorPro)
            End If
        End If
        Return lblnEsValida
    End Function
    Friend Function FblnEsValidoDscto(adecValorDscto As Decimal,
                ByRef astrMens As String) As Boolean
        GobjPanDat.SControleProcesoObj(True)
        Dim lobjNotaCr As ClsNotaCr = ObjPadre
        Dim lblnEsValido = ObjIdTipoDscto_ItemNotaCrByt.BlnEsValido AndAlso
                    ObjPrefijoFact_ItemNotaCrStr.BlnEsValido AndAlso
                    ObjIdFactura_ItemNotaCrEnt.BlnEsValido AndAlso
                    ObjIdItemFac_ItemNotaCrShr.BlnEsValido
        If lblnEsValido Then
            lblnEsValido = adecValorDscto - Int(adecValorDscto) = 0
            If Not lblnEsValido Then
                astrMens = "El Valor ingresado debe ser sin Centavos!"
            End If
        End If
        If lblnEsValido Then
            Dim lenuSev As EnuSeveridadNot = EnuSeveridadNot.None
            Dim lstrIdItemFac = ObjIdItemFac_ItemNotaCrShr.ToString
            Dim lobjItemFac As ClsItemFactura = ObjFactura.ColItemsFactura(lstrIdItemFac)
            Dim lenuTipoDscto As EnuTipoDescuento = ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
            If lenuTipoDscto >= EnuTipoDescuento.EnuReteFuente AndAlso lenuTipoDscto <=
                        EnuTipoDescuento.EnuReteCree Then
                lblnEsValido = Not lobjItemFac.FblnRetencionAplicada(lenuTipoDscto)
                If Not lblnEsValido Then
                    astrMens = "Este tipo de Retención ya fue aplicada!"
                End If
            End If
            If lblnEsValido Then
                If lenuTipoDscto = EnuTipoDescuento.EnuDsctoCapital Then
                    Dim ldecDeudaSinIva = lobjItemFac.DecDeudaSer
                    lblnEsValido = adecValorDscto <= (ldecDeudaSinIva)
                    If Not lblnEsValido Then
                        astrMens = "El Valor ingresado es mayor a la Deuda de Capital!"
                    End If
                ElseIf lenuTipoDscto = EnuTipoDescuento.EnuDsctoIntMora Then
                    lblnEsValido = adecValorDscto <= ObjItemFac.FdecDeudaIntTotal
                    If Not lblnEsValido Then
                        astrMens = "El Valor ingresado es mayor a la Deuda de Intereses!"
                    End If
                End If
            End If
            If lblnEsValido Then
                ' Valor de la deuda antes del Iva
                lblnEsValido = ObjFactura.FblnEsValidoDescuento(
                        ObjIdItemFac_ItemNotaCrShr.ObjValorPro, adecValorDscto, lenuTipoDscto,
                        astrMens, lenuSev)
            End If
        End If
        GobjPanDat.SControleProcesoObj(False)
        Return lblnEsValido
    End Function
    Friend Function FblnEsValidoItemDscto(aenuIdTipoDscto As EnuTipoDescuento,
                ByRef astrMens As String) As Boolean
        Dim lblnEsValido As Boolean = True
        GobjPanDat.SControleProcesoObj(True)
        Dim lobjNotaCr As ClsNotaCr = ObjPadre
        If lobjNotaCr.ObjIdModoNotaCrByt.ObjValorPro = EnuModoNotaCr.EnuPorFactura Then
            lblnEsValido = Not lobjNotaCr.FblnExisteNuevoItem(
                ObjPrefijoFact_ItemNotaCrStr.ObjValorPro, ObjIdFactura_ItemNotaCrEnt.ObjValorPro,
                ObjIdItemFac_ItemNotaCrShr.ObjValorPro, aenuIdTipoDscto)
            If Not lblnEsValido Then
                astrMens = "No es posible repetir un Descuento ya hecho!"
            End If
        End If
        GobjPanDat.SControleProcesoObj(False)
        Return lblnEsValido
    End Function
    Friend Function FblbEsValidDsctoPP() As Boolean
        Dim lblnEsValido = GobjParametros.ObjAnoActual.ObjAplicaDsctoPPBln.ObjValorPro
        If lblnEsValido Then
            If Not IsNothing(ObjFactura) AndAlso Not IsNothing(ObjItemFac) Then
                lblnEsValido = ObjItemFac.ObjServicio.BlnEsCuotaAdministracion
            End If
        End If
        Return lblnEsValido
    End Function
    Friend Function FenuTipoNovedad() As EnuTipoNov
        Dim lenuTipoDscto As EnuTipoDescuento = ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
        Dim lenuTipoNov As EnuTipoNov = EnuTipoNov.None
        Select Case lenuTipoDscto
            Case EnuTipoDescuento.enuDsctoCapital
                If MobjPadre.BlnAnulandoFac Then
                    If ObjEsReversionIvaBln.ObjValorPro Then
                        lenuTipoNov = EnuTipoNov.EnuRDbIva
                    Else
                        lenuTipoNov = EnuTipoNov.enuRDbCap
                    End If
                Else
                    lenuTipoNov = EnuTipoNov.enuCrDctoCap
                End If
            Case EnuTipoDescuento.enuDsctoIntMora
                lenuTipoNov = EnuTipoNov.enuCrDctoInt
            Case EnuTipoDescuento.enuReteCree
                lenuTipoNov = EnuTipoNov.enuCrRetCre
            Case EnuTipoDescuento.enuReteFuente
                lenuTipoNov = EnuTipoNov.enuCrRetFte
            Case EnuTipoDescuento.enuReteIca
                lenuTipoNov = EnuTipoNov.enuCrRetIca
            Case EnuTipoDescuento.enuReteIva
                lenuTipoNov = EnuTipoNov.enuCrRetIva
            Case EnuTipoDescuento.enuDsctoPP
                lenuTipoNov = EnuTipoNov.enuCrDctoCap
            Case EnuTipoDescuento.enuCancelaIva
                lenuTipoNov = EnuTipoNov.enuCrIvaGas
        End Select
        Return lenuTipoNov
    End Function
    Friend Function FblnEsDscto() As Boolean
        Dim lenuTipoItem As EnuTipoDescuento = ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
        Return (lenuTipoItem = EnuTipoDescuento.enuDsctoCapital) OrElse
                (lenuTipoItem = EnuTipoDescuento.enuDsctoIntMora) OrElse
                (lenuTipoItem = EnuTipoDescuento.enuDsctoPP)
    End Function
    Friend Function FblnEsRetencion() As Boolean
        Dim lenuTipoItem As EnuTipoDescuento = ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
        Return (lenuTipoItem > EnuTipoDescuento.EnuDsctoIntMora) AndAlso
                (lenuTipoItem < EnuTipoDescuento.EnuDsctoPP)
    End Function
#End Region

End Class

#Region "Clases de Propiedad"
Friend Class ClsBaseDscto_NotaCrDec
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsItemNotaCr
    Private Const MCSTRNOMBRECAMPOBD As String = "BaseDscto"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
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
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If MobjPadre.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro >= EnuTipoDescuento.EnuDsctoIntMora Then
                    HblnEsValido = (HobjValorNew > 0)
                End If
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
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsEsReversionIvaBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EsReversionIva"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Es Reversión Iva"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
    End Function
End Class

Friend Class ClsIdFactura_ItemNotaCrEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdFactura"
    Private ReadOnly MobjPadre As ClsItemNotaCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdFactura_ItemNotaCr"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HstrMens = String.Empty
                Dim lobjNotaCr As ClsNotaCr = MobjPadre.ObjPadre
                HblnEsValido = MobjPadre.ObjPrefijoFact_ItemNotaCrStr.BlnEsValido
                Dim lstrPref = MobjPadre.ObjPrefijoFact_ItemNotaCrStr.ObjValorPro
                If HblnEsValido Then
                    Dim lobjCliente As ClsCliente = lobjNotaCr.ObjClienteNotaCr
                    If lobjCliente.FblnEsFactElec(lstrPref, HobjValorNew) Then
                        HblnEsValido = lobjCliente.FblnFactEstadoEFacOk(lstrPref, HobjValorNew)
                        If Not HblnEsValido Then
                            HstrMens = "No es posible hacer una Nota Crédito a una Factura " &
                                    "no registrada!"
                            SNotifiqueDatInv()
                        End If
                    End If
                End If
                If Not String.IsNullOrEmpty(HstrMens) Then
                    SNotifiqueDatInv()
                End If
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso HblnEsValido Then
            MobjPadre.ObjValor_ItemNotaCrDec.SValide()
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

Friend Class ClsIdItemFac_ItemNotaCrShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdItemFactura"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Id Item Factura ItemNotaCr"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            Dim lobjPadre As ClsCBObjetoPan = ObjPadre
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

Friend Class ClsIdItemNotaCrShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdItemNotaCr"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdItemNotaCr"
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

Friend Class ClsIdNotaCr_ItemNotaCrEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdNotaCr"
    Private ReadOnly MobjPadre As ClsItemNotaCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdNotaCr_ItemNotaCr"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        Dim lobjNotaCr As ClsNotaCr = MobjPadre.ObjPadre
        If lobjNotaCr.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HblnEsValido = True
        Else
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido)
            If HblnEsValido Then
                HblnEsValido = (HobjValorNew = lobjNotaCr.ObjIdNotaCrEnt.ObjValorPro)
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

Friend Class ClsIdTipoDscto_ItemNotaCrByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoDescuento"
    Private ReadOnly MobjPadre As ClsItemNotaCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTipoDescuento"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuTipoDescuento.EnuDsctoCapital, EnuTipoDescuento.EnuCancelaIva,
                BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            Else
                Dim lenuTipoDscto As EnuTipoDescuento = HobjValorNew
                HblnEsValido = MobjPadre.FblnEsValidoItemDscto(lenuTipoDscto, HstrMens)
                If HblnEsValido AndAlso lenuTipoDscto = EnuTipoDescuento.EnuDsctoPP Then
                    HblnEsValido = MobjPadre.FblbEsValidDsctoPP()
                    If Not HblnEsValido Then
                        HstrMens = "El Descuento por Pronto Pago solo se aplica " &
                            " a las Cuotas de Administración!"
                    End If
                End If
                If Not String.IsNullOrEmpty(HstrMens) Then
                    If Not HblnEsValido Then
                        SNotifiqueDatInv()
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso HblnEsValido Then
            MobjPadre.ObjValor_ItemNotaCrDec.SValide()
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
            Return ClsOrionCop.FstrNombreDatoConstanteOri(
                    EnuGrupoConstantesOriDef.EnuTipoDescuento, HobjValorPro)
        End If
    End Function
End Class

Friend Class ClsPrefijoFact_ItemNotaCrStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoFactura"
    Private ReadOnly MobjPadre As ClsItemNotaCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
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
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso HblnEsValido Then
            MobjPadre.ObjValor_ItemNotaCrDec.SValide()
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

Friend Class ClsTasaDscto_ItemNotaCrDbl
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsItemNotaCr = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "TasaDscto"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Tasa de descuento"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, 1.0, BlnEsRequerido,
                EnuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew > 0)
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
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsValor_ItemNotaCrDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"
    Private ReadOnly MobjPadre As ClsItemNotaCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "ValorItem_NotaCr"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0.01, Decimal.MaxValue,
                BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = HobjValorNew = HobjValorOriginal
            Else
                Dim ldecVlrDscto As Decimal = HobjValorNew
                HblnEsValido = MobjPadre.FblnEsValidoDscto(ldecVlrDscto, HstrMens)
                If Not String.IsNullOrEmpty(HstrMens) Then
                    SNotifiqueDatInv()
                End If
            End If
        Else
            If HobjValorNew = 0 Then
                Dim lobjNotaCr As ClsNotaCr = MobjPadre.ObjPadre
                HblnEsValido = lobjNotaCr.ObjAnuladoBln.ObjValorPro
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
            Return Format(0, "c")
        Else
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
#End Region