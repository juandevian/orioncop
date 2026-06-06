Imports OPT.OrionP.OrionCopL
Friend MustInherit Class ClsCBInterfazContableOri
#Region "Definiciones"
    'Variables
    Protected Const CHSTRCOMA = ","
    Private ReadOnly MblnRegistrado As Boolean = False
#End Region
#Region "Constructores"
    Protected Sub New(aobjRegistro As Object)
        If aobjRegistro Is Nothing Then
            Throw New ModuloNoRegistradoPanException("El módulo no ha sido debidamente cargado!")
        ElseIf Not (aobjRegistro.GetType.Name = "String" AndAlso aobjRegistro = GCOBJREGISTRO) Then
            Throw New ModuloNoRegistradoPanException("El módulo no ha sido debidamente cargado!")
        Else
            MblnRegistrado = True
        End If
    End Sub
#End Region
#Region "Propiedades"
    Protected Property HstrArchivoSalidaInterfaz As String = String.Empty
    Friend MustOverride ReadOnly Property DtmFechaFinInterfazAnterior As Date

    Friend Property DtmFechaDesde As Date = GCDTMFECHANULA
    Friend Property DtmFechaHasta As Date = Date.Today
    Friend Property EntIdComprobanteInicial As Integer = 0
    Friend Property BlnIdTerceroStr As Boolean = False
    Friend ReadOnly Property BlnRegistrado As Boolean
        Get
            Return MblnRegistrado
        End Get
    End Property
    Friend Shared ReadOnly Property BlnEsValidoTipoComprobante As Boolean
        Get
            Dim lblnEsValido = True
            If GobjParametros.ObjTipoInterfazByt.ObjValorPro = EnuTipoInterfazDef.EnuPorComprobante Then
                lblnEsValido = ClsPanorama.FblnEsValidoString(StrTipoComprobante, 2, 10, True)
            End If
            Return lblnEsValido
        End Get
    End Property
    Friend ReadOnly Property BlnEsValidaFechaDesde As Boolean
        Get
            Dim ldtmFechaMin = #1/1/2010#
            Dim ldtmFechaMax = Now
            Dim lblnEsValido = ClsPanorama.FblnEsValidoFecha(DtmFechaDesde, ldtmFechaMin, ldtmFechaMax, True)
            Return lblnEsValido
        End Get
    End Property
    Friend ReadOnly Property BlnEsValidaFechaHasta As Boolean
        Get
            Dim ldtmFechaMin = Date.Today
            If BlnEsValidaFechaDesde Then
                ldtmFechaMin = DtmFechaDesde
            End If
            Dim ldtmFechaMax = Now
            Dim lblnEsValido = ClsPanorama.FblnEsValidoFecha(DtmFechaHasta, ldtmFechaMin, ldtmFechaMax, True)
            Return lblnEsValido
        End Get
    End Property
    Friend ReadOnly Property BlnEsValidaIdComprobanteInicial As Boolean
        Get
            Dim lblnEsValido = True
            If GobjParametros.ObjTipoInterfazByt.ObjValorPro = EnuTipoInterfazDef.EnuPorComprobante Then
                lblnEsValido = ClsPanorama.FblnEsValidoNumero(EntIdComprobanteInicial, 1,
                        Integer.MaxValue, True, EnuTipoValor.enuInteger)
            End If
            Return lblnEsValido
        End Get
    End Property
    Friend Shared ReadOnly Property StrTipoComprobante As String
        Get
            Dim lstrTipoCom = String.Empty
            If GobjParametros.ObjTipoInterfazByt.ObjValorPro = EnuTipoInterfazDef.EnuPorComprobante Then
                Dim lobjDoc As ClsDocumento = GobjParametros.ObjDocumento(EnuIdDocumentoDef.EnuComprobanteInterfaz)
                lstrTipoCom = lobjDoc.ObjTipoDocumentoStr.ObjValorPro
            End If
            Return lstrTipoCom
        End Get
    End Property
    Friend Shared ReadOnly Property StrPrefijoComprobante As String
        Get
            Dim lstrPrefCom = String.Empty
            If GobjParametros.ObjTipoInterfazByt.ObjValorPro = EnuTipoInterfazDef.EnuPorComprobante Then
                Dim lobjDoc As ClsDocumento = GobjParametros.ObjDocumento(EnuIdDocumentoDef.EnuComprobanteInterfaz)
                lstrPrefCom = lobjDoc.ObjPrefijo_DocStr.ObjValorPro
            End If
            Return lstrPrefCom
        End Get
    End Property
#End Region
#Region "Procedimientos"
    Friend Overridable Sub SGenerereInterfazContable(ablnFinMes As Boolean, ByRef astrMens As String)
        '
    End Sub
    Protected Function FblnEstanTodosOkInterfaz() As Boolean
        Dim lblnEstanOk = BlnEsValidaFechaDesde AndAlso BlnEsValidaFechaHasta AndAlso
                BlnEsValidaIdComprobanteInicial AndAlso BlnEsValidoTipoComprobante
        Return lblnEstanOk
    End Function
    Friend Overridable Function FentIdUltimoComprobanteInterfazAnterior() As Integer
        Return 0
    End Function
    Protected Shared Function FstrUltimoArchInt(astrNombreArch As String) As String
        Dim lstrFechaFin = String.Empty, lstrUltimoArch = String.Empty
        If My.Computer.FileSystem.DirectoryExists(GstrTrayInterfContable) Then
            Dim lcolArchivos = My.Computer.FileSystem.GetFiles(GstrTrayInterfContable,
                        FileIO.SearchOption.SearchTopLevelOnly, astrNombreArch)
            If lcolArchivos.Count > 0 Then
                Dim lstrPartesArch() As String = Array.Empty(Of String)
                For Each lstrArchivo As String In lcolArchivos
                    lstrPartesArch = lstrArchivo.Split("_")
                    If lstrPartesArch.Length = 3 Then
                        If lstrFechaFin < lstrPartesArch(2).Substring(0, 8) Then
                            lstrFechaFin = lstrPartesArch(2).Substring(0, 8)
                            lstrUltimoArch = lstrArchivo
                        End If
                    End If
                Next
            End If
        End If
        Return lstrUltimoArch
    End Function
    Protected Shared Function FstrDetalle(adrwNov As DataRow) As String
        If adrwNov Is Nothing Then
            Throw New ValorArgumentoInvalidoException("adrwNov en clsInterfazContableOri.fstrDetalle")
        End If
        Dim lstrDetalle As String, lstrDetaAdiDoc As String
        Dim lenutipo As EnuTipoNov = ClsPanorama.FobjValorCampo(adrwNov(
                ClsIdTipoNovedadByt.SstrNombreCampoBd), EnuTipoValor.enuByte)
        Dim lenuTipoDocOrigen As EnuTipoDocOri = ClsPanorama.FobjValorCampo(adrwNov(
                ClsIdTipoDocOrigenByt.SstrNombreCampoBd), EnuTipoValor.enuByte)
        Dim lstrPrefijoDoc As String = CType(ClsPanorama.FobjValorCampo(adrwNov(
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString), String)
        If IsNothing(lstrPrefijoDoc) Then lstrPrefijoDoc = String.Empty
        Dim lentIdDoc As Integer = CType(ClsPanorama.FobjValorCampo(adrwNov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
        Dim lstrNroDocOrige = ClsPanorama.FstrNumeroDcto(lstrPrefijoDoc, lentIdDoc)
        Dim lstrPredioAgr As String = ClsPanorama.FobjValorCampo(adrwNov(
                ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        If Not String.IsNullOrEmpty(lstrPredioAgr) Then
            lstrNroDocOrige += " " & lstrPredioAgr
        End If
        Dim lstrTipoDocCobro = "Cuenta de Cobro "
        If GobjParametros.BlnEFacAutorizado Then
            lstrTipoDocCobro = "Factura "
        End If
        Select Case lenuTipoDocOrigen
            Case EnuTipoDocOri.EnuFactura
                lstrDetalle = If(lenutipo = EnuTipoNov.EnuDbIva, "Iva factura " & lstrNroDocOrige,
                        lstrTipoDocCobro & lstrNroDocOrige)
            Case EnuTipoDocOri.EnuNotaCon
                lstrDetalle = "Nota aplicación Anticipos " & lstrNroDocOrige
            Case EnuTipoDocOri.EnuNotaCr
                lstrDetaAdiDoc = FstrDetaAdiNCr(lstrPrefijoDoc, lentIdDoc)
                lstrDetalle = "Nota Crédito " & lstrNroDocOrige & ". " & lstrDetaAdiDoc
                If lstrDetalle.Length > 60 Then
                    lstrDetalle = lstrDetalle.Substring(0, 60)
                End If
            Case EnuTipoDocOri.EnuNotaDb
                lstrDetalle = If(lenutipo = EnuTipoNov.EnuDbIvaInt,
                        "Iva Nota Db por Intereses " & lstrNroDocOrige,
                        "Nota Db por Intereses " & lstrNroDocOrige)
            Case EnuTipoDocOri.EnuNotaDevAnt
                lstrDetalle = "Nota Reintegro Anticipos " & lstrNroDocOrige
            Case EnuTipoDocOri.EnuReciboCaja
                lstrDetaAdiDoc = FstrDetaAdiRC(lstrPrefijoDoc, lentIdDoc)
                lstrDetalle = "Recibo de Caja " & lstrNroDocOrige & ". " & lstrDetaAdiDoc
                If lstrDetalle.Length > 60 Then
                    lstrDetalle = lstrDetalle.Substring(0, 60)
                End If
            Case EnuTipoDocOri.EnuNotaRevCr
                lstrDetaAdiDoc = FstrDetaAdiNRCr(lstrPrefijoDoc, lentIdDoc)
                lstrDetalle = "Nota Reversión Crédito " & lstrNroDocOrige & ". " & lstrDetaAdiDoc
                If lstrDetalle.Length > 60 Then
                    lstrDetalle = lstrDetalle.Substring(0, 60)
                End If
            Case EnuTipoDocOri.EnuNotaAjuste
                lstrDetalle = "Nota Ajuste Cuota Administración " & lstrNroDocOrige
            Case Else
                Throw New ErrorInesperadoPanLException("Tipo Documento Origen no esperado")
        End Select
        Return lstrDetalle
    End Function
    Private Shared Function FstrDetaAdiRC(astrPref As String, aentIdRC As Integer) As String
        Dim lstrTabla = ClsReciboCaja.SstrNombreTabla
        Dim lstrCamSel As String() = {ClsComentario_RecStr.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsPrefijo_RecStr.SstrNombreCampoBd &
                " = '" & astrPref & "' AND " & ClsIdRecCajaEnt.SstrNombreCampoBd & " = " & aentIdRC
        Dim lstrOrden = {{"", ""}}
        Dim ldtbComRC = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSel, lstrOrden, lstrFiltro)
        Dim lstrDetAdi As String = ClsPanorama.FobjValorCampo(ldtbComRC.Rows(0)(0),
                EnuTipoValor.enuString)
        Return lstrDetAdi
    End Function
    Private Shared Function FstrDetaAdiNCr(astrPref As String, aentIdNCr As Integer) As String
        Dim lstrTabla = ClsNotaCr.SstrNombreTabla
        Dim lstrCamSel As String() = {ClsComentario_NotaCrStr.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsPrefijo_NotaCrStr.SstrNombreCampoBd &
                " = '" & astrPref & "' AND " & ClsIdNotaCrEnt.SstrNombreCampoBd & " = " & aentIdNCr
        Dim lstrOrden = {{"", ""}}
        Dim ldtbComNCr = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSel, lstrOrden, lstrFiltro)
        Dim lstrDetAdi As String = ClsPanorama.FobjValorCampo(ldtbComNCr.Rows(0)(0),
                EnuTipoValor.enuString)
        Return lstrDetAdi
    End Function
    Private Shared Function FstrDetaAdiNRCr(astrPref As String, aentIdNRCr As Integer) As String
        Dim lstrTabla = ClsNotaReversionCr.SstrNombreTabla
        Dim lstrCamSel As String() = {ClsDetalle_NotaReversaCrStr.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd & " = '" & astrPref & "' AND " &
                ClsIdNotaReversaCrEnt.SstrNombreCampoBd & " = " & aentIdNRCr
        Dim lstrOrden = {{"", ""}}
        Dim ldtbComNCr = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSel, lstrOrden, lstrFiltro)
        Dim lstrDetAdi As String = ClsPanorama.FobjValorCampo(ldtbComNCr.Rows(0)(0),
                EnuTipoValor.enuString)
        Return lstrDetAdi
    End Function
    Protected Shared Function FstrIdTerceroCajaBancos(adrwMov As DataRow) As String
        If adrwMov Is Nothing Then
            Throw New ErrorInesperadoPanLException("adrwMov es nulo!")
        End If
        Dim lstrTer = String.Empty
        Select Case GobjParametros.ObjTipoTerceroCajaByt.ObjValorPro
            Case EnuTipoTerceroCajaDef.EnuSinTercero
                '
            Case EnuTipoTerceroCajaDef.EnuCliente
                lstrTer = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
            Case EnuTipoTerceroCajaDef.EnuCopropiedad
                lstrTer = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.
                        ObjTerceroCentroUtilidad.ObjIdTerceroDbl.ToString
        End Select
        Return lstrTer
    End Function
    Protected Shared Function FstrNombreTerceroCajaBancos(adrwMov As DataRow) As String
        If adrwMov Is Nothing Then
            Throw New ErrorInesperadoPanLException("adrwMov es nulo!")
        End If
        Dim lstrNombre = String.Empty
        Select Case GobjParametros.ObjTipoTerceroCajaByt.ObjValorPro
            Case EnuTipoTerceroCajaDef.EnuSinTercero
                lstrNombre = String.Empty
            Case EnuTipoTerceroCajaDef.EnuCliente
                lstrNombre = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
            Case EnuTipoTerceroCajaDef.EnuCopropiedad
                lstrNombre = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.
                        ObjTerceroCentroUtilidad.FstrNombreCompleto()
        End Select
        Return lstrNombre
    End Function
    Friend Function FblnHayDatos() As Boolean
        Dim lblnHayDatos = False
        If Not (DtmFechaDesde = GCDTMFECHANULA OrElse DtmFechaHasta = GCDTMFECHANULA) Then
            Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaDesde) & "'"
            Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaHasta) & "'"
            Dim lstrTabla = ClsNovedad.SstrNombreTabla
            Dim lstrCamposSelect As String() = {"COUNT(" & ClsIdFactura_NovEnt.SstrNombreCampoBd & ")"}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsFechaNovedadDtm.SstrNombreCampoBd & " BETWEEN " & lstrFechaDesde & " AND " &
                    lstrFechaHasta
            Dim lstrIndice As String(,) = {{"", ""}}
            Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrIndice, lstrFiltro)
            lblnHayDatos = ldtbRes.Rows(0)(0) > 0
        End If
        Return lblnHayDatos
    End Function
#End Region
#Region "Lectura datos Orión"
    Protected Function FdtbMovimiento() As DataTable 'Ok
        Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaDesde.ToString) & "'"
        Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaHasta.ToString) & "'"
        Dim lstrNombreTablaT1 = ClsNovedad.SstrNombreTabla
        Dim lstrNombreTablaT2 = ClsCliente.SstrNombreTabla
        Dim lstrCamposSelectT1 = {ClsIdTipoDocOrigenByt.SstrNombreCampoBd,
                ClsFechaNovedadDtm.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd,
                ClsIdDocOrigenEnt.SstrNombreCampoBd,
                ClsPrefijoFact_NovStr.SstrNombreCampoBd,
                ClsIdFactura_NovEnt.SstrNombreCampoBd,
                ClsIdItemFacturaShr.SstrNombreCampoBd,
                ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd,
                ClsIdCuentaDb_NovStr.SstrNombreCampoBd,
                ClsIdCuentaCr_NovStr.SstrNombreCampoBd,
                ClsIdTipoNovedadByt.SstrNombreCampoBd,
                ClsAliasCont_NovStr.SstrNombreCampoBd,
                ClsIdTercero_NovDbl.SstrNombreCampoBd,
                "SUM(" & ClsBaseDec.SstrNombreCampoBd & ")",
                "SUM(" & ClsValor_NovDec.SstrNombreCampoBd & ")"}
        Dim lstrCamposSelectT2 = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdTercero_NovDbl.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsIdTipoDocOrigenByt.SstrNombreCampoBd, "Asc"},
                          {ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd, "ASC"},
                          {ClsIdDocOrigenEnt.SstrNombreCampoBd, "ASC"},
                          {ClsIdCuentaDb_NovStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsFechaNovedadDtm.SstrNombreCampoBd & " >= " & lstrFechaDesde & " AND " &
                ClsFechaNovedadDtm.SstrNombreCampoBd & " <= " & lstrFechaHasta & " AND " &
                ClsValor_NovDec.SstrNombreCampoBd & " > 0"
        Dim lstrCamposGrupo = {ClsIdTipoDocOrigenByt.SstrNombreCampoBd,
                ClsFechaNovedadDtm.SstrNombreCampoBd, ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd,
                ClsIdDocOrigenEnt.SstrNombreCampoBd, ClsIdCuentaDb_NovStr.SstrNombreCampoBd,
                ClsIdCuentaCr_NovStr.SstrNombreCampoBd, ClsIdTipoNovedadByt.SstrNombreCampoBd,
                ClsAliasContStr.SstrNombreCampoBd,
                ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim ldtbMovimiento = ClsPanorama.FdtbDataTable(lstrNombreTablaT1, lstrCamposSelectT1,
                lstrNombreTablaT2, lstrCamposSelectT2, lstrCamposRelPri, lstrCamposRelSec,
                lstrIndice, False, lstrFiltro, lstrCamposGrupo)
        Return ldtbMovimiento
    End Function
    Protected Function FdtbMoviRC() As DataTable
        Dim lstrIndice = ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " ASC, " &
                ClsIdDocOrigenEnt.SstrNombreCampoBd & " ASC, " &
                ClsIdCuentaDb_NovStr.SstrNombreCampoBd & " ASC"
        Dim lstrSqlRCNov = FstrExpSqlRCNov()
        Dim lstrSqlRCNovAnt = FstrExpSqlRCNovAnt()
        Dim lstrsqlMovRC = "(" & lstrSqlRCNov & ") UNION ALL (" & lstrSqlRCNovAnt & ") ORDER BY " &
                lstrIndice
        Dim ldtbMoviRC = ClsPanorama.FdtbDataTable(lstrsqlMovRC)
        Return ldtbMoviRC
    End Function
    Protected Function FdtbMoviAnt(aenuDocumOrigen As EnuIdDocumentoDef) As DataTable 'Ok
        Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaDesde.ToString) & "'"
        Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaHasta.ToString) & "'"
        Dim lstrNombreTablaT1 = ClsNovedadAnticipo.SstrNombreTabla
        Dim lstrNombreTablaT2 = ClsCliente.SstrNombreTabla
        Dim lstrCamposSelectT1 = {ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd,
                ClsFechaNovedadAntDtm.SstrNombreCampoBd,
                "'' AS " & ClsPrefijoFact_NovStr.SstrNombreCampoBd,
                "0 AS " & ClsIdFactura_NovEnt.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd,
                ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd,
                " '' AS " & ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd,
                ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd,
                ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd,
                ClsIdTercero_NovAntDbl.SstrNombreCampoBd,
                ClsAliasCont_NovAntStr.SstrNombreCampoBd,
                "0 as Base", "SUM(" & ClsValor_NovAntDec.SstrNombreCampoBd & ")",
                ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd}
        Dim lstrCamposSelectT2 = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdTercero_NovAntDbl.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd, "ASC"},
                          {ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd, "ASC"},
                          {ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd & " = " & aenuDocumOrigen &
                " AND " & ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd & " <> " &
                EnuTipoNov.EnuRDbAntApl & " AND " &
                ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd & " <> " &
                EnuTipoNov.EnuDbAntApl & " AND " & ClsFechaNovedadAntDtm.SstrNombreCampoBd &
                " >= " & lstrFechaDesde & " AND " & ClsFechaNovedadAntDtm.SstrNombreCampoBd &
                " <= " & lstrFechaHasta & " AND " & ClsValor_NovDec.SstrNombreCampoBd & " > 0"
        Dim lstrCamposGrupo = {ClsFechaNovedadAntDtm.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd,
                ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd, ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd,
                ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd, ClsAliasCont_NovAntStr.SstrNombreCampoBd,
                "Base", ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim ldtbMovAntRei = ClsPanorama.FdtbDataTable(lstrNombreTablaT1, lstrCamposSelectT1,
                lstrNombreTablaT2, lstrCamposSelectT2, lstrCamposRelPri, lstrCamposRelSec,
                lstrIndice, False, lstrFiltro, lstrCamposGrupo)
        Return ldtbMovAntRei
    End Function
    Protected Function FdtbMoviNotaAjuste() As DataTable 'Ok
        Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaDesde.ToString) & "'"
        Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaHasta.ToString) & "'"
        Dim lstrNombreTablaT1 = ClsNovedadAnticipo.SstrNombreTabla
        Dim lstrNombreTablaT2 = ClsCliente.SstrNombreTabla
        Dim lstrCamposSelectT1 = {ClsIdTipoDocOrigen_NovAntByt.SstrNombreCampoBd,
                ClsFechaNovedadAntDtm.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd,
                ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd,
                ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd,
                ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd,
                ClsAliasCont_NovAntStr.SstrNombreCampoBd,
                ClsIdTercero_NovAntDbl.SstrNombreCampoBd,
                "'' AS IdPredioAgrupador", "0 AS " & ClsBaseDec.SstrNombreCampoBd,
                " SUM(" & ClsValor_NovAntDec.SstrNombreCampoBd & ")",
                ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd}
        Dim lstrCamposSelectT2 = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdTercero_NovAntDbl.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrIndice = {{"", ""}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsIdTipoDocOrigen_NovAntByt.SstrNombreCampoBd & " = " & EnuIdDocumentoDef.EnuNotaAjuste &
                " AND " & ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd & " = " & EnuTipoNov.EnuCrAntRec &
                " AND " & ClsFechaNovedadAntDtm.SstrNombreCampoBd & " >= " & lstrFechaDesde & " AND " &
                ClsFechaNovedadAntDtm.SstrNombreCampoBd & " <= " & lstrFechaHasta & " AND " &
                ClsValor_NovAntDec.SstrNombreCampoBd & " > 0"
        Dim lstrCamposGrupo = {ClsIdTipoDocOrigen_NovAntByt.SstrNombreCampoBd,
                ClsFechaNovedadAntDtm.SstrNombreCampoBd, ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd,
                ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd, ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd,
                ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd, ClsIdClienteDbl.SstrNombreCampoBd,
                ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTablaT1, lstrCamposSelectT1,
                lstrNombreTablaT2, lstrCamposSelectT2, lstrCamposRelPri, lstrCamposRelSec,
                lstrIndice, lstrFiltro, lstrCamposGrupo)
        Dim ldtbMoviNotaAjuste = ClsPanorama.FdtbDataTable(lstrSql)
        Return ldtbMoviNotaAjuste
    End Function
    Private Function FstrExpSqlRCNov() As String 'Ok
        Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaDesde.ToString) & "'"
        Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaHasta.ToString) & "'"
        Dim lstrNombreTablaT1 = ClsNovedad.SstrNombreTabla
        Dim lstrNombreTablaT2 = ClsCliente.SstrNombreTabla
        Dim lstrCamposSelectT1 = {ClsIdTipoDocOrigenByt.SstrNombreCampoBd,
                ClsFechaNovedadDtm.SstrNombreCampoBd,
                ClsPrefijoFact_NovStr.SstrNombreCampoBd,
                ClsIdFactura_NovEnt.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd,
                ClsIdDocOrigenEnt.SstrNombreCampoBd,
                ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd,
                ClsIdCuentaDb_NovStr.SstrNombreCampoBd, ClsIdCuentaCr_NovStr.SstrNombreCampoBd,
                ClsIdTercero_NovDbl.SstrNombreCampoBd, ClsAliasCont_NovStr.SstrNombreCampoBd,
                "SUM(" & ClsBaseDec.SstrNombreCampoBd & ")",
                "SUM(" & ClsValor_NovDec.SstrNombreCampoBd & ")",
                ClsIdTipoNovedadByt.SstrNombreCampoBd}
        Dim lstrCamposSelectT2 = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdTercero_NovDbl.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrIndice = {{"", ""}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " & EnuIdDocumentoDef.EnuReciboCaja &
                " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd & " >= " & lstrFechaDesde & " AND " &
                ClsFechaNovedadDtm.SstrNombreCampoBd & " <= " & lstrFechaHasta & " AND " &
                ClsValor_NovDec.SstrNombreCampoBd & " > 0"
        Dim lstrCamposGrupo = {ClsIdTipoDocOrigenByt.SstrNombreCampoBd,
                ClsFechaNovedadDtm.SstrNombreCampoBd, ClsPrefijoFact_NovStr.SstrNombreCampoBd,
                ClsIdFactura_NovEnt.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd,
                ClsIdDocOrigenEnt.SstrNombreCampoBd, ClsIdCuentaDb_NovStr.SstrNombreCampoBd,
                ClsIdCuentaCr_NovStr.SstrNombreCampoBd, ClsAliasCont_NovStr.SstrNombreCampoBd,
                ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTablaT1, lstrCamposSelectT1,
                lstrNombreTablaT2, lstrCamposSelectT2, lstrCamposRelPri, lstrCamposRelSec,
                lstrIndice, lstrFiltro, lstrCamposGrupo)
        Return lstrSql
    End Function
    Private Function FstrExpSqlRCNovAnt() As String
        Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaDesde.ToString) & "'"
        Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaHasta.ToString) & "'"
        Dim lstrNombreTablaT1 = ClsNovedadAnticipo.SstrNombreTabla
        Dim lstrNombreTablaT2 = ClsCliente.SstrNombreTabla
        Dim lstrCamposSelectT1 = {ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd,
                ClsFechaNovedadAntDtm.SstrNombreCampoBd,
                "'' AS " & ClsPrefijoFact_NovStr.SstrNombreCampoBd,
                "0 AS " & ClsIdFactura_NovEnt.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd,
                ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd,
                "'' AS " & ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd,
                ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd,
                ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd,
                ClsIdTercero_NovAntDbl.SstrNombreCampoBd,
                ClsAliasCont_NovAntStr.SstrNombreCampoBd,
                "0 AS " & ClsBaseDec.SstrNombreCampoBd,
                " SUM(" & ClsValor_NovAntDec.SstrNombreCampoBd & ")",
                ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd}
        Dim lstrCamposSelectT2 = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdTercero_NovAntDbl.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrIndice = {{"", ""}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " & EnuIdDocumentoDef.EnuReciboCaja &
                " AND (" & ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd & " = " & EnuTipoNov.EnuCrAntRec &
                " OR " & ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd & " = " & EnuTipoNov.EnuRCrAntRec &
                " OR " & ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd & " = " & EnuTipoNov.EnuDbAntDev &
                ") AND " & ClsFechaNovedadDtm.SstrNombreCampoBd & " >= " & lstrFechaDesde & " AND " &
                ClsFechaNovedadDtm.SstrNombreCampoBd & " <= " & lstrFechaHasta & " AND " &
                ClsValor_NovDec.SstrNombreCampoBd & " > 0"
        Dim lstrCamposGrupo = {ClsIdTipoDocOrigenByt.SstrNombreCampoBd,
                ClsFechaNovedadDtm.SstrNombreCampoBd, ClsPrefijoFact_NovStr.SstrNombreCampoBd,
                ClsIdFactura_NovEnt.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd,
                ClsIdDocOrigenEnt.SstrNombreCampoBd, ClsIdCuentaDb_NovStr.SstrNombreCampoBd,
                ClsIdCuentaCr_NovStr.SstrNombreCampoBd, ClsAliasCont_NovAntStr.SstrNombreCampoBd,
                ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTablaT1, lstrCamposSelectT1,
                lstrNombreTablaT2, lstrCamposSelectT2, lstrCamposRelPri, lstrCamposRelSec,
                lstrIndice, lstrFiltro, lstrCamposGrupo)
        Return lstrSql
    End Function
#End Region
#Region "Exportar .csv a .xlsx"
    Protected Shared Sub SExporteAExcel(astrNombreArchivo As String)
        If FblnExisteExcel() Then
            ' Elimina archivo si existe
            If My.Computer.FileSystem.FileExists(astrNombreArchivo & ".xlsx") Then
                My.Computer.FileSystem.DeleteFile(astrNombreArchivo & ".xlsx")
            End If
            ' Crear una nueva aplicación de Excel
            Dim lappExcel As New Microsoft.Office.Interop.Excel.Application
            ' Abrir el archivo .csv
            Dim lwbInterfaz As Microsoft.Office.Interop.Excel.Workbook =
                lappExcel.Workbooks.Open(astrNombreArchivo & ".csv")
            ' Guardar el archivo como .xlsx
            lwbInterfaz.SaveAs(astrNombreArchivo & ".xlsx",
                Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
            ' Cerrar el libro de trabajo
            lwbInterfaz.Close()
            ' Cerrar la aplicación de Excel
            lappExcel.Quit()
            ' Liberar los objetos COM
            ReleaseObject(lwbInterfaz)
            ReleaseObject(lappExcel)
            If My.Computer.FileSystem.FileExists(astrNombreArchivo & ".csv") Then
                My.Computer.FileSystem.DeleteFile(astrNombreArchivo & ".csv")
            End If
        Else
            Dim lblnNoUsado = MsgBox("No es posible generar la interfaz debido a que Excel no está instalado!",
                    vbOKOnly, "Sin Excel")
        End If
    End Sub
    Private Shared Sub ReleaseObject(obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
#End Region
End Class
