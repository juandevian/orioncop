Imports System.Drawing
Imports System.Windows.Media.Imaging
Imports ThoughtWorks.QRCode.Codec
Friend Class ClsNotaReversionCr
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriNotasReversaCr"
    ' Variables de modulo
    Private MobjPredioAgrNRevCr As ClsPredio = Nothing
    Private MobjClienteNota As ClsCliente = Nothing
    Private MobjDocReversado As ClsCBObjetoPan = Nothing
    Private MdtbNovedades As DataTable = Nothing
    Private McolNovedades As Collection = Nothing
    Private MdtbNovedadesAnt As DataTable = Nothing
    Private McolNovedadesAnt As Collection = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Nota Reversión Crédito en modo único
    ''' </summary>
    Public Sub New()
        HobjPadre = Nothing
        HblnEsCreable = False
        HblnEsModificable = False
        HblnEsAnulable = False
        HblnEsSuprimible = False
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add({"*"})
    End Sub
    ''' <summary>
    ''' Instancia un objeto Nota Crédito en modo navegable
    ''' </summary>
    Public Sub New(astrPref As String)
        HobjPadre = Nothing
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion
        lstrFiltro &= " AND " & ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd & " = '" & astrPref & "'"
        HcolFiltros.Add(lstrFiltro)
        Dim lstrCamposSelect = {StrCampoCarpeta,
            StrCampoCentroUtil, ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd,
            ClsIdNotaReversaCrEnt.SstrNombreCampoBd}
        HblnEsSuprimible = False
        HblnEsModificable = False
        HblnEsAnulable = False
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwNotaRecersaCr">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As Object, adrwNotaRecersaCr As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        HblnEsSuprimible = False
        HblnEsModificable = False
        HblnEsCreable = False
        '
        DrwRegistroActual = adrwNotaRecersaCr
        DtbTablaColeccion = DrwRegistroActual.Table
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
            Return EnuIdClasesPanDef.EnuNotaReversaCr
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Nota Reversión Crédito"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & "Nro. " & StrNumeroNotaReversaCr & Chr(34)
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjCUDEStr As New ClsCUDEStr(Me)
    Friend ReadOnly Property ObjCUDocStr As New ClsCUDocStr(Me)
    Friend ReadOnly Property ObjDetalle_NotaReversaCrStr As New ClsDetalle_NotaReversaCrStr(Me)
    Friend ReadOnly Property ObjFecha_NotaReversaCrDtm As New ClsFecha_NotaReversaCrDtm(Me)
    Friend ReadOnly Property ObjIdCarpeta_NotaReversaCrShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_NotaReversaCrShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdCliente_NotaReversaCrDbl As New ClsIdCliente_NotaReversaCrDbl(Me)
    Friend ReadOnly Property ObjIdEstadoEDocEnt As New ClsIdEstadoEDocEnt(Me)
    Friend ReadOnly Property ObjIdNotaReversaCrEnt As New ClsIdNotaReversaCrEnt(Me)
    Friend ReadOnly Property ObjIdPredioAgrupador_NotaReversaCrStr As New ClsIdPredioAgr_NotaReversaCrStr(Me)
    Friend ReadOnly Property ObjIdDoc_NotaReversaCrEnt As New ClsIdDoc_NotaReversaCrEnt(Me)
    Friend ReadOnly Property ObjIdUsuario_NotaReversaCrStr As New ClsIdUsuarioStr(Me)
    Friend ReadOnly Property ObjPrefijo_NotaReversaCrStr As New ClsPrefijo_NotaReversaCrStr(Me)
    Friend ReadOnly Property ObjPrefijoDoc_NotaReversaCrStr As New ClsPrefijoDoc_NotaReversaCrStr(Me)
    Friend ReadOnly Property ObjTipoDocReversadoByt As New ClsTipoDocReversadoByt(Me)
    Friend ReadOnly Property ObjValor_NotaReversaCrDec As New ClsValor_NotaReversaCrDec(Me)
    Friend ReadOnly Property ObjVerEFacEnt As New ClsVerEFacEnt(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjOrigenInstanciaStr)
                HcolPropiedades.Add(ObjCUDocStr)
                HcolPropiedades.Add(ObjCUDEStr)
                HcolPropiedades.Add(ObjDetalle_NotaReversaCrStr)
                HcolPropiedades.Add(ObjPrefijoDoc_NotaReversaCrStr)
                HcolPropiedades.Add(ObjTipoDocReversadoByt)
                HcolPropiedades.Add(ObjFecha_NotaReversaCrDtm)
                HcolPropiedades.Add(ObjIdCarpeta_NotaReversaCrShr)
                HcolPropiedades.Add(ObjIdCentroUtil_NotaReversaCrShr)
                HcolPropiedades.Add(ObjIdCliente_NotaReversaCrDbl)
                HcolPropiedades.Add(ObjIdEstadoEDocEnt)
                HcolPropiedades.Add(ObjIdNotaReversaCrEnt)
                HcolPropiedades.Add(ObjIdPredioAgrupador_NotaReversaCrStr)
                HcolPropiedades.Add(ObjIdDoc_NotaReversaCrEnt)
                HcolPropiedades.Add(ObjIdUsuario_NotaReversaCrStr)
                HcolPropiedades.Add(ObjPrefijo_NotaReversaCrStr)
                HcolPropiedades.Add(ObjValor_NotaReversaCrDec)
                HcolPropiedades.Add(ObjVerEFacEnt)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    ''' <summary>
    ''' Devuelve un string compuesto por el prefijo de la nota y el id de la nota separados por un
    ''' guion. Si no existe el prefijo devuelve solo el id de la nota
    ''' </summary>
    ''' <value></value>
    Friend ReadOnly Property StrNumeroNotaReversaCr As String
        Get
            Dim lstrNumeroNotaReversaCr As String = ClsPanorama.FstrNumeroDcto(ObjPrefijo_NotaReversaCrStr.ObjValorPro,
                    ObjIdNotaReversaCrEnt.ObjValorPro)
            Return lstrNumeroNotaReversaCr
        End Get
    End Property
    ''' <summary>
    ''' Devuelve un string compuesto por el prefijo del recibo de caja reversado y el 
    ''' id del recibo de caja reversado separados por un guion. 
    ''' Si no existe el prefijo devuelve solo el id del recibo de caja reversado
    ''' </summary>
    ''' <value></value>
    Friend ReadOnly Property StrNumeroDocRev As String
        Get
            Dim lstrNumeroDocRev As String = ClsPanorama.FstrNumeroDcto(
                    ObjPrefijoDoc_NotaReversaCrStr.ObjValorPro,
                    ObjIdDoc_NotaReversaCrEnt.ObjValorPro)
            Return lstrNumeroDocRev
        End Get
    End Property
    Friend Property ObjClienteNota As ClsCliente
        Get
            If MobjClienteNota Is Nothing Then
                Dim ldblIdCliente As Double = ObjIdCliente_NotaReversaCrDbl.ObjValorPro
                Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente}
                MobjClienteNota = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
                MobjClienteNota.SAbra(lobjValorLlave)
            End If
            Return MobjClienteNota
        End Get
        Set(value As ClsCliente)
            MobjClienteNota = value
            HobjPadre = MobjClienteNota
        End Set
    End Property
    Friend ReadOnly Property ObjPredioAgrNRevCr As ClsPredio
        Get
            If IsNothing(MobjPredioAgrNRevCr) Then
                If Not String.IsNullOrEmpty(ObjIdPredioAgrupador_NotaReversaCrStr.ToString) AndAlso
                        ObjIdPredioAgrupador_NotaReversaCrStr.ToString <> GCSTRSINPA Then
                    MobjPredioAgrNRevCr = New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil,
                            ObjIdPredioAgrupador_NotaReversaCrStr.ObjValorPro}
                    MobjPredioAgrNRevCr.SAbra(lobjValorLlave)
                    If Not MobjPredioAgrNRevCr.BlnExiste Then
                        MobjPredioAgrNRevCr = Nothing
                    End If
                End If
            End If
            Return MobjPredioAgrNRevCr
        End Get
    End Property
    Friend Property ObjDocReversado As ClsCBObjetoPan
        Get
            If MobjDocReversado Is Nothing Then
                If ObjTipoDocReversadoByt.BlnEsValido AndAlso
                        ObjPrefijoDoc_NotaReversaCrStr.BlnEsValido AndAlso
                        ObjIdDoc_NotaReversaCrEnt.BlnEsValido Then
                    Dim lstrPref = ObjPrefijoDoc_NotaReversaCrStr.ToString
                    Dim lentIdDoc = ObjIdDoc_NotaReversaCrEnt.ObjValorPro
                    Dim lobjvalorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdDoc}
                    If ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuReciboC Then
                        MobjDocReversado = New ClsReciboCaja()
                        MobjDocReversado.SAbra(lobjvalorLlave)
                    Else
                        MobjDocReversado = New ClsNotaCr()
                        MobjDocReversado.SAbra(lobjvalorLlave)
                    End If
                End If
            End If
            Return MobjDocReversado
        End Get
        Set(value As ClsCBObjetoPan)
            MobjDocReversado = value
        End Set
    End Property
    Friend ReadOnly Property DtmFechaDocRev As Date
        Get
            Dim ldtmFecha = GCDTMFECHANULA
            If ObjTipoDocReversadoByt.BlnEsValido AndAlso Not IsNothing(ObjDocReversado) Then
                If ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuReciboC Then
                    Dim lobjRecCaj As ClsReciboCaja = ObjDocReversado
                    ldtmFecha = lobjRecCaj.ObjFechaRecDtm.ObjValorPro
                Else
                    Dim lobjNotaCr As ClsNotaCr = ObjDocReversado
                    ldtmFecha = lobjNotaCr.ObjFecha_NotaCrDtm.ObjValorPro
                End If
            End If
            Return ldtmFecha
        End Get
    End Property
    Friend ReadOnly Property ObjFacturaAfectada As ClsFactura
        Get
            If BlnAfectaFrasRegEFac() Then
                Dim lobjNotaCr As ClsNotaCr = ObjDocReversado
                Return lobjNotaCr.ObjFacturaAfectada
            Else
                Return Nothing
            End If
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        MdtbNovedades = Nothing
        MobjClienteNota = Nothing
        McolNovedades = Nothing
        MobjDocReversado = Nothing
        MdtbNovedadesAnt = Nothing
        McolNovedadesAnt = Nothing
    End Sub
    Protected Overrides Sub SInicialiceObj()
        ObjIdCarpeta_NotaReversaCrShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtil_NotaReversaCrShr.ObjValorPro = GshrIdCentroUtil
        ObjIdUsuario_NotaReversaCrStr.ObjValorPro = GstrIdUsuario
        ObjIdDoc_NotaReversaCrEnt.ObjValorPro = 0
        ObjPrefijoDoc_NotaReversaCrStr.ObjValorPro = String.Empty
        ObjOrigenInstanciaAnuloStr.ObjValorPro = String.Empty
        ObjOrigenInstanciaStr.ObjValorPro = GstrOrigenActual
        ObjIdUsuarioAnuloStr.ObjValorPro = String.Empty
        Dim lstrPrefijo = GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuNotaRevCr)
        If IsNothing(lstrPrefijo) Then lstrPrefijo = String.Empty
        ObjPrefijo_NotaReversaCrStr.ObjValorPro = lstrPrefijo
        ObjCUDEStr.ObjValorPro = String.Empty
        ObjCUDocStr.ObjValorPro = String.Empty
        ObjVerEFacEnt.ObjValorPro = EnuVerEFac.EnuNinguna
        ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoEDoc
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        GobjPanDat.SControleProcesoObj(True)
        Try
            GobjPanDat.SInicialiceTransaccion()
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                SNumereObj()
                If ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuReciboC Then
                    Dim lobjRecCaj As ClsReciboCaja = ObjDocReversado
                    ObjValor_NotaReversaCrDec.ObjValorPro = lobjRecCaj.ObjValor_RecDec.ObjValorPro
                ElseIf ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuNotaCr Then
                    Dim lobNotaCr As ClsNotaCr = ObjDocReversado
                    ObjValor_NotaReversaCrDec.ObjValorPro = lobNotaCr.ObjValor_NotaCrDec.ObjValorPro
                Else
                    Throw New ErrorInesperadoPanLException("Tipo Documento Reversado no valido")
                End If
                SComplementeEFac()
                ObjFechaCreacionDtm.ObjValorPro = Now
                Dim lstrMens = String.Empty
                SReverseCr(lstrMens)
                If Not String.IsNullOrEmpty(lstrMens) Then
                    Throw New ErrorInesperadoPanLException(lstrMens)
                End If
            End If
            ' McolNovedades es Null cuando el recibo de caja no afecto factura alguna; solo fue un anticipo.
            If Not IsNothing(ColNovedades) Then
                ClsPanorama.SActualiceCol(McolNovedades)
            End If
            If Not IsNothing(ColNovedadesAnt) Then
                ClsPanorama.SActualiceCol(McolNovedadesAnt)
            End If
            MyBase.SActualice(ablnExigeRequeridos)
            lblnNoHayError = True
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return StrNumeroNotaReversaCr
        End Get
    End Property
#End Region
#Region "Manejo de Novedades"
    Friend ReadOnly Property ColNovedades As Collection
        Get
            If McolNovedades Is Nothing OrElse McolNovedades.Count = 0 Then
                McolNovedades = New Collection
                Dim lobjNovedad As ClsNovedad
                SCargueDtbNovedades()
                If MdtbNovedades.Rows.Count Then
                    For Each ldrwNov As DataRow In MdtbNovedades.Rows
                        lobjNovedad = New ClsNovedad(Me, ldrwNov)
                        lobjNovedad.SLeaValores(True)
                        McolNovedades.Add(lobjNovedad)
                    Next
                End If
            End If
            Return McolNovedades
        End Get
    End Property
    Friend ReadOnly Property ColNovedadesAnt As Collection
        Get
            If IsNothing(McolNovedadesAnt) Then
                McolNovedadesAnt = New Collection
                Dim lobjNovedadAnt As ClsNovedadAnticipo
                Dim lobjAnticipo As New ClsAnticipo(EnuModoInstanciaObjDef.enuUnico)
                Dim lentIdAnticipo As Integer
                SCargueDtbNovedadesAnt()
                If MdtbNovedadesAnt.Rows.Count Then
                    For Each ldrwNov As DataRow In MdtbNovedadesAnt.Rows
                        lentIdAnticipo = ClsPanorama.FobjValorCampo(ldrwNov(ClsIdAnticipo_NovEnt.SstrNombreCampoBd),
                                EnuTipoValor.enuInteger)
                        lobjAnticipo.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lentIdAnticipo})
                        lobjNovedadAnt = New ClsNovedadAnticipo(lobjAnticipo, ldrwNov)
                        lobjNovedadAnt.SLeaValores(True)
                        McolNovedadesAnt.Add(lobjNovedadAnt)
                    Next
                End If
            End If
            Return McolNovedadesAnt
        End Get
    End Property
    ''' <summary>
    ''' Devuelve un datatable que contiene información de las novedades y las novedades de anticipo
    ''' para ser visualizada en la IU.
    ''' </summary>
    Friend ReadOnly Property DtbNovedadesIU As DataTable
        Get
            GobjPanDat.SControleProcesoObj(True)
            Dim ldtbNovedadesIU = FdtbNovedadesIU()
            SComplementeTablaNov(ldtbNovedadesIU)
            GobjPanDat.SControleProcesoObj(False)
            Return ldtbNovedadesIU
        End Get
    End Property
    Private Function FdtbNovedadesIU()
        Dim ldtbNovIU As DataTable = Nothing
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            Dim lstrExpSqlNovedades = FstrExpSqlNovedades()
            Dim lstrExpSqlNovedadesAnt = FstrExpSqlNovedadesAnt()
            Dim lstrExpSql = (lstrExpSqlNovedades) & " UNION ALL " & (lstrExpSqlNovedadesAnt)
            ldtbNovIU = ClsPanorama.FdtbDataTable(lstrExpSql)
        End If
        Return ldtbNovIU
    End Function
    Private Sub SCargueDtbNovedades()
        If MdtbNovedades Is Nothing OrElse MdtbNovedades.Rows.Count = 0 Then
            Dim lstrPrefNotaRevCr = ObjPrefijo_NotaReversaCrStr.ToString
            Dim lstrIdNotaRevCr = ObjIdNotaReversaCrEnt.ToString
            If String.IsNullOrEmpty(lstrIdNotaRevCr) Then lstrIdNotaRevCr = "0"
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {ClsIdTipoDocOrigenByt.SstrNombreCampoBd, "ASC"},
                              {ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdDocOrigenEnt.SstrNombreCampoBd, "ASC"},
                              {ClsPrefijoFact_NovStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdFactura_NovEnt.SstrNombreCampoBd, "ASC"},
                              {ClsIdNovedadShr.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdTipoDocOrigenByt.SstrNombreCampoBd &
                    " = " & EnuTipoDocOri.EnuNotaRevCr & " AND " &
                    ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" & lstrPrefNotaRevCr & "' AND " &
                    ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " & lstrIdNotaRevCr & " AND " &
                    ClsValor_NovDec.SstrNombreCampoBd & " <> 0"
            Dim lstrCamposSelect() = {"*"}
            MdtbNovedades = ClsPanorama.FdtbDataTable(ClsNovedad.SstrNombreTabla, lstrCamposSelect, lstrIndice,
                    lstrFiltro)
        End If
    End Sub
    Private Sub SCargueDtbNovedadesAnt()
        If IsNothing(MdtbNovedadesAnt) Then
            Dim lstrPrefNotaRevCr = ObjPrefijo_NotaReversaCrStr.ToString
            Dim lstrIdNotaRevCr = ObjIdNotaReversaCrEnt.ToString
            If String.IsNullOrEmpty(lstrIdNotaRevCr) Then lstrIdNotaRevCr = "0"
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd, "ASC"},
                              {ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd, "ASC"},
                              {ClsIdNovedadAntShr.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd &
                    " = " & EnuTipoDocOri.EnuNotaRevCr & " AND " &
                    ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd & " = '" & lstrPrefNotaRevCr & "' AND " &
                    ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd & " = " & lstrIdNotaRevCr & " AND " &
                    ClsValor_NovAntDec.SstrNombreCampoBd & " <> 0"
            Dim lstrCamposSelect() = {"*"}
            MdtbNovedadesAnt = ClsPanorama.FdtbDataTable(ClsNovedadAnticipo.SstrNombreTabla, lstrCamposSelect, lstrIndice,
                    lstrFiltro)
        End If
    End Sub
    Private Function FstrExpSqlNovedades() As String
        Dim lstrIdNotaRevCr = 0
        Dim lstrPrefNotaRevCr = ObjPrefijo_NotaReversaCrStr.ToString
        If Not IsNothing(ObjIdNotaReversaCrEnt.ObjValorPro) Then
            lstrIdNotaRevCr = ObjIdNotaReversaCrEnt.ToString
        End If
        Dim lstrNombreTabla = ClsNovedad.SstrNombreTabla
        Dim lstrCamposSelect As String() = {ClsAnuladoBln.SstrNombreCampoBd,
                                            ClsFechaCreacionDtm.SstrNombreCampoBd,
                                            ClsIdCuentaDb_NovStr.SstrNombreCampoBd,
                                            ClsIdCuentaCr_NovStr.SstrNombreCampoBd,
                                            ClsPrefijoFact_NovStr.SstrNombreCampoBd,
                                            ClsIdFactura_NovEnt.SstrNombreCampoBd,
                                            ClsValor_NovDec.SstrNombreCampoBd,
                                            ClsIdTipoNovedadByt.SstrNombreCampoBd,
                                            "'' AS NroFac", "'' AS Detalle"}
        Dim lstrIndice = {{"", ""}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdTipoDocOrigenByt.SstrNombreCampoBd &
                       " = " & EnuTipoDocOri.EnuNotaRevCr & " AND " &
                       ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" & lstrPrefNotaRevCr & "' AND " &
                       ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " & lstrIdNotaRevCr
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelect, lstrIndice,
                lstrFiltro, Array.Empty(Of String))
        Return lstrExpSql
    End Function
    Private Function FstrExpSqlNovedadesAnt() As String
        Dim lstrPrefNotaRevCr = ObjPrefijo_NotaReversaCrStr.ToString
        Dim lstrIdNotaRevCr = 0
        If Not IsNothing(ObjIdNotaReversaCrEnt.ObjValorPro) Then
            lstrIdNotaRevCr = ObjIdNotaReversaCrEnt.ToString
        End If
        Dim lstrNombreTabla = ClsNovedadAnticipo.SstrNombreTabla
        Dim lstrCamposSelect As String() = {ClsAnuladoBln.SstrNombreCampoBd,
                                            ClsFechaCreacionDtm.SstrNombreCampoBd,
                                            ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd,
                                            ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd,
                                            "'' AS " & ClsPrefijoFact_NovStr.SstrNombreCampoBd,
                                            "0 AS " & ClsIdFactura_NovEnt.SstrNombreCampoBd,
                                            ClsValor_NovAntDec.SstrNombreCampoBd,
                                            ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd,
                                            "'' AS NroFac", "'' AS Detalle"}
        Dim lstrIndice = {{"", ""}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd &
                       " = " & EnuTipoDocOri.EnuNotaRevCr & " AND " &
                       ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd & " = '" & lstrPrefNotaRevCr & "' AND " &
                       ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd & " = " & lstrIdNotaRevCr
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelect, lstrIndice,
                lstrFiltro, Array.Empty(Of String))
        Return lstrExpSql
    End Function
    Private Shared Sub SComplementeTablaNov(adtbNovedades As DataTable)
        Dim ldrwNovedades = adtbNovedades.Select
        Dim lstrConceptoNovedad As String
        For Each ldrwNovedad As DataRow In ldrwNovedades
            Dim lstrPrefFac As String = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsPrefijoFact_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            Dim lentIdFac As Integer = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdFactura_NovEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            Dim lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPrefFac, lentIdFac)
            Dim lenuTipoNovedad As EnuTipoNov =
                    ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdTipoNovedadByt.SstrNombreCampoBd),
                    EnuTipoValor.enuByte)
            Select Case lenuTipoNovedad
                Case EnuTipoNov.EnuRCrAnApCap
                    lstrConceptoNovedad = "Reversión Anticipo aplicado a Capital"
                Case EnuTipoNov.EnuRCrAnApInt
                    lstrConceptoNovedad = "Reversión Anticipo aplicado a Int. de Mora"
                Case EnuTipoNov.EnuRCrAntRec
                    lstrConceptoNovedad = "Reversión Anticipo recibido"
                Case EnuTipoNov.EnuRCrDctoCap
                    lstrConceptoNovedad = "Reversión Descuento a Capital"
                Case EnuTipoNov.EnuRCrIvaGas
                    lstrConceptoNovedad = "Reversión Iva llevado al Gasto"
                Case EnuTipoNov.EnuRCrDctoInt
                    lstrConceptoNovedad = "Reversión Descuento a Int. de Mora"
                Case EnuTipoNov.EnuRCrPagoCap
                    lstrConceptoNovedad = "Reversión Abono a Capital"
                Case EnuTipoNov.EnuRCrPagoInt
                    lstrConceptoNovedad = "Reversión Abono a Int. de Mora"
                Case EnuTipoNov.EnuRCrRetCre
                    lstrConceptoNovedad = "Reversión ReteCree aplicado"
                Case EnuTipoNov.EnuRCrRetFte
                    lstrConceptoNovedad = "Reversión Retefuente aplicado"
                Case EnuTipoNov.EnuRCrRetIca
                    lstrConceptoNovedad = "Reversión ReteIca aplicado"
                Case EnuTipoNov.EnuRCrRetIva
                    lstrConceptoNovedad = "Reversión ReteIva aplicado"
                Case EnuTipoNov.EnuRDbCap
                    lstrConceptoNovedad = "Reversión Debito a Deuda de Capital"
                Case EnuTipoNov.EnuRDbInt
                    lstrConceptoNovedad = "Reversión Debito a Int. de Mora"
                Case EnuTipoNov.EnuRDbIva
                    lstrConceptoNovedad = "Reversión a Debito a Iva Generado."
                Case EnuTipoNov.EnuRDbIvaInt
                    lstrConceptoNovedad = "Reversión a Debito a Iva Generado Int. Mora"
                Case EnuTipoNov.EnuRDbAntApl
                    lstrConceptoNovedad = "Reversión Anticipo aplicado."
                Case EnuTipoNov.EnuRDbAntDev
                    lstrConceptoNovedad = "Reversión Anticipo Reintegrado"
                Case Else
                    Throw New ErrorInesperadoPanLException("Tipo de Novedad inconsistente!")
            End Select
            ldrwNovedad("NroFac") = lstrNroFac
            ldrwNovedad("Detalle") = lstrConceptoNovedad
        Next
    End Sub
#End Region
#Region "eFac"
    ''' <summary>
    ''' Indica si es una factura electrónica y si esta registrada 
    ''' ''' </summary>
    Friend ReadOnly Property BlnEstaRegEFac As Boolean
        Get
            Return BlnEsDocEle AndAlso ObjIdEstadoEDocEnt.ObjValorPro >= EnuEstadoEDoc.EnuRegi
        End Get
    End Property
    Friend ReadOnly Property BlnEsDocEle As Boolean
        Get
            Dim lblnEsDocEle = GobjParametros.BlnEFacAutorizado AndAlso
                    ObjIdEstadoEDocEnt.ObjValorPro <> EnuEstadoEDoc.EnuNoEDoc
            Return lblnEsDocEle
        End Get
    End Property
    Public Function BlnAfectaFrasRegEFac() As Boolean
        Dim lblnAfecta = False
        If ObjTipoDocReversadoByt.BlnEsValido AndAlso Not IsNothing(ObjDocReversado) Then
            If ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuNotaCr Then
                Dim lobjNotaCr As ClsNotaCr = ObjDocReversado
                lblnAfecta = lobjNotaCr.BlnAfectaFrasRegEFac()
            End If
        End If
        Return lblnAfecta
    End Function
    Friend Function FblnInsertarEFac() As Boolean
        Dim lblnReg = False
        If BlnEsDocEle AndAlso BlnAfectaFrasRegEFac() Then
            lblnReg = (ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoReg)
            If lblnReg Then
                lblnReg = ObjFacturaAfectada.BlnEsFacEle AndAlso
                        (Not ObjFacturaAfectada.FenuVerFacEFac = EnuVerEFac.EnuV1) AndAlso
                        ObjFacturaAfectada.ObjIdEstadoEDocEnt.ObjValorPro >=
                        EnuEstadoEDoc.EnuRegi AndAlso
                        ObjFacturaAfectada.ObjIdEstadoEDocEnt.ObjValorPro <>
                        EnuEstadoEDoc.EnuNoEDoc
            End If
        End If
        Return lblnReg
    End Function
    Friend Function FblnActualizarEstEFac() As Boolean
        Dim lblnActu = False
        If BlnEsDocEle AndAlso BlnAfectaFrasRegEFac() Then
            lblnActu = (ObjIdEstadoEDocEnt.ObjValorPro < EnuEstadoEDoc.EnuEnviada) AndAlso
                    (ObjIdEstadoEDocEnt.ObjValorPro >= EnuEstadoEDoc.EnuEnProceso)
        End If
        Return lblnActu
    End Function
    Friend Function FblnEnviarEFac() As Boolean
        Dim lblnEnv = False
        If BlnEsDocEle AndAlso BlnAfectaFrasRegEFac() Then
            lblnEnv = (ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuRegi)
        End If
        Return lblnEnv
    End Function
    Friend Sub SHabiliteProcesarEFac()
        If BlnEsDocEle Then
            If ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuInvalida OrElse
                    ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuErrorFtp Then
                If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                End If
                ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoReg
                SActualice(True)
            End If
        End If
    End Sub
    Friend Sub SPrepareParaReprocesarEfac()
        If GobjParametros.ObjAutorizaEFacBln.ObjValorPro AndAlso ObjIdEstadoEDocEnt.ObjValorPro <>
                EnuEstadoEDoc.EnuNoReg AndAlso ObjIdEstadoEDocEnt.ObjValorPro <>
                EnuEstadoEDoc.EnuNoEDoc AndAlso ObjIdEstadoEDocEnt.ObjValorPro <>
                EnuEstadoEDoc.EnuInvalida Then
            If Not String.IsNullOrEmpty(ObjCUDocStr.ToString) Then
                EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuErrorFtp
                ObjCUDEStr.ObjValorPro = ""
                ObjCUDocStr.ObjValorPro = ""
                SActualice(True)
            End If
        End If
    End Sub
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lentIdNotaRevCr As Integer
            Dim lstrPrefijo = GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuNotaRevCr)
            If IsNothing(lstrPrefijo) Then lstrPrefijo = String.Empty
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd & " = '" & lstrPrefijo & "'"
            lentIdNotaRevCr = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                    ClsIdNotaReversaCrEnt.SstrNombreCampoBd, ObjIdNotaReversaCrEnt.EnuTipoValor,
                    lstrFiltro)
            lentIdNotaRevCr += 1
            ObjPrefijo_NotaReversaCrStr.ObjValorPro = lstrPrefijo
            ObjIdNotaReversaCrEnt.ObjValorPro = lentIdNotaRevCr
        End If
    End Sub
    Private Sub SComplementeEFac()
        If GobjParametros.ObjAutorizaEFacBln.ObjValorPro Then
            If BlnAfectaFrasRegEFac() Then
                ObjVerEFacEnt.ObjValorPro = EnuVerEFac.EnuV2
                ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoReg
            End If
        End If
    End Sub
    Private Sub SReverseCr(ByRef astrMens As String)
        If ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuReciboC Then
            Dim lobjDocRev As ClsReciboCaja = ObjDocReversado
            If ClsOrionCop.FblnDocReversable(lobjDocRev.ObjPrefijo_RecStr.ObjValorPro,
                    lobjDocRev.ObjIdRecCajaEnt.ObjValorPro, ObjTipoDocReversadoByt.ObjValorPro,
                    astrMens) Then
                If Not IsNothing(lobjDocRev.ObjAnticipo) Then
                    lobjDocRev.ObjAnticipo.SReverseAnticipo(EnuTipoDocOri.EnuNotaRevCr,
                        ObjPrefijo_NotaReversaCrStr.ObjValorPro, ObjIdNotaReversaCrEnt.ObjValorPro,
                        ObjFecha_NotaReversaCrDtm.ObjValorPro)
                End If
                SReverseNov(lobjDocRev.ColNovedades)
            End If
        ElseIf ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuNotaCr Then
            Dim lobjDocRev As ClsNotaCr = ObjDocReversado
            If ClsOrionCop.FblnDocReversable(lobjDocRev.ObjPrefijo_NotaCrStr.ObjValorPro,
                    lobjDocRev.ObjIdNotaCrEnt.ObjValorPro, ObjTipoDocReversadoByt.ObjValorPro,
                    astrMens) Then
                SReverseNov(lobjDocRev.ColNovedades)
            End If
        End If
    End Sub
    Private Sub SReverseNov(acolNovedades As Collection)
        Dim lstrPrefFac As String, lentIdFact As Integer
        Dim lobjFactura As New ClsFactura()
        For Each lobjNov As ClsNovedad In acolNovedades
            lstrPrefFac = lobjNov.ObjPrefijoFact_NovStr.ObjValorPro
            lentIdFact = lobjNov.ObjIdFactura_NovEnt.ObjValorPro
            lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFac, lentIdFact})
            lobjFactura.SReverseNovedad(lobjNov, EnuTipoDocOri.EnuNotaRevCr,
                                        ObjFecha_NotaReversaCrDtm.ObjValorPro,
                                        ObjPrefijo_NotaReversaCrStr.ObjValorPro,
                                        ObjIdNotaReversaCrEnt.ObjValorPro)
            lobjFactura.SActualice(True)
        Next
    End Sub
    ''' <summary>
    ''' Devuelve un array de strings cada uno de los cuales contiene le prefijo, el número 
    ''' y el item de factura afectada, asi como el tipo y el valor del descuento sin tener 
    ''' en cuenta las retenciones
    ''' </summary>
    ''' <returns></returns>
    Friend Function FstrDsctosReversados() As String()
        Dim lstrDsctos As String() = Array.Empty(Of String), i = -1
        Dim lstrPrefFac As String, lstrIdFac As String
        Dim lstrIdItemFac As String, lstrTipoNov As String, lstrValor As String
        Dim lenuTipoNov As EnuTipoNov
        Dim lenuTipoDscto As EnuTipoDescuentoDef = EnuTipoDescuentoDef.None
        If BlnAfectaFrasRegEFac() Then
            For Each lobjNov As ClsNovedad In ColNovedades
                lenuTipoNov = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
                If lenuTipoNov = EnuTipoNov.EnuRCrDctoCap OrElse
                        lenuTipoNov = EnuTipoNov.EnuRCrDctoInt OrElse
                        lenuTipoNov = EnuTipoNov.EnuRCrIvaGas Then
                    If lenuTipoNov = EnuTipoNov.EnuRCrDctoCap OrElse
                        lenuTipoNov = EnuTipoNov.EnuRCrIvaGas Then
                        lenuTipoDscto = EnuTipoDescuentoDef.EnuDsctoCapital
                    ElseIf lenuTipoNov = EnuTipoNov.EnuRCrDctoInt Then
                        lenuTipoDscto = EnuTipoDescuentoDef.EnuDsctoIntMora
                    End If
                    i += 1
                    lstrPrefFac = lobjNov.ObjPrefijoFact_NovStr.ObjValorPro
                    lstrIdFac = lobjNov.ObjIdFactura_NovEnt.ToString
                    lstrIdItemFac = lobjNov.ObjIdItemFact_NovShr.ToString
                    lstrTipoNov = CByte(lenuTipoDscto).ToString
                    lstrValor = Format(lobjNov.ObjValor_NovDec.ObjValorPro, "#0.00")
                    Dim lstrDscto As String = lstrPrefFac & "&" & lstrIdFac & "&" & lstrIdItemFac &
                        "&" & lstrTipoNov & "&" & lstrValor
                    ReDim Preserve lstrDsctos(i)
                    lstrDsctos(i) = lstrDscto
                End If
            Next
        End If
        Return lstrDsctos
    End Function
    ''' <summary>
    ''' Devuelve el valor total de descuentos sin incluir las retenciones
    ''' </summary>
    ''' <returns></returns>
    Friend Function FdecDsctosNota() As Decimal
        Dim ldecValor = 0D
        Dim lenuTipoNov As EnuTipoNov
        For Each lobjNov As ClsNovedad In ColNovedades
            lenuTipoNov = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
            If lenuTipoNov = EnuTipoNov.EnuRCrDctoCap OrElse
                    lenuTipoNov = EnuTipoNov.EnuRCrDctoInt OrElse
                    lenuTipoNov = EnuTipoNov.EnuRCrIvaGas Then
                ldecValor += lobjNov.ObjValor_NovDec.ObjValorPro
            End If
        Next
        Return ldecValor
    End Function
    Friend Function FdecDsctoRevItemFac(astrIdItemFac As String)
        Dim ldecDscto = 0D
        For Each lobjNov As ClsNovedad In ColNovedades
            If lobjNov.ObjIdItemFact_NovShr.ToString = astrIdItemFac Then
                If lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuRCrDctoCap OrElse
                        lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuRCrDctoInt OrElse
                        lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuRCrIvaGas Then
                    ldecDscto += lobjNov.ObjValor_NovDec.ObjValorPro
                End If
            End If
        Next
        Return ldecDscto
    End Function
    Friend Function FstrAliasCon() As String
        Dim lstrAliasCon = String.Empty
        If ObjPredioAgrNRevCr IsNot Nothing Then
            lstrAliasCon = ObjPredioAgrNRevCr.ObjAliasContStr.ToString()
        End If
        If String.IsNullOrEmpty(lstrAliasCon) Then
            lstrAliasCon = ObjIdCliente_NotaReversaCrDbl.ToString
        End If
        Return lstrAliasCon
    End Function
    Friend Function FbytQRNcr() As Byte()
        Dim lentColorFondoQR As Integer = Color.FromArgb(255, 255, 255, 255).ToArgb()
        Dim lentColorQR As Integer = Color.FromArgb(255, 0, 0, 0).ToArgb()
        Dim lqreQRNdb As New QRCodeEncoder With {
            .QRCodeEncodeMode = QRCodeEncoder.ENCODE_MODE.BYTE,
            .QRCodeScale = Int32.Parse(4),
            .QRCodeErrorCorrect = QRCodeEncoder.ERROR_CORRECTION.H,
            .QRCodeVersion = 0,
            .QRCodeBackgroundColor = System.Drawing.Color.FromArgb(lentColorFondoQR),
            .QRCodeForegroundColor = System.Drawing.Color.FromArgb(lentColorQR)
       }
        Dim lstrNit As String = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.
                ObjIdTerceroCentroUtilDbl.ToString
        Dim lbytQR As Byte() = Array.Empty(Of Byte)()
        Dim ldtmFecNrcr As Date = ObjFecha_NotaReversaCrDtm.ObjValorPro
        Dim lstrFecNrcr As String = Year(ldtmFecNrcr).ToString &
                Format(Month(ldtmFecNrcr), "00") & Format(Day(ldtmFecNrcr), "00") &
                Format(Hour(ldtmFecNrcr), "00") & Format(Minute(ldtmFecNrcr), "00") &
                Format(Second(ldtmFecNrcr), "00")
        Dim lstrNdb As String = "Número:" & StrNumeroNotaReversaCr & vbCrLf &
                "Fecha:" & lstrFecNrcr & vbCrLf &
                "Nit:" & lstrNit & vbCrLf &
                "DocAdq:" & ObjIdCliente_NotaReversaCrDbl.ToString & vbCrLf &
                "ValNdb:" & Format(ObjValor_NotaReversaCrDec.ObjValorPro, "#0.00") & vbCrLf &
                "ValIva:" & Format(0, "#0.00") & vbCrLf &
                "ValOtroIm:" & "0.00" & vbCrLf &
                "ValNdbIm:" & Format(ObjValor_NotaReversaCrDec.ObjValorPro, "#0.00") & vbCrLf
        If GobjParametros.ObjIdProveedorEFacEnt.ObjValorPro > EnuProveedorEFac.None Then
            lstrNdb &= "CUDE:" & ObjCUDEStr.ObjValorPro
        End If
        Try
            Dim lbmiQRNdb As New BitmapImage
            lbmiQRNdb.BeginInit()
            Dim lbtmQR As Bitmap = lqreQRNdb.Encode(lstrNdb, System.Text.Encoding.UTF8)
            lbytQR = FBytQR(lbtmQR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly)
        End Try
        Return lbytQR
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsDetalle_NotaReversaCrStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Detalle"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Comentario"
        HshrLongitud = 500
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return HobjValorPro
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsFecha_NotaReversaCrDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaNota"
    Private ReadOnly MobjPadre As ClsNotaReversionCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha Nota Reversa Cr"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = Date.Today
        HobjValorPro = Date.Today
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
        Else
            HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
Friend Class ClsIdCliente_NotaReversaCrDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCliente"
    Private MobjCliente As ClsCliente = Nothing
    Private ReadOnly MobjPadre As ClsNotaReversionCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id Tercero Cliente Nota Reversa Cr"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC,
                BlnEsRequerido)
        HstrMens = String.Empty
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                If IsNothing(MobjCliente) Then
                    MobjCliente = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
                End If
                Dim lobjLlavePrincipal() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                MobjCliente.SAbra(lobjLlavePrincipal)
                MobjPadre.objClienteNota = MobjCliente
                If Not MobjCliente.BlnExiste Then
                    HblnEsValido = False
                    MobjCliente.SVacie()
                    HstrMens = "La Id. del Cliente ingresada, '" & HobjValorNew.ToString & "',  no es valida!"
                End If
            End If
        Else
            HstrMens = "La Id. del Cliente ingresada, '" & HobjValorNew.ToString & "',  no es válida!"
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
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
Friend Class ClsIdDoc_NotaReversaCrEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdDocumento"
    Private ReadOnly MobjPadre As ClsNotaReversionCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id Documento Reversado"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.ObjTipoDocReversadoByt.BlnEsValido Then
                If MobjPadre.ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuReciboC Then
                    Dim lobjRecCaja As ClsReciboCaja
                    If MobjPadre.ObjPrefijoDoc_NotaReversaCrStr.BlnEsValido Then
                        Dim lstrPref = MobjPadre.ObjPrefijoDoc_NotaReversaCrStr.ObjValorPro
                        lobjRecCaja = New ClsReciboCaja()
                        lobjRecCaja.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPref, HobjValorNew})
                        HblnEsValido = lobjRecCaja.BlnExiste
                        If HblnEsValido Then
                            MobjPadre.ObjDocReversado = lobjRecCaja
                        End If
                    End If
                ElseIf MobjPadre.ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuNotaCr Then
                    Dim lobjNotaCr As ClsNotaCr
                    If MobjPadre.ObjPrefijoDoc_NotaReversaCrStr.BlnEsValido Then
                        Dim lstrPref = MobjPadre.ObjPrefijoDoc_NotaReversaCrStr.ToString
                        lobjNotaCr = New ClsNotaCr()
                        lobjNotaCr.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPref, HobjValorNew})
                        HblnEsValido = lobjNotaCr.BlnExiste
                        If HblnEsValido Then
                            MobjPadre.ObjDocReversado = lobjNotaCr
                        End If
                    End If
                Else
                    Throw New ErrorInesperadoPanLException("Documento Reversado no valido")
                End If
                If Not String.IsNullOrEmpty(HstrMens) Then
                    SNotifiqueDatInv()
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
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdNotaReversaCrEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdNota"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Id Nota Reversa Cr"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido,
                       EnuTipoValor)
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
Friend Class ClsTipoDocReversadoByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "TipoDocReversado"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdModoFacturacion"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuDocReversado.EnuReciboC,
                EnuDocReversado.EnuNotaCr, BlnEsRequerido)
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
            Return ClsOrionCop.FstrNombreModoFacturacion(HobjValorPro)
        End If
    End Function
End Class
Friend Class ClsIdPredioAgr_NotaReversaCrStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredioAgrupador"
    Private ReadOnly MobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
    Private ReadOnly MobjPadre As ClsNotaReversionCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id Predio Agrupador Nota Reversa Cr"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 200
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud,
                BlnEsRequerido)
        If Not HblnEsValido AndAlso Not IsNothing(HobjValorNew) AndAlso
                HobjValorNew.ToString.Length > ShrLongitud Then
            HstrMens = "La longitud en caracteres de los Predios Agrupadores" &
                    " es mayor a lo permitido!"
            SNotifiqueDatInv()
        End If
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew <> My.Resources.Ninguno)
                If HblnEsValido Then
                    FblnAgruPrediosEsValido(HobjValorNew)
                End If
            Else
                If Not GblnActualizandoApp Then
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
                End If
            End If
        End If
    End Sub
    Friend Function FblnAgruPrediosEsValido(astrIdPredios As String)
        Dim lstrPredios() As String = astrIdPredios.Split(",")
        Dim lobjLlavePrincipal() As Object
        For Each lstrIdPredio As String In lstrPredios
            If lstrIdPredio <> "" Then
                lobjLlavePrincipal = {GshrIdCarpeta, GshrIdCentroUtil, lstrIdPredio}
                MobjPredio.SAbra(lobjLlavePrincipal)
                HblnEsValido = MobjPredio.BlnExiste
            Else
                HblnEsValido = True
            End If
            If Not HblnEsValido Then
                Exit For
            End If
        Next
        Return HblnEsValido
    End Function
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            Dim lstrIdPredioAgr = GCSTRSINPA
            If HobjValorPro <> "" Then
                lstrIdPredioAgr = HobjValorPro
            End If
            Return lstrIdPredioAgr
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsPrefijo_NotaReversaCrStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoNota"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PrefijoNotaCon"
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
End Class
Friend Class ClsPrefijoDoc_NotaReversaCrStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoDoc"
    Private ReadOnly MobjPadre As ClsNotaReversionCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Prefijo ReciboCaja"
        HshrLongitud = 5
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then
            HobjValorNew = String.Empty
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            Dim lobjPadre As ClsCBObjetoPan = ObjPadre
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Private Sub ClsPrefijoDoc_NotaReversaCrStr_evnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If HblnEsValido Then
            MobjPadre.ObjIdDoc_NotaReversaCrEnt.SValide()
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
Friend Class ClsValor_NotaReversaCrDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"
    Private ReadOnly MobjPadre As ClsNotaReversionCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Valor"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0.01,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
        If Not HblnEsValido Then
            If HobjValorNew = 0 Then
                HblnEsValido = MobjPadre.ObjAnuladoBln.ObjValorPro
            End If
        Else
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                If MobjPadre.ObjAnuladoBln.ObjValorPro Then
                    HblnEsValido = (HobjValorNew = 0)
                Else
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
#End Region