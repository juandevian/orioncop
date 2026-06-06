Imports System.Drawing
Imports System.Windows.Media.Imaging
Imports ThoughtWorks.QRCode.Codec
Friend Class ClsNotaDb
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriNotasDb"
    ' Variables de modulo
    Private MobjClienteNotaDb As ClsCliente = Nothing
    Private MobjPredioAgrNDb As ClsPredio = Nothing
    Private McolItemsNotaDb As Collection = Nothing
    Private MdtbItemsNotaDb As DataTable = Nothing
    Private MdtbNovedadesNotaDb As DataTable = Nothing
    Private MstrNroNotaDb As String = String.Empty
#End Region

#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Nota Débito en modo único
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
    ''' Instancia un objeto Nota Débito en modo navegable
    ''' </summary>
    Public Sub New(astrPref As String)
        HobjPadre = Nothing
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion
        lstrFiltro &= " AND " & ClsPrefijo_NotaDbStr.SstrNombreCampoBd & " = '" & astrPref & "'"
        HcolFiltros.Add(lstrFiltro)
        Dim lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_NotaDbStr.SstrNombreCampoBd, ClsIdNotaDbEnt.SstrNombreCampoBd}
        HblnEsCreable = False
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
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As ClsCliente, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        HblnEsSuprimible = False
        HblnEsModificable = False
        '
        DrwRegistroActual = adrwObjeto
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
            Return EnuIdClasesPanDef.EnuNotaDb
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Nota Debito"
        End Get
    End Property
#End Region

#Region "Propiedades Prop"
    Friend ReadOnly Property ObjCUDEStr As New ClsCUDEStr(Me)
    Friend ReadOnly Property ObjCUDocStr As New ClsCUDocStr(Me)
    Friend ReadOnly Property ObjFecha_NotaDbDtm As New ClsFecha_NotaDbDtm(Me)
    Friend ReadOnly Property ObjFechaAnulacion_NotaDbDtm As New ClsFechaAnulacion_NotaDbDtm(Me)
    Friend ReadOnly Property ObjIdCarpeta_NotaDbShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_NotaDbShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdEstadoEDocEnt As New ClsIdEstadoEDocEnt(Me)
    Friend ReadOnly Property ObjIdNotaDbEnt As New ClsIdNotaDbEnt(Me)
    Friend ReadOnly Property ObjIdPredioAgrupador_NotaDbStr As New _
            ClsIdPredioAgrupador_NotaDbStr(Me)
    Friend ReadOnly Property ObjIdCliente_NotaDbDbl As New ClsIdCliente_NotaDbDbl(Me)
    Friend ReadOnly Property ObjIdUsuario_NotaDbStr As New ClsIdUsuarioStr(Me)
    Friend ReadOnly Property ObjOrigenByt As New ClsOrigenByt(Me)
    Friend ReadOnly Property ObjPrefijo_NotaDbStr As New ClsPrefijo_NotaDbStr(Me)
    Friend ReadOnly Property ObjValor_NotaDbDec As New ClsValor_NotaDbDec(Me)
    Friend ReadOnly Property ObjVerEFacEnt As New ClsVerEFacEnt(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAnuladoBln)
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjIdUsuarioAnuloStr)
                HcolPropiedades.Add(ObjOrigenInstanciaStr)
                HcolPropiedades.Add(ObjOrigenInstanciaAnuloStr)
                HcolPropiedades.Add(ObjCUDEStr)
                HcolPropiedades.Add(ObjCUDocStr)
                HcolPropiedades.Add(ObjFecha_NotaDbDtm)
                HcolPropiedades.Add(ObjFechaAnulacion_NotaDbDtm)
                HcolPropiedades.Add(ObjIdCarpeta_NotaDbShr)
                HcolPropiedades.Add(ObjIdCentroUtil_NotaDbShr)
                HcolPropiedades.Add(ObjIdEstadoEDocEnt)
                HcolPropiedades.Add(ObjIdNotaDbEnt)
                HcolPropiedades.Add(ObjIdPredioAgrupador_NotaDbStr)
                HcolPropiedades.Add(ObjIdCliente_NotaDbDbl)
                HcolPropiedades.Add(ObjIdUsuario_NotaDbStr)
                HcolPropiedades.Add(ObjOrigenByt)
                HcolPropiedades.Add(ObjPrefijo_NotaDbStr)
                HcolPropiedades.Add(ObjValor_NotaDbDec)
                HcolPropiedades.Add(ObjVerEFacEnt)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region

#Region "Otras propiedades"
    ''' <summary>
    ''' Devuelve uns string compuesto por el prefijo de la nota y el id de la nota separados por un
    ''' guion. Si no existe el prefijo devuelve solo el id de la factura
    ''' </summary>
    ''' <value></value>
    Friend ReadOnly Property StrNumeroNotaDb As String
        Get
            Dim lstrNumeroNotaDb As String = ClsPanorama.FstrNumeroDcto(ObjPrefijo_NotaDbStr.ObjValorPro,
                    ObjIdNotaDbEnt.ObjValorPro)
            Return lstrNumeroNotaDb
        End Get
    End Property
    Friend ReadOnly Property ObjClienteNotaDb As ClsCliente
        Get
            If IsNothing(MobjClienteNotaDb) Then
                MobjClienteNotaDb = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
                MobjClienteNotaDb.SAbra({GshrIdCarpeta, GshrIdCentroUtil,
                        ObjIdCliente_NotaDbDbl.ObjValorPro})
            End If
            Return MobjClienteNotaDb
        End Get
    End Property
    Friend ReadOnly Property ObjPredioAgrNDb As ClsPredio
        Get
            If IsNothing(MobjPredioAgrNDb) Then
                If Not String.IsNullOrEmpty(ObjIdPredioAgrupador_NotaDbStr.ToString) AndAlso
                        ObjIdPredioAgrupador_NotaDbStr.ToString <> GCSTRSINPA Then
                    MobjPredioAgrNDb = New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil,
                            ObjIdPredioAgrupador_NotaDbStr.ObjValorPro}
                    MobjPredioAgrNDb.SAbra(lobjValorLlave)
                    If Not MobjPredioAgrNDb.BlnExiste Then
                        MobjPredioAgrNDb = Nothing
                    End If
                End If
            End If
            Return MobjPredioAgrNDb
        End Get
    End Property
    ''' <summary>
    ''' Devuelve la factura que afecta el primer item de la nota. Para facturación elecrrónica solo puede
    ''' existir una npta por factura.
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property ObjFacturaAfectada As ClsFactura
        Get
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                McolItemsNotaDb = ColItemsNotaDb
            End If
            Dim lobjItemNdb As ClsItemNotaDb = Nothing
            If McolItemsNotaDb IsNot Nothing AndAlso McolItemsNotaDb.Count > 0 Then
                lobjItemNdb = McolItemsNotaDb(1)
            End If
            Dim lobjFact As ClsFactura = Nothing
            If lobjItemNdb IsNot Nothing Then
                lobjFact = lobjItemNdb.ObjFactura
            End If
            Return lobjFact
        End Get
    End Property
    Friend Shared ReadOnly Property StrTipoConceptoDian As String
        Get
            Return CType(EnuConceptoNotaDbDian.EnuOtros, Integer).ToString
        End Get
    End Property
    Friend ReadOnly Property DecValorIvaNota As Decimal
        Get
            Dim ldecValorIva = 0D
            For Each lobjItemNotaDb As ClsItemNotaDb In ColItemsNotaDb
                ldecValorIva += lobjItemNotaDb.DecValorIva
            Next
            Return ldecValorIva
        End Get
    End Property
    Friend ReadOnly Property DecValorAntesIva As Decimal
        Get
            Dim ldecVlr = 0D
            For Each lobjItemNotaDb As ClsItemNotaDb In ColItemsNotaDb
                ldecVlr += lobjItemNotaDb.DecValorAntesIva
            Next
            Return ldecVlr
        End Get
    End Property
    Friend ReadOnly Property DecValorBaseIvaNota As Decimal
        Get
            Dim ldecValoBaserIva = 0D
            For Each lobjItemNotaDb As ClsItemNotaDb In ColItemsNotaDb
                ldecValoBaserIva += lobjItemNotaDb.DecValorBaseIva
            Next
            Return ldecValoBaserIva
        End Get
    End Property
#End Region
#End Region

#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        MdtbItemsNotaDb = Nothing
        McolItemsNotaDb = Nothing
        MdtbNovedadesNotaDb = Nothing
        MobjClienteNotaDb = Nothing
        MobjPredioAgrNDb = Nothing
        MstrNroNotaDb = String.Empty
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        GobjPanDat.SControleProcesoObj(True)
        Try
            GobjPanDat.SInicialiceTransaccion()
            If ColItemsNotaDb.Count = 0 Then
                Throw New ErrorInesperadoPanLException("Inentado actualizar Nota Db sin Items")
            End If
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                SDetermineValorNota()
                SNumereObj()
                SComplementeEFac()
                ClsPanorama.SActualiceCol(ColItemsNotaDb)
                ObjFechaCreacionDtm.ObjValorPro = Date.Now
                MyBase.SActualice(ablnExigeRequeridos)
            Else
                ClsPanorama.SActualiceCol(ColItemsNotaDb)
                MyBase.SActualice(ablnExigeRequeridos)
            End If
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
    Protected Overrides Sub SInicialiceObj()
        ObjIdCarpeta_NotaDbShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtil_NotaDbShr.ObjValorPro = GshrIdCentroUtil
        ObjIdUsuario_NotaDbStr.ObjValorPro = GstrIdUsuario
        ObjCUDocStr.ObjValorPro = String.Empty
        ObjCUDEStr.ObjValorPro = String.Empty
        ObjPrefijo_NotaDbStr.ObjValorPro = String.Empty
        ObjOrigenByt.ObjValorPro = EnuOrigenNotaDb.None
        ObjAnuladoBln.ObjValorPro = False
        ObjOrigenInstanciaStr.ObjValorPro = GstrOrigenActual
        ObjVerEFacEnt.ObjValorPro = EnuVerEFac.EnuNinguna
        ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoEDoc
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return StrNumeroNotaDb
        End Get
    End Property
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
        Dim lblnHay = False
        For Each lobjItemNdb As ClsItemNotaDb In ColItemsNotaDb
            lblnHay = lobjItemNdb.ObjFactura.BlnEsFacEle AndAlso
                    (Not lobjItemNdb.ObjFactura.FenuVerFacEFac = EnuVerEFac.EnuV1)
            If lblnHay Then Exit For
        Next
        Return lblnHay
    End Function
    Friend ReadOnly Property FblnInsertarEFac As Boolean
        Get
            Dim lblnReg = False
            If BlnEsDocEle AndAlso BlnAfectaFrasRegEFac() Then
                lblnReg = ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoReg
                If lblnReg Then
                    For Each lobjItemNdb As ClsItemNotaDb In ColItemsNotaDb
                        lblnReg = lobjItemNdb.ObjFactura.BlnEsFacEle AndAlso
                                (Not lobjItemNdb.ObjFactura.FenuVerFacEFac = EnuVerEFac.EnuV1) AndAlso
                                lobjItemNdb.ObjFactura.ObjIdEstadoEDocEnt.ObjValorPro >=
                                EnuEstadoEDoc.EnuRegi AndAlso
                                lobjItemNdb.ObjFactura.ObjIdEstadoEDocEnt.ObjValorPro <>
                                EnuEstadoEDoc.EnuNoEDoc
                        If lblnReg Then Exit For
                    Next
                End If
            End If
            Return lblnReg
        End Get
    End Property
    Friend Function FblnActualizarEstEFac() As Boolean
        Dim lblnActu = False
        If BlnEsDocEle AndAlso BlnAfectaFrasRegEFac() Then
            lblnActu = (ObjIdEstadoEDocEnt.ObjValorPro < EnuEstadoEDoc.EnuEnviada) AndAlso
                    (ObjIdEstadoEDocEnt.ObjValorPro >= EnuEstadoEDoc.EnuEnProceso)
            If lblnActu Then
                If lblnActu Then
                    For Each lobjItemNdb As ClsItemNotaDb In ColItemsNotaDb
                        lblnActu = lobjItemNdb.ObjFactura.BlnEsFacEle AndAlso
                                (Not lobjItemNdb.ObjFactura.FenuVerFacEFac = EnuVerEFac.EnuV1) AndAlso
                                lobjItemNdb.ObjFactura.ObjIdEstadoEDocEnt.ObjValorPro >= EnuEstadoEDoc.EnuRegi
                        If lblnActu Then Exit For
                    Next
                End If
            End If
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
        If GobjParametros.BlnEFacAutorizado AndAlso BlnAfectaFrasRegEFac() Then
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
            Dim lentIdNotaDb As Integer
            Dim lstrPrefijo = GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuNotaDb)
            If IsNothing(lstrPrefijo) Then lstrPrefijo = String.Empty
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsPrefijo_NotaDbStr.SstrNombreCampoBd & " = '" & lstrPrefijo & "'"
            lentIdNotaDb = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                    ClsIdNotaDbEnt.SstrNombreCampoBd, ObjIdNotaDbEnt.EnuTipoValor,
                    lstrFiltro)
            If lentIdNotaDb < GobjParametros.FentNumeracionInicialDoc(EnuTipoDocOri.EnuNotaDb) Then
                lentIdNotaDb = GobjParametros.FentNumeracionInicialDoc(EnuTipoDocOri.EnuNotaDb)
            End If
            lentIdNotaDb += 1
            ObjPrefijo_NotaDbStr.ObjValorPro = lstrPrefijo
            ObjIdNotaDbEnt.ObjValorPro = lentIdNotaDb
            For Each lobjItemNotaDb As ClsItemNotaDb In ColItemsNotaDb
                lobjItemNotaDb.ObjIdNotaDb_ItemNotaDbEnt.ObjValorPro = lentIdNotaDb
                lobjItemNotaDb.SGenereNovedades(lstrPrefijo, lentIdNotaDb,
                        ObjIdPredioAgrupador_NotaDbStr.ObjValorPro)
            Next
        End If
    End Sub

    Private Sub SComplementeEFac()
        If GobjParametros.ObjAutorizaEFacBln.ObjValorPro Then
            If BlnAfectaFrasRegEFac() Then
                ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoReg
                ObjVerEFacEnt.ObjValorPro = EnuVerEFac.EnuV2
            End If
        End If
    End Sub

    Private Sub SDetermineValorNota()
        Dim ldecValorNota = 0D
        For Each lobjItemotaDb As ClsItemNotaDb In McolItemsNotaDb
            ldecValorNota += lobjItemotaDb.ObjValor_ItemNotaDbDec.ObjValorPro
        Next
        ObjValor_NotaDbDec.ObjValorPro = ldecValorNota
    End Sub

    Friend Function FarlCorreosNDb() As ArrayList
        Dim larlListaCorreos As New ArrayList
        Dim lstrCorreoCli As String, lstrCorreoPredio As String
        If ObjClienteNotaDb.ObjRecibeDocsPorEmailBln.ObjValorPro Then
            lstrCorreoCli = ObjClienteNotaDb.ObjEmailStr.ToString
            If Not String.IsNullOrEmpty(lstrCorreoCli) Then
                larlListaCorreos.Add(lstrCorreoCli)
            End If
        End If
        If Not IsNothing(ObjPredioAgrNDb) Then
            lstrCorreoPredio = ObjPredioAgrNDb.ObjEmailAdiStr.ToString
            If Not String.IsNullOrEmpty(lstrCorreoPredio) Then
                larlListaCorreos.Add(lstrCorreoPredio)
            End If
        End If
        Return larlListaCorreos
    End Function

    Friend Function FstrAliasCon() As String
        Dim lstrAliasCon = String.Empty
        If ObjPredioAgrNDb IsNot Nothing Then
            lstrAliasCon = ObjPredioAgrNDb.ObjAliasContStr.ToString
        End If
        If String.IsNullOrEmpty(lstrAliasCon) Then
            lstrAliasCon = ObjIdCliente_NotaDbDbl.ToString()
        End If
        Return lstrAliasCon
    End Function

    Friend Function FbytQRNdb() As Byte()
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
        Dim ldtmFecNdb As Date = ObjFecha_NotaDbDtm.ObjValorPro
        Dim lstrFecNdb As String = Year(ldtmFecNdb).ToString &
                Format(Month(ldtmFecNdb), "00") & Format(Day(ldtmFecNdb), "00") &
                Format(Hour(ldtmFecNdb), "00") & Format(Minute(ldtmFecNdb), "00") &
                Format(Second(ldtmFecNdb), "00")
        Dim lstrNdb As String = "Número:" & StrNumeroNotaDb & vbCrLf &
                "Fecha:" & lstrFecNdb & vbCrLf &
                "Nit:" & lstrNit & vbCrLf &
                "DocAdq:" & ObjIdCliente_NotaDbDbl.ToString & vbCrLf &
                "ValNdb:" & Format(ObjValor_NotaDbDec.ObjValorPro, "#0.00") & vbCrLf &
                "ValIva:" & Format(0, "#0.00") & vbCrLf &
                "ValOtroIm:" & "0.00" & vbCrLf &
                "ValNdbIm:" & Format(ObjValor_NotaDbDec.ObjValorPro, "#0.00") & vbCrLf
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

    Friend Function FdtbNotasUltimoMes() As DataTable
        Dim lstrTablaPri = SstrNombreTabla
        Dim lstrTablaSec = ClsCliente.SstrNombreTabla
        Dim lstrCampSelPri As String() = {ClsFecha_NotaDbDtm.SstrNombreCampoBd,
                ClsPrefijo_NotaDbStr.SstrNombreCampoBd, ClsIdNotaDbEnt.SstrNombreCampoBd,
                ClsIdPredioAgrupador_NotaDbStr.SstrNombreCampoBd,
                ClsIdCliente_NotaDbDbl.SstrNombreCampoBd, ClsValor_NotaDbDec.SstrNombreCampoBd}
        Dim lstrCampSelSec As String() = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsIdCliente_NotaDbDbl.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrFechaIni = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today.AddMonths(-2)) &
                "'"
        Dim lstrFechaFin = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today) & "'"
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsFecha_NotaDbDtm.SstrNombreCampoBd & " BETWEEN " & lstrFechaIni & " AND " &
                lstrFechaFin
        Dim lstrOrden As String(,) = {{ClsFecha_NotaDbDtm.SstrNombreCampoBd, "DESC"},
                {ClsPrefijo_NotaDbStr.SstrNombreCampoBd, "ASC"},
                {ClsIdNotaDbEnt.SstrNombreCampoBd, "DESC"}}
        Dim ldtbNotasdb = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCampSelPri, lstrTablaSec,
                lstrCampSelSec, lstrCampRelPri, lstrCampRelSec, lstrOrden, True, lstrFiltro, {})
        Return ldtbNotasdb
    End Function
#End Region

#Region "Manejo Items Nota Db"
    Friend ReadOnly Property ColItemsNotaDb As Collection
        Get
            If IsNothing(McolItemsNotaDb) OrElse McolItemsNotaDb.Count = 0 Then
                McolItemsNotaDb = New Collection
                If ObjIdNotaDbEnt.BlnEsValido AndAlso EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                    SCargueDtbItemsNotaDb()
                    If Not IsNothing(MdtbItemsNotaDb) AndAlso MdtbItemsNotaDb.Rows.Count > 0 Then
                        Dim ldrwItemsNotaDb() As DataRow = MdtbItemsNotaDb.Select
                        For Each ldrwItemNotaDb As DataRow In ldrwItemsNotaDb
                            Dim lobjItemNotaDb As New ClsItemNotaDb(Me, ldrwItemNotaDb)
                            lobjItemNotaDb.SLeaValores(True)
                            McolItemsNotaDb.Add(lobjItemNotaDb, lobjItemNotaDb.ObjIdItemNotaDbShr.ToString)
                        Next
                    End If
                End If
            End If
            Return McolItemsNotaDb
        End Get
    End Property
    Friend ReadOnly Property DtbItemsNotaDb As DataTable
        Get
            SCargueDtbItemsNotaDb()
            SComplementeTablaItems()
            Return MdtbItemsNotaDb
        End Get
    End Property
    Friend Sub SAdicioneItemNotaDb(aentIdFactura As Integer, astrPrefijoFact As String,
            astcIntMoraFact As StcIntMoraFactura, adblTasaMora As Double)
        If IsNothing(McolItemsNotaDb) Then
            McolItemsNotaDb = New Collection
        End If
        Dim lshrIdItemNotaDb As Short = McolItemsNotaDb.Count + 1
        SCargueDtbItemsNotaDb()
        Dim ldrwItemNotaDb As DataRow = MdtbItemsNotaDb.NewRow
        Dim lobjItemNotaDb As New ClsItemNotaDb(Me, ldrwItemNotaDb)
        With lobjItemNotaDb
            .SCreeObj(Nothing)
            .ObjAnuladoBln.ObjValorPro = False
            .ObjBaseMoraDec.ObjValorPro = astcIntMoraFact.DecBaseIntereses
            .ObjFechaCausoMora_Dtm.ObjValorPro = astcIntMoraFact.DtmFechaCauso
            .ObjIdCarpeta_ItemNotaDbShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_ItemNotaDbShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdFactura_ItemNotaDbEnt.ObjValorPro = aentIdFactura
            .ObjIdItemFac_ItemNotaDbShr.ObjValorPro = astcIntMoraFact.ShrIdItemFactura
            .ObjIdItemNotaDbShr.ObjValorPro = lshrIdItemNotaDb
            .ObjPrefijo_ItemNotaDbStr.ObjValorPro =
                    GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuNotaDb)
            .ObjPrefijoFact_ItemNotaDbStr.ObjValorPro = astrPrefijoFact
            .ObjTasaMora_ItemNotaDbDbl.ObjValorPro = adblTasaMora
            .ObjValor_ItemNotaDbDec.ObjValorPro = astcIntMoraFact.DecVlrMora
            .ObjDiasMoraEnt.ObjValorPro = astcIntMoraFact.EntDiasMora
            .ObjTarifaIva_ItemNotaDbDbl.ObjValorPro = astcIntMoraFact.DblTarifaIva
        End With
        McolItemsNotaDb.Add(lobjItemNotaDb, lshrIdItemNotaDb.ToString)
    End Sub
    Friend Sub SAdicioneItemNotaDb(aentIdFactura As Integer, astrPrefijoFact As String, adecBase As Decimal,
            aentIdItemFac As Integer, adblTasaMora As Double, adecVlrMora As Decimal,
            aentDiasMora As Integer, adtmFechaCauso As Date)
        If IsNothing(McolItemsNotaDb) Then
            McolItemsNotaDb = New Collection
        End If
        Dim lshrIdItemNotaDb As Short = McolItemsNotaDb.Count + 1
        SCargueDtbItemsNotaDb()
        Dim ldrwItemNotaDb As DataRow = MdtbItemsNotaDb.NewRow
        Dim lobjItemNotaDb As New ClsItemNotaDb(Me, ldrwItemNotaDb)
        With lobjItemNotaDb
            If Not CType(.EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
                .EnuPermisosObj += EnuPermisosDef.enuCrear
            End If
            .SCreeObj(Nothing)
            .ObjAnuladoBln.ObjValorPro = False
            .ObjBaseMoraDec.ObjValorPro = adecBase
            .ObjIdCarpeta_ItemNotaDbShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_ItemNotaDbShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdFactura_ItemNotaDbEnt.ObjValorPro = aentIdFactura
            .ObjIdItemFac_ItemNotaDbShr.ObjValorPro = aentIdItemFac
            .ObjIdItemNotaDbShr.ObjValorPro = lshrIdItemNotaDb
            .ObjPrefijo_ItemNotaDbStr.ObjValorPro =
                    GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuNotaDb)
            .ObjPrefijoFact_ItemNotaDbStr.ObjValorPro = astrPrefijoFact
            .ObjTasaMora_ItemNotaDbDbl.ObjValorPro = adblTasaMora
            .ObjValor_ItemNotaDbDec.ObjValorPro = adecVlrMora
            .ObjDiasMoraEnt.ObjValorPro = aentDiasMora
            .ObjFechaCausoMora_Dtm.ObjValorPro = adtmFechaCauso
        End With
        McolItemsNotaDb.Add(lobjItemNotaDb, lshrIdItemNotaDb.ToString)
    End Sub
    Private Sub SCargueDtbItemsNotaDb()
        Dim lstrIndice = {{ClsPrefijo_NotaDbStr.SstrNombreCampoBd, "ASC"},
                          {ClsIdNotaDb_ItemNotaDbEnt.SstrNombreCampoBd, "ASC"},
                          {ClsIdItemNotaDbShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsPrefijo_NotaDbStr.SstrNombreCampoBd &
                    " = '" & ObjPrefijo_NotaDbStr.ObjValorPro & "' AND " &
                    ClsIdNotaDb_ItemNotaDbEnt.SstrNombreCampoBd & " = "
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            lstrFiltro &= "0"
        Else
            If IsNothing(ObjIdNotaDbEnt.ObjValorPro) Then
                lstrFiltro &= "0"
            Else
                lstrFiltro &= ObjIdNotaDbEnt.ObjValorPro
            End If
        End If
        Dim lstrCamposSelect() = {"*", "'' as NroFact"}
        MdtbItemsNotaDb = ClsPanorama.FdtbDataTable(ClsItemNotaDb.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
    End Sub
    Private Sub SComplementeTablaItems()
        If IsNothing(McolItemsNotaDb) Then
            McolItemsNotaDb = ColItemsNotaDb
        End If
        Dim lobjcenUtil As ClsCentroUtilidad = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual
        Dim ldrwItemsNotaDb As DataRow() = MdtbItemsNotaDb.Select()
        Dim lstrPrefFact As String : Dim lentIdFact As Integer : Dim lstrNroFact As String
        For Each ldrwItNotaDb As DataRow In ldrwItemsNotaDb
            lstrPrefFact = ClsPanorama.FobjValorCampo(ldrwItNotaDb(ClsPrefijoFact_ItemNotaDbStr.SstrNombreCampoBd),
                                EnuTipoValor.enuString)
            lentIdFact = ClsPanorama.FobjValorCampo(ldrwItNotaDb(ClsIdFactura_ItemNotaDbEnt.SstrNombreCampoBd),
                                EnuTipoValor.enuInteger)
            lstrNroFact = ClsPanorama.FstrNumeroDcto(lstrPrefFact, lentIdFact)
            ldrwItNotaDb("NroFact") = lstrNroFact
        Next
    End Sub
#End Region

#Region "Novedades"
    Friend ReadOnly Property DtbNovedadesNotaDb As DataTable
        Get
            SCargueDtbNovedadesNota()
            SComplementeTablaNov()
            Return MdtbNovedadesNotaDb
        End Get
    End Property
    Private Sub SCargueDtbNovedadesNota()
        If IsNothing(MdtbNovedadesNotaDb) Then
            Dim lstrIdNotaDb = ObjIdNotaDbEnt.ToString
            If lstrIdNotaDb = String.Empty Then lstrIdNotaDb = "0"
            Dim lstrTabla = ClsNovedad.SstrNombreTabla
            Dim lstrCamposSelect = {"*", "'' AS NroFac"}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdTipoDocOrigenByt.SstrNombreCampoBd &
                    " = " & EnuTipoDocOri.EnuNotaDb & " AND " & ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" &
                    ObjPrefijo_NotaDbStr.ObjValorPro & "' AND " & ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " &
                    lstrIdNotaDb
            Dim lstrIndice = {{ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdDocOrigenEnt.SstrNombreCampoBd, "ASC"},
                              {ClsPrefijoFact_NovStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdFactura_NovEnt.SstrNombreCampoBd, "ASC"},
                              {ClsIdNovedadShr.SstrNombreCampoBd, "ASC"}}
            MdtbNovedadesNotaDb = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrIndice, lstrFiltro)
        End If
    End Sub
    Private Sub SComplementeTablaNov()
        If MstrNroNotaDb <> StrNumeroNotaDb Then
            MstrNroNotaDb = StrNumeroNotaDb
            Dim lstrPrefFact As String : Dim lentIdFact As Integer : Dim lstrNroFact As String
            Dim ldrwNovedades = MdtbNovedadesNotaDb.Select
            For Each ldrwNovedad As DataRow In ldrwNovedades
                lstrPrefFact = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsPrefijoFact_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
                lentIdFact = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdFactura_NovEnt.SstrNombreCampoBd),
                        EnuTipoValor.enuInteger)
                lstrNroFact = ClsPanorama.FstrNumeroDcto(lstrPrefFact, lentIdFact)
                ldrwNovedad("NroFac") = lstrNroFact
            Next
        End If
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsCUDE_NDbStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "CUDE"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Código Unico Factura Electrónica"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 200
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
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
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class

Friend Class ClsFecha_NotaDbDtm
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaNota"
    Private ReadOnly MobjPadre As clsNotaDb = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaNotaDb"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = GCDTMFECHANULA
        HobjValorPro = GCDTMFECHANULA
    End Sub
    Public Overrides Sub SValide()
        Dim lobjFact As ClsFactura = MobjPadre.ObjFacturaAfectada
        Dim lenuModoCM As EnuModoCausaMora = EnuModoCausaMora.None
        If lobjFact IsNot Nothing Then
            lenuModoCM = lobjFact.FenuModoCausaMora
        End If
        Dim ldtmFecIniPer = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
        Dim ldtmFechaMin = If(GblnCausandoFM AndAlso lenuModoCM = EnuModoCausaMora.EnuUltimoDia,
                ldtmFecIniPer.AddDays(-1), ldtmFecIniPer)
        If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
            ldtmFechaMax = Now()
        Else
            ldtmFechaMax = ldtmFechaMax.AddDays(1)
        End If
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                    BlnEsRequerido)
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

Friend Class ClsFechaAnulacion_NotaDbDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaAnulacion"
    Private ReadOnly MobjPadre As ClsNotaDb = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaAnulacion_NotaDb"
        HenuTipoValor = EnuTipoValor.enuDateTime
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = GCDTMFECHANULA
        HobjValorPro = HobjValorNew
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = Not IsNothing(HobjValorNew) AndAlso IsDate(HobjValorNew)
        If HblnEsValido Then
            If HobjValorNew <> GCDTMFECHANULA Then
                HblnEsValido = MobjPadre.FblnEsAnulable
                If HblnEsValido Then
                    Dim ldtmFechaMin As Date = GCDTMFECHANULA
                    Dim ldtmFechaMax As Date = GCDTMFECHANULA
                    If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                        HblnEsValido = MobjPadre.ObjAnuladoBln.ObjValorPro
                        If HblnEsValido Then
                            ldtmFechaMin = Date.Today
                            ldtmFechaMax = Now
                        End If
                    ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                        ldtmFechaMin = HobjValorOriginal
                        ldtmFechaMax = HobjValorOriginal
                    End If
                    HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                            BlnEsRequerido)
                End If
            End If
        End If
    End Sub
    Private Sub EPosSetValor() Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando AndAlso BlnEsValido Then
            If HobjValorPro <> GCDTMFECHANULA Then
                MobjPadre.ObjValor_NotaDbDec.SValide()
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, GCSTRFMTFECHASIMPLE)
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class

Friend Class ClsIdNotaDbEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdNotaDb"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdNotaDebito"
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

Friend Class ClsIdPredioAgrupador_NotaDbStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredioAgrupador"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdPredioAgrupador NotaDb"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = "***"
        HobjValorPro = "***"
        HobjValorOriginal = "***"
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud,
                BlnEsRequerido) AndAlso HobjValorNew <> "***"
        If HblnEsValido Then
            Dim lobjPadre As ClsNotaDb = ObjPadre
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If Not String.IsNullOrEmpty(HobjValorNew) Then
                    Dim lobjLlavePrincipal() As Object = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                    Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                    lobjPredio.SAbra(lobjLlavePrincipal)
                    HblnEsValido = lobjPredio.BlnExiste
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

Friend Class ClsIdCliente_NotaDbDbl
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCliente"
    Private MstrNombreCliente As String = String.Empty
    Private ReadOnly MobjPadre As clsNotaDb = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTerceroCliente_NotaDb"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MstrNombreCliente = String.Empty
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                Dim lobjCliente As ClsCliente = MobjPadre.ObjPadre
                HblnEsValido = (HobjValorNew = lobjCliente.ObjIdClienteDbl.ObjValorPro)
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Friend ReadOnly Property StrNombreCliente As String
        Get
            If HblnEsValido AndAlso MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                If String.IsNullOrEmpty(MstrNombreCliente) Then
                    Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
                    Dim lobjvalorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorPro}
                    lobjCliente.SAbra(lobjvalorLlave)
                    MstrNombreCliente = lobjCliente.ObjNombreCompletoStr.ObjValorPro
                End If
            End If
            Return MstrNombreCliente
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

Friend Class ClsOrigenByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Origen"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Origen Nota Db"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsNotaDb = ObjPadre
        If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuOrigenNotaDb.EnuAplicacion,
                EnuOrigenNotaDb.EnuImportado, BlnEsRequerido)
        Else
            HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuOrigenNotaDb.None,
                EnuOrigenNotaDb.EnuImportado, BlnEsRequerido)
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

Friend Class ClsPrefijo_NotaDbStr
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoNotaDb"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PrefijoNotaDb"
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
        If HblnEsValido Then
            Dim lobjPadre As ClsCBObjetoPan = ObjPadre
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuNotaDb))
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

Friend Class ClsValor_NotaDbDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"
    Private ReadOnly MobjPadre As ClsNotaDb = Nothing
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