Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Friend Class ClsInterfazMisFacturas
#Region "Definiciones"
    Implements IDisposable
#Region "Variables y constantes"
    Private Const CMSTRINDENT As String = "  "
    Private Const CMSTRINICIODOC As String = "{"
    Private Const CMSTRFINCAMPO As String = ","
    ' Apertura Informacion del Cliente
    Private Const CMSTRCIERREINF As String = CMSTRINDENT & "},"
    Private Const CMSTRFALSO As String = "false"
    Private Const CMSTRCERO As String = "0"
    Private Const CMSTRCERODEC = "0.00"
    Private Const CMSTRUNODEC As String = "1.00"
    Private Const CMSTRFRMDEC As String = "#0.00"
    Private ReadOnly CMSTRNULO As String = FstrComillas("")
    Private ReadOnly CMSTRUNI As String = FstrComillas("94")
    Private ReadOnly CMSTRCODMONEDA As String = FstrComillas("COP")
    Private ReadOnly CMSTRFECHANULL As String = FstrComillas("0001-01-01T00:00:00")
    ' Objetos
    Private MblnDisposed As Boolean = False
    Private MobjFactura As ClsFactura = Nothing
    Private Mobjcliente As ClsCliente = Nothing
    Private MstrToken As String = String.Empty
    Private MstrToken_V1 As String = String.Empty
    ' Http
    Private MhttpMisFac As Http = Nothing
    '
    Private MhttpClienteMisFac As HttpClient = Nothing
    Private MhttpRespuesta As HttpResponseMessage = Nothing
    Private ReadOnly MstrUsuarioMisF As String = String.Empty
    Private ReadOnly MstrPasswordMisF As String = String.Empty
    Private ReadOnly MstrUrlApi As String = String.Empty
    Private ReadOnly MstrCompl As String = String.Empty
    Private ReadOnly MstrIdUsuario As String = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjIdTerceroCentroUtilDbl.ToString
    Friend Property StrMensajeError As String = String.Empty
#End Region
#Region "Enumeradores"
    Private Enum EnuFormaPago As Integer
        None = 0
        enuContado
        enuCredito
    End Enum
#End Region
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Panorama.
    ''' </summary>
    ''' <param name="aobjRegistro">Indica si el llamador está autorizado.</param>
    Public Sub New(aobjRegistro As Object)
        If aobjRegistro Is Nothing OrElse Not (aobjRegistro.GetType.Name = "String" AndAlso
                aobjRegistro = GCOBJREGISTRO) Then
            Throw New ModuloNoRegistradoPanException()
        End If
        MstrUrlApi = GobjParametros.ObjURLStr.ObjValorPro
        MstrUsuarioMisF = GobjParametros.ObjIdUsuarioProvEFacStr.ObjValorPro
        MstrPasswordMisF = ClsPanorama.FstrContrasena(GobjParametros.ObjContrasenaAPIEFacStr)
        MstrCompl = "SchemaID=31&IDNumber=" & MstrIdUsuario
    End Sub
#End Region
#Region "Propiedades"
    Friend ReadOnly Property BlnDisposed As Boolean
        Get
            Return MblnDisposed
        End Get
    End Property
    Friend ReadOnly Property FhttpCliMisFac As HttpClient
        Get
            If IsNothing(MhttpClienteMisFac) OrElse BlnDisposed Then
                MhttpClienteMisFac = New HttpClient
            End If
            Return MhttpClienteMisFac
        End Get
    End Property
#End Region
#Region "Procedimientos"
    Friend Async Function SInserteDoc(aobjDoc As ClsCBObjetoPan,
                aenuTipoDoc As EnuTipoDocOri) As Task
        Dim lstrRespuesta = String.Empty, lstrCUDoc = String.Empty, lstrDoc = String.Empty
        Dim lstrMensError = String.Empty, lblnNoHayError As Boolean
        Dim lstrMens = String.Empty
        Try
            Dim lstrJsDoc = FstrJsonDoc(aobjDoc, aenuTipoDoc)
            Select Case aenuTipoDoc
                Case EnuTipoDocOri.EnuFactura
                    If MobjFactura.ObjIdModoFacturacionByt.ObjValorPro =
                            EnuModoFacturacionDef.EnuContingencia Then
                        lstrRespuesta = Await SInserteJsFacCon(lstrJsDoc, 0)
                    Else
                        lstrRespuesta = Await SInserteJsFac(lstrJsDoc, 0)
                    End If
                    lstrDoc = "Fac"
                Case EnuTipoDocOri.EnuNotaCon
                    lstrRespuesta = Await SInserteJsNota(lstrJsDoc, False, 0)
                    lstrDoc = "NCon"
                Case EnuTipoDocOri.EnuNotaCr
                    lstrRespuesta = Await SInserteJsNota(lstrJsDoc, False, 0)
                    lstrDoc = "NCr"
                Case EnuTipoDocOri.EnuNotaDb
                    lstrRespuesta = Await SInserteJsNota(lstrJsDoc, True, 0)
                    lstrDoc = "NDb"
                Case EnuTipoDocOri.EnuNotaRevCr
                    lstrRespuesta = Await SInserteJsNota(lstrJsDoc, True, 0)
                    lstrDoc = "NRcr"
            End Select
            lblnNoHayError = True
        Catch ex As HttpRequestException
            lstrMensError = MhttpRespuesta.ToString()
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensError = ex.ToString()
        Catch ex As Exception
            lstrMensError = MhttpRespuesta.ToString()
        Finally
            If lblnNoHayError Then
#If PRU = 0 Then
                If Not String.IsNullOrEmpty(lstrRespuesta) Then
                    lstrCUDoc = FstrCudoc(lstrRespuesta)
                    SActualiceDoc(aobjDoc, aenuTipoDoc, lstrCUDoc)
                Else
                    SExporteJsonDoc(aobjDoc, aenuTipoDoc)
                    SRegistreError(FstrFechaPro(Now), lstrDoc, aobjDoc.StrIdObjeto,
                            "0-NoInsertado", CMSTRCERO, lstrMensError)
                    SRegistreErrorFtp(aobjDoc, aenuTipoDoc)
                End If
#Else
                lstrCUDoc = "111111111111111"
                SActualiceDoc(aobjDoc, aenuTipoDoc, lstrCUDoc)
#End If
            Else
                If String.IsNullOrEmpty(lstrMens) Then
                    lstrMens = FstrError(lstrMensError) & " SInserteDoc"
                End If
                SRegistreError(FstrFechaPro(Now), lstrDoc, aobjDoc.StrIdObjeto,
                        "0-NoInsertado", CMSTRCERO, lstrMens)
                SExporteJsonDoc(aobjDoc, aenuTipoDoc)
                SRegistreErrorFtp(aobjDoc, aenuTipoDoc)
            End If
        End Try
    End Function
#If PRU = 0 Then
    Friend Async Function SActualiceDoc(aobjDoc As ClsCBObjetoPan,
            aenuTipoDoc As EnuTipoDocOri, ablnPrimer As Boolean) As Task
#Else
    Friend Async Function SActualiceDoc(aobjDoc As ClsCBObjetoPan,
            aenuTipoDoc As EnuTipoDocOri, ablnPrimer As Boolean, ablnActualiza As Boolean) As Task
#End If
        If String.IsNullOrEmpty(MstrToken) Then
            Await FstrObtengaToken(0)
        End If
        If Not String.IsNullOrEmpty(MstrToken) Then
            Dim lblnNoHayError As Boolean, i = 0
            Dim lobjEstadoDoc As ClsEstadoDoc = Nothing
            Dim lstrDoc = aobjDoc.StrIdObjeto
            Dim lstrCudoc = FstrCudocDoc(aobjDoc, aenuTipoDoc)
            If String.IsNullOrEmpty(lstrCudoc) Then
                Throw New ErrorInesperadoPanLException("Doc. " & aobjDoc.StrNombreClase & " sin Cudoc!")
            End If
#If PRU = 0 Then
            Try
                If ablnPrimer Then
                    SEspere(0, 1, 0)
                End If
                lobjEstadoDoc = Await FobjObtengaEstado(lstrCudoc, aenuTipoDoc, lstrDoc)
                lblnNoHayError = True
            Catch ex As HttpRequestException
                Throw
            Catch ex As Exception
                Throw
            Finally
                If lblnNoHayError Then
                    SActualiceEstadoDoc(aobjDoc, lobjEstadoDoc, aenuTipoDoc)
                Else
                    SExporteJsonDoc(aobjDoc, aenuTipoDoc)
                End If
            End Try
        End If
#Else
            SEspere(0, 0, 100)
            lobjEstadoDoc = New ClsEstadoDoc With {
                .CUDE = "ijadjsdpiusd53645646asdhaosdyiah",
                .CUFE = "ewrtyunbvcx5218wdfdjl646313",
                .CustomerPartyID = "564321546",
                .CustomerParty = "AVV",
                .InvoiceNumber = "8465431",
                .DocumentNumber = aobjDoc.StrIdObjeto,
                .StatusDate = "2021-11-04T15:04:31",
                .DIANErrors = Nothing
            }
            If ablnActualiza Then
                lobjEstadoDoc.DocumentStatus = EnuEstadoEDoc.EnuRegi
            Else
                lobjEstadoDoc.DocumentStatus = EnuEstadoEDoc.EnuEnviada
            End If
            SActualiceEstadoDoc(aobjDoc, lobjEstadoDoc, aenuTipoDoc)
            lblnNoHayError = True
        End If
#End If
    End Function
    Private Function FstrJsonDoc(aobjDoc As ClsCBObjetoPan,
            aenuTipoDoc As EnuTipoDocOri) As String
        Dim lstrJsDoc = ""
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                MobjFactura = aobjDoc
                Mobjcliente = MobjFactura.ObjClienteFactura
                lstrJsDoc = FstrJsFactura()
            Case EnuTipoDocOri.EnuNotaCon
                Dim lobjNotaCon As ClsNotaCon = aobjDoc
                lstrJsDoc = FstrJsCredito(lobjNotaCon)
            Case EnuTipoDocOri.EnuNotaCr
                Dim lobjNotaCr As ClsNotaCr = aobjDoc
                lstrJsDoc = FstrJsCredito(lobjNotaCr)
            Case EnuTipoDocOri.EnuNotaDb
                Dim lobjNotaDb As ClsNotaDb = aobjDoc
                lstrJsDoc = FstrJsDebito(lobjNotaDb)
            Case EnuTipoDocOri.EnuNotaRevCr
                Dim lobjNotaRcr As ClsNotaReversionCr = aobjDoc
                lstrJsDoc = FstrJsDebito(lobjNotaRcr)
        End Select
        Return lstrJsDoc
    End Function
    Private Sub SRegistreErrorFtp(aobjDoc As ClsCBObjetoPan, aenuTipoDoc As EnuTipoDocOri)
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                Dim lobjDoc As ClsFactura = aobjDoc
                lobjDoc.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                lobjDoc.ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuErrorFtp
                lobjDoc.SActualice(True)
            Case EnuTipoDocOri.EnuNotaCon
                Dim lobjDoc As ClsNotaCon = aobjDoc
                lobjDoc.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                lobjDoc.ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuErrorFtp
                lobjDoc.SActualice(True)
            Case EnuTipoDocOri.EnuNotaCr
                Dim lobjDoc As ClsNotaCr = aobjDoc
                lobjDoc.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                lobjDoc.ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuErrorFtp
                lobjDoc.SActualice(True)
            Case EnuTipoDocOri.EnuNotaDb
                Dim lobjDoc As ClsNotaDb = aobjDoc
                lobjDoc.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                lobjDoc.ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuErrorFtp
                lobjDoc.SActualice(True)
            Case EnuTipoDocOri.EnuNotaRevCr
                Dim lobjDoc As ClsNotaReversionCr = aobjDoc
                lobjDoc.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                lobjDoc.ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuErrorFtp
                lobjDoc.SActualice(True)
        End Select
    End Sub
    Private Sub SActualiceDoc(aobjDoc As ClsCBObjetoPan, aenuTipoDoc As EnuTipoDocOri,
            astrCudoc As String)
        Dim lenuEstadoEDoc As EnuEstadoEDoc
        If aobjDoc.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            aobjDoc.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        End If
        If String.IsNullOrEmpty(astrCudoc) Then
            lenuEstadoEDoc = EnuEstadoEDoc.EnuNoReg
#If PRU = 1 Then
            aobjDoc.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando
#End If
        Else
#If PRU = 1 Then
            lenuEstadoEDoc = EnuEstadoEDoc.EnuEnProceso
#Else
            lenuEstadoEDoc = EnuEstadoEDoc.EnuEnProceso
#End If
        End If
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                Dim lobjFac As ClsFactura = aobjDoc
                lobjFac.ObjIdEstadoEDocEnt.ObjValorPro = lenuEstadoEDoc
                lobjFac.ObjCUDocStr.ObjValorPro = astrCudoc
            Case EnuTipoDocOri.EnuNotaCon
                Dim lobjNCo As ClsNotaCon = aobjDoc
                lobjNCo.ObjIdEstadoEDocEnt.ObjValorPro = lenuEstadoEDoc
                lobjNCo.ObjCUDocStr.ObjValorPro = astrCudoc
            Case EnuTipoDocOri.EnuNotaCr
                Dim lobjNCr As ClsNotaCr = aobjDoc
                lobjNCr.ObjIdEstadoEDocEnt.ObjValorPro = lenuEstadoEDoc
                lobjNCr.ObjCUDocStr.ObjValorPro = astrCudoc
            Case EnuTipoDocOri.EnuNotaDb
                Dim lobjNDb As ClsNotaDb = aobjDoc
                lobjNDb.ObjIdEstadoEDocEnt.ObjValorPro = lenuEstadoEDoc
                lobjNDb.ObjCUDocStr.ObjValorPro = astrCudoc
            Case EnuTipoDocOri.EnuNotaRevCr
                Dim lobjNRcr As ClsNotaReversionCr = aobjDoc
                lobjNRcr.ObjIdEstadoEDocEnt.ObjValorPro = lenuEstadoEDoc
                lobjNRcr.ObjCUDocStr.ObjValorPro = astrCudoc
        End Select
        aobjDoc.SActualice(True)
        If String.IsNullOrEmpty(astrCudoc) Then
            SExporteJsonDoc(aobjDoc, aenuTipoDoc)
        End If
    End Sub
    Private Function SExporteJsonDoc(aobjDoc As ClsCBObjetoPan, aenuTipoDoc As EnuTipoDocOri)
        Dim lstrJsDoc = ""
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                Dim lobjFac As ClsFactura = aobjDoc
                SExporteJsFactura(lobjFac)
            Case EnuTipoDocOri.EnuNotaCon
                Dim lobjNotaCon As ClsNotaCon = aobjDoc
                SExporteJsCredito(lobjNotaCon)
            Case EnuTipoDocOri.EnuNotaCr
                Dim lobjNotaCr As ClsNotaCr = aobjDoc
                SExporteJsCredito(lobjNotaCr)
            Case EnuTipoDocOri.EnuNotaDb
                Dim lobjNotaDb As ClsNotaDb = aobjDoc
                SExporteJsDebito(lobjNotaDb)
            Case EnuTipoDocOri.EnuNotaRevCr
                Dim lobjNotaRcr As ClsNotaReversionCr = aobjDoc
                SExporteJsDebito(lobjNotaRcr)
        End Select
        Return lstrJsDoc
    End Function
    Private Function SExporteJsonDoc(astrNroDoc As String, aenuTipoDoc As EnuTipoDocOri) As String
        Dim lstrJsDoc = ""
        Dim lstrPrefDoc = ClsPanorama.FstrPrefijoDcto(astrNroDoc)
        Dim lentIdDoc = ClsPanorama.FentIdDcto(astrNroDoc)
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                Dim lobjFac As New ClsFactura()
                lobjFac.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDoc, lentIdDoc})
                SExporteJsFactura(lobjFac)
            Case EnuTipoDocOri.EnuNotaCon
                Dim lobjNotaCon As New ClsNotaCon()
                lobjNotaCon.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDoc, lentIdDoc})
                SExporteJsCredito(lobjNotaCon)
            Case EnuTipoDocOri.EnuNotaCr
                Dim lobjNotaCr As New ClsNotaCr()
                lobjNotaCr.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDoc, lentIdDoc})
                SExporteJsCredito(lobjNotaCr)
            Case EnuTipoDocOri.EnuNotaDb
                Dim lobjNotaDb As New ClsNotaDb()
                lobjNotaDb.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDoc, lentIdDoc})
                SExporteJsDebito(lobjNotaDb)
            Case EnuTipoDocOri.EnuNotaRevCr
                Dim lobjNotaRcr As New ClsNotaReversionCr()
                lobjNotaRcr.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDoc, lentIdDoc})
                SExporteJsDebito(lobjNotaRcr)
        End Select
        Return lstrJsDoc
    End Function
    Private Function FstrCudocDoc(aobjDoc As ClsCBObjetoPan, aenuTipoDoc As EnuTipoDocOri) As String
        Dim lstrCudoc = String.Empty
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                Dim lobjFac As ClsFactura = aobjDoc
                lstrCudoc = lobjFac.ObjCUDocStr.ObjValorPro
            Case EnuTipoDocOri.EnuNotaCon
                Dim lobjNCo As ClsNotaCon = aobjDoc
                lstrCudoc = lobjNCo.ObjCUDocStr.ObjValorPro
            Case EnuTipoDocOri.EnuNotaCr
                Dim lobjNCr As ClsNotaCr = aobjDoc
                lstrCudoc = lobjNCr.ObjCUDocStr.ObjValorPro
            Case EnuTipoDocOri.EnuNotaDb
                Dim lobjNDb As ClsNotaDb = aobjDoc
                lstrCudoc = lobjNDb.ObjCUDocStr.ObjValorPro
            Case EnuTipoDocOri.EnuNotaRevCr
                Dim lobjNRcr As ClsNotaReversionCr = aobjDoc
                lstrCudoc = lobjNRcr.ObjCUDocStr.ObjValorPro
        End Select
        Return lstrCudoc
    End Function
    Private Sub SActualiceEstadoDoc(aobjDoc As ClsCBObjetoPan, aobjEstadoDoc As ClsEstadoDoc,
                aenuTipoDoc As EnuTipoDocOri)
        Dim lenuEstadoEDoc As EnuEstadoEDoc = EnuEstadoEDoc.None
        If aobjDoc.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            aobjDoc.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        End If
#If PRU = 0 Then
        lenuEstadoEDoc = FenuEstadoEDoc(aobjEstadoDoc)
        If lenuEstadoEDoc = EnuEstadoEDoc.EnuErrorFtp Then
            SregistreError(aobjEstadoDoc, aenuTipoDoc, aobjDoc.StrIdObjeto)
            SExporteJsonDoc(aobjDoc, aenuTipoDoc)
        ElseIf lenuEstadoEDoc = EnuEstadoEDoc.EnuInvalida Then
            SregistreError(aobjEstadoDoc, aenuTipoDoc, aobjDoc.StrIdObjeto)
            SExporteJsonDoc(aobjDoc, aenuTipoDoc)
        End If
#Else
        lenuEstadoEDoc = aobjEstadoDoc.DocumentStatus
#End If
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                Dim lobjFac As ClsFactura = aobjDoc
                lobjFac.ObjIdEstadoEDocEnt.ObjValorPro = lenuEstadoEDoc
                lobjFac.ObjCUFEStr.ObjValorPro = If(aobjEstadoDoc IsNot Nothing, aobjEstadoDoc.CUFE, "")
                lobjFac.ObjFechaEmisionEFacStr.ObjValorPro = aobjEstadoDoc.StatusDate
            Case EnuTipoDocOri.EnuNotaCon
                Dim lobjNCo As ClsNotaCon = aobjDoc
                lobjNCo.ObjCUDEStr.ObjValorPro = If(aobjEstadoDoc IsNot Nothing, aobjEstadoDoc.CUDE, "")
                lobjNCo.ObjIdEstadoEDocEnt.ObjValorPro = lenuEstadoEDoc
            Case EnuTipoDocOri.EnuNotaCr
                Dim lobjNCr As ClsNotaCr = aobjDoc
                lobjNCr.ObjCUDEStr.ObjValorPro = If(aobjEstadoDoc IsNot Nothing, aobjEstadoDoc.CUDE, "")
                lobjNCr.ObjIdEstadoEDocEnt.ObjValorPro = lenuEstadoEDoc
            Case EnuTipoDocOri.EnuNotaDb
                Dim lobjNDb As ClsNotaDb = aobjDoc
                lobjNDb.ObjCUDEStr.ObjValorPro = If(aobjEstadoDoc IsNot Nothing, aobjEstadoDoc.CUDE, "")
                lobjNDb.ObjIdEstadoEDocEnt.ObjValorPro = lenuEstadoEDoc
            Case EnuTipoDocOri.EnuNotaRevCr
                Dim lobjNRcr As ClsNotaReversionCr = aobjDoc
                lobjNRcr.ObjCUDEStr.ObjValorPro = If(aobjEstadoDoc IsNot Nothing, aobjEstadoDoc.CUDE, "")
                lobjNRcr.ObjIdEstadoEDocEnt.ObjValorPro = lenuEstadoEDoc
        End Select
        aobjDoc.SActualice(True)
    End Sub
    Friend Async Function SVerifiqueCUFEFact(aobjFac As ClsFactura, ablnPrimer As Boolean) As Task
        If String.IsNullOrEmpty(MstrToken) Then
            Await FstrObtengaToken(0)
        End If
        If Not String.IsNullOrEmpty(MstrToken) Then
            Dim lblnNoHayError As Boolean
            Dim lobjEstadoDoc As ClsEstadoDoc = Nothing
            Dim lstrFac = aobjFac.StrIdObjeto
            Dim lstrCudoc = FstrCudocDoc(aobjFac, EnuTipoDocOri.EnuFactura)
            If String.IsNullOrEmpty(lstrCudoc) Then
                Throw New ErrorInesperadoPanLException("Factura " & lstrFac & " sin Cudoc!")
            End If
            Try
                If ablnPrimer Then
                    SEspere(0, 0, 300)
                End If
                lobjEstadoDoc = Await FobjObtengaEstado(lstrCudoc, EnuTipoDocOri.EnuFactura, lstrFac)
                lblnNoHayError = True
            Catch ex As HttpRequestException
                Throw
            Catch ex As Exception
                Throw
            Finally
                If lblnNoHayError AndAlso lobjEstadoDoc IsNot Nothing Then
                    If aobjFac.StrIdObjeto.Replace("-", "") = lobjEstadoDoc.DocumentNumber Then
                        If aobjFac.ObjCUFEStr.ObjValorPro <> lobjEstadoDoc.CUFE Then
                            aobjFac.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                            aobjFac.ObjCUFEStr.ObjValorPro = lobjEstadoDoc.CUFE
                            aobjFac.SActualice(True)
                            ClsPanorama.SEscribaArchivoError("Inconsistencia en Factura Nro. " &
                                aobjFac.StrNumeroFactura)
                        End If
#If PRU = 0 Then
                    Else
                        ClsPanorama.SEscribaArchivoError("Cudoc no corresponde a Factura Nro. " &
                                aobjFac.StrNumeroFactura)
#End If
                    End If
                Else
                    SExporteJsonDoc(aobjFac, EnuTipoDocOri.EnuFactura)
                End If
            End Try
        End If
    End Function
#End Region
#Region "Subir Facturas a la API"
    Private ReadOnly Property FhttpMisFac As Http
        Get
            If IsNothing(MhttpMisFac) Then
                If FblnHabilitoDll() Then
                    MhttpMisFac = New Chilkat.Http
                End If
            End If
            Return MhttpMisFac
        End Get
    End Property
    Private Function FblnHabilitoDll() As Boolean
        Dim lblnHabilito = False
        Using lChilkatGlobal = New Chilkat.Global
            lblnHabilito = lChilkatGlobal.UnlockBundle("AURVLL.CB1072020_iatEz63A6RA4")
            If Not lblnHabilito Then
                StrMensajeError = lChilkatGlobal.LastErrorText
            End If
        End Using
        Return lblnHabilito
    End Function
    Private Shared Sub SExporteDoc(aobjDoc As ClsCBObjetoPan, aenuTipoDoc As EnuTipoDocOri)
        Dim lobjTerCenUtil As ClsTercero =
                GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjTerceroCentroUtilidad
        Dim lRepOri As New ClsRepOrionCop(GCOBJREGISTRO)
        If Not My.Computer.FileSystem.DirectoryExists(GstrTrayEmails) Then
            My.Computer.FileSystem.CreateDirectory(GstrTrayEmails)
        End If
        SLimpieCarpeta(GstrTrayEmails, "*.*")
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                lRepOri.SExporteFactura(aobjDoc, lobjTerCenUtil)
            Case EnuTipoDocOri.EnuNotaCr
                lRepOri.SExporteNotaCr(aobjDoc, lobjTerCenUtil)
            Case EnuTipoDocOri.EnuNotaDb
                lRepOri.SExporteNotaDb(aobjDoc, lobjTerCenUtil)
            Case EnuTipoDocOri.EnuNotaRevCr
                lRepOri.SExporteNotaRCr(aobjDoc, lobjTerCenUtil)
            Case EnuTipoDocOri.EnuNotaCon
                lRepOri.SExporteNotaCon(aobjDoc, lobjTerCenUtil)
            Case Else
                Throw New ErrorInesperadoPanLException("Tipo documento invalido!")
        End Select
    End Sub
#End Region
#Region "Construccion Json"
#Region "Json Factura"
    ''' <summary>
    ''' Genera el Json de la factura y lo exporta a la carpetas eFacturas
    ''' </summary>
    ''' <param name="aobjFactura"></param>
    Friend Sub SExporteJsFactura(aobjFactura As ClsFactura)
        MobjFactura = aobjFactura
        Mobjcliente = MobjFactura.ObjClienteFactura
        Dim lstrJsFac As String = FstrJsFactura()
        lstrJsFac = JsonConvert.DeserializeObject(lstrJsFac).ToString()
        SEscribaEstruJsonDoc(lstrJsFac, EnuTipoDocOri.EnuFactura,
                MobjFactura.StrNumeroFactura)
    End Sub
    Private Function FstrJsFactura() As String
        Dim lstrJsFac As String = CMSTRINICIODOC & FstrJsInfCliente() & FstrJsInfGralFact() &
                FstrJsInfPago() & FstrJsItemsFac() & FstrJsImptosFac() & FstrJsCargosDsctosFac() &
                FstrJsTotalFactura()
        Return lstrJsFac
    End Function
    ''' <summary>
    ''' Genera el Json de los objetos "InvoiceGeneralInformation", "Delivery" y "AdditionalDocuments"
    ''' correspondientes a la factura
    ''' </summary>
    ''' <returns></returns>
    Private Function FstrJsInfGralFact() As String
        Dim lstrNroResAut As String, lstrFechaFac = String.Empty, lstrFechaVence = String.Empty
        Dim lstrCusomId = FstrComillas("10")
        If MobjFactura.ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuContingencia Then
            lstrNroResAut = FstrComillas(GobjParametros.ObjNumeroResolContiStr.ToString)
        Else
            lstrNroResAut = FstrComillas(MobjFactura.ObjNumeroResolAutoStr.ToString)
        End If
        lstrFechaFac = FstrComillas(FstrFechaPro(MobjFactura.ObjFechaFacturaDtm.ObjValorPro))
        lstrFechaVence = FstrComillas(FstrFechaPro(MobjFactura.ObjFechaVencimientoDtm.ObjValorPro))
        Dim lstrNroFac As String = FstrComillas(MobjFactura.ObjIdFacturaEnt.ToString)
        Dim lstrDiasVence As String = ClsPanorama.FentDiasEntreFechas(
                MobjFactura.ObjFechaFacturaDtm.ObjValorPro,
                MobjFactura.ObjFechaVencimientoDtm.ObjValorPro).ToString
        lstrDiasVence = FstrComillas(lstrDiasVence)
        Dim lstrComen As String = Mobjcliente.ObjComentario_ClienteStr.ObjValorPro
        Dim lstrIdPedido = CMSTRNULO
        If Not String.IsNullOrEmpty(lstrComen) AndAlso lstrComen.Length > 7 Then
            lstrIdPedido = If(lstrComen.Substring(0, 7).ToUpper = "PEDIDO:",
                FstrComillas(lstrComen.Substring(7, lstrComen.Length - 7)), CMSTRNULO)
        End If
        Dim lstrInfGralFac = CMSTRINDENT & My.Resources.InvGralInf &
                CMSTRINDENT & My.Resources.NroResAut & lstrNroResAut & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.NroPreFac & lstrNroFac & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.NroFac & lstrNroFac & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.IssDate & lstrFechaFac & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.DiasVen & lstrDiasVence & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.CodMoneda & CMSTRCODMONEDA & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.ExchRate & CMSTRCERODEC & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.ExchRateDate & CMSTRFECHANULL & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.CustomId & lstrCusomId & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.SalPer & CMSTRNULO & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Nota & CMSTRNULO & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.ExtGR & "true" & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.FecVen & lstrFechaVence & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.AddProFac &
                CMSTRINDENT & "]" & CMSTRCIERREINF &
                CMSTRINDENT & My.Resources.Entrega &
                CMSTRINDENT & My.Resources.AddLin & CMSTRNULO & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.PaisDir & CMSTRNULO & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.CouNam & CMSTRNULO & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.SubCod & CMSTRNULO & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.SubNam & CMSTRNULO & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.CitCod & CMSTRNULO & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.CitNam & CMSTRNULO & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.ConPer & CMSTRNULO & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.DelDate & CMSTRFECHANULL & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.DelCom & CMSTRNULO &
                CMSTRCIERREINF &
                CMSTRINDENT & My.Resources.AddDocs &
                CMSTRINDENT & My.Resources.OrdRef & lstrIdPedido & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.FecOrdCompra & CMSTRFECHANULL & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.DesDocRef & CMSTRNULO & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.FecOrdDespacho & CMSTRFECHANULL & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.RecDocRef & CMSTRNULO & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.FeccAvisRecibo & CMSTRFECHANULL & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.AddDoc
        If MobjFactura.ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuContingencia Then
            lstrInfGralFac &= FstrInfCont()
        Else
            lstrInfGralFac &= vbCrLf
        End If
        lstrInfGralFac &= CMSTRINDENT & "]" & CMSTRCIERREINF
        Return lstrInfGralFac
    End Function
    Private Function FstrInfCont() As String
        Dim lstrInfCon As String
        Dim lstrDocCode As String = FstrComillas("R")
        Dim lentIdInfCon As Integer = MobjFactura.ObjIdInformeCont_FacEnt.ObjValorPro
        Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lentIdInfCon}
        Dim lobjInfCon As New ClsInformeCont(EnuModoInstanciaObjDef.EnuUnico)
        lobjInfCon.SAbra(lobjValorLlave)
        Dim lstrDocNum As String = FstrComillas(lobjInfCon.ObjIdInformeContEnt.ToString)
        Dim lstrTipoDoc As String = FstrComillas(ClsInformeCont.StrNombreInfContingencia)
        Dim lstrFechaRad As String = FstrComillas(FstrFechaPro(lobjInfCon.ObjFechaRadicoDtm.ObjValorPro))
        lstrInfCon = CMSTRINDENT & CMSTRINICIODOC &
                CMSTRINDENT & My.Resources.DocNum & lstrDocNum & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.DocCod & lstrDocCode & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.TipoDoc & lstrTipoDoc & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.IssDate & lstrFechaRad &
                CMSTRINDENT & "}"
        Return lstrInfCon
    End Function
    Private Function FstrJsInfPago() As String
        Dim lstrFormaPago = MobjFactura.ObjIdFormaPagoByt.ToString
        Dim lstrMedioPago = String.Empty
        If MobjFactura.ObjIdFormaPagoByt.ObjValorPro = EnuFormaPago.enuContado Then
            lstrMedioPago = FstrMedioPagoDian(MobjFactura.ObjIdMedioPagoByt.ObjValorPro)
        Else
            lstrMedioPago = FstrComillas("ZZZ")
        End If
        Dim lstrNotaFac = FstrComillas(MobjFactura.ObjPieFacturaDos_FactStr.ToString)
        Dim lstrInfPag = CMSTRINDENT & My.Resources.ResuPago &
                CMSTRINDENT & My.Resources.PayTyp & lstrFormaPago & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.PayMeans & lstrMedioPago & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.PayNot & lstrNotaFac & CMSTRCIERREINF
        Return lstrInfPag
    End Function
    Private Function FstrJsItemsFac() As String
        Dim lstrItemsFac = CMSTRINDENT & My.Resources.InfItems
        Dim lobjItemFac As ClsItemFactura = Nothing
        For i = 1 To MobjFactura.ColItemsFactura.Count
            lobjItemFac = MobjFactura.ColItemsFactura(i)
            lstrItemsFac &= FstrJsInfItemFac(lobjItemFac, i = MobjFactura.ColItemsFactura.Count)
        Next
        lstrItemsFac &= CMSTRINDENT & "],"
        Return lstrItemsFac
    End Function
    Private Function FstrJsInfItemFac(aobjItemFac As ClsItemFactura, ablnUltimo As Boolean) As String
        Dim lstrItemReference = aobjItemFac.StrItemRef
        Dim lstrName = aobjItemFac.StrName
        Dim lstrPrice = aobjItemFac.StrPrice
        Dim lstrLineExtensionAmount = aobjItemFac.StrLineExtensionAmount
        Dim lstrLineAllowanceTotal = CMSTRCERODEC
        Dim lstrLineChargeTotal = CMSTRCERODEC
        Dim lstrLineTotalTaxes = aobjItemFac.StrLineTotalTaxes
        Dim lstrLineTotal = aobjItemFac.StrLineTotal
        Dim lstrMeasureUnitCode = CMSTRUNI
        Dim lstrJsInfGraItemFac As String = FstrJsInfItemFacFac(lstrItemReference, lstrName, lstrPrice,
                lstrLineExtensionAmount, lstrLineAllowanceTotal, lstrLineChargeTotal, lstrLineTotalTaxes,
                lstrLineTotal, lstrMeasureUnitCode)
        Dim lstrJsInfItemFac As String = lstrJsInfGraItemFac &
                CMSTRINDENT & My.Resources.InfImptosFac &
                FstrJsImptosItemFac(aobjItemFac) & FstrJsCargosDsctosItem()
        If Not ablnUltimo Then
            lstrJsInfItemFac &= CMSTRINDENT & "},"
        Else
            lstrJsInfItemFac &= CMSTRINDENT & "}"
        End If
        Return lstrJsInfItemFac
    End Function
    ''' <summary>
    ''' Devuelve el json con la información del item de la factura que es común a la factura y a las notas
    ''' </summary>
    ''' <returns></returns>
    Private Function FstrJsInfItemFacFac(astrItemRef As String, astrName As String,
            astrPrice As String, astrLinExtAmo As String, astrLinAllTot As String, astrLinChrTot As String,
            astrLinTotTax As String, astrLinTot As String, astrMesUniCod As String) As String
        Dim lstrJsInfGralItemFac As String = CMSTRINICIODOC &
                CMSTRINDENT & My.Resources.ItemRef & astrItemRef & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Nombre & astrName & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Cant & CMSTRUNODEC & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Precio & astrPrice & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.LinExtAmo & astrLinExtAmo & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.LinAllTotal & astrLinAllTot & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.LinChrTotal & astrLinChrTot & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.LinTotTax & astrLinTotTax & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.LinTot & astrLinTot & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.MesUniCod & astrMesUniCod & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.FreOfChargeInd & CMSTRFALSO & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Nota & CMSTRNULO & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.AddRef &
                vbCrLf &
                CMSTRINDENT & "]," &
                CMSTRINDENT & My.Resources.AddPro &
                vbCrLf &
                CMSTRINDENT & "],"
        Return lstrJsInfGralItemFac
    End Function
    Private Shared Function FstrJsImptosItemFac(aobjItemFac As ClsItemFactura) As String
        Dim lstrJsImptosItemFac As String = String.Empty
        Dim ldecVlrIva = aobjItemFac.FdecIvaServicio
        If ldecVlrIva > 0 Then
            Dim lstrIdTipoImpto As String = FstrComillas("01")
            Dim lstrTasa As String = Format(aobjItemFac.ObjTarifaIva_ItemFactDbl.ObjValorPro * 100, "#0")
            Dim lstrBaseImpto As String = Format(aobjItemFac.FdecBaseIvaServicio, CMSTRFRMDEC)
            Dim lstrVlrImpto As String = Format(aobjItemFac.FdecIvaServicio, CMSTRFRMDEC)
            lstrJsImptosItemFac &= FstrJsInfImptoItem(lstrIdTipoImpto, lstrTasa, lstrBaseImpto,
                    lstrVlrImpto, CMSTRFALSO, True)
        End If
        If String.IsNullOrEmpty(lstrJsImptosItemFac) Then
            lstrJsImptosItemFac &= vbCrLf
        End If
        lstrJsImptosItemFac &= CMSTRINDENT & "],"
        Return lstrJsImptosItemFac
    End Function
    Private Shared Function FstrJsInfImptoItem(astrIdTipoImpto As String, astrTasa As String,
            astrBaseImpto As String, astrVlrImpto As String, astrEsImpto As String,
            ablnEsElUltimo As Boolean) As String
        Dim lstrInfImtpJson As String = CMSTRINICIODOC &
                CMSTRINDENT & My.Resources.Id & astrIdTipoImpto & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.EsImpto & astrEsImpto & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.BaseImpto & astrBaseImpto & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.VlrImpto & astrVlrImpto & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Tasa & astrTasa & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.BasUniMeas & CMSTRCERODEC & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.PerUniAmo & CMSTRCERODEC &
                CMSTRINDENT
        If ablnEsElUltimo Then
            lstrInfImtpJson &= "}"
        Else
            lstrInfImtpJson &= "},"
        End If
        Return lstrInfImtpJson
    End Function
    Private Shared Function FstrJsCargosDsctosItem()
        Dim lstrJsCrgsDsctsItem = CMSTRINDENT & My.Resources.AllChar &
                CMSTRINDENT & "]"
        Return lstrJsCrgsDsctsItem
    End Function
    Private Function FstrJsImptosFac() As String
        Dim lstrTasasIvaFac() = MobjFactura.FstrIvasFactura
        Dim lstrImpsFac = CMSTRINDENT & My.Resources.InvTaxTot
        Dim lstrIdTipoImpto = CMSTRNULO, lstrTasa = String.Empty, lstrBase = String.Empty, lstrVlrImpto = String.Empty
        Dim i = 0
        If lstrTasasIvaFac.Length > 0 Then
            For Each lstrIva As String In lstrTasasIvaFac
                i += 1
                lstrTasa = lstrIva.Split("&")(0)
                lstrBase = lstrIva.Split("&")(1)
                lstrVlrImpto = lstrIva.Split("&")(2)
                lstrIdTipoImpto = FstrComillas("01")
                lstrImpsFac &= FstrImptoFac(lstrIdTipoImpto, CMSTRFALSO, lstrTasa, lstrBase, lstrVlrImpto,
                        i = lstrTasasIvaFac.Length)
            Next
        Else
            lstrImpsFac &= vbCrLf
        End If
        lstrImpsFac &= CMSTRINDENT & "],"
        Return lstrImpsFac
    End Function
    Private Shared Function FstrImptoFac(astrIdTipoImpto As String, astrEsImpto As String,
            astrTasa As String, astrBase As String, astrVlrImpto As String, ablnUltimo As Boolean) As String
        Dim lstrImpFac = CMSTRINICIODOC &
                    CMSTRINDENT & My.Resources.Id & astrIdTipoImpto & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.EsImpto & astrEsImpto & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.BaseImpto & astrBase & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.VlrImpto & astrVlrImpto & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.Tasa & astrTasa & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.BasUniMeas & CMSTRCERODEC & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.PerUniAmo & CMSTRCERODEC
        If ablnUltimo Then
            lstrImpFac &= CMSTRINDENT & "}"
        Else
            lstrImpFac &= CMSTRINDENT & "},"
        End If
        Return lstrImpFac
    End Function
    Private Shared Function FstrJsCargosDsctosFac() As String
        Dim lstrJsCargosDsctosFac = CMSTRINDENT & My.Resources.InvAllChar &
                CMSTRINDENT & "],"
        Return lstrJsCargosDsctosFac
    End Function
    Private Function FstrJsTotalFactura()
        Dim ldecVlrTotFac As Decimal = MobjFactura.ObjValor_FactDec.ObjValorPro
        Dim ldecVlrPagarSinIva = MobjFactura.FdecValorServicios
        Dim ldecAntAplicado = MobjFactura.FdecAnticipoAplicado
        Dim ldecTotalAPagar = ldecVlrTotFac
        Dim lstrLinExtAmo = Format(ldecVlrPagarSinIva, CMSTRFRMDEC)
        Dim lstrTaxExcAmo = Format(MobjFactura.FdecBaseIvaCapital, CMSTRFRMDEC)
        Dim lstrTaxIncAmo = Format(ldecTotalAPagar, CMSTRFRMDEC)
        Dim lstrAllTotAmo = CMSTRCERODEC
        Dim lstrPrePaiAmo = Format(ldecAntAplicado, CMSTRFRMDEC)
        Dim lstrPayAmo = Format(ldecTotalAPagar, CMSTRFRMDEC)
        Dim lstrJsTotFac = CMSTRINDENT & My.Resources.InvTotal &
                CMSTRINDENT & My.Resources.LinExtAmo & lstrLinExtAmo & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.TaxExcAmo & lstrTaxExcAmo & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.TaxIncAmo & lstrTaxIncAmo & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.AllTotAmo & lstrAllTotAmo & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.ChaTotAmo & CMSTRCERODEC & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.PrePaiAmo & lstrPrePaiAmo & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.PayAmo & lstrPayAmo &
                 CMSTRINDENT & "}" & "}"
        Return lstrJsTotFac
    End Function
#End Region
#Region "Json Nota"
#Region "Js Nota NotaDb"
    ''' <summary>
    ''' Genera el Json y lo escribe en un archivo
    ''' </summary>
    ''' <param name="aobjNotaDb"></param>
    Friend Sub SExporteJsDebito(aobjNotaDb As ClsNotaDb)
        Dim lstrJsNota = FstrJsDebito(aobjNotaDb)
        lstrJsNota = JsonConvert.DeserializeObject(lstrJsNota).ToString()
        SEscribaEstruJsonDoc(lstrJsNota, EnuTipoDocOri.EnuNotaDb,
                aobjNotaDb.StrNumeroNotaDb)
    End Sub
    Private Function FstrJsDebito(aobjNotaDb As ClsNotaDb) As String
        Mobjcliente = aobjNotaDb.ObjClienteNotaDb
        Dim lstrJsInfCliente = FstrJsInfCliente()
        Dim lstrInfGralNota = FstrJsInfGralNotaDb(aobjNotaDb)
        Dim lstrJsInfItems = FstrJsItemsNotaDb(aobjNotaDb)
        Dim lstrJsImpNot As String = FstrJsImptosNDb(aobjNotaDb)
        Dim lstrJsTotalNota As String = FstrJsTotalNotaDb(aobjNotaDb)
        Dim lstrJsNota = CMSTRINICIODOC & lstrJsInfCliente & lstrInfGralNota & lstrJsInfItems &
                lstrJsImpNot & lstrJsTotalNota
        Return lstrJsNota
    End Function
    Private Function FstrJsInfGralNotaDb(aobjNotaDb As ClsNotaDb) As String
        Dim lstrPrefNota As String = FstrComillas(aobjNotaDb.ObjPrefijo_NotaDbStr.ToString())
        Dim lstrNroNota = FstrComillas(aobjNotaDb.ObjIdNotaDbEnt.ToString())
        Dim ldtmFechaNota As Date = aobjNotaDb.ObjFecha_NotaDbDtm.ObjValorPro
        Dim lstrFecFra As String, lstrNroFac As String, lstrCufe As String
        Dim lstrDetalle As String = FstrComillas("Nota de Intereses causados a la fecha")
        Dim lstrConceptoDian = ClsNotaDb.StrTipoConceptoDian
        Dim lstrCustomId As String
        Dim lstrFecFinPer = CMSTRNULO
        Dim lstrFecIniPer As String
        lstrFecIniPer = FstrFechaIniPeriodo(ldtmFechaNota, lstrFecFinPer)
        lstrCustomId = FstrComillas("32")
        lstrCufe = FstrComillas("")
        lstrNroFac = FstrComillas("")
        lstrFecFra = FstrComillas("0001-01-01T00:00:00")
        '
        Dim lstrJsInfGralNDb = FstrJsInfGralNota(lstrPrefNota, lstrNroNota, lstrFecFra,
                lstrDetalle, lstrCufe, lstrNroFac, lstrConceptoDian, lstrCustomId, lstrFecIniPer,
                lstrFecFinPer)
        Return lstrJsInfGralNDb
    End Function
    Private Shared Function FstrJsImptosNDb(aobjNotaDb As ClsNotaDb) As String
        Dim lstrJsImpNotDb As String
        If aobjNotaDb.DecValorIvaNota > 0 Then
            Dim lstrImpuestosNDb As ArrayList = FstrImpuestosNdb(aobjNotaDb)
            If lstrImpuestosNDb.Count > 0 Then
                lstrJsImpNotDb = FstrJsImptosNota(lstrImpuestosNDb)
            Else
                lstrJsImpNotDb = FstrJsImptosNota()
            End If
        Else
            lstrJsImpNotDb = FstrJsImptosNota()
        End If
        Return lstrJsImpNotDb
    End Function
    Private Shared Function FstrImpuestosNdb(aobjNotaDb As ClsNotaDb) As ArrayList
        Dim lstrImptosNDb As New ArrayList
        Dim ldblTarsIvaFac = FdblTarsIvaNdb(aobjNotaDb)
        Dim lstrCodIm As String
        If ldblTarsIvaFac.Count > 0 Then
            lstrCodIm = FstrIdImptoDian(True, EnuTipoDescuentoDef.None)
            Dim lstrInfIva As String, lstrTasa As String
            For Each ldblTar As Double In ldblTarsIvaFac
                ' idimpto, tasa,base,vlrimpto,esret
                lstrTasa = Format(ldblTar * 100, CMSTRFRMDEC)
                lstrInfIva = lstrCodIm & "&" & lstrTasa & "&" &
                        FstrInfIvaNDb(aobjNotaDb, ldblTar) & "&" & CMSTRFALSO
                lstrImptosNDb.Add(lstrInfIva)
            Next
        End If
        Return lstrImptosNDb
    End Function
    Private Shared Function FdblTarsIvaNdb(aobjNotaDb As ClsNotaDb) As ArrayList
        Dim ldblTarIva As Double
        Dim ldblTarsIva As New ArrayList
        For Each lobjItemNDb As ClsItemNotaDb In aobjNotaDb.ColItemsNotaDb
            ldblTarIva = lobjItemNDb.ObjTarifaIva_ItemNotaDbDbl.ObjValorPro
            If ldblTarIva > 0 AndAlso Not ldblTarsIva.Contains(ldblTarIva) Then
                ldblTarsIva.Add(ldblTarIva)
            End If
        Next
        ldblTarsIva.Sort()
        Return ldblTarsIva
    End Function
    Private Shared Function FstrInfIvaNDb(aobjNotadb As ClsNotaDb, adblTar As Double) As String
        Dim lstrInfIvaNDb As String, ldecBase = 0D, ldecVlrIva = 0D
        Dim lstrBase = CMSTRCERODEC, lstrVlrIva = CMSTRCERODEC
        For Each lobjItemNDb As ClsItemNotaDb In aobjNotadb.ColItemsNotaDb
            If lobjItemNDb.ObjTarifaIva_ItemNotaDbDbl.ObjValorPro = adblTar Then
                ldecBase += lobjItemNDb.DecValorAntesIva
                ldecVlrIva += lobjItemNDb.DecValorIva
            End If
        Next
        If ldecBase > 0 Then
            lstrBase = Format(ldecBase, CMSTRFRMDEC)
            lstrVlrIva = Format(ldecVlrIva, CMSTRFRMDEC)
        End If
        lstrInfIvaNDb = lstrBase & "&" & lstrVlrIva
        Return lstrInfIvaNDb
    End Function
    Private Shared Function FstrJsTotalNotaDb(aobjNotaDb As ClsNotaDb) As String
        Dim lstrJsTotNotDb As String
        If aobjNotaDb.DecValorIvaNota > 0 Then
            lstrJsTotNotDb = FstrJsTotNDbIva(aobjNotaDb)
        Else
            Dim lstrVlrNota As String = Format(aobjNotaDb.ObjValor_NotaDbDec.ObjValorPro, CMSTRFRMDEC)
            lstrJsTotNotDb = FstrJsTotalNota(lstrVlrNota, CMSTRCERODEC, lstrVlrNota, lstrVlrNota)
        End If
        Return lstrJsTotNotDb
    End Function
    Private Shared Function FstrJsTotNDbIva(aobjNotaDb As ClsNotaDb) As String
        Dim ldecVlrIva = aobjNotaDb.DecValorIvaNota
        Dim ldecVlrAntesIva = aobjNotaDb.DecValorAntesIva
        Dim ldecVlrBaseGravable = aobjNotaDb.DecValorBaseIvaNota
        Dim lstrVlrAntesImpto As String = Format(aobjNotaDb.DecValorAntesIva)
        Dim lstrVlrBaseGravable As String = Format(ldecVlrBaseGravable, CMSTRFRMDEC)
        Dim lstrVlrTotalNota As String = Format(aobjNotaDb.ObjValor_NotaDbDec.ObjValorPro,
                CMSTRFRMDEC)
        Dim lstrVlrTotalPagar As String = lstrVlrTotalNota
        Dim lstrJsTotNDbIva = FstrJsTotalNota(lstrVlrAntesImpto, lstrVlrBaseGravable,
                lstrVlrTotalNota, lstrVlrTotalPagar)
        Return lstrJsTotNDbIva
    End Function
    Private Sub SAbraFacturaNotaDbIva(aobjNotaDb As ClsNotaDb)
        If IsNothing(MobjFactura) Then
            MobjFactura = New ClsFactura()
        End If
        Dim lstrPrefFac = String.Empty, lentIdFact As Integer, lstrIdFac As String
        Dim larlFacts As New ArrayList
        For Each lobjItemNDb As ClsItemNotaDb In aobjNotaDb.ColItemsNotaDb
            lstrPrefFac = lobjItemNDb.ObjPrefijoFact_ItemNotaDbStr.ObjValorPro
            lentIdFact = lobjItemNDb.ObjIdFactura_ItemNotaDbEnt.ObjValorPro
            lstrIdFac = ClsPanorama.FstrNumeroDcto(lstrPrefFac, lentIdFact)
            If Not larlFacts.Contains(lstrIdFac) Then
                larlFacts.Add(lstrIdFac)
            Else
                Throw New ErrorInesperadoPanLException("Nota Db " & aobjNotaDb.StrNumeroNotaDb &
                        " tiene asociadas dos facturas!")
            End If
        Next
        SAbraFac(lstrPrefFac, lentIdFact)
    End Sub
    Private Function FobjFacturaItemNotaDb(aobjItemNDb As ClsItemNotaDb) As ClsFactura
        Dim lobjFac As New ClsFactura()
        Dim lstrPrefFac As String = aobjItemNDb.ObjPrefijoFact_ItemNotaDbStr.ObjValorPro
        Dim lentIdFac As Integer = aobjItemNDb.ObjIdFactura_ItemNotaDbEnt.ObjValorPro
        Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFac, lentIdFac}
        lobjFac.SAbra(lobjValorLlave)
        Return lobjFac
    End Function
#Region "Js Items Nota Db"
    Private Function FstrJsItemsNotaDb(aobjNotaDb As ClsNotaDb) As String
        Dim lstrJsItemNDb = String.Empty, i = 0
        Dim lstrJsItemTaxInf As String
        Dim lstrJsItemsNDb = CMSTRINDENT & My.Resources.InfItems
        For Each lobjItemNDb As ClsItemNotaDb In aobjNotaDb.ColItemsNotaDb
            i += 1
            lstrJsItemTaxInf = FstrJsInfTaxItemNdb(lobjItemNDb, i = aobjNotaDb.ColItemsNotaDb.Count)
            lstrJsItemNDb = FstrJsItemNDb(lobjItemNDb) & lstrJsItemTaxInf
            If i = aobjNotaDb.ColItemsNotaDb.Count Then
                lstrJsItemNDb &= CMSTRINDENT & "}"
            Else
                lstrJsItemNDb &= CMSTRINDENT & "},"
            End If
            lstrJsItemsNDb &= lstrJsItemNDb
        Next
        lstrJsItemsNDb &= CMSTRINDENT & "],"
        Return lstrJsItemsNDb
    End Function
    Private Shared Function FstrJsInfTaxItemNdb(aobjItemNDb As ClsItemNotaDb,
            ablnEsElUltimo As Boolean) As String
        Dim lstrJsInfTaxItem As String
        If aobjItemNDb.DecValorIva > 0 Then
            Dim lstrIdTax = FstrIdImptoDian(True, EnuTipoDescuentoDef.None)
            Dim lstrTaxEvidInd = CMSTRFALSO
            Dim lstrTaxableAmou = Format(aobjItemNDb.DecValorBaseIva, CMSTRFRMDEC)          ' Base gravable
            Dim lstrTaxAmou = Format(aobjItemNDb.DecValorIva, CMSTRFRMDEC)                   ' Valor Iva
            Dim lstrPerc = Format(aobjItemNDb.ObjTarifaIva_ItemNotaDbDbl.ObjValorPro * 100,
                    CMSTRFRMDEC)
            Dim lstrBaseUnitMe = CMSTRCERODEC
            Dim lstrPerUniAmou = CMSTRCERODEC
            lstrJsInfTaxItem = CMSTRINDENT & My.Resources.TaxInf &
                    CMSTRINICIODOC &
                    CMSTRINDENT & My.Resources.Id & lstrIdTax & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.EsImpto & lstrTaxEvidInd & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.BaseImpto & lstrTaxableAmou & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.VlrImpto & lstrTaxAmou & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.Tasa & lstrPerc & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.BasUniMeas & lstrBaseUnitMe & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.PerUniAmo & lstrPerUniAmou &
                    CMSTRINDENT & "}" &
                    CMSTRINDENT & "]"
        Else
            lstrJsInfTaxItem = CMSTRINDENT & My.Resources.TaxInf &
                CMSTRINDENT & "]"
        End If
        'If ablnEsElUltimo Then
        '    lstrJsInfTaxItem &= CMSTRINDENT & "}"
        'Else
        '    lstrJsInfTaxItem &= CMSTRINDENT & "},"
        'End If
        Return lstrJsInfTaxItem
    End Function
    Private Function FstrJsItemNDb(aobjItemNDb As ClsItemNotaDb) As String
        Dim lobjFactItemNDb = FobjFacturaItemNotaDb(aobjItemNDb)
        Dim lstrIdItemFac = aobjItemNDb.ObjIdItemFac_ItemNotaDbShr.ToString
        Dim lobjItemFac As ClsItemFactura = lobjFactItemNDb.ColItemsFactura(lstrIdItemFac)
        Dim lstrItemRef = lobjItemFac.StrItemRef
        Dim lstrName = lobjItemFac.StrName
        Dim lstrQua = CMSTRUNODEC
        Dim lstrPrice = aobjItemNDb.StrPrice
        Dim lstrLinExtAmo = aobjItemNDb.StrLinExtAmo
        Dim lstrLinTot = aobjItemNDb.StrLinTot
        Dim lstrMeaUniCod = CMSTRUNI
        Dim lstrFreOfChaInd = CMSTRFALSO
        Dim lstrLinAllTotal = CMSTRCERODEC
        Dim lstrLinChaTotal = CMSTRCERODEC
        Dim lstrLineTotalTax = aobjItemNDb.StrLinTotTax
        Dim lstrJsItemNDb = FstrJsItemNota(lstrItemRef, lstrName, lstrQua, lstrPrice, lstrLinExtAmo,
                lstrLinAllTotal, lstrLinChaTotal, lstrLineTotalTax,
                lstrLinTot, lstrMeaUniCod, lstrFreOfChaInd)
        Return lstrJsItemNDb
    End Function
#End Region
#End Region
#Region "Js Nota NotaCr"
    Friend Sub SExporteJsCredito(aobjNotaCr As ClsNotaCr)
        Dim lstrJsNota = FstrJsCredito(aobjNotaCr)
        lstrJsNota = JsonConvert.DeserializeObject(lstrJsNota).ToString()
        SEscribaEstruJsonDoc(lstrJsNota, EnuTipoDocOri.EnuNotaCr,
                    aobjNotaCr.StrNumeroNotaCr)
    End Sub
    Private Function FstrJsCredito(aobjNotaCr As ClsNotaCr) As String
        Mobjcliente = aobjNotaCr.ObjClienteNotaCr
        SAbraFacturaNotaCr(aobjNotaCr)
        Dim lstrJsInfCliente = FstrJsInfCliente() 'CustomerInformation Ok
        Dim lstrInfGralNota = FstrJsInfGralNotaCr(aobjNotaCr) 'NoteGeneralInformation Ok
        Dim lstrJsInfItems = FstrJsItemsNotaCr(aobjNotaCr) 'ItemInformation Ok
        Dim lstrJsImpNot As String = FstrJsImptosNCr(aobjNotaCr) 'Impuestos de la nota Ok
        Dim lstrJsTotalNota As String = FstrJsTotalNotaCr(aobjNotaCr) ' Total de la Nota Ok
        Dim lstrJsNota = CMSTRINICIODOC & lstrJsInfCliente & lstrInfGralNota & lstrJsInfItems &
                lstrJsImpNot & lstrJsTotalNota
        Return lstrJsNota
    End Function
    Private Function FstrJsInfGralNotaCr(aobjNotaCr As ClsNotaCr) As String
        Dim lstrPrefNota As String = FstrComillas(aobjNotaCr.ObjPrefijo_NotaCrStr.ToString())
        Dim lstrNroNota = FstrComillas(aobjNotaCr.ObjIdNotaCrEnt.ToString())
        Dim ldtmFechaNota As Date = aobjNotaCr.ObjFecha_NotaCrDtm.ObjValorPro
        Dim lstrDetalle As String = FstrComillas(aobjNotaCr.ObjComentario_NotaCrStr.ObjValorPro)
        Dim lstrFecFra As String, lstrNroFac As String, lstrCufe As String
        Dim lstrTipoCorreNcr = aobjNotaCr.StrTipoConceptoDian
        Dim lstrCustomId As String
        Dim lstrFecFinPer = CMSTRNULO
        Dim lstrFecIniPer = CMSTRNULO
        If lstrTipoCorreNcr = CType(EnuConceptoNotaCrDian.EnuAnulacion, Integer).ToString Then
            Dim ldtmFechaFac As Date
            If String.IsNullOrEmpty(MobjFactura.ObjFechaEmisionEFacStr.ToString) Then
                ldtmFechaFac = MobjFactura.ObjFechaFacturaDtm.ObjValorPro
            Else
                ldtmFechaFac = MobjFactura.ObjFechaEmisionEFacStr.ObjValorPro
            End If
            lstrFecFra = FstrComillas(FstrFechaPro(ldtmFechaFac))
            lstrNroFac = FstrComillas(MobjFactura.StrNumeroFactura)
            lstrNroFac = lstrNroFac.Replace("-", String.Empty)
            lstrCufe = FstrComillas(MobjFactura.ObjCUFEStr.ObjValorPro)
            lstrCustomId = FstrComillas("20")
        Else
            lstrFecIniPer = FstrFechaIniPeriodo(ldtmFechaNota, lstrFecFinPer)
            lstrCustomId = FstrComillas("22")
            lstrCufe = FstrComillas("")
            lstrNroFac = FstrComillas("")
            lstrFecFra = FstrComillas("0001-01-01T00:00:00")
        End If
        Dim lstrJsInfGralNCr = FstrJsInfGralNota(lstrPrefNota, lstrNroNota,
                    lstrFecFra, lstrDetalle, lstrCufe, lstrNroFac, lstrTipoCorreNcr, lstrCustomId,
                    lstrFecIniPer, lstrFecFinPer)
        Return lstrJsInfGralNCr
    End Function
    Private Shared Function FstrJsImptosNCr(aobjNotaCr As ClsNotaCr) As String
        Dim lstrJsImpNotCr As String
        If aobjNotaCr.ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuAnulaFac Then
            Dim lstrImpuestosNCr As ArrayList = FstrImpuestosNCr(aobjNotaCr)
            If lstrImpuestosNCr.Count > 0 Then
                lstrJsImpNotCr = FstrJsImptosNota(lstrImpuestosNCr)
            Else
                lstrJsImpNotCr = FstrJsImptosNota()
            End If
        Else
            lstrJsImpNotCr = FstrJsImptosNota()
        End If
        Return lstrJsImpNotCr
    End Function
    Private Function FstrJsTotalNotaCr(aobjNotaCr As ClsNotaCr) As String
        Dim lstrJsTotNotCr = String.Empty
        If aobjNotaCr.ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuDescuento Then
            lstrJsTotNotCr = FstrJsTotNcrDscto(aobjNotaCr)
        ElseIf aobjNotaCr.ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuAnulaFac Then
            lstrJsTotNotCr = FstrJsTotNcrAnu()
        End If
        Return lstrJsTotNotCr
    End Function
    Private Function FstrJsTotNcrDscto(aobjNotaCr As ClsNotaCr) As String
        Dim ldecVlrIvaRev = aobjNotaCr.FdecIvaReversado
        Dim ldecVlrAntesIva = aobjNotaCr.DecDsctosNota - ldecVlrIvaRev
        Dim ldecVlrBaseGravable = aobjNotaCr.FdecBaseGrabable(MobjFactura)
        Dim lstrVlrAntesImpto As String = Format(ldecVlrAntesIva, CMSTRFRMDEC)
        Dim lstrVlrBaseGravable As String = Format(ldecVlrBaseGravable, CMSTRFRMDEC)
        Dim lstrVlrTotalNota As String = Format(aobjNotaCr.DecDsctosNota, CMSTRFRMDEC)
        Dim lstrVlrTotalPagar As String = lstrVlrTotalNota
        Dim lstrJsTotDscto = FstrJsTotalNota(lstrVlrAntesImpto, lstrVlrBaseGravable, lstrVlrTotalNota,
                lstrVlrTotalPagar)
        Return lstrJsTotDscto
    End Function
    ' AVV Revisar Debe incluir IVA intereses?
    Private Function FstrJsTotNcrAnu() As String
        Dim ldecVlrIva = MobjFactura.FdecIvaServicios
        Dim ldecVlrAntesIva = MobjFactura.FdecValorServicios
        Dim ldecVlrBaseGravable = MobjFactura.FdecBaseIvaTotal
        Dim lstrVlrAntesImpto As String = Format(ldecVlrAntesIva, CMSTRFRMDEC)
        Dim lstrVlrBaseGravable As String = Format(ldecVlrBaseGravable, CMSTRFRMDEC)
        Dim lstrVlrTotalNota As String = Format(MobjFactura.FdecValorServicios +
                MobjFactura.FdecIvaTotal, CMSTRFRMDEC)
        Dim lstrVlrTotalPagar As String = lstrVlrTotalNota
        Dim lstrJsTotDscto = FstrJsTotalNota(lstrVlrAntesImpto, lstrVlrBaseGravable,
                lstrVlrTotalNota, lstrVlrTotalPagar)
        Return lstrJsTotDscto
    End Function
    Private Sub SAbraFacturaNotaCr(aobjNotaCr As ClsNotaCr)
        If IsNothing(MobjFactura) Then
            MobjFactura = New ClsFactura()
        End If
        Dim lobjItemNCr As ClsItemNotaCr = aobjNotaCr.ColItemsNotaCr("1")
        Dim lstrPrefFac As String = lobjItemNCr.ObjPrefijoFact_ItemNotaCrStr.ObjValorPro
        Dim lentIdFact As Integer = lobjItemNCr.ObjIdFactura_ItemNotaCrEnt.ObjValorPro
        SAbraFac(lstrPrefFac, lentIdFact)
    End Sub
    ''' <summary>
    ''' Devueleve un ArrayList de strings cada uno de los cuales contiene: 
    ''' id del impuesto, tasa, base grabable, valor de impuesto y false separados por ampesand
    ''' </summary>
    ''' <param name="aobjNotaCr"></param>
    ''' <returns></returns>
    Private Shared Function FstrImpuestosNCr(aobjNotaCr As ClsNotaCr) As ArrayList
        Dim lstrImptosNCr As New ArrayList
        Dim ldblTarsIvaFac = FdblTarsIvaFac(aobjNotaCr)
        Dim lstrCodIm As String
        If ldblTarsIvaFac.Count > 0 Then
            lstrCodIm = FstrIdImptoDian(True, EnuTipoDescuentoDef.None)
            Dim lstrInfIva As String, lstrTasa As String
            For Each ldblTar As Double In ldblTarsIvaFac
                ' idimpto, tasa,base,vlrimpto,esret
                lstrTasa = Format(ldblTar * 100, CMSTRFRMDEC)
                lstrInfIva = lstrCodIm & "&" & lstrTasa & "&" & FstrInfIvaNCr(aobjNotaCr, ldblTar) &
                        "&" & CMSTRFALSO
                lstrImptosNCr.Add(lstrInfIva)
            Next
        End If
        Return lstrImptosNCr
    End Function
    Private Shared Function FdblTarsIvaFac(aobjNotaCr As ClsNotaCr) As ArrayList ' No las tarifas deben ser sacadas de los items NCr
        Dim ldblTarIva As Double
        Dim ldblTarsIva As New ArrayList
        For Each lobjItemNCr As ClsItemNotaCr In aobjNotaCr.ColItemsNotaCr
            If lobjItemNCr.ObjEsReversionIvaBln.ObjValorPro Then
                ldblTarIva = lobjItemNCr.ObjItemFac.ObjTarifaIva_ItemFactDbl.ObjValorPro
                If ldblTarIva > 0 AndAlso Not ldblTarsIva.Contains(ldblTarIva) Then
                    ldblTarsIva.Add(ldblTarIva)
                End If
            End If
        Next
        ldblTarsIva.Sort()
        Return ldblTarsIva
    End Function
    ''' <summary>
    '''  Devuelve un string que contiene la base gravable y el valor del iva separador por el
    '''  caracter ampersand y es utilizado solo para la anulación de la factura
    ''' </summary>
    ''' <param name="aobjNotaCr">Nota crédito que se genera al anular la factura</param>
    ''' <param name="adblTar">Tarifa del iva urilizado por cada uno de los items de factura</param>
    ''' <returns></returns>
    Private Shared Function FstrInfIvaNCr(aobjNotaCr As ClsNotaCr, adblTar As Double) As String
        Dim lstrInfIvaNCr As String, ldecBase = 0D, ldecVlrIva = 0D
        Dim lstrBase = CMSTRCERODEC, lstrVlrIva = CMSTRCERODEC
        For Each lobjItemFac As ClsItemFactura In aobjNotaCr.ObjFacturaAfectada.ColItemsFactura
            If lobjItemFac.ObjTarifaIva_ItemFactDbl.ObjValorPro = adblTar Then
                ldecBase += lobjItemFac.FdecBaseIvaServicio + lobjItemFac.FdecBaseIvaInt
                ldecVlrIva += lobjItemFac.FdecIvaServicio + lobjItemFac.FdecIvaInt
            End If
        Next
        If ldecBase > 0 Then
            lstrBase = Format(ldecBase, CMSTRFRMDEC)
            lstrVlrIva = Format(ldecVlrIva, CMSTRFRMDEC)
        End If
        lstrInfIvaNCr = lstrBase & "&" & lstrVlrIva
        Return lstrInfIvaNCr
    End Function
#Region "Js Items Nota Cr"
    Private Function FstrJsItemsNotaCr(aobjNotaCr As ClsNotaCr) As String
        Dim lstrJsItemsNCr = String.Empty
        If aobjNotaCr.ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuDescuento Then
            lstrJsItemsNCr = FstrJsItemsNotaCrDscto(aobjNotaCr)
        ElseIf aobjNotaCr.ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuAnulaFac Then
            lstrJsItemsNCr = FstrJsItemsNotaCrAnu()
        End If
        Return lstrJsItemsNCr
    End Function
    Private Function FstrJsItemsNotaCrDscto(aobjNotaCR As ClsNotaCr) As String ' OK
        Dim larlTiposDscto = FarlTiposDscto(aobjNotaCR), i = 0
        Dim lstrJsItemNCr = String.Empty
        Dim lstrJsItemTaxInf = CMSTRINDENT & My.Resources.TaxInf &
                CMSTRINDENT & "]"
        Dim lstrJsItemsNCr = CMSTRINDENT & My.Resources.InfItems
        For Each lenuTipoDscto As EnuTipoDescuentoDef In larlTiposDscto
            i += 1
            lstrJsItemsNCr &= FstrJsItemDsctoNCr(aobjNotaCR, lenuTipoDscto, i) & lstrJsItemTaxInf
            If i = larlTiposDscto.Count Then
                lstrJsItemNCr = CMSTRINDENT & "}"
            Else
                lstrJsItemNCr = CMSTRINDENT & "},"
            End If
            lstrJsItemsNCr &= lstrJsItemNCr
        Next
        lstrJsItemsNCr &= CMSTRINDENT & "],"
        Return lstrJsItemsNCr
    End Function
    Private Shared Function FarlTiposDscto(aobjNotaCr As ClsNotaCr) As ArrayList
        Dim larlTiposDscto As New ArrayList
        Dim lenuTipoDscto As EnuTipoDescuentoDef
        For Each lobjItemNcr As ClsItemNotaCr In aobjNotaCr.ColItemsNotaCr
            If lobjItemNcr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro =
                    EnuTipoDescuentoDef.EnuDsctoCapital OrElse
                    lobjItemNcr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro =
                    EnuTipoDescuentoDef.EnuDsctoIntMora OrElse
                    lobjItemNcr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro =
                    EnuTipoDescuentoDef.EnuDsctoPP Then
                lenuTipoDscto = lobjItemNcr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
                If Not larlTiposDscto.Contains(lenuTipoDscto) Then
                    larlTiposDscto.Add(lenuTipoDscto)
                End If
            End If
        Next
        larlTiposDscto.Sort()
        Return larlTiposDscto
    End Function
    Private Function FstrJsItemDsctoNCr(aobjNotaCr As ClsNotaCr, aenuTipoDscto As EnuTipoDescuentoDef,
            aentItemDscto As Integer) As String
        Dim lstrName As String
        Dim lstrQuan = CMSTRUNODEC
        Dim lstrPrice = aobjNotaCr.FdecValorDescuento(aenuTipoDscto)
        Select Case aenuTipoDscto
            Case EnuTipoDescuentoDef.EnuDsctoCapital
                lstrName = "Descuento a Capital"
            Case EnuTipoDescuentoDef.EnuDsctoPP
                lstrName = "Descuento a Capital por Pronto Pago"
            Case EnuTipoDescuentoDef.EnuDsctoIntMora
                lstrName = "Decuento a Intereses de Mora"
            Case Else
                Throw New ErrorInesperadoPanLException("Tipo de descuento no valido!")
        End Select
        lstrName = FstrComillas(lstrName)
        Dim lstrLinExtAmo = lstrPrice
        Dim lstrLinTot = lstrPrice
        Dim lstrMeaUniCod = CMSTRUNI
        Dim lstrFreOfChaInd = CMSTRFALSO
        Dim lstrLinAllTotal = CMSTRCERODEC
        Dim lstrLinChaTotal = CMSTRCERODEC
        Dim lstrLineTotalTax = CMSTRCERODEC
        Dim lstrJsonDscto = FstrJsItemNota(aentItemDscto, lstrName, lstrQuan, lstrPrice,
                lstrLinExtAmo, lstrLinAllTotal, lstrLinChaTotal, lstrLineTotalTax,
                lstrLinTot, lstrMeaUniCod, lstrFreOfChaInd)
        Return lstrJsonDscto
    End Function
    Private Function FstrJsItemsNotaCrAnu() As String
        Dim lstrJsItemsNCr = CMSTRINDENT & My.Resources.InfItems
        Dim lstrJsItFac = String.Empty, lblnItemIva = False
        Dim lblnEsElUltimo = False, ldblTarIva = 0.0
        For Each lobjItemFac As ClsItemFactura In MobjFactura.ColItemsFactura
            lblnItemIva = lobjItemFac.ObjTarifaIva_ItemFactDbl.ObjValorPro > 0
            If lblnItemIva Then
                ldblTarIva = lobjItemFac.ObjTarifaIva_ItemFactDbl.ObjValorPro
                Exit For
            End If
        Next
        For Each lobjItemFac As ClsItemFactura In MobjFactura.ColItemsFactura
            If Not lblnItemIva Then
                lblnEsElUltimo = (lobjItemFac.ObjIdItemFacturaShr.ObjValorPro =
                    MobjFactura.ColItemsFactura.Count)
            End If
            lstrJsItFac &= FstrJsItemFacNCrAnu(lobjItemFac, lblnEsElUltimo)
        Next
        If lblnItemIva Then
            lstrJsItFac &= FstrJsItemIntNCrAnu(ldblTarIva)
        End If
        lstrJsItemsNCr &= lstrJsItFac & CMSTRINDENT & "],"
        Return lstrJsItemsNCr
    End Function
    Private Function FstrJsItemFacNCrAnu(aobjItemFac As ClsItemFactura, ablnEsElUltimo As Boolean) As String
        Dim lstrItemRef = aobjItemFac.StrItemRef
        Dim lstrName = aobjItemFac.StrName
        Dim lstrQuan = CMSTRUNODEC
        ' lstrPrice = total de debitos a capital + total debitos a intereses que causaron IVA
        Dim lstrPrice As String = aobjItemFac.ObjValor_ItemFactDec.ObjValorPro -
                aobjItemFac.FdecIvaServicio
        Dim lstrLinExtAmo As String = lstrPrice
        ' TlstrLinTot = lstrprice + ivas
        Dim lstrLinTot As String = aobjItemFac.ObjValor_ItemFactDec.ObjValorPro
        Dim lstrMeaUniCod = CMSTRUNI
        Dim lstrFreOfChaInd = CMSTRFALSO
        Dim lstrLinAllTotal = CMSTRCERODEC
        Dim lstrLinChaTotal = CMSTRCERODEC
        Dim lstrLineTotalTax As String = aobjItemFac.FdecIvaServicio
        Dim lstrJsonDscto = FstrJsItemNota(lstrItemRef, lstrName, lstrQuan, lstrPrice,
                lstrLinExtAmo, lstrLinAllTotal, lstrLinChaTotal, lstrLineTotalTax,
                lstrLinTot, lstrMeaUniCod, lstrFreOfChaInd)
        Dim lstrJsTaxItem = FstrJsInfTaxItem(aobjItemFac, ablnEsElUltimo)
        lstrJsonDscto &= lstrJsTaxItem
        Return lstrJsonDscto
    End Function
    Private Shared Function FstrJsInfTaxItem(aobjItemFac As ClsItemFactura,
            ablnEsElUltimo As Boolean) As String
        Dim lstrJsInfTaxItem As String
        If aobjItemFac.FdecIvaServicio > 0 Then
            Dim lstrIdTax = FstrIdImptoDian(True, EnuTipoDescuentoDef.None)
            Dim lstrTaxEvidInd = CMSTRFALSO
            Dim lstrTaxableAmou = Format(aobjItemFac.FdecBaseIvaServicio, CMSTRFRMDEC)          ' Base gravable
            Dim lstrTaxAmou = Format(aobjItemFac.FdecIvaServicio, CMSTRFRMDEC)                   ' Valor Iva
            Dim lstrPerc = Format(aobjItemFac.ObjTarifaIva_ItemFactDbl.ObjValorPro * 100, CMSTRFRMDEC)
            Dim lstrBaseUnitMe = CMSTRCERODEC
            Dim lstrPerUniAmou = CMSTRCERODEC
            lstrJsInfTaxItem = CMSTRINDENT & My.Resources.TaxInf &
                    CMSTRINICIODOC &
                    CMSTRINDENT & My.Resources.Id & lstrIdTax & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.EsImpto & lstrTaxEvidInd & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.BaseImpto & lstrTaxableAmou & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.VlrImpto & lstrTaxAmou & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.Tasa & lstrPerc & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.BasUniMeas & lstrBaseUnitMe & CMSTRFINCAMPO &
                    CMSTRINDENT & My.Resources.PerUniAmo & lstrPerUniAmou &
                    CMSTRINDENT & "}" &
                    CMSTRINDENT & "]"
        Else
            lstrJsInfTaxItem = CMSTRINDENT & My.Resources.TaxInf &
                CMSTRINDENT & "]"
        End If
        If ablnEsElUltimo Then
            lstrJsInfTaxItem &= CMSTRINDENT & "}"
        Else
            lstrJsInfTaxItem &= CMSTRINDENT & "},"
        End If
        Return lstrJsInfTaxItem
    End Function
    Private Function FstrJsItemIntNCrAnu(adblTarifaIva As Double) As String
        Dim lstrItemRef = FstrComillas(GCSHRIDMORA.ToString())
        Dim lstrName = FstrComillas("Interese de Mora")
        Dim lstrQuan = CMSTRUNODEC
        ' lstrPrice = total base Iva
        Dim lstrPrice As String = MobjFactura.FdecBaseIvaInt
        Dim lstrLinExtAmo As String = lstrPrice
        ' TlstrLinTot = lstrprice + ivas
        Dim lstrLinTot As String = MobjFactura.FdecBaseIvaInt + MobjFactura.FdecIvaInt
        Dim lstrMeaUniCod = CMSTRUNI
        Dim lstrFreOfChaInd = CMSTRFALSO
        Dim lstrLinAllTotal = CMSTRCERODEC
        Dim lstrLinChaTotal = CMSTRCERODEC
        Dim lstrLineTotalTax As String = MobjFactura.FdecIvaInt
        Dim lstrJsonDscto = FstrJsItemNota(lstrItemRef, lstrName, lstrQuan, lstrPrice,
                lstrLinExtAmo, lstrLinAllTotal, lstrLinChaTotal, lstrLineTotalTax,
                lstrLinTot, lstrMeaUniCod, lstrFreOfChaInd)
        Dim lstrJsTaxItem = FstrJsInfTaxItemInt(adblTarifaIva)
        lstrJsonDscto &= lstrJsTaxItem
        Return lstrJsonDscto
    End Function
    Private Function FstrJsInfTaxItemInt(adblTarifaIva As Double) As String
        Dim lstrJsInfTaxItem As String
        Dim lstrIdTax = FstrIdImptoDian(True, EnuTipoDescuentoDef.None)
        Dim lstrTaxEvidInd = CMSTRFALSO
        Dim lstrTaxableAmou = Format(MobjFactura.FdecBaseIvaInt, CMSTRFRMDEC)   ' Base gravable
        Dim lstrTaxAmou = Format(MobjFactura.FdecIvaInt, CMSTRFRMDEC)           ' Valor Iva
        Dim lstrPerc = Format(adblTarifaIva * 100, CMSTRFRMDEC)
        Dim lstrBaseUnitMe = CMSTRCERODEC
        Dim lstrPerUniAmou = CMSTRCERODEC
        lstrJsInfTaxItem = CMSTRINDENT & My.Resources.TaxInf &
                CMSTRINICIODOC &
                CMSTRINDENT & My.Resources.Id & lstrIdTax & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.EsImpto & lstrTaxEvidInd & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.BaseImpto & lstrTaxableAmou & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.VlrImpto & lstrTaxAmou & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Tasa & lstrPerc & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.BasUniMeas & lstrBaseUnitMe & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.PerUniAmo & lstrPerUniAmou &
                CMSTRINDENT & "}" &
                CMSTRINDENT & "]"
        lstrJsInfTaxItem &= CMSTRINDENT & "}"
        Return lstrJsInfTaxItem
    End Function
#End Region
#End Region
#Region "Js NotaRcr"
    ''' <summary>
    ''' Genera el Json y lo escribe en un archivo
    ''' </summary>
    ''' <param name="aobjNotaRcr"></param>
    ''' <returns></returns>
    Friend Sub SExporteJsDebito(aobjNotaRcr As ClsNotaReversionCr)
        Dim lstrJsNota = FstrJsDebito(aobjNotaRcr)
        lstrJsNota = JsonConvert.DeserializeObject(lstrJsNota).ToString()
        SEscribaEstruJsonDoc(lstrJsNota, EnuTipoDocOri.EnuNotaRevCr,
                    aobjNotaRcr.StrNumeroNotaReversaCr)
    End Sub
    Private Function FstrJsDebito(aobjNotaRcr As ClsNotaReversionCr) As String
        Mobjcliente = aobjNotaRcr.ObjClienteNota
        SAbraFacturaNotaRCr(aobjNotaRcr)
        Dim lstrJsInfCliente = FstrJsInfCliente()
        Dim lstrInfGralNota = FstrJsInfGralNotaDb(aobjNotaRcr)
        Dim lstrJsInfItem = FstrJsItemsNotaDb(aobjNotaRcr)
        Dim lstrJsImpNot As String = FstrJsImptosNota()
        Dim lstrJsTotalNota As String = FstrJsTotalNotaRCr(aobjNotaRcr)
        Dim lstrJsNota = CMSTRINICIODOC & lstrJsInfCliente & lstrInfGralNota & lstrJsInfItem &
                lstrJsImpNot & lstrJsTotalNota
        Return lstrJsNota
    End Function
    Private Function FstrJsInfGralNotaDb(aobjNotaRCr As ClsNotaReversionCr) As String
        Dim lstrPrefNota As String = aobjNotaRCr.ObjPrefijo_NotaReversaCrStr.ToString()
        If String.IsNullOrEmpty(lstrPrefNota) Then
            lstrPrefNota = "RCR"
        End If
        lstrPrefNota = FstrComillas(lstrPrefNota)
        Dim lstrNroNota = FstrComillas(aobjNotaRCr.ObjIdNotaReversaCrEnt.ToString())
        Dim lstrDetalle As String = FstrComillas(aobjNotaRCr.ObjDetalle_NotaReversaCrStr.ObjValorPro)
        Dim lstrFecFra As String, lstrNroFac As String, lstrCufe As String
        Dim lstrTipoCorreNrcr = "4"
        Dim lstrCustomId As String
        Dim lstrFecFinPer = String.Empty
        Dim lstrFecIniPer As String = FstrFechaIniPeriodo(
                    MobjFactura.ObjFechaEmisionEFacStr.ObjValorPro, lstrFecFinPer)
        lstrCustomId = FstrComillas("32")
        lstrCufe = FstrComillas("")
        lstrNroFac = FstrComillas("")
        lstrFecFra = FstrComillas("0001-01-01T00:00:00")
        Dim lstrJsInfGralNCr = FstrJsInfGralNota(lstrPrefNota, lstrNroNota,
                lstrFecFra, lstrDetalle, lstrCufe, lstrNroFac, lstrTipoCorreNrcr, lstrCustomId,
                lstrFecIniPer, lstrFecFinPer)
        Return lstrJsInfGralNCr
    End Function
    Private Function FstrJsItemsNotaDb(aobjNotaRCr As ClsNotaReversionCr) As String
        Dim larlItemsFacAfectado = FarlItemsFacConDsctoRev(aobjNotaRCr), i = 0
        Dim lstrJsItemNCr = String.Empty
        Dim lstrJsItemTaxInf = CMSTRINDENT & My.Resources.TaxInf &
                CMSTRINDENT & "]"
        Dim lstrJsItemsNCr = CMSTRINDENT & My.Resources.InfItems
        For Each lstrIdItemFac As String In larlItemsFacAfectado
            i += 1
            lstrJsItemsNCr &= FstrJsItemFacNCr(aobjNotaRCr, lstrIdItemFac) & lstrJsItemTaxInf
            If i = larlItemsFacAfectado.Count Then
                lstrJsItemNCr = CMSTRINDENT & "}"
            Else
                lstrJsItemNCr = CMSTRINDENT & "},"
            End If
            lstrJsItemsNCr &= lstrJsItemNCr
        Next
        lstrJsItemsNCr &= CMSTRINDENT & "],"
        Return lstrJsItemsNCr
    End Function
    Private Shared Function FstrJsTotalNotaRCr(aobjNotaRCr As ClsNotaReversionCr) As String
        Dim ldecVlrAntesIva = aobjNotaRCr.FdecDsctosNota
        Dim ldecVlrBaseGravable = 0D
        Dim lstrVlrAntesImpto As String = Format(ldecVlrAntesIva, CMSTRFRMDEC)
        Dim lstrVlrBaseGravable As String = Format(ldecVlrBaseGravable, CMSTRFRMDEC)
        Dim lstrVlrTotalNota As String = Format(ldecVlrAntesIva, CMSTRFRMDEC)
        Dim lstrVlrTotalPagar As String = lstrVlrTotalNota
        Dim lstrJsTotNotDb = FstrJsTotalNota(lstrVlrAntesImpto, lstrVlrBaseGravable, lstrVlrTotalNota,
                lstrVlrTotalPagar)
        Return lstrJsTotNotDb
    End Function
    Private Sub SAbraFacturaNotaRCr(aobjNotaRCr As ClsNotaReversionCr)
        If IsNothing(MobjFactura) Then
            MobjFactura = New ClsFactura()
        End If
        Dim lobjNov As ClsNovedad = aobjNotaRCr.ColNovedades(1)
        Dim lstrPrefFac As String = lobjNov.ObjPrefijoFact_NovStr.ObjValorPro
        Dim lentIdFact As Integer = lobjNov.ObjIdFactura_NovEnt.ObjValorPro
        SAbraFac(lstrPrefFac, lentIdFact)
    End Sub
    ''' <summary>
    ''' Devuelve el array de los Id de los items de factura afectados con dsctos por la NRCr
    ''' </summary>
    Private Shared Function FarlItemsFacConDsctoRev(aobjNotaRCr As ClsNotaReversionCr) As ArrayList
        Dim larlItems As New ArrayList
        Dim lstrIdItemFac As String
        For Each lobjNov As ClsNovedad In aobjNotaRCr.ColNovedades
            If lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuRCrDctoCap OrElse
                    lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuRCrDctoInt OrElse
                    lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuRCrIvaGas Then
                lstrIdItemFac = lobjNov.ObjIdItemFact_NovShr.ToString
                If Not larlItems.Contains(lstrIdItemFac) Then
                    larlItems.Add(lstrIdItemFac)
                End If
            End If
        Next
        larlItems.Sort()
        Return larlItems
    End Function
    Private Function FstrJsItemFacNCr(aobjNotaRCr As ClsNotaReversionCr, astrIdItemFac As String) As String
        Dim lobjItemFac As ClsItemFactura = MobjFactura.ColItemsFactura(astrIdItemFac)
        Dim lstrItemRef = lobjItemFac.StrItemRef
        Dim lstrName = lobjItemFac.StrName
        Dim lstrQuan = CMSTRUNODEC
        Dim lstrPrice = aobjNotaRCr.FdecDsctoRevItemFac(astrIdItemFac)
        Dim lstrLinExtAmo = lstrPrice
        Dim lstrLinTot = lstrPrice
        Dim lstrMeaUniCod = CMSTRUNI
        Dim lstrFreOfChaInd = CMSTRFALSO
        Dim lstrLinAllTotal = CMSTRCERODEC
        Dim lstrLinChaTotal = CMSTRCERODEC
        Dim lstrLineTotalTax = CMSTRCERODEC
        Dim lstrJsonDscto = FstrJsItemNota(lstrItemRef, lstrName, lstrQuan, lstrPrice,
                lstrLinExtAmo, lstrLinAllTotal, lstrLinChaTotal, lstrLineTotalTax,
                lstrLinTot, lstrMeaUniCod, lstrFreOfChaInd)
        Return lstrJsonDscto
    End Function
#End Region
#Region "Js Nota NotaCon"
    Friend Sub SExporteJsCredito(aobjNotaCon As ClsNotaCon)
        Dim lstrJsNota = FstrJsCredito(aobjNotaCon)
        lstrJsNota = JsonConvert.DeserializeObject(lstrJsNota).ToString()
        SEscribaEstruJsonDoc(lstrJsNota, EnuTipoDocOri.EnuNotaCon,
                    aobjNotaCon.StrNumeroNotaCon)
    End Sub
    Private Function FstrJsCredito(aobjNotaCon As ClsNotaCon) As String
        Mobjcliente = aobjNotaCon.ObjClienteNotaCon
        SAbraFacturaNotaCon(aobjNotaCon)
        Dim lstrJsInfCliente = FstrJsInfCliente() 'CustomerInformation Ok
        Dim lstrInfGralNota = FstrJsInfGralNotaCon(aobjNotaCon)
        Dim lstrJsInfItems = FstrJsItemsNotaConDscto(aobjNotaCon)
        Dim lstrJsImpNot As String = FstrJsImptosNota()
        Dim lstrJsTotalNota As String = FstrJsTotNCon(aobjNotaCon)
        Dim lstrJsNota = CMSTRINICIODOC & lstrJsInfCliente & lstrInfGralNota & lstrJsInfItems &
                lstrJsImpNot & lstrJsTotalNota
        SEscribaEstruJsonDoc(lstrJsNota, EnuTipoDocOri.EnuNotaCon,
                    aobjNotaCon.StrNumeroNotaCon)
        Return lstrJsNota
    End Function
    Private Function FstrJsInfGralNotaCon(aobjNotaCon As ClsNotaCon) As String
        Dim lstrPrefNota As String = aobjNotaCon.ObjPrefijo_NotaConStr.ToString()
        Dim lstrNroNota = FstrComillas(aobjNotaCon.ObjIdNotaConEnt.ToString())
        If String.IsNullOrEmpty(lstrPrefNota) Then
            lstrPrefNota = "NCON"
        End If
        lstrPrefNota = FstrComillas(lstrPrefNota)
        Dim lstrFecNot As String = FstrComillas(FstrFechaPro(aobjNotaCon.ObjFecha_NotaConDtm.ObjValorPro))
        Dim lstrFecFra As String
        If String.IsNullOrEmpty(MobjFactura.ObjFechaEmisionEFacStr.ToString) Then
            lstrFecFra = FstrComillas(FstrFechaPro(MobjFactura.ObjFechaFacturaDtm.ObjValorPro))
        Else
            lstrFecFra = FstrComillas(FstrFechaPro(MobjFactura.ObjFechaEmisionEFacStr.ObjValorPro))
        End If
        Dim lstrDetalle As String = FstrComillas("Ajuste a la Cuota de Administración")
        Dim lstrNroFac = FstrComillas(MobjFactura.StrNumeroFactura)
        lstrNroFac = lstrNroFac.Replace("-", String.Empty)
        Dim lstrCufe As String = FstrComillas(MobjFactura.ObjCUFEStr.ObjValorPro)
        Dim lstrTipoCorreNcr = CType(EnuConceptoNotaCrDian.EnuDevolucionParcial, Integer).ToString
        Dim lstrCustomId = FstrComillas("20")
        Dim lstrJsInfGralNCr = FstrJsInfGralNota(lstrPrefNota, lstrNroNota,
                lstrFecFra, lstrDetalle, lstrCufe, lstrNroFac, lstrTipoCorreNcr, lstrCustomId,
                CMSTRNULO, CMSTRNULO)
        Return lstrJsInfGralNCr
    End Function
    Private Shared Function FstrJsTotNCon(aobjNotaCon As ClsNotaCon) As String
        Dim ldecVlrAntesIva = aobjNotaCon.ObjValor_NotaConDec.ObjValorPro
        Dim ldecVlrBaseGravable = 0
        Dim lstrVlrAntesImpto As String = Format(ldecVlrAntesIva, CMSTRFRMDEC)
        Dim lstrVlrBaseGravable As String = Format(ldecVlrBaseGravable, CMSTRFRMDEC)
        Dim lstrVlrTotalNota As String = Format(ldecVlrAntesIva, CMSTRFRMDEC)
        Dim lstrVlrTotalPagar As String = lstrVlrTotalNota
        Dim lstrJsTotDscto = FstrJsTotalNota(lstrVlrAntesImpto, lstrVlrBaseGravable, lstrVlrTotalNota,
                lstrVlrTotalPagar)
        Return lstrJsTotDscto
    End Function
    Private Sub SAbraFacturaNotaCon(aobjNotaCon As ClsNotaCon)
        If IsNothing(MobjFactura) Then
            MobjFactura = New ClsFactura()
        End If
        Dim lobjItemNCon As ClsItemNotaCon = aobjNotaCon.ColItemsNotaCon("1")
        Dim lstrPrefFac As String = lobjItemNCon.ObjPrefijoFact_ItemNotaConStr.ObjValorPro
        Dim lentIdFact As Integer = lobjItemNCon.ObjIdFactura_ItemNotaConEnt.ObjValorPro
        SAbraFac(lstrPrefFac, lentIdFact)
    End Sub
#Region "Js Items Nota Con"
    Private Function FstrJsItemsNotaConDscto(aobjNotaCon As ClsNotaCon) As String
        Dim larlItemsFacAfectado = FarlItemsFacConDscto(aobjNotaCon), i = 0
        Dim lstrJsItemNCr = String.Empty
        Dim lstrJsItemTaxInf = CMSTRINDENT & My.Resources.TaxInf &
                CMSTRINDENT & "]"
        Dim lstrJsItemsNCon = CMSTRINDENT & My.Resources.InfItems
        For Each lstrIdItemFac As String In larlItemsFacAfectado
            i += 1
            lstrJsItemsNCon &= FstrJsItemFacNCon(aobjNotaCon, lstrIdItemFac) & lstrJsItemTaxInf
            If i = larlItemsFacAfectado.Count Then
                lstrJsItemNCr = CMSTRINDENT & "}"
            Else
                lstrJsItemNCr = CMSTRINDENT & "},"
            End If
            lstrJsItemsNCon &= lstrJsItemNCr
        Next
        lstrJsItemsNCon &= CMSTRINDENT & "],"
        Return lstrJsItemsNCon
    End Function
    ''' <summary>
    ''' Devuelve el array de los Id de los items de factra afectados por la NCr
    ''' </summary>
    Private Shared Function FarlItemsFacConDscto(aobjNotaCon As ClsNotaCon) As ArrayList
        Dim larlItems As New ArrayList
        Dim lstrIdItemFac As String
        For Each lobjItemNCon As ClsItemNotaCon In aobjNotaCon.ColItemsNotaCon
            lstrIdItemFac = lobjItemNCon.ObjIdItemFac_ItemNotaConShr.ToString
            If Not larlItems.Contains(lstrIdItemFac) Then
                larlItems.Add(lstrIdItemFac)
            End If
        Next
        larlItems.Sort()
        Return larlItems
    End Function
    Private Function FstrJsItemFacNCon(aobjNotaCon As ClsNotaCon, astrIdItemFac As String) As String
        Dim lobjItemFac As ClsItemFactura = MobjFactura.ColItemsFactura(astrIdItemFac)
        Dim lstrItemRef = lobjItemFac.StrItemRef
        Dim lstrName = lobjItemFac.StrName
        Dim lstrQuan = CMSTRUNODEC
        Dim lstrPrice = aobjNotaCon.FdecVlrAntApliItemFac(astrIdItemFac)
        Dim lstrLinExtAmo = lstrPrice
        Dim lstrLinTot = lstrPrice
        Dim lstrMeaUniCod = CMSTRUNI
        Dim lstrFreOfChaInd = CMSTRFALSO
        Dim lstrLinAllTotal = CMSTRCERODEC
        Dim lstrLinChaTotal = CMSTRCERODEC
        Dim lstrLineTotalTax = CMSTRCERODEC
        Dim lstrJsonDscto = FstrJsItemNota(lstrItemRef, lstrName, lstrQuan, lstrPrice,
                lstrLinExtAmo, lstrLinAllTotal, lstrLinChaTotal, lstrLineTotalTax,
                lstrLinTot, lstrMeaUniCod, lstrFreOfChaInd)
        Return lstrJsonDscto
    End Function
#End Region
#End Region
#Region "Js General para notas"
    Private Function FstrJsInfGralNota(astrPrefNota As String, astrNroNota As String,
            astrFechaFac As String, astrNota As String, astrCUFE As String, astrNroFac As String,
           astrTipoCorreccion As String, astrCustomId As String, astrFechaIniPer As String,
           astrFechaFinPer As String) As String
        Dim lstrJsonInfGralNota = CMSTRINDENT & My.Resources.NotGralInf
        If Not String.IsNullOrEmpty(astrPrefNota) Then
            lstrJsonInfGralNota &= CMSTRINDENT & My.Resources.Pref & astrPrefNota & CMSTRFINCAMPO
        End If
        lstrJsonInfGralNota &= CMSTRINDENT & My.Resources.NotNum & astrNroNota & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Cufe & astrCUFE & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.RefId & astrNroFac & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.IssDateInv & astrFechaFac & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.CustomId & astrCustomId & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.DisCod & astrTipoCorreccion & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.PeriodoStart & astrFechaIniPer & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.PeriodoEnd & astrFechaFinPer & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.CodMoneda & CMSTRCODMONEDA & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Nota & astrNota & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.ExtGR & "true" & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.ExchRate & CMSTRCERODEC & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.ExchRateDate & CMSTRFECHANULL & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.CustomIdInv & CMSTRNULO &
                CMSTRCIERREINF
        Return lstrJsonInfGralNota
    End Function
    Private Shared Function FstrJsItemNota(astrItemRef As String, astrName As String, astrQuant As String,
                astrPrice As String, astrLinExtAmo As String, astrLinAllTotal As String,
                astrLinChaTotal As String, astrLinTotalTax As String, astrLinTot As String,
                astrMeaUniCod As String, astrFreChaInd As String) As String
        Dim lstrJsItemNota = CMSTRINICIODOC &
                CMSTRINDENT & My.Resources.ItemRef & astrItemRef & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Name & astrName & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Cant & astrQuant & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Precio & astrPrice & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.LinExtAmo & astrLinExtAmo & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.LinAllTotal & astrLinAllTotal & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.LinChrTotal & astrLinChaTotal & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.LinTotTax & astrLinTotalTax & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.LinTot & astrLinTot & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.MesUniCod & astrMeaUniCod & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.FreOfChargeInd & astrFreChaInd & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.AddRef & "]," &
                CMSTRINDENT & My.Resources.AddPro & "],"
        Return lstrJsItemNota
    End Function
    ''' <summary>
    ''' Devuelve el Json de los impuestos de la nota cuando el valor es cero
    ''' </summary>
    Private Shared Function FstrJsImptosNota() As String
        Dim lstrImptosNota = CMSTRINDENT & My.Resources.NotTaxTot &
                CMSTRINDENT & "],"
        Return lstrImptosNota
    End Function
    ''' <summary>
    ''' Devuelve el Json de los impuestos de la nota cuando el valor no es cero
    ''' </summary>
    ''' <param name="astrImptosNota"></param>
    ''' <returns></returns>
    Private Shared Function FstrJsImptosNota(astrImptosNota As ArrayList) As String
        Dim lstrImptosNota = CMSTRINDENT & My.Resources.NotTaxTot, i = 0
        For Each lstrImptoNota As String In astrImptosNota
            i += 1
            lstrImptosNota &= FstrJsImptoNota(lstrImptoNota, i = astrImptosNota.Count)
        Next
        lstrImptosNota &= CMSTRINDENT & "],"
        Return lstrImptosNota
    End Function
    Private Shared Function FstrJsImptoNota(astrImptoNota As String, ablnUltimo As Boolean) As String
        ' astrImptoNota es un string que debe traer la información de cada impto de la nota y 
        ' cuyos valores deben venir separados por &
        Dim lstrInfImpto As String() = astrImptoNota.Split("&")
        Dim lstrIdImpto As String = lstrInfImpto(0)
        Dim lstrTasa As String = lstrInfImpto(1)
        Dim lstrBaseImpto As String = lstrInfImpto(2)
        Dim lstrVlrImpto As String = lstrInfImpto(3)
        Dim lstrEsRet As String = lstrInfImpto(4)
        Dim lstrTotalImtoNota = CMSTRINICIODOC &
                CMSTRINDENT & My.Resources.Id & lstrIdImpto & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.EsImpto & lstrEsRet & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.BaseImpto & lstrBaseImpto & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.VlrImpto & lstrVlrImpto & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Tasa & lstrTasa & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.BasUniMeas & CMSTRCERODEC & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.PerUniAmo & CMSTRCERODEC
        If ablnUltimo Then
            lstrTotalImtoNota &= CMSTRINDENT & "}"
        Else
            lstrTotalImtoNota &= CMSTRINDENT & "},"
        End If
        Return lstrTotalImtoNota
    End Function
    Private Shared Function FstrJsTotalNota(astrTotalAntesImptos As String, astrTotalBaseGrabable As String,
                astrTotalNota As String, astrTotalPagar As String) As String
        Dim lstrJsTotNotDb = CMSTRINDENT & My.Resources.NotTot &
                CMSTRINDENT & My.Resources.LinExtAmo & astrTotalAntesImptos & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.TaxExcAmo & astrTotalBaseGrabable & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.TaxIncAmo & astrTotalNota & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.PayAmo & astrTotalPagar &
                CMSTRINDENT & "}" & "}"
        Return lstrJsTotNotDb
    End Function
#End Region
#End Region
#Region "Generales"
    ' Json con la información del cliente
    Private Function FstrJsInfCliente() As String
        Dim lstrTipoIdCli As String = FstrTipoDocIdDian(Mobjcliente.ObjTipoDocIdentidadByt.ObjValorPro)
        Dim lstrIdCliente = FstrComillas(Mobjcliente.ObjIdClienteDbl.ToString)
        Dim lstrDigVerificacion = CMSTRCERO
        Dim lentDigVerificacion As Integer = Mobjcliente.ObjIdTerceroDbl.SbyDigitoVerificacion
        If lentDigVerificacion > 0 Then
            lstrDigVerificacion = lentDigVerificacion.ToString
        End If
        Dim lstrNombreCli = FstrComillas(Mobjcliente.ObjNombreCompletoStr.ObjValorPro)
        Dim lstrPaisDir = FstrComillas(Mobjcliente.ObjPaisDirStr.ObjValorPro)
        Dim lstrNombrePais = Mobjcliente.FstrNombrePaisDir()
        Dim lstrIdSubdPais = Mobjcliente.FstrIdSubdivisionPais
        Dim lstrNomSubPais = Mobjcliente.FstrNombreSubPais
        Dim lstrIdCiudad = Mobjcliente.FstrIdCiudad()
        Dim lstrNomCiudad = Mobjcliente.FstrNombreCiudad
        Dim lstrDir = FstrDireccionCliente(Mobjcliente)
        Dim lstrTelCliente = Mobjcliente.FstrTelCliente
        Dim lstrEmail = FstrComillas(Mobjcliente.ObjEmailStr.ToString)
#If DES = 1 Then
        If lstrEmail <> FstrComillas("aureliovv47@gmail.com") Then Stop
#End If
        Dim lstrRegIva = Mobjcliente.FstrRegimenIva()
        Dim lstrTipoPersonaCli As String = FstrTipoPersonaDian(Mobjcliente.ObjTipoTerceroByt.ObjValorPro)
        Dim lstrCodCliente = FstrComillas("")
        Dim lstrResponsFiscal = FstrComillas(Mobjcliente.StrResponsFiscal)
        Dim lstrCodigoEsquemaTax = FstrComillas("ZZ")
        Dim lstrNombreEsquemaTax = FstrComillas("No aplica")
        Dim lstrCodPostal = FstrComillas(Mobjcliente.ObjCodigoPostalStr.ToString)
        Dim lStrInfCliente = CMSTRINDENT & My.Resources.InfCli &
                CMSTRINDENT & My.Resources.TipoId & lstrTipoIdCli & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Ident & lstrIdCliente & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.DV & lstrDigVerificacion & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Regist & lstrNombreCli & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.PaisDir & lstrPaisDir & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.CouNam & lstrNombrePais & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.SubCod & lstrIdSubdPais & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.SubNam & lstrNomSubPais & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.CitCod & lstrIdCiudad & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.CitNam & lstrNomCiudad & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.AddLin & lstrDir & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Teleph & lstrTelCliente & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.Email & lstrEmail & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.TaxLev & lstrRegIva & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.AddAcc & lstrTipoPersonaCli & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.CusCod & lstrCodCliente & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.FisRes & lstrResponsFiscal & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.TaxSchCod & lstrCodigoEsquemaTax & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.TaxSchName & lstrNombreEsquemaTax & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.PosZon & lstrCodPostal & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.PartPer & "100" & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.AddCus & "[" &
                CMSTRINDENT & "]" & CMSTRCIERREINF
        Return lStrInfCliente
    End Function
    ''' <summary>
    ''' Indica si alguna de las facturas referenciadas en el parametro fue ingresada a la API de
    ''' MisFacturas
    ''' </summary>
    ''' <param name="astrNrosFacs">Numeros de las facturas a examinar</param>
    ''' <returns></returns>
    Friend Function FblnHayFacAfectadas(astrNrosFacs As String()) As Boolean
        Dim lblnHay = False, lstrPrefFac As String, lentIdFact As Integer
        For Each lstrNroFac As String In astrNrosFacs
            lstrPrefFac = Split(lstrNroFac, ",")(0)
            lentIdFact = CType(Split(lstrNroFac, ",")(1), Integer)
            If IsNothing(MobjFactura) Then
                MobjFactura = New ClsFactura()
            End If
            MobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFac, lentIdFact})
            lblnHay = Not String.IsNullOrEmpty(MobjFactura.ObjCUFEStr.ToString)
            If lblnHay Then Exit For
        Next
        Return lblnHay
    End Function
    Private Sub SAbraFac(astrPrefFac As String, aentIdFac As Integer)
        Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, astrPrefFac, aentIdFac}
        MobjFactura.SAbra(lobjValorLlave)
        If Not MobjFactura.BlnExiste Then
            Throw New ErrorInesperadoPanLException("Factura no existe!")
        End If
    End Sub
    Private Shared Function FstrCudoc(astrRespuesta As String) As String
        Dim lobjInsertRespuesta As ClsInsertRespuesta
        lobjInsertRespuesta = JsonConvert.DeserializeObject(Of ClsInsertRespuesta)(astrRespuesta)
        Dim lstrCudoc = lobjInsertRespuesta.DocumentId
        If lobjInsertRespuesta Is Nothing OrElse String.IsNullOrEmpty(lstrCudoc) Then
            SRegistreError(FstrFechaPro(Now), String.Empty, String.Empty,
                            "0-NoEstadoApi", CMSTRCERO, "Doc No Insertado!")
        End If
        Return lstrCudoc
    End Function
    Private Shared Function FstrJsonStatusDoc_V1(astrIddoc As String,
                aenuTipoDoc As EnuTipoDocOri) As String
        Dim lstrjsonDocStat = CMSTRINICIODOC, lentIdTipoDoc = 0
        If aenuTipoDoc = EnuTipoDocOri.EnuFactura Then
            lentIdTipoDoc = 1
        ElseIf aenuTipoDoc = EnuTipoDocOri.EnuNotaCr Then
            lentIdTipoDoc = 2
        ElseIf aenuTipoDoc = EnuTipoDocOri.EnuNotaDb Then
            lentIdTipoDoc = 3
        ElseIf aenuTipoDoc = EnuTipoDocOri.EnuNotaRevCr Then
            lentIdTipoDoc = 3
        Else
            Throw New ErrorInesperadoPanLException("Tipo documento no esperado")
        End If
        lstrjsonDocStat &= CMSTRINDENT & My.Resources.IdDoc & astrIddoc & CMSTRFINCAMPO &
                CMSTRINDENT & My.Resources.TipoDoc & lentIdTipoDoc & "}"
        Return lstrjsonDocStat
    End Function
    Private Shared Function FstrFechaIniPeriodo(adtmFechaDoc As DateTime,
            ByRef astrFechaFinPeriodo As String) As String
        Dim lentIdAno = adtmFechaDoc.Year
        Dim lobjAno As ClsAno = GobjParametros.ColAnos(lentIdAno.ToString())
        Dim lobjPeriodo As ClsPeriodo = lobjAno.ColPeriodos(Format(adtmFechaDoc.Month, "0#"))
        Dim ldtmFecIniPer As Date = lobjPeriodo.DtmFechaInicioPeriodo
        Dim ldtmFecFinPer As Date = lobjPeriodo.DtmFechaFinPeriodo
        Dim lstrFecIniPer As String = FstrComillas(FstrFechaPro(ldtmFecIniPer))
        astrFechaFinPeriodo = FstrComillas(FstrFechaPro(ldtmFecFinPer))
        Return lstrFecIniPer
    End Function
#End Region
#End Region
#Region "Comunicación con API"
    Friend Async Function FstrObtengaToken(i As Integer) As Task(Of String)
        Dim lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            If String.IsNullOrEmpty(MstrToken) Then
                Dim lstrCompl = "username=" & MstrUsuarioMisF & "&password=" & MstrPasswordMisF
                Dim lstrUri = MstrUrlApi & "login?" & lstrCompl
                Dim lstcQueryString As New StringContent(lstrCompl)
                MhttpRespuesta = Await FhttpCliMisFac.PostAsync(lstrUri, lstcQueryString)
                MhttpRespuesta.EnsureSuccessStatusCode()
                Dim lstrResponseBody As String = Await MhttpRespuesta.Content.ReadAsStringAsync()
                MstrToken = lstrResponseBody
                MstrToken = MstrToken.Replace(Chr(34), "")
            End If
            lblnNoHayError = True
        Catch ex As HttpRequestException
            lstrMensEx = ex.Message
        Catch ex As Exception
            lstrMensEx = ex.Message
        Finally
            If Not lblnNoHayError AndAlso i > 2 Then
                SRegistreError(FstrFechaPro(Now), "", "",
                        "Excepción en FstrObtengaToken", CMSTRCERO, lstrMensEx)
                MstrToken = String.Empty
            End If
        End Try
        If String.IsNullOrEmpty(MstrToken) AndAlso i <= 2 Then
            SEspere(0, 0, 500)
            i += 1
            MstrToken = Await FstrObtengaToken(i)
        End If
        Return MstrToken
    End Function
    Friend Async Function FstrObtengaToken_V1() As Task(Of String)
        Try
            If String.IsNullOrEmpty(MstrToken_V1) Then
                Dim lstrCompl = "username=" & MstrUsuarioMisF & "&password=" & MstrPasswordMisF
                Dim lstrUri = CMSTRURL_V1 & "login?" & lstrCompl
                Dim lstcQueryString As New StringContent(lstrCompl)
                MhttpRespuesta = Await FhttpCliMisFac.PostAsync(lstrUri,
                        lstcQueryString)
                MhttpRespuesta.EnsureSuccessStatusCode()
                Dim lstrResponseBody As String = Await MhttpRespuesta.Content.ReadAsStringAsync()
                MstrToken_V1 = lstrResponseBody
                MstrToken_V1 = MstrToken_V1.Replace(Chr(34), "")
            End If
        Catch ex As HttpRequestException
            Throw
        Catch ex As Exception
            Throw
        End Try
        Return MstrToken_V1
    End Function
    Friend Async Function FobjObtengaEstado(astrCUDoc As String, aenuTipoDoc As EnuTipoDocOri,
                astrNroDoc As String) As Task(Of ClsEstadoDoc)
        Dim lobjEstado As ClsEstadoDoc = Nothing, i = 0
        Do While i <= 5
            lobjEstado = Await FobjEstado(astrCUDoc, aenuTipoDoc, astrNroDoc)
            If lobjEstado Is Nothing Then
                SEspere(0, 0, 500)
            Else
                Exit Do
            End If
            i += 1
        Loop
        If i > 5 AndAlso lobjEstado Is Nothing Then
            SRegistreError(FstrFechaPro(Now), "", "", "Excepción en FobjObtengaEstado",
                            CMSTRCERO, "Solicitud rechazada por API")
        End If
        Return lobjEstado
    End Function
    Private Async Function FobjEstado(astrCUDoc As String, aenuTipoDoc As EnuTipoDocOri,
                astrNroDoc As String) As Task(Of ClsEstadoDoc)
        Dim lobjEstadoDoc As ClsEstadoDoc = Nothing
        If Not String.IsNullOrEmpty(astrCUDoc) Then
            Dim lstrMensEx = String.Empty
            Dim lblnNoHayError = False
            Dim lctrCudoc As String = astrCUDoc.Replace(Chr(34), "")
            Dim lstrCompl = MstrCompl & "&DocumentID=" & lctrCudoc
            If aenuTipoDoc = EnuTipoDocOri.EnuFactura Then
                lstrCompl &= "&DocumentType=1"
            Else
                lstrCompl &= "&DocumentType=2"
            End If
#If PRU = 0 Then
            Dim lstrUri = MstrUrlApi & "GetDocumentStatus?" & lstrCompl
            Try
                If String.IsNullOrEmpty(MstrToken) Then
                    Await FstrObtengaToken(0)
                End If
                If Not String.IsNullOrEmpty(MstrToken) Then
                    FhttpCliMisFac.DefaultRequestHeaders.Clear()
                    FhttpCliMisFac.DefaultRequestHeaders.Add("Authorization", "misfacturas " &
                            MstrToken)
                    Dim lstcQueryString As New StringContent(lstrCompl)
                    MhttpRespuesta = Await FhttpCliMisFac.PostAsync(lstrUri,
                        lstcQueryString)
                    MhttpRespuesta.EnsureSuccessStatusCode()
                    Dim lstrResponseBody = Await MhttpRespuesta.Content.ReadAsStringAsync()
                    lobjEstadoDoc = JsonConvert.DeserializeObject(Of ClsEstadoDoc)(lstrResponseBody)
                    If lobjEstadoDoc IsNot Nothing AndAlso lobjEstadoDoc.DocumentStatus = 70 Then
                        SregistreError(lobjEstadoDoc, aenuTipoDoc, astrNroDoc)
                        SExporteJsonDoc(astrNroDoc, aenuTipoDoc)
                    End If
                End If
                lblnNoHayError = True
            Catch ex As Exception
                lstrMensEx = ex.ToString()
            Finally
                If Not lblnNoHayError Then
                    lobjEstadoDoc = Nothing
                End If
            End Try
#Else
            SEspere(0, 0, 100)
            lobjEstadoDoc = New ClsEstadoDoc With {
            .DocumentStatus = 5,
            .CUDE = "ijadjsdpiusd53645646asdhaosdyiah",
            .CUFE = "ewrtyunbvcx5218wdfdjl646313",
            .CustomerPartyID = "564321546",
            .CustomerParty = "AVV",
            .InvoiceNumber = "8465431",
            .DocumentNumber = "111111",
            .StatusDate = "2021-11-04T15:04:31",
            .DIANErrors = Nothing
        }
#End If
        End If
        Return lobjEstadoDoc
    End Function
    Friend Async Function FstrObjtengaDocsXFecha(astrFechaDesde As String,
            astrFechaHasta As String) As Task(Of String)
        Dim lstrRespuesta = String.Empty
        Dim lstrComp = MstrCompl & "&StartDate=" & astrFechaDesde & "&EndDate=" & astrFechaHasta &
                "&DocumentType=1"
        Dim lstrUri = MstrUrlApi & "GetDocumentsByDates?" & lstrComp
        Try
            If String.IsNullOrEmpty(MstrToken) Then
                Await FstrObtengaToken(0)
            End If
            If Not String.IsNullOrEmpty(MstrToken) Then
                FhttpCliMisFac.DefaultRequestHeaders.Clear()
                FhttpCliMisFac.DefaultRequestHeaders.Add("Authorization", "misfacturas " & MstrToken)
                MhttpRespuesta = Await FhttpCliMisFac.GetAsync(lstrUri)
                MhttpRespuesta.EnsureSuccessStatusCode()
                lstrRespuesta = Await MhttpRespuesta.Content.ReadAsStringAsync()
            End If
        Catch ex As Exception
            Throw
        End Try
        Return lstrRespuesta
    End Function
    Friend Async Function FstrObtengaEstado_V1(astrCUDoc As String,
            aenuTipoDoc As EnuTipoDocOri) As Task(Of String)
        If String.IsNullOrEmpty(astrCUDoc) Then
            Return String.Empty
        End If
        Dim lstrUri = CMSTRURL_V1 & "GetDocument"
        Dim lstrJson = FstrJsonStatusDoc_V1(astrCUDoc, aenuTipoDoc)
        Dim lstrResponseBody = String.Empty, lblnNoHayError As Boolean
        Try
            If String.IsNullOrEmpty(MstrToken_V1) Then
                Await FstrObtengaToken_V1()
            End If
            If Not String.IsNullOrEmpty(MstrToken_V1) Then
                MhttpClienteMisFac.DefaultRequestHeaders.Clear()
                MhttpClienteMisFac.DefaultRequestHeaders.Add("Authorization", "misfacturas " & MstrToken_V1)
                Dim lstcQueryString As New StringContent(lstrJson, Encoding.UTF8,
                        "application/json")
                MhttpRespuesta = Await FhttpCliMisFac.PostAsync(lstrUri,
                                        lstcQueryString)
                MhttpRespuesta.EnsureSuccessStatusCode()
                lstrResponseBody = Await MhttpRespuesta.Content.ReadAsStringAsync()
            End If
            lblnNoHayError = True
        Catch ex As HttpRequestException
            Throw
        Catch ex As Exception
            Throw
        End Try
        Return lstrResponseBody
    End Function
    ''' <summary>
    ''' Envia el Json de la factura a la API y devuelve el CUDoc
    ''' </summary>
    ''' <param name="astrJsonFac">Json de la factura</param>
    ''' <returns></returns>
    Private Async Function SInserteJsFac(astrJsonFac As String, i As Integer) As Task(Of String)
        Dim lstrRespuesta = String.Empty
        Dim lstrNroFac = MobjFactura.StrNumeroFactura
        Dim lstrUri = MstrUrlApi & "insertinvoice?" & MstrCompl & "&TemplateID=73"
        Dim lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            If String.IsNullOrEmpty(MstrToken) Then
                Await FstrObtengaToken(0)
            End If
#If PRU = 0 Then
            If Not String.IsNullOrEmpty(MstrToken) Then
                FhttpCliMisFac.DefaultRequestHeaders.Clear()
                FhttpCliMisFac.DefaultRequestHeaders.Add("Authorization", "misfacturas " & MstrToken)
                Dim lstcQueryString As New StringContent(astrJsonFac,
                    Encoding.UTF8, "application/json")
                MhttpRespuesta = Await FhttpCliMisFac.PostAsync(lstrUri,
                    lstcQueryString)
                MhttpRespuesta.EnsureSuccessStatusCode()
                lstrRespuesta = Await MhttpRespuesta.Content.ReadAsStringAsync()
            End If
#Else
            SEspere(0, 0, 100)
            lstrRespuesta = "Ok"
#End If
            lblnNoHayError = True
        Catch ex As HttpRequestException
            lstrMensEx = ex.Message
        Catch ex As Exception
            lstrMensEx = ex.Message
        Finally
            If Not lblnNoHayError AndAlso i > 2 Then
                SRegistreError(FstrFechaPro(Now), "", "",
                        "Excepción en SInserteJsFac", CMSTRCERO, lstrMensEx)
            End If
        End Try
        If String.IsNullOrEmpty(lstrRespuesta) AndAlso i <= 5 Then
            SEspere(0, 0, 500)
            i += 1
            Await SInserteJsFac(astrJsonFac, i)
        End If
        Return lstrRespuesta
    End Function
    ''' <summary>
    ''' Envia el Json de la factura por contingencia a la API y devuelve el CUDoc
    ''' </summary>
    ''' <param name="astrJsonFacCon">Json de la factura</param>
    Private Async Function SInserteJsFacCon(astrJsonFacCon As String, i As Integer) As Task(Of String)
        Dim lstrRespuesta = String.Empty
        Dim lstrNroFac = MobjFactura.StrNumeroFactura
        Dim lstrUri = MstrUrlApi & "contingency?" & MstrCompl & "&TemplateID=73"
        Dim lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            If String.IsNullOrEmpty(MstrToken) Then
                Await FstrObtengaToken(0)
            End If
#If PRU = 0 Then
            If Not String.IsNullOrEmpty(MstrToken) Then
                FhttpCliMisFac.DefaultRequestHeaders.Clear()
                FhttpCliMisFac.DefaultRequestHeaders.Add("Authorization", "misfacturas " & MstrToken)
                Dim lstcQueryString As New StringContent(astrJsonFacCon,
                        Encoding.UTF8, "application/json")
                MhttpRespuesta = Await FhttpCliMisFac.PostAsync(lstrUri,
                        lstcQueryString)
                MhttpRespuesta.EnsureSuccessStatusCode()
                lstrRespuesta = Await MhttpRespuesta.Content.ReadAsStringAsync()
            End If
#Else
            SEspere(0, 0, 100)
            lstrRespuesta = "Ok"
#End If
            lblnNoHayError = True
        Catch ex As HttpRequestException
            lstrMensEx = ex.Message
        Catch ex As Exception
            lstrMensEx = ex.Message
        Finally
            If Not lblnNoHayError AndAlso i > 2 Then
                SRegistreError(FstrFechaPro(Now), "", "",
                        "Excepción en SInserteJsFacCon", CMSTRCERO, lstrMensEx)
            End If
        End Try
        If String.IsNullOrEmpty(lstrRespuesta) AndAlso i <= 5 Then
            SEspere(0, 0, 500)
            i += 1
            Await SInserteJsFacCon(astrJsonFacCon, i)
        End If
        Return lstrRespuesta
    End Function
    ''' <summary>
    ''' Envia el Json de la Nota a la API y devuelve el CUDOC "Código Unico Documento"
    ''' </summary>
    ''' <param name="astrJsonNota">Json de la Nota</param>
    ''' <returns></returns>
    Private Async Function SInserteJsNota(astrJsonNota As String, ablnDb As Boolean,
                i As Integer) As Task(Of String)
        Dim lstrRespuesta = String.Empty
        Dim lstrUri = MstrUrlApi & "insertnote?" & MstrCompl
        If ablnDb Then
            lstrUri &= "&NoteType=92"
        Else
            lstrUri &= "&NoteType=91"
        End If
        Dim lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            If String.IsNullOrEmpty(MstrToken) Then
                Await FstrObtengaToken(0)
            End If
#If PRU = 0 Then
            If Not String.IsNullOrEmpty(MstrToken) Then
                FhttpCliMisFac.DefaultRequestHeaders.Clear()
                FhttpCliMisFac.DefaultRequestHeaders.Add("Authorization", "misfacturas " & MstrToken)
                Dim lstcQueryString As New StringContent(astrJsonNota,
                                Encoding.UTF8, "application/json")
                MhttpRespuesta = Await FhttpCliMisFac.PostAsync(lstrUri,
                                lstcQueryString)
                MhttpRespuesta.EnsureSuccessStatusCode()
                lstrRespuesta = Await MhttpRespuesta.Content.ReadAsStringAsync()
            End If
#Else
            SEspere(0, 0, 100)
            lstrRespuesta = "OK"
#End If
            lblnNoHayError = True
        Catch ex As HttpRequestException
            lstrMensEx = ex.Message
        Catch ex As Exception
            lstrMensEx = ex.Message
        Finally
            If Not lblnNoHayError AndAlso i > 2 Then
                SRegistreError(FstrFechaPro(Now), "", "",
                        "Excepción en SInserteJsNota", CMSTRCERO, lstrMensEx)
            End If
        End Try
        If String.IsNullOrEmpty(lstrRespuesta) AndAlso i <= 5 Then
            SEspere(0, 0, 500)
            i += 1
            Await SInserteJsNota(astrJsonNota, ablnDb, i)
        End If
        Return lstrRespuesta
    End Function
    Friend Async Function SEnvieAdjunto(aobjDoc As ClsCBObjetoPan,
            aenuTipoDoc As EnuTipoDocOri) As Task
        Dim lstrCudoc As String, lstrArchivo As String
        Dim lstrPrefDoc As String, lentIdDoc As Integer
        If String.IsNullOrEmpty(MstrToken) Then
            Await FstrObtengaToken(0)
        End If
        If aenuTipoDoc = EnuTipoDocOri.EnuFactura Then
            Dim lobjFac As ClsFactura = aobjDoc
            lstrPrefDoc = lobjFac.ObjPrefijo_FactStr.ObjValorPro
            lentIdDoc = lobjFac.ObjIdFacturaEnt.ObjValorPro
            lstrCudoc = lobjFac.ObjCUDocStr.ObjValorPro
        ElseIf aenuTipoDoc = EnuTipoDocOri.EnuNotaCr Then
            Dim lobjNotaCr As ClsNotaCr = aobjDoc
            lstrPrefDoc = lobjNotaCr.ObjPrefijo_NotaCrStr.ObjValorPro
            lentIdDoc = lobjNotaCr.ObjIdNotaCrEnt.ObjValorPro
            lstrCudoc = lobjNotaCr.ObjCUDocStr.ObjValorPro
        ElseIf aenuTipoDoc = EnuTipoDocOri.EnuNotaDb Then
            Dim lobjNotaDb As ClsNotaDb = aobjDoc
            lstrPrefDoc = lobjNotaDb.ObjPrefijo_NotaDbStr.ObjValorPro
            lentIdDoc = lobjNotaDb.ObjIdNotaDbEnt.ObjValorPro
            lstrCudoc = lobjNotaDb.ObjCUDocStr.ObjValorPro
        ElseIf aenuTipoDoc = EnuTipoDocOri.EnuNotaRevCr Then
            Dim lobjNotaRCr As ClsNotaReversionCr = aobjDoc
            lstrPrefDoc = lobjNotaRCr.ObjPrefijo_NotaReversaCrStr.ObjValorPro
            lentIdDoc = lobjNotaRCr.ObjIdNotaReversaCrEnt.ObjValorPro
            lstrCudoc = lobjNotaRCr.ObjCUDocStr.ObjValorPro
        ElseIf aenuTipoDoc = EnuTipoDocOri.EnuNotaCon Then
            Dim lobjNCon As ClsNotaCon = aobjDoc
            lstrPrefDoc = lobjNCon.ObjPrefijo_NotaConStr.ObjValorPro
            lentIdDoc = lobjNCon.ObjIdNotaConEnt.ObjValorPro
            lstrCudoc = lobjNCon.ObjCUDocStr.ObjValorPro
        Else
            Throw New ErrorInesperadoPanLException("Tipo de documento no válido")
        End If
#If PRU = 0 Then
        Dim lstrTipoDoc As String
        If aenuTipoDoc = EnuTipoDocOri.EnuFactura Then
            lstrTipoDoc = "1"
        Else
            lstrTipoDoc = "2"
        End If
        lstrArchivo = ClsOrionCop.FstrArchivoPdfDcto(lstrPrefDoc, lentIdDoc, aenuTipoDoc)
        SExporteDoc(aobjDoc, aenuTipoDoc)
        If My.Computer.FileSystem.FileExists(lstrArchivo) Then
            If Not String.IsNullOrEmpty(MstrToken) Then
                lstrCudoc = lstrCudoc.Replace(Chr(34), "")
                Dim lhttpReq As New Chilkat.HttpRequest With {
                    .HttpVerb = "POST",
                    .ContentType = "multipart/form-data"
                }
                lhttpReq.AddHeader("Authorization", "misfacturas " & MstrToken)
                lhttpReq.AddHeader("Expect", "100-continue")
                lhttpReq.Path = "integrationAPI_2/api/AttachRG?" & MstrCompl & "&DocumentID=" &
                        lstrCudoc & "&DocumentType=" & lstrTipoDoc
                Dim lblnLogro = lhttpReq.AddFileForUpload("File", lstrArchivo)
                If Not lblnLogro Then
                    Throw New ErrorInesperadoPanLException("No cargo Archivo para ser enviado a API")
                End If
                Dim lstrDom = FstrDominioProvEFac()
                FhttpMisFac.SynchronousRequest(lstrDom, 443, True, lhttpReq)
                If (FhttpMisFac.LastMethodSuccess <> True) Then
                    Dim lstrError = FhttpMisFac.LastErrorText
                    SRegistreError(FstrFechaPro(Now), lstrTipoDoc, aobjDoc.StrIdObjeto,
                            "0-NoEnviado", CMSTRCERO, lstrError)
                End If
            End If
        End If
#Else
        SEspere(0, 0, 100)
#End If
    End Function
    Friend Shared Sub SEscribaEstruJsonDoc(astrEstJsonFac As String,
            aenuTipoDoc As EnuTipoDocOri, astrNroDoc As String)
        If Not String.IsNullOrEmpty(astrEstJsonFac) Then
            Dim lstrTrya = GstrTrayEFac
            If aenuTipoDoc = EnuTipoDocOri.EnuFactura Then
                lstrTrya &= "\JsonFac_" & astrNroDoc & ".txt"
            ElseIf aenuTipoDoc = EnuTipoDocOri.EnuNotaDb Then
                lstrTrya &= "\JsonNdb_" & astrNroDoc & ".txt"
            ElseIf aenuTipoDoc = EnuTipoDocOri.EnuReciboCaja Then
                lstrTrya &= "\JsonRec_" & astrNroDoc & ".txt"
            ElseIf aenuTipoDoc = EnuTipoDocOri.EnuNotaCr Then
                lstrTrya &= "\JsonNcr_" & astrNroDoc & ".txt"
            ElseIf aenuTipoDoc = EnuTipoDocOri.EnuNotaRevCr Then
                lstrTrya &= "\JsonNRcr_" & astrNroDoc & ".txt"
            ElseIf aenuTipoDoc = EnuTipoDocOri.EnuNotaCon Then
                lstrTrya &= "\JsonNCon_" & astrNroDoc & ".txt"
            End If
            If File.Exists(lstrTrya) Then
                File.Delete(lstrTrya)
            End If
            Using lswArchivo As New StreamWriter(lstrTrya)
                lswArchivo.WriteLine(astrEstJsonFac)
            End Using
        End If
    End Sub
    Friend Shared Function FstrDominioProvEFac() As String
        Dim lstrUrl As String = GobjParametros.ObjURLStr.ToString
        Dim lstrDom As String
        If lstrUrl.Contains("www") Then
            lstrDom = lstrUrl.Substring(lstrUrl.IndexOf(".") + 1)
        Else
            lstrDom = lstrUrl.Substring(lstrUrl.IndexOf("//") + 2)
        End If
        lstrDom = lstrDom.Substring(0, lstrDom.IndexOf("/"))
        Return lstrDom
    End Function
    Private Shared Function FstrError(astrMensaje As String)
        Dim lstrError = String.Empty
        If Not String.IsNullOrEmpty(astrMensaje) Then
            If astrMensaje.Contains(",") Then
                Dim lstrPartes As String() = astrMensaje.Split(",")
                lstrError = lstrPartes(0).Trim & ", " & lstrPartes(1)
            Else
                lstrError = astrMensaje
            End If
        End If
        Return lstrError
    End Function
#End Region
#Region "Dispose"
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overridable Overloads Sub Dispose(ablnDisposing As Boolean)
        If Not MblnDisposed Then
            If ablnDisposing Then
                If Not IsNothing(FhttpCliMisFac) Then
                    MhttpClienteMisFac.Dispose()
                End If
                If Not IsNothing(MhttpRespuesta) Then
                    MhttpRespuesta.Dispose()
                End If
                If Not IsNothing(MhttpMisFac) Then
                    MhttpMisFac.Dispose()
                End If
                MblnDisposed = True
            End If
        End If
    End Sub
#End Region
End Class
#Region "Clases para descerializar Json de GetDocumentStatus"
Public Class ClsEstadoDoc
    Public DocumentUUID As String
    Public DocumentType As Integer
    Public DocumentStatus As Integer
    Public StatusDate As String
    Public CUFE As String
    Public InvoiceNumber As String
    Public CUDE As String
    Public DocumentNumber As String
    Public CustomerParty As String
    Public CustomerPartyID As String
    Public DIANErrors As List(Of ClsErrorDian)
End Class
Public Class ClsErrorDian
    Public Code As String
    Public Description As String
End Class
Public Class ClsInsertRespuesta
    Public DocumentId As String
    Public MessageValidation As String
End Class
#End Region