Imports System.Threading.Tasks
Friend Class ClsEFactura
#Region "Definiciones"
    Implements IDisposable
    '
    Private MIntMisFacturas As ClsInterfazMisFacturas = Nothing
    Private MobjFactura As ClsFactura = Nothing
    Private MblnDisposed As Boolean = False
    Public Event EvnInicio As EventHandler(Of ClsPanEventArgs)
    Public Event EvnAvance As EventHandler(Of ClsPanEventArgs)
    Public Event EvnFin As EventHandler(Of ClsPanEventArgs)
    Private MobjArgumentoEventoPan As ClsPanEventArgs = Nothing
#If PRU Then
    Dim MblnInsertando As Boolean = False
#End If
#End Region

#Region "Generales"
    ''' <summary>
    ''' Indica que objeto esta aceptado para informar su avance en la p{agina principal
    ''' </summary>
    ''' <returns></returns>
    Friend Property BlnAceptado As Boolean = False
    Friend ReadOnly Property BlnDisposed As Boolean
        Get
            Return MblnDisposed
        End Get
    End Property
    Private ReadOnly Property ObjIntMisFact As ClsInterfazMisFacturas
        Get
            If IsNothing(MIntMisFacturas) Then
                MIntMisFacturas = New ClsInterfazMisFacturas(GCOBJREGISTRO)
            End If
            Return MIntMisFacturas
        End Get
    End Property
    Private ReadOnly Property ObjFactura As ClsFactura
        Get
            If IsNothing(MobjFactura) Then
                MobjFactura = New ClsFactura()
            End If
            Return MobjFactura
        End Get
    End Property
    Friend ReadOnly Property ObjArgumentoEventoPan As ClsPanEventArgs
        Get
            If IsNothing(MobjArgumentoEventoPan) Then
                MobjArgumentoEventoPan = New ClsPanEventArgs With {
                    .BlnCancele = False,
                    .BlnVaciandoObjeto = False,
                    .BlnProcesoOk = False,
                    .DblCantProcesada = 0.0,
                    .DblCantAProcesar = 0.0,
                    .EnuProceso = EnuProcesoDef.None
                }
            End If
            Return MobjArgumentoEventoPan
        End Get
    End Property
    Private Async Function FenuEstadoDocEFac(aobjDoc As ClsCBObjetoPan,
            aenuTipoDoc As EnuTipoDocOri) As Task(Of EnuEstadoEDoc)
        Dim lstrNroDoc As String, lstrCudoc As String
        Dim lenuEstDoc As EnuEstadoEDoc
        Dim lobjEstado As ClsEstadoDoc = Nothing
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                Dim lobjFac As ClsFactura = aobjDoc
                lstrCudoc = lobjFac.ObjCUDocStr.ObjValorPro
            Case EnuTipoDocOri.EnuNotaCr
                Dim lobjNCr As ClsNotaCr = aobjDoc
                lstrCudoc = lobjNCr.ObjCUDocStr.ObjValorPro
            Case EnuTipoDocOri.EnuNotaDb
                Dim lobjNdb As ClsNotaDb = aobjDoc
                lstrCudoc = lobjNdb.ObjCUDocStr.ObjValorPro
            Case EnuTipoDocOri.EnuNotaRevCr
                Dim lobjNRCr As ClsNotaReversionCr = aobjDoc
                lstrCudoc = lobjNRCr.ObjCUDocStr.ObjValorPro
            Case EnuTipoDocOri.EnuNotaCon
                Dim lobjNCon As ClsNotaCon = aobjDoc
                lstrCudoc = lobjNCon.ObjCUDocStr.ObjValorPro
            Case Else
                lstrCudoc = String.Empty
        End Select
        lstrNroDoc = aobjDoc.StrIdObjeto
#If PRU = 0 Then
        If Not String.IsNullOrEmpty(lstrCudoc) Then
            lstrCudoc = lstrCudoc.Replace(Chr(34), "")
            Try
                lobjEstado = Await ObjIntMisFact.FobjObtengaEstado(lstrCudoc,
                    aenuTipoDoc, lstrNroDoc)
            Finally
                If lobjEstado IsNot Nothing Then
                    lenuEstDoc = FenuEstadoEDoc(lobjEstado)
                Else
                    lenuEstDoc = EnuEstadoEDoc.EnuErrorFtp
                End If
            End Try
        Else
            lenuEstDoc = EnuEstadoEDoc.EnuNoReg
        End If
#Else
        Dim lentIdEstado As Integer
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                Dim lobjFac As ClsFactura = aobjDoc
                lstrCudoc = lobjFac.ObjCUDocStr.ObjValorPro
                lentIdEstado = lobjFac.ObjIdEstadoEDocEnt.ObjValorPro
            Case EnuTipoDocOri.EnuNotaCr
                Dim lobjNCr As ClsNotaCr = aobjDoc
                lstrCudoc = lobjNCr.ObjCUDocStr.ObjValorPro
                lentIdEstado = lobjNCr.ObjIdEstadoEDocEnt.ObjValorPro
            Case EnuTipoDocOri.EnuNotaDb
                Dim lobjNdb As ClsNotaDb = aobjDoc
                lstrCudoc = lobjNdb.ObjCUDocStr.ObjValorPro
                lentIdEstado = lobjNdb.ObjIdEstadoEDocEnt.ObjValorPro
            Case EnuTipoDocOri.EnuNotaRevCr
                Dim lobjNRCr As ClsNotaReversionCr = aobjDoc
                lstrCudoc = lobjNRCr.ObjCUDocStr.ObjValorPro
                lentIdEstado = lobjNRCr.ObjIdEstadoEDocEnt.ObjValorPro
            Case EnuTipoDocOri.EnuNotaCon
                Dim lobjNCon As ClsNotaCon = aobjDoc
                lstrCudoc = lobjNCon.ObjCUDocStr.ObjValorPro
                lentIdEstado = lobjNCon.ObjIdEstadoEDocEnt.ObjValorPro
            Case Else
                lstrCudoc = String.Empty
        End Select
        If MblnInsertando Then
            lenuEstDoc = EnuEstadoEDoc.EnuNoReg
        Else
            If lentIdEstado = 1 Then
                lenuEstDoc = EnuEstadoEDoc.EnuEnProceso
            ElseIf lentIdEstado < 5 Then
                lenuEstDoc = lentIdEstado + 1
            End If
        End If
#End If
        Return lenuEstDoc
    End Function
#End Region

#Region "API MisFacturas V2"
    Friend Async Function FstrProceseDocsEfac(aenuTipoDoc As EnuTipoDocOri,
            adtbDocsAProcesar As DataTable) As Task
        Dim lblnNoHayError = False
        Dim lblnProcesar As Boolean, lstrFiltro As String, lstrFiltroTipoDoc As String
        GobjPanDat.SControleProcesoObj(True)
        Try
            lstrFiltroTipoDoc = "TipoDocu = " & aenuTipoDoc
            If adtbDocsAProcesar.Select(lstrFiltroTipoDoc).Count > 0 Then
                Dim lstrToken As String = Await ObjIntMisFact.FstrObtengaToken(0)
                If Not String.IsNullOrEmpty(lstrToken) Then
                    ' 1er paso Insertar
                    lstrFiltro = lstrFiltroTipoDoc & " AND " & ClsIdEstadoEDocEnt.SstrNombreCampoBd &
                            " = " & EnuEstadoEDoc.EnuNoReg
                    Dim ldrwDocsAProcesar = adtbDocsAProcesar.Select(lstrFiltro)
                    lblnProcesar = ldrwDocsAProcesar.Length > 0
                    If lblnProcesar Then
                        Await SInserteDocsAPI(ldrwDocsAProcesar, aenuTipoDoc)
                    End If
                    If ObjArgumentoEventoPan.BlnCancele Then
                        Exit Function
                    End If
                    ' 2o paso Actualizar estado
                    lstrFiltro = lstrFiltroTipoDoc & " AND " & ClsIdEstadoEDocEnt.SstrNombreCampoBd &
                        " = " & EnuEstadoEDoc.EnuEnProceso
                    ldrwDocsAProcesar = adtbDocsAProcesar.Select(lstrFiltro)
                    lblnProcesar = ldrwDocsAProcesar.Length > 0
                    If lblnProcesar Then
                        Await SActualiceDocsAPI(ldrwDocsAProcesar, aenuTipoDoc)
                        If aenuTipoDoc = EnuTipoDocOri.EnuFactura Then
                            Await SVerifiqueCUFE(ldrwDocsAProcesar)
                        End If
                    End If
                    If ObjArgumentoEventoPan.BlnCancele Then
                        Exit Function
                    End If
                    ' 3er Paso Enviar doc
                    lstrFiltro = lstrFiltroTipoDoc & " AND " & ClsIdEstadoEDocEnt.SstrNombreCampoBd &
                        " >= " & EnuEstadoEDoc.EnuRegi
                    ldrwDocsAProcesar = adtbDocsAProcesar.Select(lstrFiltro)
                    lblnProcesar = ldrwDocsAProcesar.Length > 0
                    If lblnProcesar Then
                        Await SEnvieDocsAPI(ldrwDocsAProcesar, aenuTipoDoc)
                    End If
                    If ObjArgumentoEventoPan.BlnCancele Then
                        Exit Function
                    End If
                    ' 4o Paso Actualizar enviados
                    lblnProcesar = ldrwDocsAProcesar.Length > 0
                    If lblnProcesar Then
                        Await SActualiceDocsAPI(ldrwDocsAProcesar, aenuTipoDoc)
                    End If
                    If ObjArgumentoEventoPan.BlnCancele Then
                        Exit Function
                    End If
                End If
            End If
            lblnNoHayError = True
        Catch ex As Exception
            Throw
        Finally
            If Not lblnNoHayError Then
                ObjArgumentoEventoPan.BlnProcesoOk = False
                RaiseEvent EvnFin(Me, ObjArgumentoEventoPan)
            End If
        End Try
    End Function

    Friend Async Function FstrDocsApiXFecha(astrFechaDesde As String,
                astrFechaHasta As String) As Task(Of String)
        Return Await ObjIntMisFact.FstrObjtengaDocsXFecha(astrFechaDesde, astrFechaHasta)
    End Function

#Region "Registrar EFac"
    Private Async Function SInserteDocsAPI(adrwsDocs As DataRow(),
            aenuTipoDoc As EnuTipoDocOri) As Task
        Dim lobjValorLlave As Object(), lstrPref = String.Empty, lentIdDoc As Integer
        Dim lstrMens = String.Empty, i = 0
        Dim lobjDoc = FobjNewDoc(aenuTipoDoc)
        Dim lenuProIns = FenuProcesoIns(aenuTipoDoc)
#If PRU = 1 Then
        MblnInsertando = True
#End If
        For Each ldrwDoc As DataRow In adrwsDocs
            i += 1
            If Not BlnAceptado Then
                ObjArgumentoEventoPan.EnuProceso = lenuProIns
                ObjArgumentoEventoPan.DblCantAProcesar = adrwsDocs.Length
                RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
            End If
            lstrPref = ClsPanorama.FobjValorCampo(ldrwDoc("Prefijo"), EnuTipoValor.enuString)
            lentIdDoc = ClsPanorama.FobjValorCampo(ldrwDoc("IdDocu"), EnuTipoValor.enuInteger)
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdDoc}
            lobjDoc.SAbra(lobjValorLlave)
            Dim lenuEstadoAPI = Await FenuEstadoDocEFac(lobjDoc, aenuTipoDoc)
            If lenuEstadoAPI > EnuEstadoEDoc.EnuNoReg AndAlso lenuEstadoAPI <> EnuEstadoEDoc.EnuInvalida Then
#If PRU = 0 Then
                Await ObjIntMisFact.SActualiceDoc(lobjDoc, aenuTipoDoc, False)
#Else
                Await ObjIntMisFact.SActualiceDoc(lobjDoc, aenuTipoDoc, False, False)
#End If
            Else
                If FblnInsertarDoc(lobjDoc, aenuTipoDoc) AndAlso (lenuEstadoAPI =
                            EnuEstadoEDoc.EnuNoReg OrElse lenuEstadoAPI =
                            EnuEstadoEDoc.EnuInvalida OrElse lenuEstadoAPI =
                            EnuEstadoEDoc.EnuErrorFtp) Then
                    Await ObjIntMisFact.SInserteDoc(lobjDoc, aenuTipoDoc)
                End If
            End If
            If BlnAceptado Then
                ObjArgumentoEventoPan.DblCantProcesada = i
                RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                If ObjArgumentoEventoPan.BlnCancele Then
                    Exit For
                End If
            End If
        Next
#If PRU = 1 Then
        MblnInsertando = False
#End If
        If BlnAceptado Then
            RaiseEvent EvnFin(Me, ObjArgumentoEventoPan)
        End If
    End Function
    Private Async Function SActualiceDocsAPI(adrwsDocs As DataRow(),
            aenuTipoDoc As EnuTipoDocOri) As Task
        Dim lblnNoHayError As Boolean
        Dim lobjValorLlave As Object(), lstrPref = String.Empty, lentIdDoc As Integer, i = 0
        Dim lobjDoc = FobjNewDoc(aenuTipoDoc)
        Dim lenuProAct As EnuProcesoDef = FenuProcesoAct(aenuTipoDoc)
        Try
            For Each ldrwDoc As DataRow In adrwsDocs
                i += 1
                If Not BlnAceptado Then
                    ObjArgumentoEventoPan.EnuProceso = lenuProAct
                    ObjArgumentoEventoPan.DblCantAProcesar = adrwsDocs.Length
                    RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
                End If
                lstrPref = ClsPanorama.FobjValorCampo(ldrwDoc("Prefijo"), EnuTipoValor.enuString)
                lentIdDoc = ClsPanorama.FobjValorCampo(ldrwDoc("IdDocu"), EnuTipoValor.enuInteger)
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdDoc}
                lobjDoc.SAbra(lobjValorLlave)
                If FblnActualizarDoc(lobjDoc, aenuTipoDoc) Then
#If PRU = 0 Then
                    Await ObjIntMisFact.SActualiceDoc(lobjDoc, aenuTipoDoc, i = 1)
#Else
                    Await ObjIntMisFact.SActualiceDoc(lobjDoc, aenuTipoDoc, i = 1, True)
#End If
                End If
                If BlnAceptado Then
                    ObjArgumentoEventoPan.DblCantProcesada = i
                    RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                    If ObjArgumentoEventoPan.BlnCancele Then
                        Exit For
                    End If
                End If
            Next
            If BlnAceptado Then
                RaiseEvent EvnFin(Me, ObjArgumentoEventoPan)
            End If
            lblnNoHayError = True
        Catch ex As Exception
            Throw
        End Try
    End Function
    Friend Async Function SVerifiqueCUFE(adrwFacts As DataRow()) As Task
        Dim lblnNoHayError As Boolean
        Dim lobjValorLlave As Object(), lstrPref As String, lentIdDoc As Integer, i = 0
        Dim lobjFac = FobjNewDoc(EnuTipoDocOri.EnuFactura)
        Dim lenuProAct As EnuProcesoDef = FenuProcesoAct(EnuTipoDocOri.EnuFactura)
        Try
            For Each ldrwDoc As DataRow In adrwFacts
                i += 1
                If Not BlnAceptado Then
                    ObjArgumentoEventoPan.EnuProceso = lenuProAct
                    ObjArgumentoEventoPan.DblCantAProcesar = adrwFacts.Length
                    RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
                End If
                lstrPref = ClsPanorama.FobjValorCampo(ldrwDoc("Prefijo"), EnuTipoValor.enuString)
                lentIdDoc = ClsPanorama.FobjValorCampo(ldrwDoc("IdDocu"), EnuTipoValor.enuInteger)
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdDoc}
                lobjFac.SAbra(lobjValorLlave)
                Await ObjIntMisFact.SVerifiqueCUFEFact(lobjFac, i = 1)
                If BlnAceptado Then
                    ObjArgumentoEventoPan.DblCantProcesada = i
                    RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                    If ObjArgumentoEventoPan.BlnCancele Then
                        Exit For
                    End If
                End If
            Next
            If BlnAceptado Then
                RaiseEvent EvnFin(Me, ObjArgumentoEventoPan)
            End If
            lblnNoHayError = True
        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Function FblnInsertarDoc(aobjDoc As ClsCBObjetoPan, aenuTipoDoc As EnuTipoDocOri) As Boolean
        Dim lblnIns = False
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                Dim lobjFac As ClsFactura = aobjDoc
                lblnIns = lobjFac.FblnInsertarEFac()
            Case EnuTipoDocOri.EnuNotaCon
                Dim lobjNco As ClsNotaCon = aobjDoc
                lblnIns = lobjNco.FblnInsertarEFac()
            Case EnuTipoDocOri.EnuNotaCr
                Dim lobjNCr As ClsNotaCr = aobjDoc
                lblnIns = lobjNCr.FblnInsertarEFac()
            Case EnuTipoDocOri.EnuNotaDb
                Dim lobjNDb As ClsNotaDb = aobjDoc
                lblnIns = lobjNDb.FblnInsertarEFac()
            Case EnuTipoDocOri.EnuNotaRevCr
                Dim lobjNRcr As ClsNotaReversionCr = aobjDoc
                lblnIns = lobjNRcr.FblnInsertarEFac()
        End Select
        Return lblnIns
    End Function
    Private Function FblnActualizarDoc(aobjDoc As ClsCBObjetoPan, aenuTipoDoc As EnuTipoDocOri) As Boolean
        Dim lblnAct = False
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                Dim lobjFac As ClsFactura = aobjDoc
                lblnAct = lobjFac.FblnActualizarEstEFac()
            Case EnuTipoDocOri.EnuNotaCon
                Dim lobjNco As ClsNotaCon = aobjDoc
                lblnAct = lobjNco.FblnActualizarEstEFac()
            Case EnuTipoDocOri.EnuNotaCr
                Dim lobjNCr As ClsNotaCr = aobjDoc
                lblnAct = lobjNCr.FblnActualizarEstEFac()
            Case EnuTipoDocOri.EnuNotaDb
                Dim lobjNDb As ClsNotaDb = aobjDoc
                lblnAct = lobjNDb.FblnActualizarEstEFac()
            Case EnuTipoDocOri.EnuNotaRevCr
                Dim lobjNRcr As ClsNotaReversionCr = aobjDoc
                lblnAct = lobjNRcr.FblnActualizarEstEFac()
        End Select
        Return lblnAct
    End Function
#End Region
#End Region
#Region "Actualizar docs estado cero"
    Friend Async Function FstrProceseDocsEstadoCero(aenuTipoDoc As EnuTipoDocOri,
            adrwsDocsAProcesar As DataRow(), astrDocsApi As String) As Task
        Dim lblnNoHayError = False
        GobjPanDat.SControleProcesoObj(True)
        Try
            Dim lstrToken As String = Await ObjIntMisFact.FstrObtengaToken(0)
            If Not String.IsNullOrEmpty(astrDocsApi) Then
                If aenuTipoDoc = EnuTipoDocOri.EnuFactura Then
                    If Not String.IsNullOrEmpty(lstrToken) Then
                        SActuliceFacsEstCero(adrwsDocsAProcesar, astrDocsApi)
                    End If
                ElseIf aenuTipoDoc = EnuTipoDocOri.EnuNotaDb OrElse
                        aenuTipoDoc = EnuTipoDocOri.EnuNotaRevCr Then
                    SActuliceNotasDbEstCero(adrwsDocsAProcesar, astrDocsApi)
                ElseIf aenuTipoDoc = EnuTipoDocOri.EnuNotaCr OrElse
                        aenuTipoDoc = EnuTipoDocOri.EnuNotaCon Then
                    SActuliceNotasCrEstCero(adrwsDocsAProcesar, astrDocsApi)
                End If
                If ObjArgumentoEventoPan.BlnCancele Then
                    Exit Function
                End If
            End If
            lblnNoHayError = True
        Catch ex As Exception
            Throw
        Finally
            If Not lblnNoHayError Then
                ObjArgumentoEventoPan.BlnProcesoOk = False
                RaiseEvent EvnFin(Me, ObjArgumentoEventoPan)
            End If
        End Try
    End Function

    Private Sub SActuliceFacsEstCero(adrwsFactsEstadoCero As DataRow(), astrDocsApi As String)
        Dim lstrPref As String, lentIdFact As Integer, lstrFac As String
        Dim lobjFact As New ClsFactura(), lenuEstadoEDoc As EnuEstadoEDoc
        Dim lobjValorLlave As Object(), lobjEstadoDoc As ClsEstadoDoc
        For Each adrwFact As DataRow In adrwsFactsEstadoCero
            lstrPref = ClsPanorama.FobjValorCampo(adrwFact("Pref"), EnuTipoValor.enuString)
            lentIdFact = ClsPanorama.FobjValorCampo(adrwFact("IdDoc"), EnuTipoValor.enuInteger)
            lstrFac = lstrPref & lentIdFact.ToString()
            lobjEstadoDoc = FobjEstadoDocEnApi(EnuTipoDocOri.EnuFactura, astrDocsApi, lstrFac)
            If lobjEstadoDoc IsNot Nothing Then
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFact}
                lobjFact.SAbra(lobjValorLlave)
                lobjFact.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                lenuEstadoEDoc = FenuEstadoEDoc(lobjEstadoDoc)
                If lenuEstadoEDoc >= EnuEstadoEDoc.EnuRegi Then
                    lenuEstadoEDoc = EnuEstadoEDoc.EnuRegi
                End If
                lobjFact.ObjIdEstadoEDocEnt.ObjValorPro = lenuEstadoEDoc
                lobjFact.ObjCUDocStr.ObjValorPro = lobjEstadoDoc.DocumentUUID
                lobjFact.ObjCUFEStr.ObjValorPro = lobjEstadoDoc.CUFE
                lobjFact.SActualice(True)
            End If
        Next
    End Sub

    Private Sub SActuliceNotasCrEstCero(adrwsNotasCrEstadoCero As DataRow(),
                astrDocsApi As String)
        Dim lstrPref As String, lentIdNotaCr As Integer, lstrCudoc As String, lstrNCr As String
        Dim lobjEstadoDoc As ClsEstadoDoc
        Dim lobjValorLlave As Object(), lenuEstadoEDoc As EnuEstadoEDoc
        For Each adrwNotaCr As DataRow In adrwsNotasCrEstadoCero
            lstrPref = ClsPanorama.FobjValorCampo(adrwNotaCr("Pref"),
                    EnuTipoValor.EnuString)
            lentIdNotaCr = ClsPanorama.FobjValorCampo(adrwNotaCr("IdDoc"),
                    EnuTipoValor.EnuInteger)
            lstrNCr = lstrPref & lentIdNotaCr.ToString()
            lobjEstadoDoc = FobjEstadoDocEnApi(EnuTipoDocOri.EnuNotaCr, astrDocsApi, lstrNCr)
            If lobjEstadoDoc IsNot Nothing Then
                lstrCudoc = lobjEstadoDoc.DocumentUUID
                If lstrPref = "NCON" Then
                    Dim lobjNota As New ClsNotaCon()
                    lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, "", lentIdNotaCr}
                    lobjNota.SAbra(lobjValorLlave)
                    lobjNota.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                    lenuEstadoEDoc = FenuEstadoEDoc(lobjEstadoDoc)
                    If lenuEstadoEDoc >= EnuEstadoEDoc.EnuRegi Then
                        lenuEstadoEDoc = EnuEstadoEDoc.EnuEnProceso
                    End If
                    lobjNota.ObjIdEstadoEDocEnt.ObjValorPro = lenuEstadoEDoc
                    lobjNota.ObjCUDocStr.ObjValorPro = lstrCudoc
                    lobjNota.SActualice(True)
                Else
                    Dim lobjNota As New ClsNotaCr()
                    lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdNotaCr}
                    lobjNota.SAbra(lobjValorLlave)
                    lobjNota.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                    lenuEstadoEDoc = FenuEstadoEDoc(lobjEstadoDoc)
                    If lenuEstadoEDoc >= EnuEstadoEDoc.EnuRegi Then
                        lenuEstadoEDoc = EnuEstadoEDoc.EnuEnProceso
                    End If
                    lobjNota.ObjIdEstadoEDocEnt.ObjValorPro = lenuEstadoEDoc
                    lobjNota.ObjCUDocStr.ObjValorPro = lstrCudoc
                    lobjNota.SActualice(True)
                End If
            End If
        Next
    End Sub

    Private Sub SActuliceNotasDbEstCero(adrwsNotasDbEstadoCero As DataRow(), astrDocsApi As String)
        Dim lstrPref As String, lentIdNotaDb As Integer, lstrCudoc As String, lstrNDb As String
        Dim lobjEstadoDoc As ClsEstadoDoc
        Dim lobjValorLlave As Object(), lenuEstadoEDoc As EnuEstadoEDoc
        For Each adrwFact As DataRow In adrwsNotasDbEstadoCero
            lstrPref = ClsPanorama.FobjValorCampo(adrwFact("Pref"),
                    EnuTipoValor.EnuString)
            lentIdNotaDb = ClsPanorama.FobjValorCampo(adrwFact("IdDoc"),
                    EnuTipoValor.EnuInteger)
            lstrNDb = lstrPref & lentIdNotaDb.ToString()
            lobjEstadoDoc = FobjEstadoDocEnApi(EnuTipoDocOri.EnuNotaDb, astrDocsApi, lstrNDb)
            If lobjEstadoDoc IsNot Nothing Then
                lstrCudoc = lobjEstadoDoc.DocumentUUID
                If lstrPref = "RCR" Then
                    Dim lobjNota As New ClsNotaReversionCr()
                    lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, "", lentIdNotaDb}
                    lobjNota.SAbra(lobjValorLlave)
                    lobjNota.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                    lenuEstadoEDoc = FenuEstadoEDoc(lobjEstadoDoc)
                    If lenuEstadoEDoc >= EnuEstadoEDoc.EnuRegi Then
                        lenuEstadoEDoc = EnuEstadoEDoc.EnuEnProceso
                    End If
                    lobjNota.ObjIdEstadoEDocEnt.ObjValorPro = lenuEstadoEDoc
                    lobjNota.ObjCUDocStr.ObjValorPro = lstrCudoc
                    lobjNota.SActualice(True)
                Else
                    Dim lobjNota As New ClsNotaDb()
                    lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdNotaDb}
                    lobjNota.SAbra(lobjValorLlave)
                    lobjNota.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                    lenuEstadoEDoc = FenuEstadoEDoc(lobjEstadoDoc)
                    If lenuEstadoEDoc >= EnuEstadoEDoc.EnuRegi Then
                        lenuEstadoEDoc = EnuEstadoEDoc.EnuEnProceso
                    End If
                    lobjNota.ObjIdEstadoEDocEnt.ObjValorPro = lenuEstadoEDoc
                    lobjNota.ObjCUDocStr.ObjValorPro = lstrCudoc
                    lobjNota.SActualice(True)
                End If
            End If
        Next
    End Sub
#End Region
    Private Function FobjEstadoDocEnApi(aenuTiopoDoc As EnuTipoDocOri, astrDocsApi As String,
            astrNroDoc As String) As ClsEstadoDoc
        Dim lobjEstadoDoc As ClsEstadoDoc = Nothing, lblnEncontrado = False
        If Not String.IsNullOrEmpty(astrDocsApi) Then
            Dim lstrDocsApi As String = (astrDocsApi.Substring(1, astrDocsApi.Length - 2)).
                Replace("},", "};")
            Dim lstrDocsEnApi As String() = lstrDocsApi.Split(";")
            For Each lstrDocApi As String In lstrDocsEnApi
                If Not String.IsNullOrEmpty(lstrDocApi) Then
                    lobjEstadoDoc = JsonConvert.DeserializeObject(Of ClsEstadoDoc)(lstrDocApi)
                    If aenuTiopoDoc = EnuTipoDocOri.EnuFactura AndAlso
                            lobjEstadoDoc.DocumentType = 1 Then
                        lblnEncontrado = lobjEstadoDoc.InvoiceNumber = astrNroDoc
                        If lblnEncontrado Then Exit For
                    ElseIf (aenuTiopoDoc = EnuTipoDocOri.EnuNotaDb OrElse aenuTiopoDoc =
                            EnuTipoDocOri.EnuNotaRevCr) AndAlso lobjEstadoDoc.DocumentType = 92 Then
                        lblnEncontrado = lobjEstadoDoc.InvoiceNumber = astrNroDoc
                        If lblnEncontrado Then Exit For
                    ElseIf (aenuTiopoDoc = EnuTipoDocOri.EnuNotaCr OrElse aenuTiopoDoc =
                            EnuTipoDocOri.EnuNotaCon) AndAlso lobjEstadoDoc.DocumentType = 91 Then
                        lblnEncontrado = lobjEstadoDoc.InvoiceNumber = astrNroDoc
                        If lblnEncontrado Then Exit For
                    End If
                End If
            Next
        End If
        If Not lblnEncontrado Then
            lobjEstadoDoc = Nothing
        End If
        Return lobjEstadoDoc
    End Function

    Friend Function FobjEstadoDoc(astrDocApi As String) As ClsEstadoDoc
        Dim lobjEstadoDoc As ClsEstadoDoc
        lobjEstadoDoc = JsonConvert.DeserializeObject(Of ClsEstadoDoc)(astrDocApi)
        Return lobjEstadoDoc
    End Function
    Private Async Function SEnvieDocsAPI(adrwDocsAEnviar As DataRow(),
            aenuTipoDoc As EnuTipoDocOri) As Task
        Dim lobjValorLlave As Object(), lstrPref As String, lentIdDoc As Integer, i = 0
        Dim lobjDoc = FobjNewDoc(aenuTipoDoc)
        Dim lenuProEnv = FenuProcesoEnv(aenuTipoDoc)
        For Each ldrwDoc As DataRow In adrwDocsAEnviar
            i += 1
            If Not BlnAceptado Then
                ObjArgumentoEventoPan.EnuProceso = lenuProEnv
                ObjArgumentoEventoPan.DblCantAProcesar = adrwDocsAEnviar.Length
                RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
            End If
            lstrPref = ClsPanorama.FobjValorCampo(ldrwDoc("Prefijo"), EnuTipoValor.enuString)
            lentIdDoc = ClsPanorama.FobjValorCampo(ldrwDoc("IdDocu"), EnuTipoValor.enuInteger)
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdDoc}
            lobjDoc.SAbra(lobjValorLlave)
            Await ObjIntMisFact.SEnvieAdjunto(lobjDoc, aenuTipoDoc)
            If BlnAceptado Then
                ObjArgumentoEventoPan.DblCantProcesada = i
                RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                If ObjArgumentoEventoPan.BlnCancele Then
                    Exit For
                End If
            End If
        Next
        If BlnAceptado AndAlso Not ObjArgumentoEventoPan.BlnCancele Then
            RaiseEvent EvnFin(Me, ObjArgumentoEventoPan)
        End If
    End Function

    Friend Async Function SEnvieDocAPI(aobjDoc As ClsCBObjetoPan,
            aenuTipoDoc As EnuTipoDocOri) As Task
        Dim lenuProEnv = FenuProcesoEnv(aenuTipoDoc)
        Await ObjIntMisFact.SEnvieAdjunto(aobjDoc, aenuTipoDoc)
    End Function

    Private Function FobjNewDoc(aenuTipoDoc As EnuTipoDocOri) As ClsCBObjetoPan
        Dim lobjDocPan As ClsCBObjetoPan
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                lobjDocPan = ObjFactura
            Case EnuTipoDocOri.EnuNotaCon
                lobjDocPan = New ClsNotaCon()
            Case EnuTipoDocOri.EnuNotaCr
                lobjDocPan = New ClsNotaCr()
            Case EnuTipoDocOri.EnuNotaDb
                lobjDocPan = New ClsNotaDb()
            Case EnuTipoDocOri.EnuNotaRevCr
                lobjDocPan = New ClsNotaReversionCr()
            Case Else
                lobjDocPan = Nothing
        End Select
        Return lobjDocPan
    End Function

    Private Function FenuProcesoIns(aenuTipoDoc As EnuTipoDocOri) As EnuProcesoDef
        Dim lenuPro As EnuProcesoDef = EnuProcesoDef.None
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                lenuPro = EnuProcesoDef.EnuInsFacApi
            Case EnuTipoDocOri.EnuNotaCon
                lenuPro = EnuProcesoDef.EnuInsNConApi
            Case EnuTipoDocOri.EnuNotaCr
                lenuPro = EnuProcesoDef.EnuInsNCrApi
            Case EnuTipoDocOri.EnuNotaDb
                lenuPro = EnuProcesoDef.EnuInsNDbApi
            Case EnuTipoDocOri.EnuNotaRevCr
                lenuPro = EnuProcesoDef.EnuInsNRcrApi
        End Select
        Return lenuPro
    End Function

    Private Function FenuProcesoAct(aenuTipoDoc As EnuTipoDocOri) As EnuProcesoDef
        Dim lenuPro As EnuProcesoDef = EnuProcesoDef.None
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                lenuPro = EnuProcesoDef.EnuActFacApi
            Case EnuTipoDocOri.EnuNotaCon
                lenuPro = EnuProcesoDef.EnuActNConApi
            Case EnuTipoDocOri.EnuNotaCr
                lenuPro = EnuProcesoDef.EnuActNCrApi
            Case EnuTipoDocOri.EnuNotaDb
                lenuPro = EnuProcesoDef.EnuActNDbApi
            Case EnuTipoDocOri.EnuNotaRevCr
                lenuPro = EnuProcesoDef.EnuActNRcrApi
        End Select
        Return lenuPro
    End Function

    Private Function FenuProcesoEnv(aenuTipoDoc As EnuTipoDocOri) As EnuProcesoDef
        Dim lenuPro As EnuProcesoDef = EnuProcesoDef.None
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                lenuPro = EnuProcesoDef.EnuEnvFacApi
            Case EnuTipoDocOri.EnuNotaCon
                lenuPro = EnuProcesoDef.EnuEnvNConApi
            Case EnuTipoDocOri.EnuNotaCr
                lenuPro = EnuProcesoDef.EnuEnvNCrApi
            Case EnuTipoDocOri.EnuNotaDb
                lenuPro = EnuProcesoDef.EnuEnvNDbApi
            Case EnuTipoDocOri.EnuNotaRevCr
                lenuPro = EnuProcesoDef.EnuEnvNRcrApi
        End Select
        Return lenuPro
    End Function

#Region "Dispose"
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overridable Overloads Sub Dispose(ablnDisposing As Boolean)
        If Not MblnDisposed Then
            If ablnDisposing Then
                If Not IsNothing(MIntMisFacturas) Then
                    MIntMisFacturas.Dispose()
                    MIntMisFacturas = Nothing
                End If
                MblnDisposed = True
            End If
        End If
    End Sub
#End Region
End Class