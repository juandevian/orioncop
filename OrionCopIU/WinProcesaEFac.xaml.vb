Imports Microsoft.Win32
Imports System.Threading
Imports System.Windows.Threading
Imports System.ComponentModel
Public Class WinProcesaEFac
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
    'Variables
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomEFac
    Private WithEvents MbgwPro As BackgroundWorker = Nothing
    Private MstrProceso As String = String.Empty
    Private MstrResultado As String = String.Empty
    Private MblnInformando As Boolean = False
    Private MdblCantAProcesar As Double = 0
    Private MdblCantProcesados As Double = 0
    Private MblnReiniciarBGW As Boolean = False
    Private MblnCancelando As Boolean = False
    Private MenuDocEnProceso As EnuDocEnProceso = EnuDocEnProceso.None
    Private WithEvents MobjEFactura As ClsEFactura = Nothing
    Private ReadOnly MwinVenPrin As MWOrionCop = Nothing
    '
    Private MblnInicioSesion As Boolean = False
#Region "Enumeradores"
    Private Enum EnuDocEnProceso As Byte
        None
        EnuFactura
        EnuNotas
    End Enum
#End Region
#Region "Delegados"
    Private Delegate Sub SdgtAvanceProgressBar(dp As _
                 System.Windows.DependencyProperty,
                 Value As Object)
    Private Delegate Sub SdgtMinProgressBar(dp As _
                 System.Windows.DependencyProperty,
                 Minimum As Object)
    Private Delegate Sub SdgtActualizaLabel(dp As _
                 System.Windows.DependencyProperty,
                 Content As Object)
    Private Delegate Sub SdgtAvanceLabel(dp As _
                 System.Windows.DependencyProperty,
                 Content As Object)
    Private MdgtPgbAvance As SdgtAvanceProgressBar = Nothing
    Private MdgtPgbMin As SdgtMinProgressBar = Nothing
    Private MdgtLblActualiza As SdgtActualizaLabel = Nothing
    Private MdgtLblAvance As SdgtAvanceLabel = Nothing
#End Region
#End Region

#Region "Constructor"
    Friend Sub New(awinMW As MWOrionCop)
        MblnInicioSesion = True
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuEFac
        MwinVenPrin = awinMW
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SCargueForma(EnuElementosAdicionalesDef.None, 0,
                       Nothing, Nothing, True)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        MbgwPro = New BackgroundWorker With {
            .WorkerSupportsCancellation = True
        }
        MdgtLblActualiza = New SdgtActualizaLabel(AddressOf txtProceso.SetValue)
        MdgtLblAvance = New SdgtAvanceLabel(AddressOf txtAvance.SetValue)
        MdgtPgbAvance = New SdgtAvanceProgressBar(AddressOf pgbAvance.SetValue)
        MdgtPgbMin = New SdgtMinProgressBar(AddressOf pgbAvance.SetValue)
    End Sub
    Protected Overrides ReadOnly Property StrNombreVentana As String
        Get
            Return MstrNombreVentana
        End Get
    End Property
    Protected Overrides ReadOnly Property Enuidventana As EnuIdVentanaDef
        Get
            Return HenuIdVentana
        End Get
    End Property
    Protected Overrides Sub SInicialiceObjeto()
        ObjObjetoWin = GobjParametros
        EnuTipoPermisoObjWin = GobjParametros.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        HbttAceptar.Visibility = Visibility.Collapsed
        HbttCancelar.Content = My.Resources.Cancelar
        '
        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub
    Protected Overrides Sub SMuestreDatos()
        '
    End Sub
    Protected Overrides Sub SValide()
        '
    End Sub
    Protected Overrides Sub SRegistre()
        '
    End Sub
    Protected Overrides Sub SConfigureMenuesPropios()
        '
    End Sub
#End Region

#Region "Procedimientos invalidantes"
    Protected Overrides Sub SCancele()
        MblnCancelando = True
        MblnInicioSesion = False
    End Sub
#End Region

#Region "Métodos"
    Friend Sub SProceseEDocs()
        If MenuDocEnProceso <> EnuDocEnProceso.None Then
            If MenuDocEnProceso = EnuDocEnProceso.EnuNotas AndAlso
                        (BlnFactAuto OrElse BlnFactConti) Then
                MblnCancelando = True
            End If
        Else
            MblnCancelando = False
            If MbgwPro.IsBusy Then
                MblnReiniciarBGW = True
            End If
            SInicieCtrlProcesoEFac()
        End If
    End Sub

    Private Async Function SProceseDocs() As Task(Of String)
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False, i = 0
        Dim lblnSinRed = False
        HbttCancelar.IsEnabled = True
        Try
            SLevanteEveNoti("Procesamiento de Documentos Electrónicos en curso!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
            If MobjEFactura Is Nothing Then
                MobjEFactura = New ClsEFactura
            End If
            Dim ldtbDocsPorProcesar = ClsOrionCop.FdtbFacsPorProcesar()
            Do While ldtbDocsPorProcesar.Rows.Count > 0
                lblnSinRed = Not FblnEstaConectado(GobjParametros.ObjURLStr.ObjValorPro, lstrMens)
                i += 1
                If MblnCancelando OrElse lblnSinRed Then
                    Exit Do
                End If
                Await SProceseFacturas(ldtbDocsPorProcesar)
                ldtbDocsPorProcesar = ClsOrionCop.FdtbFacsPorProcesar()
                If i >= 5 Then
                    Exit Do
                ElseIf i > 1 Then
                    SEspere(0, 1, 0)
                End If
            Loop
            If Not (MblnCancelando OrElse lblnSinRed) Then
                If i < 5 Then
                    If BlnFactAuto Then
                        SImprimaFactAut(False)
                        BlnFactAuto = False
                    ElseIf BlnFactConti Then
                        SImprimaFactAut(True)
                        BlnFactConti = False
                    End If
                End If
                i = 0
                ldtbDocsPorProcesar = ClsOrionCop.FdtbNotasPorProcesar
                Do While ldtbDocsPorProcesar.Rows.Count > 0
                    i += 1
                    If MblnCancelando Then
                        Exit Do
                    End If
                    Await SProceseNotas(ldtbDocsPorProcesar)
                    ldtbDocsPorProcesar = ClsOrionCop.FdtbNotasPorProcesar
                    If i >= 5 OrElse ldtbDocsPorProcesar.Rows.Count = 0 Then
                        Exit Do
                    End If
                Loop
            End If
            If MbgwPro.IsBusy Then
                MblnReiniciarBGW = True
            Else
                If MblnCancelando AndAlso Not lblnSinRed Then
                    If BlnFactAuto OrElse BlnFactConti Then
                        MblnInicioSesion = True
                        SInicieCtrlProcesoEFac()
                    End If
                End If
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            HbttCancelar.IsEnabled = False
            If lblnNoHayError Then
                Dim lblnNotificar = True
                If Not String.IsNullOrEmpty(lstrMens) Then
                    lstrMens &= " Por favor informar a Soporte!"
                ElseIf Not MblnCancelando Then
                    If ClsOrionCop.FblnDocPorProcesarEFac Then
                        lstrMens = "Hay problemas en el prosesamiento de Facturación Electrónica!"
                        If GobjParametros.EnuEstadoAplicacion =
                                EnuEstadoAplicacionDef.EnuParaCierreMes Then
                            lstrMens &= " - Antes debe cerrar el mes!"
                        ElseIf ClsOrionCop.FblnHayFacsFueraFecha Then
                            lstrMens &= " - Hay facturas fuera de fecha!"
                        End If
                    Else
                        lstrMens = "Proceso terminado. Esperando por un nuevo Ciclo!"
                    End If
                Else
                    lstrMens = "Proceso cancelado por el Usuario!"
                    lblnNotificar = False
                End If
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    If lblnNotificar Then
                        MwinVenPrin.SLevanteEveNoti(lstrMens, String.Empty, 0,
                            EnuSeveridadNot.EnuInformacion)
                    End If
                End If
            Else
                Dim lenuSeveNoti = EnuSeveridadNot.EnuExcep
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, lenuSeveNoti)
                GblnPosteando = False
            End If
        End Try
        Return lstrMens
    End Function

    Private Async Function SProceseDocsEstadoCero(adtbDocsPorProcesar As DataTable) As Task(Of Boolean)
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Dim ldtmFechaDesde = GCDTMFECHANULA, ldtmFechaHasta = GCDTMFECHANULA, ldtmFechaDoc As Date
        Dim ldrwsDocPorPro As DataRow(), lstrfiltro As String
        HbttCancelar.IsEnabled = True
        Try
            SLevanteEveNoti("Procesamiento de Documentos Electrónicos en curso!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
            If MobjEFactura Is Nothing Then
                MobjEFactura = New ClsEFactura
            End If
            If adtbDocsPorProcesar.Rows.Count > 0 Then
                For Each ldrwDoc As DataRow In adtbDocsPorProcesar.Rows
                    ldtmFechaDoc = ClsPanorama.FobjValorCampo(ldrwDoc("Fecha"),
                            EnuTipoValor.EnuDate)
                    If ldtmFechaDesde = GCDTMFECHANULA Then
                        ldtmFechaDesde = ldtmFechaDoc
                    ElseIf ldtmFechaDoc < ldtmFechaDesde Then
                        ldtmFechaDesde = ldtmFechaDoc
                    End If
                    If ldtmFechaHasta = GCDTMFECHANULA Then
                        ldtmFechaHasta = ldtmFechaDoc
                    ElseIf ldtmFechaDoc > ldtmFechaHasta Then
                        ldtmFechaHasta = ldtmFechaDoc
                    End If
                Next
                lstrfiltro = "TipoDoc > 1"
                If adtbDocsPorProcesar.Select(lstrfiltro).Length > 0 Then
                    ldtmFechaDesde = ldtmFechaDesde.AddDays(-7)
                End If
                If ldtmFechaDesde <= Today.AddDays(-30) Then
                    ldtmFechaDesde = Today.AddDays(-29)
                End If
                If ldtmFechaHasta >= ldtmFechaDesde Then
                    Dim lstrFechaDesde = ldtmFechaDesde.Year.ToString & "-" &
                            Format(ldtmFechaDesde.Month, "0#") & "-" &
                            Format(ldtmFechaDesde.Day - 1, "0#")
                    Dim lstrFechaHasta = ldtmFechaHasta.Year.ToString & "-" &
                            Format(ldtmFechaHasta.Month, "0#") & "-" &
                            Format(ldtmFechaHasta.Day, "0#")
                    Dim lstrDocsApi As String = Await MobjEFactura.FstrDocsApiXFecha(
                            lstrFechaDesde, lstrFechaHasta)
                    If Not String.IsNullOrEmpty(lstrDocsApi) Then
                        ' Facturas
                        lstrfiltro = "TipoDoc = " & EnuTipoDocOri.EnuFactura
                        ldrwsDocPorPro = adtbDocsPorProcesar.Select(lstrfiltro)
                        If ldrwsDocPorPro.Length > 0 Then
                            Await SProceseDocsEstadoCero(ldrwsDocPorPro,
                                    EnuTipoDocOri.EnuFactura, lstrDocsApi)
                        End If
                        ' Notas Db, Rev Cr
                        lstrfiltro = "TipoDoc = " & EnuTipoDocOri.EnuNotaDb & " OR " &
                                "TipoDoc = " & EnuTipoDocOri.EnuNotaRevCr
                        ldrwsDocPorPro = adtbDocsPorProcesar.Select(lstrfiltro)
                        If ldrwsDocPorPro.Length > 0 Then
                            Await SProceseDocsEstadoCero(ldrwsDocPorPro, EnuTipoDocOri.EnuNotaDb,
                                lstrDocsApi)
                        End If
                        ' Notas Cr, Con (notas ajuste)
                        lstrfiltro = "TipoDoc = " & EnuTipoDocOri.EnuNotaCr & " OR " &
                            "TipoDoc = " & EnuTipoDocOri.EnuNotaCon
                        ldrwsDocPorPro = adtbDocsPorProcesar.Select(lstrfiltro)
                        If ldrwsDocPorPro.Length > 0 Then
                            Await SProceseDocsEstadoCero(ldrwsDocPorPro, EnuTipoDocOri.EnuNotaCr,
                                lstrDocsApi)
                        End If
                    End If
                End If
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            HbttCancelar.IsEnabled = False
            If lblnNoHayError Then
                Dim lblnNotificar = True
                If MblnCancelando Then
                    lstrMens = "Proceso cancelado por el Usuario!"
                    lblnNotificar = False
                End If
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    If lblnNotificar Then
                        MwinVenPrin.SLevanteEveNoti(lstrMens, String.Empty, 0,
                            EnuSeveridadNot.EnuInformacion)
                    End If
                End If
            Else
                Dim lenuSeveNoti = EnuSeveridadNot.EnuExcep
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, lenuSeveNoti)
                GblnPosteando = False
            End If
        End Try
        Return lblnNoHayError
    End Function

    Private Async Sub SReenviDocsEstadoCero(adtbDocsEstadoCero As DataTable)
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Dim ldtmFechaDesde = GCDTMFECHANULA, ldtmFechaHasta = GCDTMFECHANULA
        Dim ldtmFechaDoc As Date, lobjValorLlave As Object()
        Dim lstrPrefDoc As String, lentIdDoc As Integer, lenuTipoDoc As EnuTipoDocOri
        If MobjEFactura Is Nothing Then
            MobjEFactura = New ClsEFactura
        End If
        If adtbDocsEstadoCero.Rows.Count > 0 Then
            Try
                For Each ldrwDoc As DataRow In adtbDocsEstadoCero.Rows
                    ldtmFechaDoc = ClsPanorama.FobjValorCampo(ldrwDoc("Fecha"), EnuTipoValor.EnuDate)
                    If ldtmFechaDesde = GCDTMFECHANULA Then
                        ldtmFechaDesde = ldtmFechaDoc
                    ElseIf ldtmFechaDoc < ldtmFechaDesde Then
                        ldtmFechaDesde = ldtmFechaDoc
                    End If
                    If ldtmFechaHasta = GCDTMFECHANULA Then
                        ldtmFechaHasta = ldtmFechaDoc
                    ElseIf ldtmFechaDoc > ldtmFechaHasta Then
                        ldtmFechaHasta = ldtmFechaDoc
                    End If
                Next
                Dim lstrfiltro = "TipoDoc > 1"
                If adtbDocsEstadoCero.Select(lstrfiltro).Length > 0 Then
                    ldtmFechaDesde = ldtmFechaDesde.AddDays(-7)
                End If
                For Each ldrwDoc As DataRow In adtbDocsEstadoCero.Rows
                    lenuTipoDoc = ClsPanorama.FobjValorCampo(ldrwDoc("TipoDoc"),
                            EnuTipoValor.EnuInteger)
                    If lenuTipoDoc = EnuTipoDocOri.EnuFactura Then
                        ldtmFechaDesde = ldtmFechaDesde.AddDays(-7)
                    End If
                    If ldtmFechaDesde < Today.AddDays(-30) Then
                        ldtmFechaDesde = Today.AddDays(-28)
                    End If
                    ldtmFechaDoc = ClsPanorama.FobjValorCampo(ldrwDoc("Fecha"), EnuTipoValor.EnuDate)
                    If ldtmFechaDoc >= ldtmFechaDesde AndAlso ldtmFechaDoc <= ldtmFechaHasta Then
                        lstrPrefDoc = ClsPanorama.FobjValorCampo(ldrwDoc("Pref"), EnuTipoValor.EnuString)
                        lentIdDoc = ClsPanorama.FobjValorCampo(ldrwDoc("IdDoc"), EnuTipoValor.EnuInteger)
                        lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDoc, lentIdDoc}
                        Dim lobjDoc As ClsCBObjetoPan
                        Select Case lenuTipoDoc
                            Case EnuTipoDocOri.EnuFactura
                                lobjDoc = New ClsFactura
                                lobjDoc.SAbra(lobjValorLlave)
                                Await MobjEFactura.SEnvieDocAPI(lobjDoc, EnuTipoDocOri.EnuFactura)
                            Case EnuTipoDocOri.EnuNotaCr
                                lobjDoc = New ClsNotaCr()
                                lobjDoc.SAbra(lobjValorLlave)
                                Dim lobjNcr As ClsNotaCr = lobjDoc
                                If Not String.IsNullOrEmpty(lobjNcr.ObjCUDocStr.ToString) Then
                                    Await MobjEFactura.SEnvieDocAPI(lobjDoc, EnuTipoDocOri.EnuNotaCr)
                                End If
                            Case EnuTipoDocOri.EnuNotaDb
                                lobjDoc = New ClsNotaDb()
                                lobjDoc.SAbra(lobjValorLlave)
                                Dim lobjNdb As ClsNotaDb = lobjDoc
                                If Not String.IsNullOrEmpty(lobjNdb.ObjCUDocStr.ToString) Then
                                    Await MobjEFactura.SEnvieDocAPI(lobjDoc, EnuTipoDocOri.EnuNotaDb)
                                End If
                            Case EnuTipoDocOri.EnuNotaCon
                                lobjDoc = New ClsNotaCon()
                                lobjDoc.SAbra(lobjValorLlave)
                                Dim lobjNcon As ClsNotaCon = lobjDoc
                                If Not String.IsNullOrEmpty(lobjNcon.ObjCUDocStr.ToString) Then
                                    Await MobjEFactura.SEnvieDocAPI(lobjDoc, EnuTipoDocOri.EnuNotaCon)
                                End If
                            Case EnuTipoDocOri.EnuNotaRevCr
                                lobjDoc = New ClsNotaReversionCr()
                                lobjDoc.SAbra(lobjValorLlave)
                                Dim lobjRcr As ClsNotaReversionCr = lobjDoc
                                If Not String.IsNullOrEmpty(lobjRcr.ObjCUDocStr.ToString) Then
                                    Await MobjEFactura.SEnvieDocAPI(lobjDoc,
                                            EnuTipoDocOri.EnuNotaRevCr)
                                End If
                        End Select
                    End If
                Next
                lblnNoHayError = True
            Catch ex As ErrorInesperadoPanDatException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As ErrorInesperadoPanLException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As PanDatException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As PanLException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As ArgumentException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As Exception
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Finally
                HbttCancelar.IsEnabled = False
                If lblnNoHayError Then
                    Dim lblnNotificar = True
                    If MblnCancelando Then
                        lstrMens = "Proceso cancelado por el Usuario!"
                        lblnNotificar = False
                    End If
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                        If lblnNotificar Then
                            MwinVenPrin.SLevanteEveNoti(lstrMens, String.Empty, 0,
                                EnuSeveridadNot.EnuInformacion)
                        End If
                    End If
                Else
                    Dim lenuSeveNoti = EnuSeveridadNot.EnuExcep
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, lenuSeveNoti)
                    GblnPosteando = False
                End If
            End Try
        End If
    End Sub

    Private Async Function SProceseFacturas(adtbFacsPorProcesar As DataTable) As Task
        Dim ldtbAProcesar As DataTable
        ldtbAProcesar = adtbFacsPorProcesar
        If ldtbAProcesar.Rows.Count > 0 Then
            MenuDocEnProceso = EnuDocEnProceso.EnuFactura
            GblnPosteando = True
            Await MobjEFactura.FstrProceseDocsEfac(EnuTipoDocOri.EnuFactura,
                    ldtbAProcesar)
            GblnPosteando = False
            MenuDocEnProceso = EnuDocEnProceso.None
        End If
    End Function

    Private Async Function SProceseNotas(adtbNotsPorProcesar As DataTable) As Task
        If adtbNotsPorProcesar.Rows.Count > 0 Then
            GblnPosteando = True
            MenuDocEnProceso = EnuDocEnProceso.EnuNotas
            Await MobjEFactura.FstrProceseDocsEfac(EnuTipoDocOri.EnuNotaDb, adtbNotsPorProcesar)
            If Not MblnCancelando Then
                Await MobjEFactura.FstrProceseDocsEfac(EnuTipoDocOri.EnuNotaCr,
                        adtbNotsPorProcesar)
            End If
            If Not MblnCancelando Then
                Await MobjEFactura.FstrProceseDocsEfac(EnuTipoDocOri.EnuNotaRevCr,
                        adtbNotsPorProcesar)
            End If
            If Not MblnCancelando Then
                Await MobjEFactura.FstrProceseDocsEfac(EnuTipoDocOri.EnuNotaCon,
                        adtbNotsPorProcesar)
            End If
            GblnPosteando = False
            MenuDocEnProceso = EnuDocEnProceso.None
        End If
    End Function

    Private Async Function SProceseDocsEstadoCero(adrwsDocsPorProcesar As DataRow(),
            aenuTipoDoc As EnuTipoDocOri, astrDocsApi As String) As Task
        If MenuDocEnProceso = EnuDocEnProceso.None Then
            If adrwsDocsPorProcesar.Count > 0 Then
                GblnPosteando = True
                MenuDocEnProceso = aenuTipoDoc
                Await MobjEFactura.FstrProceseDocsEstadoCero(aenuTipoDoc, adrwsDocsPorProcesar,
                        astrDocsApi)
                GblnPosteando = False
                MenuDocEnProceso = EnuDocEnProceso.None
            End If
        End If
    End Function

    Private Async Sub SActualiceFacturasRechazadas()
        Dim ldtmFechadesde = Today.AddDays(-6), ldtmFechaHasta = Today
        Dim lstrFechaDesde = ldtmFechadesde.Year.ToString & "-" &
                            Format(ldtmFechadesde.Month, "0#") & "-" &
                            Format(ldtmFechadesde.Day - 1, "0#")
        Dim lstrFechaHasta = ldtmFechaHasta.Year.ToString & "-" &
                            Format(ldtmFechaHasta.Month, "0#") & "-" &
                            Format(ldtmFechaHasta.Day, "0#")
        If MobjEFactura Is Nothing Then
            MobjEFactura = New ClsEFactura
        End If
        Dim lenuEstadoEDoc As EnuEstadoEDoc
        Dim lstrDocsApi As String = Await MobjEFactura.FstrDocsApiXFecha(
                            lstrFechaDesde, lstrFechaHasta)
        Dim lstrDocumentosEnApi As String = (lstrDocsApi.Substring(1, lstrDocsApi.Length - 2)).
                Replace("},", "};")
        Dim lstrDocsEnApi As String() = lstrDocumentosEnApi.Split(";")
        Dim lobjEstadoDoc As ClsEstadoDoc
        Dim i = 0, lstrIdDoc As String, lstrPrefDoc As String, lstrNumDoc As String
        For Each lstrDocApi As String In lstrDocsEnApi
            lstrPrefDoc = String.Empty
            lstrNumDoc = String.Empty
            If Not String.IsNullOrEmpty(lstrDocApi) Then
                lobjEstadoDoc = MobjEFactura.FobjEstadoDoc(lstrDocApi)
                lenuEstadoEDoc = FenuEstadoEDoc(lobjEstadoDoc)
                If lenuEstadoEDoc = EnuEstadoEDoc.EnuRechazada Then
                    i += 1
                    lstrIdDoc = lobjEstadoDoc.InvoiceNumber
                    For Each lstrLetra As Char In lstrIdDoc
                        If Char.IsLetter(lstrLetra) Then
                            lstrPrefDoc &= lstrLetra
                        ElseIf Char.IsDigit(lstrLetra) Then
                            lstrNumDoc &= lstrLetra
                        End If
                    Next
                    If lobjEstadoDoc.DocumentType = 1 Then
                        ClsOrionCop.SEstadoRechazado(EnuTipoDocOri.EnuFactura,
                                lstrPrefDoc, lstrNumDoc)
                    ElseIf lobjEstadoDoc.DocumentType = 91 Then
                        ClsOrionCop.SEstadoRechazado(EnuTipoDocOri.EnuNotaCr,
                                lstrPrefDoc, lstrNumDoc)
                    ElseIf lobjEstadoDoc.DocumentType = 92 Then
                        ClsOrionCop.SEstadoRechazado(EnuTipoDocOri.EnuNotaDb,
                                lstrPrefDoc, lstrNumDoc)
                    End If
                End If
            End If
        Next
        If i > 0 Then
            MsgBox("Hay documentos rechazados por el cliente!" & vbCrLf &
                   "Debe ir al documento para ser anulado!", vbOKOnly, "Documentos Rechazados")
        End If
    End Sub
#End Region

#Region "Métodos barra de progreso"
    Private Sub SEvnInicio(aobjSender As Object, e As ClsPanEventArgs) Handles MobjEFactura.EvnInicio
        Dim lenuProceso = e.EnuProceso
        Dim lobjEFac As ClsEFactura, lblnProcesar = True
        If lenuProceso >= EnuProcesoDef.enuInsFacApi Then
            lobjEFac = aobjSender
            If Not MblnInformando Then
                lobjEFac.BlnAceptado = True
                MblnInformando = True
            Else
                lblnProcesar = False
            End If
        End If
        If lblnProcesar Then
            MdblCantAProcesar = e.DblCantAProcesar
            Select Case lenuProceso
                Case EnuProcesoDef.enuInsFacApi
                    MstrProceso = "Insertando Facturas"
                Case EnuProcesoDef.enuActFacApi
                    MstrProceso = "Actualizando Facturas"
                Case EnuProcesoDef.enuEnvFacApi
                    MstrProceso = "Enviando Facturas"
                Case EnuProcesoDef.enuInsNDbApi
                    MstrProceso = "Insertando Notas Int. Mora"
                Case EnuProcesoDef.enuActNDbApi
                    MstrProceso = "Actualizando Notas Int. Mora"
                Case EnuProcesoDef.enuEnvNDbApi
                    MstrProceso = "Enviando Notas Int. Mora"
                Case EnuProcesoDef.enuInsNCrApi
                    MstrProceso = "Insertando Notas Crédito"
                Case EnuProcesoDef.enuActNCrApi
                    MstrProceso = "Actualizando Notas Crédito"
                Case EnuProcesoDef.enuEnvNCrApi
                    MstrProceso = "Enviando Notas Crédito"
                Case EnuProcesoDef.enuInsNRcrApi
                    MstrProceso = "Insertando Notas Rev. Cr."
                Case EnuProcesoDef.enuActNRcrApi
                    MstrProceso = "Actualizando Notas Rev. Cr."
                Case EnuProcesoDef.enuEnvNRcrApi
                    MstrProceso = "Enviando Notas Rev. Cr."
                Case EnuProcesoDef.enuInsNConApi
                    MstrProceso = "Insertando Notas de Ajuste"
                Case EnuProcesoDef.enuActNConApi
                    MstrProceso = "Actualizando Notas de Ajuste"
                Case EnuProcesoDef.enuEnvNConApi
                    MstrProceso = "Enviando Notas de Ajuste"
            End Select
            MdblCantProcesados = 0
            Dispatcher.Invoke(MdgtPgbAvance,
                System.Windows.Threading.DispatcherPriority.ContextIdle,
                New Object() {ProgressBar.ValueProperty, MdblCantProcesados})
            Dispatcher.Invoke(MdgtPgbMin,
                System.Windows.Threading.DispatcherPriority.ContextIdle,
                New Object() {ProgressBar.ValueProperty, 0.0})
        End If
    End Sub
    Private Sub SEvnAvance(aobjSender As Object, e As ClsPanEventArgs) Handles _
            MobjEFactura.EvnAvance
        If MblnCancelando Then
            e.BlnCancele = True
            e.EnuProceso = EnuProcesoDef.None
            SEvnInicio(Me, e)
            MenuDocEnProceso = EnuDocEnProceso.None
            Exit Sub
        End If
        MdblCantProcesados = e.DblCantProcesada
        Dim lentAvance = Int((MdblCantProcesados / MdblCantAProcesar) * 100)
        Dim lstrAvance As String = lentAvance.ToString & " %"
        If GblnPosteando Then
            MstrResultado = MstrProceso & " / "
        Else
            MstrResultado = My.Resources.EleProce
        End If
        MstrResultado &= Format(MdblCantProcesados, "##0") &
                My.Resources.De & Format(MdblCantAProcesar, "##0")
        Dispatcher.Invoke(MdgtPgbAvance,
                System.Windows.Threading.DispatcherPriority.ContextIdle,
                New Object() {ProgressBar.ValueProperty, lentAvance})
        Dispatcher.Invoke(MdgtLblActualiza,
                System.Windows.Threading.DispatcherPriority.ContextIdle,
                New Object() {Label.ContentProperty, MstrResultado})
        Dispatcher.Invoke(MdgtLblAvance,
                System.Windows.Threading.DispatcherPriority.ContextIdle,
                New Object() {Label.ContentProperty, lstrAvance})
    End Sub
    Private Sub SEvnFin(aobjSender As Object, e As ClsPanEventArgs) Handles MobjEFactura.EvnFin
        Dim lobjEFac As ClsEFactura
        If TypeOf aobjSender Is ClsEFactura Then
            lobjEFac = aobjSender
            lobjEFac.BlnAceptado = False
            MblnInformando = False
        End If
        MenuDocEnProceso = EnuDocEnProceso.None
        If MblnCancelando Then
            MstrResultado = "Proceso cancelado por el Usuario!"
        Else
            MstrResultado = "Fin del Proceso"
        End If
        Dispatcher.Invoke(MdgtLblActualiza,
                System.Windows.Threading.DispatcherPriority.ContextIdle,
                New Object() {Label.ContentProperty, MstrResultado})
        MdblCantAProcesar = 0.0
        MdblCantProcesados = 0.0
        e.SLimpie()
        e.BlnCancele = MblnCancelando
    End Sub
#End Region

#Region "Proceso dispara automáticamente revisión documentos electrónicos"
    Private Sub SInicieCtrlProcesoEFac()
        If Not MbgwPro.IsBusy Then
            MbgwPro.RunWorkerAsync()
        End If
    End Sub

    Private Sub Bgw_DoWork(sender As Object, e As DoWorkEventArgs) Handles MbgwPro.DoWork
        SInicieControl()
    End Sub

    Private Sub SInicieControl()
        If MblnInicioSesion Then
            SEspereInicio(0, 1)
            MblnInicioSesion = False
        Else
            SEspereInicio(60, 0)
        End If
    End Sub

    Private Sub SEspereInicio(aentMinutos As Integer, aentSegundos As Integer)
        Dim ldtmFechaFin = Now.AddMinutes(aentMinutos).AddSeconds(aentSegundos)
        Dim ldtmFechaActual As Date
        Do While ldtmFechaActual < ldtmFechaFin
            ldtmFechaActual = Now
            If MblnReiniciarBGW Then
                Exit Do
            End If
            SEspere(0, 1, 0)
        Loop
    End Sub

    Private Async Sub Bgw_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) _
            Handles MbgwPro.RunWorkerCompleted
        MblnCancelando = False
        If GobjParametros.EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuDocPorProEFac OrElse
                ClsOrionCop.FblnDocPorProcesarEFac Then
            MwinVenPrin.SInformeProcEfac(True)
            Dim ldtbDocsPorProcesar = ClsOrionCop.FdtbDocsEstadoCero()
            If ldtbDocsPorProcesar.Rows.Count > 0 Then
                Await SProceseDocsEstadoCero(ldtbDocsPorProcesar)
            End If
            Await SProceseDocs()
            If ldtbDocsPorProcesar.Rows.Count > 0 Then
                SReenviDocsEstadoCero(ldtbDocsPorProcesar)
            End If
            MwinVenPrin.SInformeProcEfac(False)
        Else
            If Not GblnEstadoRechazado Then
                SActualiceFacturasRechazadas()
                GblnEstadoRechazado = True
            End If
            SLevanteEveNoti(String.Empty, String.Empty, 0, EnuSeveridadNot.EnuOk)
        End If
            MblnReiniciarBGW = False
        SInicieCtrlProcesoEFac()
    End Sub
#End Region

#Region "Eventos de la Ventana"
    Private Sub ClsFormInterface_Closing(sender As Object, e As CancelEventArgs)
        e.Cancel = True
        SLevanteEveNoti("Esta Ventana no puede ser cerrada desde aquí!", "", 0,
                EnuSeveridadNot.EnuInformacion)
    End Sub
#End Region
End Class