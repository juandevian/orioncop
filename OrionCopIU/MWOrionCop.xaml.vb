Imports System.Timers
Imports System.Windows.Media
Imports System.Windows.Threading
Public Class MWOrionCop
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuEstadoWinAyuda As Byte
        None
        EnuAbierto
        EnuCerrado
        EnuEnEspera
    End Enum
    Private Enum EnuEstadoAPP As Byte
        EnuInstalando
        EnuActivo
        EnuPorVencer
        EnuVenceHoy
        EnuVencidoPropio
        EnuVencido
        EnuSuspendido
        EnuInactivo
    End Enum
#End Region
    ' Manejo copia de seguridad 
    Private MentDiaBk As Integer = -1
    Private MentHoraBK As Integer = -1
    Private MentMinBK As Integer = -1
    Private ReadOnly MdispBKIni As New DispatcherTimer
    Private ReadOnly MdispBKFin As New DispatcherTimer
    Private WithEvents MbgwBk As BackgroundWorker = Nothing
    ' Manejo conexión BD
    Private MtmrVerifConec As Timer = Nothing
    Private WithEvents MobjObjetoWin As ClsCentroUtilOriCop
    ' Variables
    Private WithEvents MwinVentana As ClsFormInterface = Nothing
    Private MfrmAyuda As FrmAyuda = Nothing
    ' Menus principales
    Private MnuClientes As MenuItemPan = Nothing
    Private MnuPredios As MenuItemPan = Nothing
    Private MnuFacturacion As MenuItemPan = Nothing
    Private MnuRecibos As MenuItemPan = Nothing
    Private MnuNotasDoc As MenuItemPan = Nothing
    Private MnuEFactura As MenuItemPan = Nothing
    ' Items del menu Acciones
    Private MnuParametrizar As MenuItemPan = Nothing
    Private MnuVerificarIntegridad As MenuItemPan = Nothing
    Private MnuImportarDatosIni As MenuItem = Nothing
    Private MnuImportarFacturasIniciales As MenuItem = Nothing
    Private MnuNoImportarFacIniciales As MenuItem = Nothing
    Private ReadOnly MsepImportarOrion As New Separator
    Private MnuCerrarMes As MenuItemPan = Nothing
    Private MnuCausarInt As MenuItemPan = Nothing
    Private MnuCuentaClientes As MenuItem = Nothing
    Private MnuCuentaPredioAgr As MenuItem = Nothing
    Private MnuCuentaCobro As MenuItem = Nothing
    Private MnuProgramacion As MenuItem = Nothing
    Private MnuFactAuto As MenuItem = Nothing
    Private MnuEnviarFacCorreo As MenuItemPan = Nothing
    Private MnuGenerarPreFacs As MenuItem = Nothing
    Private MnuReversarPreFacs As MenuItem = Nothing
    Private MnuGenerarFacs As MenuItem = Nothing
    Private MnuImportarFactCon As MenuItem = Nothing
    Private MnuInformeCont As MenuItem = Nothing
    Private MnuReimprimirFactu As MenuItem = Nothing
    Private MnuReimprimirFactuCont As MenuItem = Nothing
    Private MnuReimprimirCtasCobro As MenuItem = Nothing
    Private MnuFichaFactura As MenuItem = Nothing
    Private MnuExportarFacsMes As MenuItem = Nothing
    Private MnuExportarFacsFechas As MenuItem = Nothing
    Private MnuFichaNotaCr As MenuItem = Nothing
    Private MnuFichaReintegroAnticipo As MenuItem = Nothing
    Private MnuFichaNotaRevCr As MenuItem = Nothing
    Private MnuFichaNotaIntMora As MenuItem = Nothing
    Private MnuFichaNotaAplAnt As MenuItem = Nothing
    Private MnuFichaNotaAjuste As MenuItem = Nothing
    'EFactura
    Private MnuProceseFactInv As MenuItem = Nothing
    Private MnuProceseFactFF As MenuItem = Nothing
    Private MnuProceseNotasInv As MenuItem = Nothing
    Private MnuProcesarDocsEFac As MenuItem = Nothing
    Private MnuRepDocsNoRegEFac As MenuItem = Nothing
    Private MnuControlarWinEFac As MenuItem = Nothing
    ' Herramientas
    Private MnuInterfazContable As MenuItemPan = Nothing
    Private MnuECorreo As MenuItemPan = Nothing
    Private MnuCopiaSeguridad As MenuItemPan = Nothing
    Private MnuConsultaSql As MenuItem = Nothing
    Private MnuLogApp As MenuItemPan = Nothing
    Private MnuCarpetaReportes As MenuItem = Nothing
    Private MnuBaseDatos As MenuItem = Nothing
    Private MnuRevisarNovs As MenuItem = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomVenMW

    Private WithEvents MobjImpoFacOri As ClsImportarOrion = Nothing
    Private WithEvents MobjReportesOrion As ClsRepOrionCop = Nothing
    Private WithEvents MobjEFacFac As ClsEFactura = Nothing
    Private WithEvents MobjEFacNcr As ClsEFactura = Nothing
    Private WithEvents MobjEFacNdb As ClsEFactura = Nothing
    Private WithEvents MobjEFacNrcr As ClsEFactura = Nothing
    Private WithEvents MobjEFacNCon As ClsEFactura = Nothing
    '   
    Private MdblCantAProcesar As Double = 0.0
    Private MdblCantProcesados As Double = 0.0
    Private MblnInformando As Boolean = False
    Private MstrResultado As String = String.Empty
    Private MstrProceso As String = String.Empty
    ' 
    Private MblnVerificoEstadoHoy As Boolean = False
    Private MblnOcultarAyuda As Boolean = False
#Region "Delegados"
    Private Delegate Sub SdgtActualizaProgressBar(dp As DependencyProperty, value As Object)
    Private Delegate Sub SdgtActualizaLabel(dp As DependencyProperty, Content As Object)
    Private MdgtPgbActualiza As SdgtActualizaProgressBar = Nothing
    Private MdgtLblActualiza As SdgtActualizaLabel = Nothing
#End Region
#End Region

#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuMWOrionCop
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        If GobjAdministrador.FcolCarpetas(True).Count = 0 Then
            MsgBox("No hay Carpetas habilitadas. El programa se cerrará!", vbOKOnly + MsgBoxStyle.Exclamation,
                    "Información")
            SSalirClic()
            Exit Sub
        End If
        If FblnYaEstaCorriendo() Then
            MsgBox("Orión Plus ya esta corriendo en este Equipo!", MsgBoxStyle.OkOnly +
                    MsgBoxStyle.Information, "Información")
            SSalirClic()
            Exit Sub
        End If
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Dim lblnLogon = False
        Try
            SCargueForma(EnuElementosAdicionalesDef.None, 0, Nothing, Nothing, True)
            SLeaArchivoIni()
            SRefresqueBotones()
            lblnLogon = FblnLogOn()
            If lblnLogon Then
                SRegistreLogOn()
                SPuebleBarraEstado()
                MobjObjetoWin.SVerifiqueApp(False, True)
                MnuImportarDatosIni.IsChecked = MobjObjetoWin.BlnImportarFacturas
                SHabiliteMenues()
                If GobjParametros.ObjTipoInterfazByt.ObjValorPro = 0 Then
                    MnuInterfazContable.Visibility = Visibility.Collapsed
                End If
                SInicieControlBK()
                MblnOcultarAyuda = MobjObjetoWin.ObjNoMostrarAyudaBln.ObjValorPro
                SMensajesInicio()
                Dim lblnEstaConectado = False
                SVerifiqueInternet(lblnEstaConectado)
                SVerifiqueContrato(lblnEstaConectado)
                SVerifiqueMenues()
                SInicieControlBD()
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                If Not lblnLogon Then
                    Close()
                End If
            Else
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub

    Protected Overrides ReadOnly Property StrNombreVentana As String
        Get
            Return MstrNombreVentana
        End Get
    End Property

    Protected Overrides ReadOnly Property EnuIdVentana As EnuIdVentanaDef
        Get
            Return HenuIdVentana
        End Get
    End Property

    Protected Overrides Sub SInicialiceObjeto()
        '
    End Sub

    Protected Overrides Sub SInicialiceControles()
        '
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
        SDefinaMenusMW()
        ' Acciones
        Dim lsepAcciones As New Separator
        Dim lsepAcciones2 As New Separator
        Dim lsepNotas As New Separator
        Dim lsepTareasFac As New Separator
        Dim lsepTareasEnvio As New Separator
        Dim lsepEfac As New Separator
        Dim lsepEfac1 As New Separator
        Dim lsepEfac2 As New Separator
        Dim lsepFacFicha As New Separator With {
            .Name = "lsepFacFicha"
        }
        Dim lsepFacFicha_1 As New Separator With {
            .Name = "lsepFacFicha_1"
        }
        Dim lsepHer As New Separator With {
            .Name = "lsepHer1"
        }
        Dim lsepHer2 As New Separator With {
            .Name = "lsepHer2"
        }
        Dim lsepHer3 As New Separator With {
        .Name = "lsepHer3"
        }

        HmnuSalir = MnuSalir
        HmnuRefrescar = MnuRefrescar
        With MenuVen
            .Items.Insert(1, MnuClientes)
            .Items.Insert(2, MnuPredios)
            .Items.Insert(3, MnuFacturacion)
            .Items.Insert(4, MnuRecibos)
            .Items.Insert(5, MnuNotasDoc)
            .Items.Insert(6, MnuEFactura)
        End With
        MnuAcciones.Items.Insert(3, MnuParametrizar)
        MnuAcciones.Items.Insert(4, lsepAcciones2)
        MnuAcciones.Items.Insert(5, MnuCerrarMes)
        MnuAcciones.Items.Insert(6, MnuCausarInt)
        MnuAcciones.Items.Insert(7, lsepAcciones)
        ' Importar datos iniciales
        MnuImportarDatosIni.Items.Add(MnuNoImportarFacIniciales)
        MnuImportarDatosIni.Items.Add(MnuImportarFacturasIniciales)
        MnuAcciones.Items.Insert(7, MsepImportarOrion)
        MnuAcciones.Items.Insert(8, MnuImportarDatosIni)
        ' Verificar Integridad
        MnuAcciones.Items.Insert(9, MnuVerificarIntegridad)
        '
        MnuFactAuto.Items.Add(MnuProgramacion)
        MnuFactAuto.Items.Add(lsepFacFicha_1)
        MnuFactAuto.Items.Add(MnuGenerarPreFacs)
        MnuFactAuto.Items.Add(MnuReversarPreFacs)
        MnuFactAuto.Items.Add(MnuGenerarFacs)
        MnuFactAuto.Items.Add(lsepTareasFac)
        MnuFactAuto.Items.Add(MnuReimprimirFactu)
        MnuFactAuto.Items.Add(MnuReimprimirCtasCobro)
        MnuFactAuto.Items.Add(lsepTareasEnvio)
        If ClsPanorama.FblnEmailsHabilitado Then
            MnuFactAuto.Items.Add(MnuEnviarFacCorreo)
        End If
        '
        MnuFacturacion.Items.Add(MnuFichaFactura)
        MnuFacturacion.Items.Add(MnuFactAuto)
        MnuFacturacion.Items.Add(lsepFacFicha)
        MnuFacturacion.Items.Add(MnuExportarFacsMes)
        MnuFacturacion.Items.Add(MnuExportarFacsFechas)
        ' Estados de Cuenta
        MnuEstadoCuenta.Items.Add(MnuCuentaClientes)
        MnuEstadoCuenta.Items.Add(MnuCuentaPredioAgr)
        MnuEstadoCuenta.Items.Add(MnuCuentaCobro)
        MnuNotasDoc.Items.Add(MnuFichaNotaCr)
        MnuNotasDoc.Items.Add(MnuFichaReintegroAnticipo)
        MnuNotasDoc.Items.Add(MnuFichaNotaRevCr)
        MnuNotasDoc.Items.Add(lsepNotas)
        MnuNotasDoc.Items.Add(MnuFichaNotaIntMora)
        MnuNotasDoc.Items.Add(MnuFichaNotaAplAnt)
        MnuNotasDoc.Items.Add(MnuFichaNotaAjuste)
        HmnuHerramientas.Items.Add(lsepHer)
        'eFactura
        MnuEFactura.Items.Add(MnuProcesarDocsEFac)
        MnuEFactura.Items.Add(lsepEfac)
        MnuEFactura.Items.Add(MnuProceseFactInv)
        MnuEFactura.Items.Add(MnuProceseFactFF)
        MnuEFactura.Items.Add(MnuProceseNotasInv)
        MnuEFactura.Items.Add(lsepEfac1)
        MnuEFactura.Items.Add(MnuInformeCont)
        MnuEFactura.Items.Add(MnuImportarFactCon)
        MnuEFactura.Items.Add(MnuReimprimirFactuCont)
        MnuEFactura.Items.Add(lsepEfac2)
        MnuEFactura.Items.Add(MnuControlarWinEFac)
        MnuEFactura.Items.Add(MnuRepDocsNoRegEFac)
        ' ECorreo
        If ClsPanorama.FblnEmailsHabilitado Then
            HmnuHerramientas.Items.Add(MnuECorreo)
        End If
        ' Interfaz Contable
        HmnuHerramientas.Items.Add(MnuInterfazContable)
        ' Separador
        HmnuHerramientas.Items.Add(lsepHer2)
        ' Copia de seguridad
        HmnuHerramientas.Items.Add(MnuCopiaSeguridad)
        ' Separador
        HmnuHerramientas.Items.Add(lsepHer3)
        ' Log App
        HmnuHerramientas.Items.Add(MnuLogApp)
        ' Reportes Sql
        HmnuHerramientas.Items.Add(MnuConsultaSql)
        ' Abrir Exploador de archivos en la carpeta que contiene los Reportes
        HmnuHerramientas.Items.Add(MnuCarpetaReportes)
        HmnuHerramientas.Items.Add(MnuRevisarNovs)
        'Base De Datos
        HmnuAyuda.Items.Add(MnuBaseDatos)
    End Sub
#End Region

#Region "Procedimientos invalidantes"
    Protected Overrides Sub SHabiliteMenues()
        MyBase.SHabiliteMenues()
        SHabiliteMenuPpal()
        SHabiliteMenuImportarIniciales()
        SHabiliteMenuFinMes()
        If GobjParametros.BlnEFacAutorizado Then
            MnuEFactura.Visibility = Visibility.Visible
        Else
            MnuEFactura.Visibility = Visibility.Collapsed
        End If
        If Not GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
            MnuCausarInt.Visibility = Visibility.Collapsed
        End If
        If GstrIdUsuario <> GCSTRUSUARIOU Then
            MnuRevisarNovs.Visibility = Visibility.Collapsed
        End If
        SComplementeMenuReportes()
    End Sub

    ''' <summary>
    ''' Determina si los items del menu principal definidos en la ventana tienen permiso para 
    ''' el usuario actual asi com el envio de la factura actual
    ''' </summary>
    Private Sub SHabiliteMenuPpal()
        Dim lmnuItemPpal As MenuItem
        Dim lblnPuedeHabiltar As Boolean, lblnTienePermiso As Boolean
        Dim lblnNoInstalando =
                GobjParametros.EnuEstadoInstalacion = EnuEstadoInstalacion.Todos
        With GobjPanorama.ObjUsuarioActual
            For Each lobjObjeto As Object In HmnuMiMenu.Items
                If TypeOf lobjObjeto Is MenuItem Then
                    lmnuItemPpal = lobjObjeto
                    Dim lblnProceda = True
                    Select Case lmnuItemPpal.Name
                        Case "MnuClientes"
                            lblnPuedeHabiltar = MobjObjetoWin.EnuEstadoInstalacion And
                                    EnuEstadoInstalacion.Terceros
                            lblnTienePermiso = lblnPuedeHabiltar AndAlso MnuClientes.IsEnabled
                        Case "MnuPredios"
                            lblnPuedeHabiltar = (MobjObjetoWin.EnuEstadoInstalacion And
                                    EnuEstadoInstalacion.Clientes)
                            lblnTienePermiso = lblnPuedeHabiltar AndAlso MnuPredios.IsEnabled
                        Case "MnuFacturacion"
                            lblnTienePermiso = lblnNoInstalando AndAlso
                                    MnuFacturacion.IsEnabled AndAlso Not ClsOrionCop.FblnCrearAno
                        Case "MnuFichaFactura"
                            lblnTienePermiso = lblnNoInstalando AndAlso MnuFichaFactura.IsEnabled
                        Case "MnuRecibos"
                            lblnTienePermiso = lblnPuedeHabiltar AndAlso MnuRecibos.IsEnabled AndAlso
                                    Not ClsOrionCop.FblnCrearAno
                        Case "MnuNotasDoc"
                            lblnTienePermiso = lblnPuedeHabiltar AndAlso
                                    MnuNotasDoc.IsEnabled AndAlso Not ClsOrionCop.FblnCrearAno
                        Case Else
                            lblnProceda = False
                    End Select
                    If lblnProceda Then
                        SHabiliteMenu(lblnTienePermiso, lmnuItemPpal)
                    End If
                End If
            Next
        End With
    End Sub

    Protected Overrides Sub SRefresqueWin()
        If Not IsNothing(GobjParametros) Then
            Mouse.OverrideCursor = Cursors.Wait
            MobjObjetoWin.SRefresqueObj()
            GobjAdministrador.SRefresque()
            GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.SRefresqueObj()
            GobjPanorama.ObjCarpetaActual.SRefresqueObj()
            MnuImportarFacturasIniciales.IsChecked = Not MobjObjetoWin.BlnImportarFacturas
            SPuebleBarraEstado()
            SMensajesInicio()
            SHabiliteMenues()
            Dim lblnConextado = False
            SVerifiqueInternet(lblnConextado)
            SVerifiqueContrato(lblnConextado)
            SVerifiqueMenues()
            Mouse.OverrideCursor = Cursors.Arrow
        End If
    End Sub

    Protected Overrides Function FblnNotificaOk(aenuIdMensNot As EnuIdMens) As Boolean
        Dim lblnOk = False, lstrMens = String.Empty
        If aenuIdMensNot = EnuIdMens.EnuNoPreFac Then
            lblnOk = ClsOrionCop.FblnPuedePrefacturar(lstrMens)
        End If
        Return lblnOk
    End Function
#End Region

#Region "Procedimientos Propios"
    Private Sub SInicieControlBD()
        MtmrVerifConec = New Timer With {
         .Interval = 3600000,
         .AutoReset = True
         }
        AddHandler MtmrVerifConec.Elapsed, AddressOf SVerificaCnn
        MtmrVerifConec.Start()
    End Sub

    Private Sub SVerificaCnn(sender As Object, e As ElapsedEventArgs)
        GobjPanDat.SDespierteCnn()
    End Sub

    Private Sub SDefinaMenusMW()
        Dim lstrMens As String
        ' Acciones
        MnuParametrizar = FmnuiMenuItemPan("MnuParametrizar", "Parametri_zar", 1, "Ctrl+Z")
        MnuCausarInt = FmnuiMenuItemPan("MnuCausarInt", "Causar _Intereses de Mora", 2, "")
        MnuCerrarMes = FmnuiMenuItemPan("MnuCerrarMes", "C_errar Mes", 3, "")
        MnuVerificarIntegridad = FmnuiMenuItemPan("MnuVerificarIntegridad", "_Verificar Integridad", 4, "")
        ' Importar Datos Iniciales
        MnuImportarDatosIni = FmnuiMenuItem("MnuImportarDatosIni", "_Importar Datos Iniciales",
                "RecMnuItemSec")
        MnuNoImportarFacIniciales = FmnuiMenuItem("MnuNoImportarFacIniciales",
                "No _Importar Facturas Iniciales", "RecMnuItemSec")
        MnuImportarFacturasIniciales = FmnuiMenuItem("MnuImportarFacturasIniciales",
                "Importar _Facturas Iniciales", "RecMnuItemSec")
        MnuNoImportarFacIniciales.IsCheckable = True
        ' Clientes
        MnuClientes = FmnuiMenuItemPan("MnuClientes", "_Clientes", 5, "", True)
        ' Predios
        MnuPredios = FmnuiMenuItemPan("MnuPredios", "_Predios", 6, "", True)
        ' Facturación
        MnuFacturacion = FmnuiMenuItemPan("MnuFacturacion", "_Cobranza", 7, "", True)
        MnuFichaFactura = FmnuiMenuItemPan("MnuFichaFactura", "_Cuentas de cobro", 8, "")
        MnuFactAuto = FmnuiMenuItemPan("MnuFactAuto", "Cobranza _Automática", 9, "")
        MnuFactAuto.ToolTip = "Aqui se llevan a cabo todas las tareas relacionadas con" & vbCrLf &
                "la generación de las Cuentas de Cobro de los servicios programados."
        MnuGenerarPreFacs = FmnuiMenuItemPan("MnuGenerarPreFacs", "Generar _Pre-Cuentas de Cobro", 11, "")
        MnuReversarPreFacs = FmnuiMenuItemPan("MnuReversarPreFacs", "Reversar P_re-Cuentas de Cobro", 12, "")
        MnuGenerarFacs = FmnuiMenuItemPan("MnuGenerarFacs", "Generar _Cuentas de Cobro Definitivas", 13, "")
        MnuReimprimirFactu = FmnuiMenuItem("MnuReimprimirFactu", "Reimprimir Cuentas de Cobro",
                "RecMnuItemSec")
        MnuReimprimirFactu.ToolTip = "Reimprime las Cuentas de Cobro del mes generadas automáticamente"
        MnuEnviarFacCorreo = FmnuiMenuItemPan("MnuEnviarFacCorreo", "Enviar Cuentas de Cobro por E_mail",
                14, "")
        MnuEnviarFacCorreo.ToolTip = "Envía por Correo Electrónico las Cuentas de Cobro automáticas del Mes aún no enviadas"
        MnuExportarFacsMes = FmnuiMenuItemPan("MnuExportarFacsMes", "E_xportar Cuentas de Cobro del Mes a PDF",
                15, "")
        MnuExportarFacsMes.ToolTip = "Exporta cada Factura del Mes actual a un archivo PDF!"
        MnuExportarFacsFechas = FmnuiMenuItemPan("MnuExportarFacsFechas",
                "E_xportar Cuentas de Cobro entre Fechas a PDF", 16, "")
        MnuExportarFacsFechas.ToolTip = "Exporta las Facturas generadas entre dos Fechas a un archivo PDF!"
        MnuProgramacion = FmnuiMenuItemPan("MnuProgramacion",
                "_Programación de Cobros automáticos", 10, "")
        MnuProgramacion.ToolTip = "Aqui se llevan a cabo todas las tareas relacionadas con la" & vbCrLf &
            "programación del cobro de los servicios."
        MnuReimprimirCtasCobro = FmnuiMenuItem("MnuReimprimirCtasCobro",
                "Reimprimir Cu_entas de Cobro adicionales", "RecMnuItemSec")
        MnuReimprimirCtasCobro.ToolTip = "Reimprime los últimas Cuentas de Cobro adicionales"
        ' Recibos Caja
        MnuRecibos = FmnuiMenuItemPan("MnuRecibos", "_Recibos de Caja", 17, "", True)
        ' Notas
        MnuNotasDoc = FmnuiMenuItemPan("MnuNotasDoc", "N_otas", 18, "", True)
        MnuFichaNotaCr = FmnuiMenuItemPan("MnuFichaNotaCr", "_Notas Crédito", 19, "")
        MnuFichaNotaCr.ToolTip = "Permite Consultar, Crear, Anular e Imprimir Notas Crédito."
        MnuFichaReintegroAnticipo = FmnuiMenuItemPan("MnuFichaReintegroAnticipo",
                "Notas Reinte_gro Anticipos", 20, "")
        MnuFichaReintegroAnticipo.ToolTip = "Permite Consultar, Crear, Anular e Imprimir Notas por Reintegro de Anticipos!"
        MnuFichaNotaRevCr = FmnuiMenuItemPan("MnuFichaNotaRevCr", "Notas Re_versión Créditos", 21, "")
        MnuFichaNotaRevCr.ToolTip = "Permite reversar el Movimiento Contable generado" & vbCrLf &
                "por un Recibo de Caja o Nota Cr con Fecha anterior al Periodo actual." & vbCrLf &
                "Además permite consultar e imprimir cualquiera de las Notas ya generadas!"
        MnuFichaNotaIntMora = FmnuiMenuItemPan("MnuFichaNotaIntMora", "Notas Intereses por _Mora",
                22, "")
        MnuFichaNotaIntMora.ToolTip = "Permite Consultar e Imprimir Notas Debito por Intereses de " &
                "Mora causados."
        MnuFichaNotaAplAnt = FmnuiMenuItemPan("MnuFichaNotaAplAnt", "Notas Aplicación de _Anticipos",
                23, "")
        MnuFichaNotaAplAnt.ToolTip = "Permite Consultar e Imprimir Notas de Contabilidad generadas " &
                "al aplicar Anticipos."
        MnuFichaNotaAjuste = FmnuiMenuItemPan("MnuFichaNotaAjuste", "Notas Ajuste Cuota Administración",
                24, "")
        MnuFichaNotaAjuste.ToolTip = "Permite Consultar e Imprimir Notas de Contabilidad generadas al ajustar" & vbCrLf &
                "las Cuotas de Administración."
        ' EFactura
        MnuEFactura = FmnuiMenuItemPan("MnuEFactura", "EFact_ura", 25, "", True)
        MnuEFactura.Visibility = Visibility.Collapsed
        MnuProcesarDocsEFac = FmnuiMenuItemPan("MnuProcesarDocsEFac",
                "Procesar Documentos Elec_trónicos", 26, "")
        lstrMens = "Procesa todas los Documentos en Estado " & Chr(34) & "1 - No Registrada" &
                Chr(34) & vbCrLf & "o en Estado" & Chr(34) & "3 - En Proceso" & Chr(34) &
                " o en Estado" & Chr(34) & "4 - Registrado" & Chr(34)
        MnuProcesarDocsEFac.ToolTip = lstrMens
        MnuProceseFactInv = FmnuiMenuItemPan("MnuProceseFactInv",
                "Procesar Facturas In_validas", 27, "")
        lstrMens = "Procesa todas las Facturas en " & "Estado " & Chr(34) & "2 - Invalido" & Chr(34) & vbCrLf &
                " cuando el Estado en el Proveedor" & " es " & Chr(34) & "Esperando RG" & Chr(34)
        MnuProceseFactInv.ToolTip = lstrMens
        MnuProceseFactFF = FmnuiMenuItemPan("MnuProceseFactFF",
                "Procesar Facturas fuera de Fecha", 27, "")
        lstrMens = "Procesa las Facturas con fecha anterior al día de hoy y no registradas " &
                "como electrónicas."
        MnuProceseFactFF.ToolTip = lstrMens
        MnuProceseNotasInv = FmnuiMenuItemPan("MnuProceseNotasInv",
                "Procesar Notas Invali_das", 28, "")
        lstrMens = "Procesa todas las Notas en " & "Estado " & Chr(34) & "2 - Invalido" & Chr(34) & vbCrLf &
                " cuando el Estado en el Proveedor" & " es " & Chr(34) & "Esperando RG" & Chr(34)
        MnuProceseNotasInv.ToolTip = lstrMens
        MnuInformeCont = FmnuiMenuItemPan("MnuInformeCont", "In_forme de Contingencia",
                29, "")
        MnuImportarFactCon = FmnuiMenuItemPan("MnuImportarFactCon", "_Importar Facturas Contingencia",
                30, "")
        MnuReimprimirFactuCont = FmnuiMenuItem("MnuReimprimirFactuCont",
                "Reimprimir Fact_uras de Contingencia", "RecMnuItemSec")
        MnuReimprimirFactuCont.ToolTip = "Reimprime las últimas Facturas de Contingencia"
        MnuRepDocsNoRegEFac = FmnuiMenuItem("MnuRepDocsNoRegEFac",
                "Reporte Doc_umentos no Registrados", "RecMnuItemSec")
        lstrMens = "Muestra una relación de los Documentos no Registrados!"
        MnuRepDocsNoRegEFac.ToolTip = lstrMens
        MnuControlarWinEFac = FmnuiMenuItem("MnuControlarWinEFac", "Mostrar / Ocultar Ventana EFac",
                "RecMnuItemSec")
        MnuControlarWinEFac.ToolTip = "Cambia la visibilidad de la Ventana!"
        ' Cuentas
        MnuCuentaClientes = FmnuiMenuItem("MnuCuentaClientes", "Del _Cliente", "RecMnuItemSec")
        MnuCuentaPredioAgr = FmnuiMenuItem("MnuCuentaPredioAgr", "Del _Predio Agrupador",
                "RecMnuItemSec")
        MnuCuentaCobro = FmnuiMenuItem("MnuCuentaCobro", "C_uentas de Cobro", "RecMnuItemSec")
        ' Herramientas
        MnuECorreo = FmnuiMenuItemPan("MnuECorreo", "Correo Elec_trónico", 32, "")
        MnuInterfazContable = FmnuiMenuItemPan("MnuInterfazContable", "_Generar Interfaz Contable", 33, "")
        MnuCopiaSeguridad = FmnuiMenuItemPan("MnuCopiaSeguridad", "Copia de Segur_idad", 34, "")
        MnuLogApp = FmnuiMenuItemPan("MnuLogApp", "_Log de la Aplicación", 35, "")
        MnuConsultaSql = FmnuiMenuItemPan("MnuConsultaSql", "Consultas _SQL", 36, "")
        MnuCarpetaReportes = FmnuiMenuItem("MnuCarpetaReportes", "Abrir Ca_rpeta Reportes", "RecMnuItemSec")
        MnuBaseDatos = FmnuiMenuItem("MnuBaseDatos", "_Base de Datos", "RecMnuItemSec")
        MnuRevisarNovs = FmnuiMenuItem("MnuRevisarNovs", "Re_visar integridad novedades", "RecMnuItemSec")
    End Sub

    Private Function FblnLogOn() As Boolean
        Do While True
            MwinVentana = New WinLogOn With {
                .WinPadre = Me
            }
            If Visibility = Visibility.Visible Then
                Visibility = Visibility.Hidden
            End If
            MwinVentana.ShowDialog()
            If Visibility = Visibility.Hidden Then
                Visibility = Visibility.Visible
            End If
            If GblnOK Then
                If GobjPanorama.ObjUsuarioActual.FblnDebeCambiarContrasena() Then
                    SCambioConstrasena()
                    If Not GblnOK Then
                        FblnLogOn()
                    Else
                        FblnLogOn()
                        If GblnOK Then
                            Exit Do
                        End If
                    End If
                Else
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop
        MwinVentana = Nothing
        Return GblnOK
    End Function

    Private Sub SCambieRepLegal()
        MwinVentana = New WinRepLegal With {
            .WinPadre = Me,
            .EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
        }
        MwinVentana.ShowDialog()
        If GblnOK Then
            Dim lstrMens = "El Representante Legal fue cambiado exitosamente!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Private Sub SCambioConstrasena()
        Me.Visibility = Visibility.Hidden
        Dim lwinCambioContrasena = New WinCambioContrasena() With {
            .EnuOperacionEnWin = EnuOperacionEnVentana.CenuModificando
            }
        lwinCambioContrasena.ShowDialog()
        If GblnOK Then
            Dim lstrMens = "La Contraseña fue cambiada exitosamente!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
        Me.Visibility = Visibility.Visible
    End Sub

    Private Sub SCambieUbicacion()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            If GblnPosteando Then
                lstrMens = "No es posible cambiar de Ubicación mientras haya un proceso " &
                        "de  EFac activo!"
            Else
                GobjPanDat.SControleProcesoObj(True)
                GobjPanorama.ObjAppActual.SRegistreLogOffOrigen(My.Computer.Name,
                EnuOrigenInstanciamientoDef.EnuEstacionTrabajo)
                GobjPanorama.ObjUsuarioActual.SRegistreLogOff()
                GblnOK = True
                If FblnLogOn() Then
                    SRegistreLogOn()
                End If
                GobjPanDat.SControleProcesoObj(False)
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                SRefrescarClic()
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuInformacion)
                End If
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub

    Private Sub SRegistreLogOn()
        Dim lstrMens = String.Empty
        GobjPanorama.ObjAppActual.SRegistreLogOnOrigen(My.Computer.Name,
                EnuOrigenInstanciamientoDef.EnuEstacionTrabajo, lstrMens)
        HblnLogOnRegistrado = True
        GstrOrigenActual = My.Computer.Name
        ClsOrionCop.SInstancieCentroUtilOriCop()
        MobjObjetoWin = GobjParametros
        ObjObjetoWin = MobjObjetoWin
        If Not String.IsNullOrEmpty(lstrMens) Then
            Dim lstrMensNot = lstrMens.Replace(vbCrLf, " ")
            SLevanteEveNoti(lstrMensNot, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            MsgBox(lstrMens, vbOKOnly, "Información")
        End If
    End Sub

    Private Sub SVerifiqueContrato(ablnEstaConectado As Boolean)
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            Dim lstrTituloMens = String.Empty, lstrMensBox = String.Empty
            Dim lblnCerrar = False, lblnRegistro = False, lblnRegistrado = False
            Dim lenuEstadoAPP As EnuEstadoAPP
            Dim lobjCarpeta As ClsCarpeta = GobjPanorama.ObjCarpetaActual
            Dim lobjCentroUtil As ClsCentroUtilidad = lobjCarpeta.ObjCentroUtilidadActual
            Dim ldtmFechaVence As Date = lobjCentroUtil.ObjFechaVenceContratoDtm.ObjValorPro
            Dim lenuEstado As EnuEstadoContrato = lobjCentroUtil.ObjEstadoContratoByt.ObjValorPro
            Dim lblnEsPropio As Boolean = lobjCentroUtil.ObjEstaLicenciadaBln.ObjValorPro
            Dim ldtmFechaVerificoCon As Date = lobjCentroUtil.ObjFechaVerificoContratoDtm.ObjValorPro
            Dim lentdiasVerificoCon = ClsPanorama.FentDiasEntreFechas(ldtmFechaVerificoCon, Today)
            Dim lentDiasVence = ClsPanorama.FentDiasEntreFechas(Today, ldtmFechaVence)
            If Not GstrIdUsuario.Trim = GCSTRUSUARIOU Then
                If ldtmFechaVence = GCDTMFECHANULA OrElse lentDiasVence < 10 OrElse
                        lentDiasVence < -30 OrElse lentdiasVerificoCon > 31 OrElse lenuEstado <>
                        EnuEstadoContrato.EnuActivo Then
                    If ablnEstaConectado AndAlso Not MblnVerificoEstadoHoy Then
                        MblnVerificoEstadoHoy = True
                        Mouse.OverrideCursor = Cursors.Wait
                        lobjCentroUtil.SActuliceAuriga()
                        ldtmFechaVence = lobjCentroUtil.ObjFechaVenceContratoDtm.ObjValorPro
                        lblnEsPropio = lobjCentroUtil.ObjEstaLicenciadaBln.ObjValorPro
                        lenuEstado = lobjCentroUtil.ObjEstadoContratoByt.ObjValorPro
                        lentDiasVence = ClsPanorama.FentDiasEntreFechas(Today, ldtmFechaVence)
                        Mouse.OverrideCursor = Cursors.Arrow
                    End If
                End If
                If lentDiasVence < 0 AndAlso Not ablnEstaConectado AndAlso Not lblnEsPropio Then
                    If lenuEstado > EnuEstadoContrato.EnuInstalando Then
                        lstrMens = "Debido a que su contrato está vencido, es necesario que este " &
                            "equipo esté conectado a internet para hacer uso de él!"
                        lstrTituloMens = "Vencido"
                        lenuEstadoAPP = EnuEstadoAPP.EnuVencido
                        lstrMens = String.Empty
                        lblnCerrar = True
                    Else
                        lenuEstadoAPP = EnuEstadoAPP.EnuInstalando
                        lstrMens = "En proceso de instalación"
                    End If
                Else
                    If lenuEstado = EnuEstadoContrato.EnuInstalando Then
                        SRegistreNit(lblnRegistro)
                        If lblnRegistro Then
                            lobjCentroUtil.SActuliceAuriga()
                            lenuEstado = EnuEstadoContrato.EnuRegistradoNit
                            lstrMens = "Registrado, en proceso de instalación"
                        Else
                            lblnCerrar = True
                            lstrMens = "En proceso de instalación"
                        End If
                    ElseIf lenuEstado = EnuEstadoContrato.EnuRegistradoNit Then
                        lstrMens = "Registrado, en proceso de instalación"
                    ElseIf lenuEstado = EnuEstadoContrato.EnuActivo Then
                        If lentDiasVence > 0 And lentDiasVence <= 8 Then
                            lstrMens = "Vence en " & lentDiasVence.ToString & " días!"
                            lenuEstadoAPP = EnuEstadoAPP.EnuPorVencer
                        ElseIf lentDiasVence = 0 Then
                            lstrMens = "Vence en el día de hoy!"
                            lenuEstadoAPP = EnuEstadoAPP.EnuVenceHoy
                        ElseIf lentDiasVence > 8 Then
                            lstrMens = "Activo"
                            lenuEstadoAPP = EnuEstadoAPP.EnuActivo
                        Else
                            If lblnEsPropio Then
                                lstrMens = "Su contrato de soporte venció hace " &
                                        (lentDiasVence * -1).ToString() & " dias; sin embargo estamos " &
                                        "en disposición de atenderlo en el momento que lo requiera."
                                lenuEstadoAPP = EnuEstadoAPP.EnuVencidoPropio
                            Else
                                lstrMens = "Su contrato venció hace " &
                                        (lentDiasVence * -1).ToString() &
                                        " dias. Por favor registre su pago en " & Chr(34) &
                                        "optimusoft@outlook.com" & Chr(34) & " para continuar " &
                                        "haciendo uso de Orión Plus. Gracias"
                                lstrMensBox = lstrMens
                                lenuEstadoAPP = EnuEstadoAPP.EnuVencido
                            End If
                        End If
                    ElseIf lenuEstado = EnuEstadoContrato.EnuSuspendido Then
                        lstrMens = "El contrato está susprendido. Para reactivarlo, por favor " &
                                "comuniquese con el área de soporte de OPTIMUSOFT al teléfono " &
                                 Chr(34) & "311 630 0406" & Chr(34)
                        lenuEstadoAPP = EnuEstadoAPP.EnuSuspendido
                        lblnCerrar = True
                        lstrTituloMens = "Suspendido"
                    ElseIf lenuEstado = EnuEstadoContrato.EnuInactivo Then
                        lstrMens = "El contrato está inactivo. Para reactivarlo comuniquese " &
                                "con el área de soporte de OPTIMUSOFT al teléfono " & Chr(34) &
                                "311 630 0406" & Chr(34)
                        lstrTituloMens = "Inactivo"
                        lenuEstadoAPP = EnuEstadoAPP.EnuInactivo
                        lblnCerrar = True
                    End If
                End If
            Else
                lenuEstadoAPP = EnuEstadoAPP.EnuActivo
                lstrMens = "Soporte"
            End If
            SCambieColorEstado(lenuEstadoAPP)
            lblEstadoCon.Text = lstrMens
            If lblnCerrar Then
                If lenuEstado = EnuEstadoContrato.EnuSuspendido Then
                    MnuFactAuto.IsEnabled = False
                Else
                    lstrMensBox = String.Empty
                    HtlbMiBarraHerramientas.IsEnabled = False
                    MenuVen.IsEnabled = False
                End If
            End If
            If Not String.IsNullOrEmpty(lstrMensBox) Then
                MsgBox(lstrMensBox, vbOKOnly, lstrTituloMens)
            End If
            lblnNoHayError = True
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub

    Private Sub SCambieColorEstado(aenuEstadoAPP As EnuEstadoAPP)
        Select Case aenuEstadoAPP
            Case EnuEstadoAPP.EnuInstalando
                grdInfo.Background = New SolidColorBrush(Colors.White) With {
                .Opacity = 1
                }
            Case EnuEstadoAPP.EnuActivo
                grdInfo.Background = New SolidColorBrush(Colors.Transparent) With {
                .Opacity = 1
                }
            Case EnuEstadoAPP.EnuPorVencer
                grdInfo.Background = New SolidColorBrush(Colors.GreenYellow) With {
                .Opacity = 0.2
                }
            Case EnuEstadoAPP.EnuVenceHoy
                grdInfo.Background = New SolidColorBrush(Colors.Red) With {
                .Opacity = 0.2
                }
            Case EnuEstadoAPP.EnuVencidoPropio
                '
            Case EnuEstadoAPP.EnuVencido
                grdInfo.Background = New SolidColorBrush(Colors.MediumVioletRed) With {
                .Opacity = 0.2
                }
            Case EnuEstadoAPP.EnuSuspendido
                grdInfo.Background = New SolidColorBrush(Colors.BlueViolet) With {
                .Opacity = 0.5
                }
            Case EnuEstadoAPP.EnuInactivo
                grdInfo.Background = New SolidColorBrush(Colors.DarkRed) With {
                .Opacity = 0.5
                }
        End Select
    End Sub

    Private Sub SVerifiqueInternet(ByRef ablnConecatdo As Boolean)
        Mouse.OverrideCursor = Cursors.Wait
        Dim lstrMens As String
        ablnConecatdo = FblnHayInternet()
        If Not ablnConecatdo Then
            lstrMens = "Sin conexión"
            lblEstadoInt.Background = New SolidColorBrush(Colors.Red)
            lblEstadoInt.Foreground = New SolidColorBrush(Colors.White)
        Else
            lstrMens = "Conectado"
            lblEstadoInt.Background = New SolidColorBrush(Colors.Transparent)
            lblEstadoInt.Foreground = New SolidColorBrush(Colors.Black)
        End If
        lblEstadoInt.Content = lstrMens
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub

    Private Sub SRegistreNit(ByRef ablnRegistro As Boolean)
        Dim lwinRegistro As New WinRegistroCopropiedad()
        Dim llngNit As Long
        lwinRegistro.ShowDialog()
        If GblnOK Then
            llngNit = lwinRegistro.DblNit
            ClsPanorama.SRegistreNit(llngNit)
            ablnRegistro = True
        Else
            MsgBox("Es indispensable registrar el NIT, de lo contrario " & vbCrLf &
                    "la aplicación no se activará!", vbOKOnly, "NIT no registrado")
            ablnRegistro = False
        End If
    End Sub

    Private Sub SPuebleBarraEstado()
        lblVersion.Content = My.Resources.Espacio & My.Resources.Ver &
                My.Resources.DosPuntosEspacio & GstrVersionApp
        With GobjPanorama.ObjCarpetaActual
            lblEmpresa.Content = .ObjIdCarpetaShr.ToString & " - " & .ObjNombreStr.ToString
            lblCentroUtilidad.Content = .ObjCentroUtilidadActual.ObjIdCentroUtilShr.ToString & " - " &
                    .ObjCentroUtilidadActual.ObjNombreCentroUtilStr.ToString
        End With
        lblUsuario.Content = GobjPanorama.ObjUsuarioActual.ObjNombreUsuarioStr.ObjValorPro
        lblEstacion.Content = My.Computer.Name
        lblFecha.Content = Date.Today.ToLongDateString
        If Not IsNothing(GobjParametros.ObjAnoActual) Then
            lblPeriodo.Content = GobjParametros.ObjAnoActual.StrNombrePeriodoActual
        End If
    End Sub

    Private Sub SAbraVentana()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            MwinVentana.WinPadre = Me
            SOculteVenProcesoEFac()
            MwinVentana.ShowDialog()
            If MwinVentana.EnuIdVentana = EnuIdVentanaDef.EnuCliente OrElse
                    MwinVentana.EnuIdVentana = EnuIdVentanaDef.EnuPredio Then
                GobjParametros.SVerifiqueInstalacion(lstrMens)
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            MwinVentana = Nothing
            SPuebleBarraEstado()
            If Not lblnNoHayError Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub

    Private Sub SHabiliteMenuImportarIniciales()
        With GobjParametros
            If .EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuListoImportar Then
                Dim lblnHabilite = ClsImportarOrion.FblnTienePermisos
                MnuImportarDatosIni.Visibility = Visibility.Visible
                MsepImportarOrion.Visibility = Visibility.Visible
                SHabiliteMenuItem(lblnHabilite, MnuImportarFacturasIniciales)
            Else
                MnuImportarDatosIni.Visibility = Visibility.Collapsed
                MsepImportarOrion.Visibility = Visibility.Collapsed
            End If
        End With
        If GstrIdUsuario = GCSTRUSUARIOU Then
            SHabiliteMenuItem(False, MnuCambiarContrasena)
        End If
    End Sub

    Private Sub SHabiliteMenuFinMes()
        MyBase.SHabiliteMenues()
        Dim lblnTienePermiso = MnuCerrarMes.IsEnabled AndAlso FblnPuedeCerrarMes()
        SHabiliteMenuItemPan(lblnTienePermiso, MnuCerrarMes)
        lblnTienePermiso = MnuCausarInt.IsEnabled
        lblnTienePermiso = lblnTienePermiso AndAlso GobjParametros.EnuEstadoInstalacion =
                EnuEstadoInstalacion.Todos AndAlso FblnPuedeCausarMora()
        SHabiliteMenuItemPan(lblnTienePermiso, MnuCausarInt)
    End Sub

    Private Sub SCierreMes()
        If FblnPuedeCerrarMes() Then
            If ClsOrionCop.FblnDebeCrearAno() Then
                Dim lstrMens = "Antes de cerrar mes, debe crear el nuevo año!"
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                MsgBox(lstrMens, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Información")
                MwinVentana = Nothing
            Else
                MwinVentana = New WinCausaMora(True)
            End If
        End If
    End Sub

    Private Function FblnPuedeCerrarMes() As Boolean
        Dim lblnPuede = GobjParametros.EnuEstadoInstalacion = EnuEstadoInstalacion.Todos
        lblnPuede = lblnPuede AndAlso ((GobjParametros.EnuEstadoAplicacion =
                EnuEstadoAplicacionDef.EnuParaCierreMes AndAlso
                Not GobjParametros.BlnImportarFacturas) OrElse
                GobjParametros.EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuListoImpNDb)
        If lblnPuede Then
            lblnPuede = Not ClsOrionCop.FblnHayItemsPorFacturar()
        End If
        Return lblnPuede
    End Function

    Private Sub SCauseMora()
        Dim lstrMens = String.Empty
        If FblnPuedeCausarMora(lstrMens) Then
            MwinVentana = New WinCausaMora(False)
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Private Shared Function FblnPuedeCausarMora(ByRef astrMens As String) As Boolean
        Dim lblnCause As Boolean
        If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
            lblnCause = Not ClsOrionCop.FblnHacerCierreMes
            If lblnCause Then
                lblnCause = (GobjParametros.ObjFechaUltCausacionGralDtm.ObjValorPro < Date.Today)
                If Not lblnCause Then
                    astrMens = "No es posible volver a causar Intereses de Mora. " &
                            "Ya están causados al Día de Hoy!"
                End If
            End If
            If lblnCause Then
                lblnCause = Not ClsOrionCop.FblnHayPrefacturas
                If Not lblnCause Then
                    astrMens = "No es posible causar Mora cuando hay Pre-Facturas generadas!"
                End If
            End If
        Else
            ' Aqui llega solo después de cerrar mes y debe causar los intereses de mora
            lblnCause = True
        End If
        Return lblnCause
    End Function

    Private Shared Function FblnPuedeCausarMora() As Boolean
        Dim lblnCause = True
        If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
            lblnCause = Not ClsOrionCop.FblnHacerCierreMes
            If lblnCause Then
                lblnCause = (GobjParametros.ObjFechaUltCausacionGralDtm.ObjValorPro < Date.Today)
            End If
            If lblnCause Then
                lblnCause = Not (ClsOrionCop.FblnHayPrefacturas)
            End If
        End If
        Return lblnCause
    End Function

    Friend Sub SMensajesInicio()
        If MobjObjetoWin.ObjAnoActual IsNot Nothing Then
            If MobjObjetoWin.ObjAnoActual.ObjTipoCalculoCuotaByt.ObjValorPro =
                    EnuTipoBaseCalculo.EnuCoeficientePro Then
                lblCCA.Visibility = Visibility.Collapsed
            Else
                lblCCA.Visibility = Visibility.Visible
            End If
        End If
        MobjObjetoWin.SVerifiqueApp(True, True)
        Dim lstrMens = String.Empty
        If GobjParametros.BlnEFacAutorizado Then
            If ClsOrionCop.FblnClientesSinEmail(lstrMens) Then
                MsgBox(lstrMens, vbInformation + vbOKOnly, "Clientes sin email")
            End If
            Dim lstrDocsRechazados = ClsOrionCop.FstrDocsRechazados()
            If Not String.IsNullOrEmpty(lstrDocsRechazados) Then
                MsgBox("Los siguientes Documentos fueron rechazados por el Cliente:" & vbCrLf &
                        lstrDocsRechazados & vbCrLf & "Estos documentos deben ser anulados!",
                       vbInformation + vbOKOnly, "Documentos Rechazados")
            End If
        End If
        If Not MblnOcultarAyuda Then
            If GobjParametros.EnuEstadoAplicacion <> EnuEstadoAplicacionDef.EnuNormal AndAlso
                    GobjParametros.EnuEstadoAplicacion <> EnuEstadoAplicacionDef.None Then
                SMuestreAyuda(True, False)
            End If
        End If
    End Sub

    Friend Sub SMuestreAyuda(ablnVerifiqueEstado As Boolean, ablnNotifique As Boolean)
        MobjObjetoWin.SVerifiqueApp(True, ablnVerifiqueEstado)
        Dim lstrTituloAyuda = String.Empty
        MblnOcultarAyuda = False
        Dim lstrMensAyuda = FstrMensajeAyuda(lstrTituloAyuda)
        If String.IsNullOrEmpty(lstrMensAyuda) Then
            If ablnNotifique Then
                SLevanteEveNoti("No hay Ayuda para mostrar!", "", 0,
                        EnuSeveridadNot.EnuInformacion)
            End If
            brdAyuda.Visibility = Visibility.Hidden
            EnuEstadoAyuda = EnuEstadoAyudaDef.EnuOff
        Else
            brdAyuda.Visibility = Visibility.Visible
            lblAyuda.Content = lstrTituloAyuda
            txtAyuda.Text = lstrMensAyuda
            bttMostrarAyuda.Visibility = Visibility.Hidden
            bttMostrarAyuda.IsEnabled = True
            bttOcultarAyuda.Visibility = Visibility.Visible
            If MfrmAyuda IsNot Nothing Then
                MfrmAyuda.SCierre()
                MfrmAyuda = Nothing
            End If
            EnuEstadoAyuda = EnuEstadoAyudaDef.EnuOn
        End If
    End Sub

    Private Sub SOculteAyuda()
        If brdAyuda.Visibility = Visibility.Visible Then
            brdAyuda.Visibility = Visibility.Hidden
        End If
        bttOcultarAyuda.Visibility = Visibility.Hidden
        bttMostrarAyuda.Visibility = Visibility.Visible
        MblnOcultarAyuda = True
        EnuEstadoAyuda = EnuEstadoAyudaDef.EnuOff
    End Sub

    Friend Sub SAbraWinAyuda()
        MobjObjetoWin.SVerifiqueApp(True, False)
        Dim lstrTituloAyuda = String.Empty
        Dim lstrMensAyuda = FstrMensajeAyuda(lstrTituloAyuda)
        If Not String.IsNullOrEmpty(lstrMensAyuda) Then
            EnuEstadoAyuda = EnuEstadoAyudaDef.EnuFrmOn
            SOculteAyuda()
            If MfrmAyuda Is Nothing Then
                MfrmAyuda = New FrmAyuda(Me) With {
                    .StrTitulo = lstrTituloAyuda,
                    .StrMensaje = lstrMensAyuda
                    }
                MfrmAyuda.Show()
                ClsOrionCop.SiempreEncima(MfrmAyuda.Handle.ToInt32)
            End If
        Else
            SLevanteEveNoti("No hay Ayuda para mostrar!", "", 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Private Sub SCierreWinAyuda()
        If MfrmAyuda IsNot Nothing Then
            MfrmAyuda.SCierre()
            MfrmAyuda = Nothing
        End If
    End Sub

    Friend Sub SRestablezcaWinAyuda()
        If EnuEstadoAyuda <> EnuEstadoAyudaDef.EnuOff Then
            SMuestreAyuda(False, True)
        End If
    End Sub

    Private Sub SGenerePrefacturas()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            If ClsOrionCop.FblnPuedePrefacturar(lstrMens) Then
                SOculteVenProcesoEFac()
                MwinVentana = New WinGenPrefacturas() With {
                    .WinPadre = Me
                }
                MwinVentana.ShowDialog()
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, EnuIdMens.EnuNoPreFac, EnuSeveridadNot.EnuFalta)
                End If
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            MwinVentana = Nothing
        End Try
    End Sub

    Private Sub SReversePrefacturas()
        Dim lstrMens = String.Empty
        If ClsOrionCop.FblnHayPrefacturas Then
            Dim lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                SOculteVenProcesoEFac()
                Dim lwinVentana = New WinReversarPreFacturas With {
                    .WinPadre = Me
                }
                lwinVentana.ShowDialog()
                lblnNoHayError = True
            Catch ex As ErrorInesperadoPanLException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As PanDatException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As PanLException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As Exception
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Finally
                SRefresqueWin()
                If Not lblnNoHayError Then
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        Else
            If GobjParametros.BlnEFacAutorizado Then
                lstrMens = "No hay Pre-Facturas "
            Else
                lstrMens = "No hay Pre-Cuentas de Cobro "
            End If
            lstrMens &= "generadas que puedan ser reversadas!!!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Private Sub SGenereFacturas()
        Dim lstrMens = String.Empty
        If ClsOrionCop.FblnPuedeCrear(EnuTipoDocOri.EnuFactura, False, lstrMens) Then
            Dim lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                SOculteVenProcesoEFac()
                MwinVentana = New WinGenFacturas(Me) With {
                    .WinPadre = Me
                }
                MwinVentana.ShowDialog()
                lblnNoHayError = True
            Catch ex As ErrorInesperadoPanLException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As PanDatException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As PanLException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As Exception
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Finally
                If Not lblnNoHayError Then
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
                MwinVentana = Nothing
            End Try
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If

    End Sub

    Private Sub SImporteFrasCont()
        If FblnPuedeImpoFrasCont() Then
            If MsgBox("Esta segura(o) de importar Facturas de Contingencia", vbYesNo,
                      "ImportaFacturas") = vbYes Then
                Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                Mouse.OverrideCursor = Cursors.Wait
                Try
                    GobjPanDat.SControleProcesoObj(True)
                    Dim lobjOrionCop As New ClsOrionCop(GCOBJREGISTRO, False)
                    If lobjOrionCop.FblnImportoFrasCon(lstrMens) Then
                        BlnFactConti = True
                        SProceseEFac(lstrMens)
                    End If
                    lblnNoHayError = True
                Catch ex As PanLException
                    lstrMens = ex.Message
                    lstrMensEx = ex.ToString
                Catch ex As PanDatException
                    lstrMens = ex.Message
                    lstrMensEx = ex.ToString
                Catch ex As ArgumentNullException
                    lstrMens = ex.Message
                    lstrMensEx = ex.ToString
                Catch ex As Exception
                    lstrMens = ex.Message
                    lstrMensEx = ex.ToString
                Finally
                    If lblnNoHayError Then
                        GobjPanDat.SConfirmeTransaccion()
                        GobjPanDat.SControleProcesoObj(False)
                        If String.IsNullOrEmpty(lstrMens) Then
                            lstrMens = ("Las Facturas de Contingencia fueron procesadas exitosamente!")
                        End If
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    Else
                        GobjPanDat.SAborteTransaccion()
                        GobjPanDat.SControleProcesoObj(False, True)
                        SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                    End If
                    Mouse.OverrideCursor = Cursors.Arrow
                End Try
            End If
        End If
    End Sub

    Private Function FblnPuedeImpoFrasCont() As Boolean
        Dim lblnPuede = GobjParametros.BlnEFacAutorizado
        If lblnPuede Then
            Dim lstrArchivo = GstrTrayDatPrg & "PlantillaFrasCon_OrionPLus.xlsx"
            lblnPuede = My.Computer.FileSystem.FileExists(lstrArchivo)
        End If
        Return lblnPuede
    End Function

    Private Sub SGenereRepCuotasAdminProp(astrIdAno As String)
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Mouse.OverrideCursor = Cursors.Wait
        Try
            GobjPanDat.SControleProcesoObj(True)
            Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO)
            Dim lobjParametrosRep As New ClsParametrosReportesDocs("", CType(astrIdAno, Integer), 0)
            With lobjRep
                .ObjParRepDocs = lobjParametrosRep
                .EnuReporte = EnuReporteDef.enuCuotasAdminPropi
                .SGenereReporte()
            End With
            lblnNoHayError = True
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub

    Private Shared Function FstrUltimosAnos() As String()
        Dim lobjAno As ClsAno
        Dim lstrUltAnos() As String = Array.Empty(Of String)()
        Dim j = 0
        For i = GobjParametros.ColAnos.Count To GobjParametros.ColAnos.Count - 3 Step -1
            If i > 0 Then
                lobjAno = GobjParametros.ColAnos(i)
                ReDim Preserve lstrUltAnos(j)
                lstrUltAnos(j) = lobjAno.ObjIdAnoShr.ToString
                j += 1
            End If
        Next
        Return lstrUltAnos
    End Function

    Private Sub SComplementeMenuReportes()
        ' Reportes
        Dim lstrUltAnos = FstrUltimosAnos()
        Dim lmnuItem As MenuItem
        Dim lstrNombreMenuI
        For Each lstrAno As String In lstrUltAnos
            lstrNombreMenuI = "MnuAno" & lstrAno
            lmnuItem = FmnuiMenuItem(lstrNombreMenuI, lstrAno, "RecMnuItemSec")
            MnuRepCuotasAdminProp.Items.Add(lmnuItem)
        Next
    End Sub

    Private Sub SReimpFacsAutoMes()
        Dim lblnExcluirFacEnvEmail = False
        If ClsOrionCop.FblnHayFacs(True) Then
            Dim lobjParaFact As New ClsParametrosReportesDocs(String.Empty, 0, 0)
            If Not GobjParametros.BlnEFacAutorizado Then
                If ClsPanorama.FblnEmailsHabilitado Then
                    If MsgBox("Excluir Facturas enviadas por Email?", vbYesNo,
                              "Imprimir Facturas") = vbYes Then
                        lblnExcluirFacEnvEmail = True
                    End If
                End If
            End If
            Mouse.OverrideCursor = Cursors.Wait
            lobjParaFact.BlnExcluirFacEnvEmail = lblnExcluirFacEnvEmail
            Dim lobjRep As New ClsRepOrionCop(GCOBJREGISTRO) With {
                .EnuReporte = EnuReporteDef.enuFactAutoMes,
                .ObjParRepDocs = lobjParaFact
            }
            lobjRep.SGenereReporte()
            Mouse.OverrideCursor = Cursors.Arrow
        Else
            SLevanteEveNoti("No se han generado Facturas automáticas en el presente Mes!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Private Sub SEnvieEmailsFacturas()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            SEnvieCorreo(EnuTipoCorreoE.EnuFactAuto, 0, "", "", lstrMens)
            lblnNoHayError = True
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            Else
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub

    Private Sub SExporteFrasMes()
        Dim lblnNoFayError = False
        If Not ClsOrionCop.FblnHayFacs(False) Then
            SLevanteEveNoti("Aún no se ha generado Facturas en el presente Mes!", "", 0, EnuSeveridadNot.EnuInformacion)
        Else
            GobjPanDat.SControleProcesoObj(True)
            Mouse.OverrideCursor = Cursors.Wait
            Dim ldtbFacrurasaExportar = ClsOrionCop.FdtbFacsMesExportar
            If ldtbFacrurasaExportar.Rows.Count > 0 Then
                MobjReportesOrion = New ClsRepOrionCop(GCOBJREGISTRO)
                MdgtPgbActualiza = New SdgtActualizaProgressBar(AddressOf pgbProcesos.SetValue)
                MdgtLblActualiza = New SdgtActualizaLabel(AddressOf lblResultado.SetValue)
                SVisibiliceCtls(True)
                Try
                    SLevanteEveNoti("Exportando Facturas!", "", 0, EnuSeveridadNot.EnuInformacion)
                    ClsOrionCop.SCreeCarpetaFras()
                    MobjReportesOrion.SExporteFacsMes(ldtbFacrurasaExportar)
                    SLevanteEveNoti("Las Facturas fueron exportadas exitosamente!", "", 0,
                            EnuSeveridadNot.EnuInformacion)
                    lblnNoFayError = True
                Catch ex As Exception
                    SLevanteEveNoti(ex.Message, ex.ToString(), 0, EnuSeveridadNot.EnuExcep)
                Finally
                    If lblnNoFayError Then
                        GobjPanDat.SControleProcesoObj(False)
                    Else
                        GobjPanDat.SControleProcesoObj(False, True)
                    End If
                    SVisibiliceCtls(False)
                End Try
            Else
                GobjPanDat.SControleProcesoObj(False)
                SLevanteEveNoti("No hay Facturas para ser exportadas!", "", 0,
                        EnuSeveridadNot.EnuInformacion)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End If
    End Sub

    Private Sub SExporteFacsFechas()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        GobjPanDat.SControleProcesoObj(True)
        Try
            GblnOK = True
            ClsOrionCop.SCreeCarpetaFras()
            Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
                .EnuReporte = EnuReporteDef.enuExpFacsFechas
                }
            lobjRep.SGenereReporte()
            If GblnOK Then
                SLevanteEveNoti("Las Facturas fueron exportadas exitosamente!", "", 0,
                        EnuSeveridadNot.EnuInformacion)
            Else
                SLevanteEveNoti("La Exportación fue cancelada por el Usuario!", "", 0,
                        EnuSeveridadNot.EnuInformacion)
            End If
            lblnNoHayError = True
        Catch ex As ConexionBdPanException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ErrorInesperadoPanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If Not lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            Else
                GobjPanDat.SControleProcesoObj(False)
            End If
        End Try
    End Sub

    Private Sub SVerifiqueMenues()
        If GobjParametros.BlnEFacAutorizado Then
            MnuFacturacion.Header = "_Facturación"
            MnuFichaFactura.Header = "_Facturas"
            MnuFactAuto.Header = "Facturación _Automática"
            MnuFactAuto.ToolTip = "Aqui se llevan a cabo todas las tareas relacionadas con" & vbCrLf &
                    "la generación de las Facturas de los servicios programados."
            MnuGenerarPreFacs.Header = "Generar _Pre-Facturas"
            MnuReversarPreFacs.Header = "Reversar P_re-Facturas"
            MnuGenerarFacs.Header = "Generar _Facturas Definitivas"
            MnuReimprimirFactu.Header = "Reimprimir Fact_uras"
            MnuReimprimirFactu.ToolTip = "Reimprime las Facturas del mes generadas automáticamente"
            MnuEnviarFacCorreo.Header = "Enviar Facturas por E_mail"
            MnuEnviarFacCorreo.ToolTip = "Envía por Correo Electrónico las Facturas automáticas del Mes aún no enviadas"
            MnuExportarFacsMes.Header = "E_xportar Facturas del Mes a PDF"
            MnuExportarFacsMes.ToolTip = "Exporta cada Factura del Mes actual a un archivo PDF!"
            MnuExportarFacsFechas.Header = "E_xportar Facturas entre Fechas a PDF"
            MnuExportarFacsFechas.ToolTip = "Exporta las Facturas generadas entre dos Fechas a un archivo PDF!"
            SInicieProcEFac(Me)
            If GobjParametros.EnuEstadoAplicacion =
                    EnuEstadoAplicacionDef.EnuDocPorProEFac OrElse
                    ClsOrionCop.FblnDocPorProcesarEFac OrElse Not GblnEstadoRechazado Then
                Dim lstrMens = String.Empty
                SProceseEFac(lstrMens)
            End If
        Else
            MnuFacturacion.Header = "_Cobranza"
            MnuFichaFactura.Header = "_Cuentas de cobro"
            MnuFactAuto.Header = "Cobranza _Automática"
            MnuFactAuto.ToolTip = "Aqui se llevan a cabo todas las tareas relacionadas con" & vbCrLf &
                    "la generación de las Cuentas de Cobro de los servicios programados."
            MnuGenerarPreFacs.Header = "Generar Pre-Cuentas de Cobro"
            MnuReversarPreFacs.Header = "Reversar P_re-Cuentas de Cobro"
            MnuGenerarFacs.Header = "Generar _Cuentas de Cobro Definitivas"
            MnuReimprimirFactu.Header = "Reimprimir C_uentas de Cobro"
            MnuReimprimirFactu.ToolTip = "Reimprime las Cuentas de Cobro del mes generadas automáticamente"
            MnuEnviarFacCorreo.Header = "Enviar Cuentas de Cobro por E_mail"
            MnuEnviarFacCorreo.ToolTip = "Envía por Correo Electrónico las Facturas automáticas del Mes aún no enviadas"
            MnuExportarFacsMes.Header = "E_xportar Cuentas de Cobro del Mes a PDF"
            MnuExportarFacsMes.ToolTip = "Exporta cada Cuenta de Cobro del Mes actual a un archivo PDF!"
            MnuExportarFacsFechas.Header = "E_xportar Cuentas de Cobro entre Fechas a PDF"
            MnuExportarFacsFechas.ToolTip = "Exporta las Cuentas de Cobro generadas entre dos Fechas a un archivo PDF!"
        End If
    End Sub

    Friend Sub SInformeProcEfac(ablnProcesando As Boolean)
        If ablnProcesando Then
            lblProcEFac.Content = "Procesando documentos electrónicos"
        Else
            lblProcEFac.Content = String.Empty
        End If
    End Sub
#End Region

#Region "Copia Seguridad automática programada"
    Private Sub SInicieControlBK()
        Dim lobjApp = ClsAdministrador.FobjAppActual
        With lobjApp
            If .ObjActivaProgramaBKBln.ObjValorPro Then
                AddHandler MdispBKIni.Tick, AddressOf DispatcherTimerIni_Tick
                AddHandler MdispBKFin.Tick, AddressOf DispatcherTimerFin_Tick
                MentDiaBk = .ObjDiaCopiaSeguridadEnt.ObjValorPro
                MentHoraBK = .ObjHoraCopiaSeguridadEnt.ObjValorPro
                MentMinBK = .ObjMinutosCopiaSeguridadEnt.ObjValorPro
                SInicieControlBK(True, 0, 0, 0, 10)
            End If
        End With
    End Sub

    Private Sub SInicieControlBK(ablnInicio As Boolean, aentDias As Integer,
            ByRef aentHoras As Integer, ByRef aentMinutos As Integer, ByRef aentSegundos As Integer)
        If ablnInicio Then
            MdispBKIni.Interval = New TimeSpan(aentDias, aentHoras, aentMinutos, aentSegundos)
            MdispBKIni.Start()
        Else
            MdispBKFin.Interval = New TimeSpan(aentHoras, aentMinutos, aentSegundos)
            MdispBKFin.Start()
        End If
    End Sub

    Private Sub DispatcherTimerIni_Tick(sender As Object, e As EventArgs)
        If Today.DayOfWeek = MentDiaBk Then
            Dim lentHoraActual = Date.Now.Hour
            Dim lentMinuActual = Date.Now.Minute
            Dim lentMinutos As Integer
            Dim lentHoras = 0
            If MentHoraBK = 0 Then
                MentHoraBK = 24
            End If
            If lentMinuActual > MentMinBK Then
                lentMinutos = 60 - lentMinuActual + MentMinBK
                lentHoras = -1
            Else
                lentMinutos = MentMinBK - lentMinuActual
            End If
            If lentHoraActual < MentHoraBK Then
                lentHoras += (MentHoraBK - lentHoraActual)
            End If
            If lentHoras >= 0 AndAlso lentMinutos >= 0 Then
                SInicieControlBK(False, 0, lentHoras, lentMinutos, 0)
            End If
            MdispBKIni.Interval = New TimeSpan(1, 0, 0, 0)
        Else
            MdispBKIni.Interval = New TimeSpan(1, 0, 0, 0)
        End If
    End Sub

    Private Sub DispatcherTimerFin_Tick(sender As Object, e As EventArgs)
        MbgwBk = New BackgroundWorker
        MbgwBk.RunWorkerAsync()
        MdispBKFin.Stop()
    End Sub

    Private Sub Bgw_DoWork(sender As System.Object,
            e As DoWorkEventArgs) Handles MbgwBk.DoWork
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SGenereBkPan(ClsOrionCop.FstrNombreArchivoCopia(False))
            lblnNoHayError = True
        Catch ex As ConexionBdPanException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ErrorInesperadoPanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If Not lblnNoHayError Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
#End Region

#Region "Importación Facturas Iniciales"
    Private Sub SImporteFacturasIniciales()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Dim lblnImpFacIni = False
        Try
            Mouse.OverrideCursor = Cursors.Wait
            MdgtPgbActualiza = New SdgtActualizaProgressBar(AddressOf pgbProcesos.SetValue)
            MdgtLblActualiza = New SdgtActualizaLabel(AddressOf lblResultado.SetValue)
            MobjImpoFacOri = New ClsImportarOrion
            lstrMens = "Importando Facturas Saldos Iniciales!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            lstrMens = String.Empty
            lblnImpFacIni = MobjImpoFacOri.FblnImportoFacturasIniciales(lstrMens)
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentNullException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                If lblnImpFacIni Then
                    SHabiliteMenuImportarIniciales()
                    MobjImpoFacOri = Nothing
                    SVisibiliceCtls(False)
                    SRefresqueWin()
                    lstrMens = "Las Facturas fueron importadas exitosamente!"
                Else
                    lstrMens &= " El Proceso no se llevo a cabo!"
                End If
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub

    Private Sub SVisibiliceCtls(ablnVisible As Boolean)
        If ablnVisible Then
            lblAccion.Visibility = Visibility.Visible
            lblResultado.Visibility = Visibility.Visible
            pgbProcesos.Visibility = Visibility.Visible
        Else
            lblAccion.Visibility = Visibility.Hidden
            lblResultado.Visibility = Visibility.Hidden
            pgbProcesos.Visibility = Visibility.Hidden
        End If
    End Sub
#End Region

#Region "Eventos en la Ventana"
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            MwinVentana = Nothing
            If Not FblnEjecutoMenu(lelmElemento) Then
                If Not FblnEjecutoReporte(lelmElemento) Then
                    Dim lstrMens = String.Empty
                    Select Case lelmElemento.Name
                        Case "MnuCausarInt"
                            SCauseMora()
                        Case "MnuAbrirReportesGen"
                            Dim lwinVentana As New WinAbrirRepGenerado
                            lwinVentana.ShowDialog()
                        Case "MnuReimprimirFactu"
                            SReimpFacsAutoMes()
                        Case "MnuReimprimirFactuCont"
                            SImprimaFactAut(True)
                        Case "MnuProcesarDocsEFac"
                            BlnFactAuto = False
                            SProceseEFac(lstrMens)
                        Case "MnuReimprimirCtasCobro"
                            Dim lstrUltimasCtaCob = ClsOrionCop.FstrIdUltimasCtasCobro
                            If Not String.IsNullOrEmpty(lstrUltimasCtaCob) Then
                                Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO)
                                SImprimaCtasCobro(lobjRep)
                            Else
                                lstrMens = "No hay Cuentas de Cobro para imprimir!"
                            End If
                        Case "MnuEnviarFacCorreo"
                            SEnvieEmailsFacturas()
                        Case "MnuExportarFacsMes"
                            SExporteFrasMes()
                        Case "MnuExportarFacsFechas"
                            SExporteFacsFechas()
                        Case "MnuInterfazContable"
                            SOculteVenProcesoEFac()
                            Dim lwinVentana = New WinInterfazCont(False) With {
                                .WinPadre = Me
                            }
                            lwinVentana.ShowDialog()
                        Case "MnuCopiaSeguridad"
                            SOculteVenProcesoEFac()
                            Dim lwinVentana = New WinCopiaSeg With {
                                .WinPadre = Me
                            }
                            lwinVentana.ShowDialog()
                            If lwinVentana.BlnReiniciar Then
                                SSalirClic()
                            End If
                        Case "MnuBaseDatos"
                            Dim lwinVentana = New WinBaseDatos
                            lwinVentana.Show()
                        Case "MnuRevisarNovs"
                            Dim lwinVentana = New WinRevisarNovedades
                            lwinVentana.Show()
                        Case Else
                            If IsNumeric(Right(lelmElemento.Name, 4)) Then
                                SGenereRepCuotasAdminProp(Right(lelmElemento.Name, 4))
                            End If
                    End Select
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    End If
                End If
            End If
            If Not IsNothing(MwinVentana) Then
                SAbraVentana()
            End If
        End If
    End Sub

    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            Dim lbttAdicional As Button = lelmElemento
            If lbttAdicional.Name = "bttMostrarAyuda" Then
                SMuestreAyuda(True, True)
            ElseIf lbttAdicional.Name = "bttOcultarAyuda" Then
                EnuEstadoAyuda = EnuEstadoAyudaDef.EnuOff
                SOculteAyuda()
            ElseIf lbttAdicional.Name = "bttAbrirWinAyuda" Then
                If MfrmAyuda Is Nothing OrElse MfrmAyuda.IsDisposed Then
                    SAbraWinAyuda()
                Else
                    MfrmAyuda.Activate()
                    MfrmAyuda.Show()
                End If
            ElseIf lbttAdicional.Name = "bttRefConexion" Then
                Dim lblnEstaConectado = False
                SVerifiqueInternet(lblnEstaConectado)
            End If
        End If
    End Sub

    Private Function FblnEjecutoMenu(amnuMenuItem As MenuItem) As Boolean
        Dim lblnEjecuto = FblnEjecutoMenuAdi(amnuMenuItem)
        If Not lblnEjecuto Then
            lblnEjecuto = True
            Select Case amnuMenuItem.Name
                Case "MnuGenerarPreFacs"
                    SGenerePrefacturas()
                Case "MnuReversarPreFacs"
                    SReversePrefacturas()
                Case "MnuNoImportarFacIniciales"
                    If MnuNoImportarFacIniciales.IsChecked Then
                        MnuImportarFacturasIniciales.Visibility = Visibility.Collapsed
                    Else
                        MnuImportarFacturasIniciales.Visibility = Visibility.Visible
                    End If
                    GobjParametros.BlnImportarFacturas =
                            Not MnuNoImportarFacIniciales.IsChecked
                    SRefresqueWin()
                Case "MnuImportarFacturasIniciales"
                    SImporteFacturasIniciales()
                Case "MnuAuxContable"
                    MwinVentana = New WinAuxiliarCont With {
                    .WinPadre = Me
                    }
                Case "MnuReversarPrefact"
                    SReversePrefacturas()
                Case "MnuGenerarFacs"
                    SGenereFacturas()
                Case "MnuImportarFactCon"
                    SImporteFrasCont()
                Case "MnuInformeCont"
                    MwinVentana = New WinInformeCont With {
                    .WinPadre = Me
                    }
                Case "MnuECorreo"
                    Dim lstrMens = String.Empty
                    SEnvieCorreo(EnuTipoCorreoE.None, 0, "", "", lstrMens)
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        SLevanteEveNoti(lstrMens, String.Empty, 0,
                                EnuSeveridadNot.EnuInformacion)
                    End If
                Case "MnuParametrizar"
                    SCierreWinAyuda()
                    MwinVentana = New WinParametrizacion()
                Case "MnuClientes"
                    MwinVentana = New WinClientes()
                Case "MnuCuentaClientes"
                    MwinVentana = New WinCuentaClientes
                Case "MnuCuentaPredioAgr"
                    MwinVentana = New WinCuentaPredios
                Case "MnuCuentaCobro"
                    MwinVentana = New WinCuentasCobro
                Case "MnuCerrarMes"
                    SCierreMes()
                Case "MnuPredios"
                    MwinVentana = New WinPredios With {
                        .WinPadre = Me
                    }
                Case "MnuProgramacion"
                    MwinVentana = New WinProgramacionFacturas
                Case "MnuFichaFactura"
                    MwinVentana = New WinFacturas(Me)
                Case Else
                    lblnEjecuto = False
            End Select
            If Not IsNothing(MwinVentana) Then
                SAbraVentana()
            End If
        End If
        Return lblnEjecuto
    End Function

    Private Function FblnEjecutoMenuAdi(amnuMenuItem As MenuItem) As Boolean
        Dim lblnEjecuto = True, lstrMens = String.Empty
        Select Case amnuMenuItem.Name
            Case "MnuCambiarRepLegal"
                SCambieRepLegal()
                MwinVentana = Nothing
            Case "MnuCambiarContrasena"
                If GstrIdUsuario <> GCSTRUSUARIOU Then
                    SCambioConstrasena()
                End If
            Case "MnuCambiarUbicacion"
                SCambieUbicacion()
                MwinVentana = Nothing
            Case "MnuRecibos"
                MwinVentana = New WinRecibosCaja()
            Case "MnuAnticipos"
                MwinVentana = New WinAnticipos
            Case "MnuFichaReintegroAnticipo"
                MwinVentana = New WinNotasDevAnt
            Case "MnuFichaNotaRevCr"
                MwinVentana = New WinNotasReversionCr(Me)
            Case "MnuFichaNotaIntMora"
                MwinVentana = New WinNotasIntMora(Me)
            Case "MnuFichaNotaAplAnt"
                MwinVentana = New WinNotasAplicaAnt
            Case "MnuFichaNotaCr"
                MwinVentana = New WinNotasCr(Me)
            Case "MnuFichaNotaAjuste"
                MwinVentana = New WinNotasAjuste
            Case "MnuLogApp"
                SOculteVenProcesoEFac()
                Using lwinVentana = New WinLogApp With {
                    .WinPadre = Me
                }
                    lwinVentana.ShowDialog()
                End Using
            Case "MnuConsultaSql"
                Dim lwinVentana = New WinConsultasSql
                lwinVentana.Show()
            Case "MnuCarpetaReportes"
                Dim NoUsado = Process.Start("explorer.exe", GstrTrayReportes)
            Case "MnuVerificarIntegridad"
                MwinVentana = New WinIntegridad() With {
                    .WinPadre = Me
                }
            Case "MnuProceseFactInv"
                ClsOrionCop.SHabiliteFras()
                BlnFactAuto = False
                SProceseEFac(lstrMens)
            Case "MnuProceseFactFF"
                ClsOrionCop.SProcesarFacsFueraFecha(lstrMens)
                SRefresqueWin()
            Case "MnuProceseNotasInv"
                ClsOrionCop.SHabiliteNotas()
                SProceseEFac(lstrMens)
            Case "MnuCoeficientesPorpP"
                MwinVentana = New WinCoeficientesPropPropietarios()
            Case Else
                lblnEjecuto = False
                If Not IsNothing(MwinVentana) Then
                    SAbraVentana()
                End If
        End Select
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
        Return lblnEjecuto
    End Function

    Private Function FblnEjecutoReporte(amnuMenuItem As MenuItem) As Boolean
        Dim lblnEjecuto = True
        If amnuMenuItem.Name.StartsWith("Mnu") Then
            Mouse.OverrideCursor = Cursors.Wait
            Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO)
            Select Case amnuMenuItem.Name
                Case "MnuRCReversados"
                    lobjRep.EnuReporte = EnuReporteDef.enuRecCajaReversados
                    lobjRep.SGenereReporte()
                Case "MnuRCFechas"
                    lobjRep.EnuReporte = EnuReporteDef.enuRCFechas
                    lobjRep.SGenereReporte()
                Case "MnuValoresFact"
                    lobjRep.EnuReporte = EnuReporteDef.enuValoresFactTodos
                    lobjRep.SGenereReporte()
                Case "MnuResumenMovCont"
                    lobjRep.EnuReporte = EnuReporteDef.enuResumenMovCont
                    lobjRep.SGenereReporte()
                Case "MnuRelacionDocs"
                    lobjRep.EnuReporte = EnuReporteDef.enuRelDocs
                    lobjRep.SGenereReporte()
                Case "MnuCajaBancos"
                    lobjRep.EnuReporte = EnuReporteDef.enuCajaBancos
                    lobjRep.SGenereReporte()
                Case "MnuInfDiario"
                    lobjRep.EnuReporte = EnuReporteDef.enuInformeDiario
                    lobjRep.SGenereReporte()
                Case "MnuCarteraCliente"
                    lobjRep.EnuReporte = EnuReporteDef.enuCarteraPorCliente
                    lobjRep.SGenereReporte()
                Case "MnuCxCDetPorSer"
                    lobjRep.EnuReporte = EnuReporteDef.enuCxCDetPorSer
                    Dim lobjParaRep As New ClsParametrosReportesDocs("", 0, 0) With {
                                .StrIdPredioAgr = String.Empty,
                                .DblIdTercero = 0
                            }
                    lobjRep.ObjParRepDocs = lobjParaRep
                    lobjRep.SGenereReporte()
                Case "MnuCarteraPredAgr"
                    lobjRep.EnuReporte = EnuReporteDef.enuCarteraPorPredioAgr
                    lobjRep.SGenereReporte()
                Case "MnuCarteraPredio"
                    lobjRep.EnuReporte = EnuReporteDef.enuCarteraPorPredio
                    lobjRep.SGenereReporte()
                Case "MnuCarteraServicio"
                    lobjRep.EnuReporte = EnuReporteDef.enuCarteraPorServicio
                    lobjRep.SGenereReporte()
                Case "MnuEdadCartera"
                    lobjRep.EnuReporte = EnuReporteDef.enuEdadCartera
                    lobjRep.SGenereReporte()
                Case "MnuRepEstadoCuenta"
                    lobjRep.EnuReporte = EnuReporteDef.enuEstadoCuentas
                    lobjRep.SGenereReporte()
                Case "MnuRepPrediosSector"
                    lobjRep.EnuReporte = EnuReporteDef.enuPrediosSector
                    lobjRep.SGenereReporte()
                Case "MnuRepPrediosCliente"
                    lobjRep.EnuReporte = EnuReporteDef.enuPrediosPropietario
                    lobjRep.SGenereReporte()
                Case "MnuRepPropiPorCP"
                    lobjRep.EnuReporte = EnuReporteDef.enuPropietariosXCP
                    lobjRep.SGenereReporte()
                Case "MnuRepPropiPorCP_Res"
                    lobjRep.EnuReporte = EnuReporteDef.enuPropietariosXCP_Res
                    lobjRep.SGenereReporte()
                Case "MnuRepDocsNoRegEFac"
                    lobjRep.EnuReporte = EnuReporteDef.enuDocsNoRegEFac
                    lobjRep.SGenereReporte()
                Case "MnuControlarWinEFac"
                    SCambieVisibilidad()
                    SMensajesInicio()
                Case Else
                    lblnEjecuto = False
            End Select
            Mouse.OverrideCursor = Cursors.Arrow
        Else
            lblnEjecuto = False
        End If
        Return lblnEjecuto
    End Function

    Private Sub ClsFormInterface_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyboardDevice.Modifiers = ModifierKeys.Control Then
            If e.Key = Key.Z Then
                FblnEjecutoMenu(MnuParametrizar)
            End If
        End If
    End Sub

    Private Sub ClsFormInterface_Closing(sender As Object, e As CancelEventArgs)
        Dim lblnCierre As Boolean = Not GblnPosteando
        Dim lstrMens = String.Empty
        If Not lblnCierre Then
            lstrMens = "No es posible cerrar la Ventana. Se están enviando" &
                    " Documentos a la API de Facturación Electrónica!"
        Else
            lblnCierre = Not GblnEnviandoEmail
            If Not lblnCierre Then
                lstrMens = "No es posible cerrar la Ventana. Se están enviando" &
                        " Correos electrónicos!"
            End If
        End If
        If Not lblnCierre Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            e.Cancel = True
        Else
            MtmrVerifConec?.Stop()
            If HblnLogOnRegistrado Then
                GobjPanDat.SControleProcesoObj(True)
                GobjPanorama.ObjAppActual.SRegistreLogOffOrigen(My.Computer.Name,
                        EnuOrigenInstanciamientoDef.EnuEstacionTrabajo)
                GobjPanorama.ObjUsuarioActual.SRegistreLogOff()
                GobjPanDat.SControleProcesoObj(False)
                SEscribaArchivoIni()
            End If
            End
        End If
    End Sub

    Private Sub SEvnInicio(aobjSender As Object, e As ClsPanEventArgs) Handles _
            MobjImpoFacOri.EvnInicioImportacion, MobjReportesOrion.EvnInicioExportacion,
            MobjEFacFac.EvnInicio, MobjEFacNcr.EvnInicio, MobjEFacNdb.EvnInicio, MobjEFacNrcr.EvnInicio
        Dim lenuProceso = e.EnuProceso
        Dim lobjEFac As ClsEFactura = Nothing, lblnProcesar = True
        If lenuProceso >= EnuProcesoDef.EnuInsFacApi Then
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
            SVisibiliceCtls(True)
            Select Case lenuProceso
                Case EnuProcesoDef.EnuImpoFras
                    lblAccion.Content = My.Resources.ImpFras
                Case EnuProcesoDef.EnuExpFras
                    lblAccion.Content = My.Resources.ExpFras
                Case EnuProcesoDef.EnuInsFacApi
                    MstrProceso = "Insertando Facturas Electrónicas"
                Case EnuProcesoDef.EnuActFacApi
                    MstrProceso = "Actualizando Facturas Electrónicas"
                Case EnuProcesoDef.EnuEnvFacApi
                    MstrProceso = "Enviando Facturas Electrónicas"
                Case EnuProcesoDef.EnuInsNDbApi
                    MstrProceso = "Insertando Notas Db. Electrónicas"
                Case EnuProcesoDef.EnuActNDbApi
                    MstrProceso = "Actualizando Notas Db. Electrónicas"
                Case EnuProcesoDef.EnuEnvNDbApi
                    MstrProceso = "Enviando Notas Db. Electrónicas"
                Case EnuProcesoDef.EnuInsNCrApi
                    MstrProceso = "Insertando Notas Cr. Electrónicas"
                Case EnuProcesoDef.EnuActNCrApi
                    MstrProceso = "Actualizando Notas Cr. Electrónicas"
                Case EnuProcesoDef.EnuEnvNCrApi
                    MstrProceso = "Enviando Notas Cr. Electrónicas"
                Case EnuProcesoDef.EnuInsNRcrApi
                    MstrProceso = "Insertando Notas Rev. Cr. Electrónicas"
                Case EnuProcesoDef.EnuActNRcrApi
                    MstrProceso = "Actualizando Notas Rev. Cr. Electrónicas"
                Case EnuProcesoDef.EnuEnvNRcrApi
                    MstrProceso = "Enviando Notas Rev. Cr. Electrónicas"
                Case EnuProcesoDef.EnuInsNConApi
                    MstrProceso = "Insertando Notas de Ajuste Electrónicas"
                Case EnuProcesoDef.EnuActNConApi
                    MstrProceso = "Actualizando Notas de Ajuste Electrónicas"
                Case EnuProcesoDef.EnuEnvNConApi
                    MstrProceso = "Enviando Notas de Ajuste Electrónicas"
            End Select
            MdblCantProcesados = 0
            pgbProcesos.Minimum = 0.0
            pgbProcesos.Maximum = MdblCantAProcesar
            pgbProcesos.Value = 0.0
        End If
    End Sub

    Private Sub SEvnAvance(aobjSender As Object, e As ClsPanEventArgs) Handles _
            MobjImpoFacOri.EvnAvance, MobjReportesOrion.EvnAvance
        MdblCantProcesados = e.DblCantProcesada
        MstrResultado = My.Resources.EleProce
        MstrResultado = Format(MdblCantProcesados, "##0") &
                My.Resources.De & Format(MdblCantAProcesar, "##0")
        Dispatcher.Invoke(MdgtPgbActualiza,
                System.Windows.Threading.DispatcherPriority.ContextIdle,
                New Object() {ProgressBar.ValueProperty, MdblCantProcesados})
        Dispatcher.Invoke(MdgtLblActualiza,
                System.Windows.Threading.DispatcherPriority.ContextIdle,
                New Object() {ContentProperty, MstrResultado})
    End Sub

    Private Sub SEvnFin(aobjSender As Object, e As ClsPanEventArgs) Handles _
            MobjEFacFac.EvnFin, MobjEFacNcr.EvnFin, MobjEFacNdb.EvnFin, MobjEFacNrcr.EvnFin
        Dim lobjEFac As ClsEFactura
        If TypeOf aobjSender Is ClsEFactura Then
            lobjEFac = aobjSender
            lobjEFac.BlnAceptado = False
            MblnInformando = False
        End If
        e.SLimpie()
        SVisibiliceCtls(False)
        SRefrescarClic()
    End Sub
#End Region
End Class