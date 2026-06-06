Imports System.Windows.Controls
Imports System.ComponentModel
Public Class WinNotasIntMora
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Delegados"
    Private Delegate Sub SdgtActualizaLabel(dp As System.Windows.DependencyProperty,
                 Content As Object)
    Private MdgtLblActualiza As SdgtActualizaLabel = Nothing
#End Region

#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuPrefijo
        enuIdNota
    End Enum
#End Region
    Private WithEvents MobjReportes As New ClsRepOrionCop(GCOBJREGISTRO)
    ' Variables
    Private MobjObjetoWin As ClsNotaDb = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomNotInt
    Private MnuEnviarPorCorreo As MenuItem = Nothing
    Private MnuEstadoProvEfac As MenuItem = Nothing
    Private MnuEDoc As MenuItem = Nothing
    Private MnuImportarNsDb As MenuItemPan = Nothing
    Private MnuImprimirNotasMes As MenuItem = Nothing
    Private MnuReprocesarEDoc As MenuItem = Nothing
    Private MnuMostrarLista As MenuItem = Nothing
    Private ReadOnly MwinMW As MWOrionCop = Nothing
    Private MblnPoblandoCombo As Boolean = False
#End Region

#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuNotaIntMora
    End Sub
    Public Sub New(awinMW As MWOrionCop)
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuNotaIntMora
        MwinMW = awinMW
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicioneControlRestringido(dgrItemsNotaDb)
        SAdicioneControlRestringido(dgrNovedades)
        SAdicioneControlRestringido(dgrNotas)
        Dim lcolControlesLlave As New Collection From {
            cboPref,
            txtIdNota
        }
        SCargueForma(EnuElementosAdicionalesDef.enuImprimir, 2,
                lcolControlesLlave, Nothing, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SCambieNombresDocCobro()
        SHabiliteMenuEFac()
        SHabiliteReprocesarEDoc()
        SMuestreNotas(Not BlnVentanaAux)
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
        If IsNothing(ObjObjetoWin) Then
            Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuNotaIntMora)
            ObjObjetoWin = New ClsNotaDb(lstrPref)
            If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                ObjObjetoWin.SVayaAlUltimo()
            End If
        End If
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub

    Protected Overrides Sub SInicialiceControles()
        MdgtLblActualiza = New SdgtActualizaLabel(AddressOf txtNombreCliente.SetValue)
        StcValidaControl(EnuValidEntradaDef.enuPrefijo) = lblIdNotaDb
        StcValidaControl(EnuValidEntradaDef.enuIdNota) = lblIdNotaDb
        SPuebleCombos()
        If Not GobjParametros.BlnEFacAutorizado Then
            lblIva.Visibility = Visibility.Hidden
            txtValorIva.Visibility = Visibility.Hidden
        End If
        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub

    Protected Overrides Sub SMuestreDatos()
        If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos Then
            SLevanteEveNoti("No hay Notas para ser mostradas!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
            cboPref.IsEnabled = False
            txtIdNota.IsEnabled = False
        Else
            With MobjObjetoWin
                cboPref.SelectedItem = .ObjPrefijo_NotaDbStr.ObjValorPro
                txtIdNota.Text = .ObjIdNotaDbEnt.ToString()
                txtIdCliente.Content = .ObjIdCliente_NotaDbDbl.ObjValorPro
                txtNombreCliente.Content = .ObjIdCliente_NotaDbDbl.StrNombreCliente
                txtIdPredioAgru.Content = .ObjIdPredioAgrupador_NotaDbStr.ToString
                txtFechaNotaDb.Content = Format(.ObjFecha_NotaDbDtm.ObjValorPro,
                        GCSTRFMTFECHASIMPLE)
                txtValorNotaDb.Content = Format(.ObjValor_NotaDbDec.ObjValorPro, "c")
                txtValorIva.Content = Format(.DecValorIvaNota, "c")
                If .ObjOrigenByt.ObjValorPro = EnuOrigenNotaDb.EnuAplicacion Then
                    txtOrigen.Content = "Sistema"
                ElseIf .ObjOrigenByt.ObjValorPro = EnuOrigenNotaDb.EnuImportado Then
                    txtOrigen.Content = "Importado"
                Else
                    txtOrigen.Content = ""
                End If
            End With
        End If
        SMuestreEstado()
        SHabiliteMenuEFac()
        SHabiliteReprocesarEDoc()
        Title = My.Resources.FichaNDb
        If Not String.IsNullOrEmpty(txtIdNota.Text) Then
            Title &= txtIdNota.Text & My.Resources.De & txtNombreCliente.Content
        End If
        SValide()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            SMuestreUsuario()
            SEstablezcaDataContext()
            txtIdNota.Focus()
        End If
        MnuMostrarLista.IsEnabled = EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
    End Sub

    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.cenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(EnuValidEntradaDef.enuPrefijo) = .ObjPrefijo_NotaDbStr.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuIdNota) = .ObjIdNotaDbEnt.BlnEsValido
            End With
        End If
        '
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub

    Protected Overrides Sub SRegistre()
        SValide()
    End Sub

    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana 
    ''' y al objeto de la ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        HmnuImprimir = FmnuiMenuItem("MnuImprimir", "Im_primir", "RecMnuItemSec")
        MnuImprimirNotasMes = FmnuiMenuItemPan("MnuImprimirNotasMes", "_Imprimir Notas del Mes",
                1, "")
        MnuEnviarPorCorreo = FmnuiMenuItemPan("MnuEnviarPorCorreo", "_Enviar por eMail", 2, "")
        Dim lentPosicion = HmnuAcciones.Items.Count - 2
        Dim lsepSeparad As New Separator
        If ClsPanorama.FblnEmailsHabilitado Then
            HmnuAcciones.Items.Insert(lentPosicion, MnuEnviarPorCorreo)
        End If
        HmnuAcciones.Items.Insert(lentPosicion, MnuImprimirNotasMes)
        HmnuAcciones.Items.Insert(lentPosicion, HmnuImprimir)
        HmnuAcciones.Items.Insert(lentPosicion, lsepSeparad)
        Dim lsepSeparadEFac As New Separator
        Dim lsepSepEFac As New Separator
        MnuImportarNsDb = FmnuiMenuItemPan("MnuImportarNsDb", "_Importar Notas Intereses Mora", 3, "")
        If GobjParametros.BlnEFacAutorizado Then
            MnuEDoc = FmnuiMenuItemPan("MnuEDoc", "eDocumento", 0, "", True)
            HmnuMiMenu.Items.Insert(2, MnuEDoc)
            MnuEstadoProvEfac = FmnuiMenuItem("MnuEstadoProvEfac", "Estado en Pro_veedor eFactura",
                    "RecMnuItemSec")
            MnuReprocesarEDoc = FmnuiMenuItem("MnuReprocesarEDoc",
                    "_Reprocesar Documento Electrónico", "RecMnuItemSec")
            MnuEDoc.Items.Add(MnuEstadoProvEfac)
            MnuEDoc.Items.Add(MnuReprocesarEDoc)
        End If
        lentPosicion = HmnuAcciones.Items.Count - 2
        HmnuAcciones.Items.Insert(lentPosicion, MnuImportarNsDb)
        HmnuAcciones.Items.Insert(lentPosicion, lsepSepEFac)
        MnuMostrarLista = FmnuiMenuItem("MnuMostrarLista", "  Mostrar lista Notas",
                "RecMnuItemPriInf")
        HmnuMiMenu.Items.Add(MnuMostrarLista)
    End Sub
#End Region

#Region "Procedimientos invalidantes"
    Protected Overrides Sub SImprima()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Dim lblnPuede = MobjObjetoWin.BlnEstaRegEFac OrElse Not MobjObjetoWin.BlnEsDocEle
            If lblnPuede Then
                Mouse.OverrideCursor = Cursors.Wait
                If MobjObjetoWin.BlnExiste Then
                    SLevanteEveNoti("Imprimiendo", String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    Dim lstrPrefNotaDb = MobjObjetoWin.ObjPrefijo_NotaDbStr.ObjValorPro
                    Dim lentIdNotaDbPrimera = MobjObjetoWin.ObjIdNotaDbEnt.ObjValorPro
                    Dim lentIdNotaDbUltima = MobjObjetoWin.ObjIdNotaDbEnt.ObjValorPro
                    Dim lobjParaNotaDb As New ClsParametrosReportesDocs(lstrPrefNotaDb,
                            lentIdNotaDbPrimera, lentIdNotaDbUltima)
                    Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
                        .ObjParRepDocs = lobjParaNotaDb,
                        .EnuReporte = EnuReporteDef.enuNotasDb
                        }
                    lobjRep.SGenereReporte()
                End If
            Else
                lstrMens = "La Nota no se puede imprimir porque aún no esta registrada en API de EFac!"
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
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                Else
                    SLevanteEveNoti(String.Empty, String.Empty, 0, EnuSeveridadNot.EnuOk)
                End If
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub

    Protected Overrides Sub SRefresqueWin()
        MyBase.SRefresqueWin()
        SCambieNombresDocCobro()
    End Sub
#End Region

#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        MyBase.SBuscar()
        If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable Then
            If BlnBusquedaOk AndAlso StrResutadosBusqueda.Length > 0 Then
                cboPref.SelectedItem = StrResutadosBusqueda(0)
                txtIdNota.Text = StrResutadosBusqueda(1)
                SAbraNotaDb()
            End If
        End If
    End Sub
    ''' <summary>
    ''' Invalida la funcion "fblnDefinioBusqueda" de la clase base.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        SDefineNombreCliente()
        Return True
    End Function
    Private Sub SDefineNombreCliente()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsNotaDb.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsIdPredioAgrupador_NotaDbStr.SstrNombreCampoBd,
                                            ClsFecha_NotaDbDtm.SstrNombreCampoBd,
                                            ClsPrefijo_NotaDbStr.SstrNombreCampoBd,
                                            ClsIdNotaDbEnt.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdCliente_NotaDbDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCamposRetornar As String() = {ClsPrefijo_NotaDbStr.SstrNombreCampoBd,
                                             ClsIdNotaDbEnt.SstrNombreCampoBd}
        Dim lstrFiltro = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta &
                " AND P." & StrCampoCentroUtil & " = " & GshrIdCentroUtil
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCamposRetornar, lstrFiltro, False)
    End Sub
#End Region

#Region "Procedimientos Propios"
    Private Sub SInicialiceNota()
        If MobjObjetoWin IsNot Nothing Then
            If cboPref.SelectedItem Is Nothing Then
                cboPref.SelectedIndex = 0
            End If
            Dim lstrPref As String = cboPref.SelectedItem
            If MobjObjetoWin.ObjPrefijo_NotaDbStr.ToString <> lstrPref Then
                ObjObjetoWin = New ClsNotaDb(lstrPref)
                MobjObjetoWin = ObjObjetoWin
                If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                    MobjObjetoWin.SVayaAlUltimo()
                End If
            End If
        End If
    End Sub

    Private Sub SPuebleCombos()
        MblnPoblandoCombo = True
        Dim ldrwConst = ClsOrionCop.FdrwPrefDoc(EnuTipoDocOri.EnuNotaDb)
        SPuebleComboBox(ldrwConst, cboPref)
        MblnPoblandoCombo = False
    End Sub

    Private Sub SMuestreEstado()
        If MobjObjetoWin.ObjAnuladoBln.ObjValorPro Then
            txtEstado.Style = FindResource("RecDocAnulado")
        Else
            txtEstado.Style = FindResource("RecDocNormal")
        End If
    End Sub

    Private Sub SMuestreUsuario()
        With MobjObjetoWin
            If MobjObjetoWin.BlnExiste Then
                txtUsuarioGenero.Content = MobjObjetoWin.ObjIdUsuario_NotaDbStr.ObjValorPro
                If .ObjAnuladoBln.ObjValorPro Then
                    lblUsuarioAnulo.Visibility = Visibility.Visible
                    txtUsuarioAnulo.Visibility = Visibility.Visible
                    txtUsuarioAnulo.Content = .ObjIdUsuarioAnuloStr.ObjValorPro
                Else
                    lblUsuarioAnulo.Visibility = Visibility.Collapsed
                    txtUsuarioAnulo.Visibility = Visibility.Collapsed
                    txtUsuarioAnulo.Content = String.Empty
                End If
            End If
        End With
    End Sub

    Private Sub SEstablezcaDataContext()
        If Not IsNothing(MobjObjetoWin) Then
            If tbiDetalles.IsSelected Then
                cnvDetalles.DataContext = MobjObjetoWin.DtbItemsNotaDb
                SOrdeneDataGrid(dgrItemsNotaDb, dgrItemsNotaDb.Columns(0),
                        ClsIdItemNotaDbShr.SstrNombreCampoBd,
                        ListSortDirection.Ascending)
            ElseIf tbiNovedades.IsSelected Then
                dgrNovedades.DataContext = MobjObjetoWin.DtbNovedadesNotaDb
                SOrdeneDataGrid(dgrNovedades, dgrNovedades.Columns(1),
                        ClsIdNovedadShr.SstrNombreCampoBd, ListSortDirection.Ascending)
            End If
        End If
    End Sub

    Private Sub SAbraNotaDb()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                If txtIdNota.Text <> MobjObjetoWin.ObjIdNotaDbEnt.ToString Then
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, cboPref.SelectedItem,
                            txtIdNota.Text}
                    MobjObjetoWin.SAbra(lobjValorLlave)
                    SMuestreDatos()
                End If
                lblnNoHayError = True
            Catch ex As PanLException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As PanDatException
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
        End If
    End Sub

    Private Sub SEnvieEmailNotaDb()
        Dim lstrMens = String.Empty
        If ClsPanorama.FblnEmailsHabilitado Then
            If MobjObjetoWin.ObjClienteNotaDb.ObjRecibeDocsPorEmailBln.ObjValorPro Then
                SEnvieCorreo(EnuTipoCorreoE.EnuNDB,
                        MobjObjetoWin.ObjClienteNotaDb.ObjIdClienteDbl.ObjValorPro,
                        MobjObjetoWin.ObjIdPredioAgrupador_NotaDbStr.ObjValorPro,
                        MobjObjetoWin.StrNumeroNotaDb, lstrMens)
            Else
                lstrMens = "El Cliente de la Nota de Intereses no tiene habilitado " &
                        "el envio de Documentos por Email!"
            End If
        Else
            lstrMens = "Aún no tiene instalada la Aplicación para el envío de Documentos por Email!"
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Private Sub SImporteNotasDb()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Dim lblnImporto = False
        Try
            lstrMens = "Validando los Datos de Origen!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            SInforme(lstrMens)
            lstrMens = String.Empty
            Mouse.OverrideCursor = Cursors.Wait
            If ClsOrionCop.FblnPuedeImpNsDb(lstrMens) Then
                If MsgBox("Esta segura(o) de importar Notas de Intereses de Mora?", vbYesNo,
                      "Importar Notas") = vbYes Then
                    Dim lblnCierreMes = False, ldtmFechaCausa = GCDTMFECHANULA
                    lstrMens = "Generando Reporte Edad de la Cartera!"
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    lstrMens = MobjReportes.SGenereEdadCartera(GentLimite1, GentLimite2,
                            GentLimite3, GentLimite4, EnuTipoRepEdadCartera.enuResumido,
                            ldtmFechaCausa, lblnCierreMes, True)
                    If String.IsNullOrEmpty(lstrMens) Then
                        Mouse.OverrideCursor = Cursors.Wait
                        lstrMens = "Importando Notas Débito!"
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                        SInforme(lstrMens)
                        lblnImporto = ClsOrionCop.FblnImportoNotasDb()
                        lstrMens = String.Empty
                    End If
                    HbttCancelar.IsEnabled = True
                Else
                    lstrMens = "Proceso fue omitido por el Usuario!"
                End If
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
                MobjObjetoWin.SRefresqueObj()
                MobjObjetoWin.SVayaAlUltimo()
                SMuestreDatos()
                If String.IsNullOrEmpty(lstrMens) Then
                    lstrMens = If(lblnImporto, "La Importación terminó exitosamente!",
                            "La Importación no se llevo a cabo!")
                End If
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub

    Private Sub SInforme(ByRef astrMens As String)
        Dispatcher.Invoke(MdgtLblActualiza,
                System.Windows.Threading.DispatcherPriority.Background,
                New Object() {Label.ContentProperty, astrMens})
    End Sub

    Private Sub SCambieNombresDocCobro()
        If GobjParametros.BlnEFacAutorizado Then
            dgtNroFac.Header = "Número Factura"
            dgtItemFac.Header = "Item Fac."
            dgtNroFacNov.Header = "Número Factura"
        Else
            dgtNroFac.Header = "Nro. Cta. Cobro"
            dgtItemFac.Header = "Item C.C."
            dgtNroFacNov.Header = "Número Cta. Cobro"
        End If
    End Sub

    Private Sub SMuestreNotas(ablnMostrar As Boolean)
        If MobjObjetoWin IsNot Nothing Then
            Dim ldtbNotas = MobjObjetoWin.FdtbNotasUltimoMes
            If ldtbNotas.Rows.Count = 0 Then
                cnvNotas.Visibility = Visibility.Hidden
                grdNotas.Visibility = Visibility.Visible
                MnuMostrarLista.Visibility = Visibility.Collapsed
            Else
                If ablnMostrar Then
                    cnvNotas.Visibility = Visibility.Visible
                    grdNotas.Visibility = Visibility.Hidden
                    MnuMostrarLista.Visibility = Visibility.Hidden
                    dgrNotas.DataContext = ldtbNotas
                    HbttAnular.IsEnabled = False
                    HbttAlPrimero.IsEnabled = False
                    HbttAlAnterior.IsEnabled = False
                    HbttAlSiguiente.IsEnabled = False
                    HbttAlUltimo.IsEnabled = False
                    HbttImprimir.IsEnabled = False
                    HmnuAnular.IsEnabled = False
                    HmnuAlPrimero.IsEnabled = False
                    HmnuAlSiguiente.IsEnabled = False
                    HmnuAlAnterior.IsEnabled = False
                    HmnuAlUltimo.IsEnabled = False
                    HmnuImprimir.IsEnabled = False
                    SLevanteEveNoti("Doble clic o clic contrario abre la nota seleccionada!",
                            String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                Else
                    cnvNotas.Visibility = Visibility.Hidden
                    grdNotas.Visibility = Visibility.Visible
                    MnuMostrarLista.Visibility = Visibility.Visible
                    HbttAnular.IsEnabled = True
                    HbttAlPrimero.IsEnabled = True
                    HbttAlAnterior.IsEnabled = True
                    HbttAlSiguiente.IsEnabled = True
                    HbttAlUltimo.IsEnabled = True
                    HbttImprimir.IsEnabled = True
                    HmnuAnular.IsEnabled = True
                    If Not BlnVentanaAux Then
                        HmnuAlPrimero.IsEnabled = True
                        HmnuAlSiguiente.IsEnabled = True
                        HmnuAlAnterior.IsEnabled = True
                        HmnuAlUltimo.IsEnabled = True
                        HmnuImprimir.IsEnabled = True
                    End If
                    SLevanteEveNoti(String.Empty, String.Empty, 0, EnuSeveridadNot.EnuOk)
                End If
            End If
        End If
    End Sub
#End Region

#Region "EFactura"
    Private Sub SHabiliteMenuEFac()
        If MnuEstadoProvEfac IsNot Nothing Then
            MnuEstadoProvEfac.Visibility = Visibility.Collapsed
            MnuReprocesarEDoc.Visibility = Visibility.Collapsed
        End If
        If GobjParametros.BlnEFacAutorizado AndAlso MobjObjetoWin.BlnExiste Then
            If Not HblnCargandoForma AndAlso MobjObjetoWin.BlnEsDocEle Then
                MnuEDoc.Visibility = Visibility.Visible
                MnuEstadoProvEfac.Visibility = Visibility.Visible
                SVerifiqueEstadoEDoc()
            ElseIf MnuEDoc IsNot Nothing Then
                MnuEDoc.Visibility = Visibility.Collapsed
            End If
        End If
    End Sub

    Private Sub SHabiliteReprocesarEDoc()
        If MnuReprocesarEDoc IsNot Nothing Then
            MnuReprocesarEDoc.Visibility = Visibility.Collapsed
        End If
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If Not HblnCargandoForma AndAlso MobjObjetoWin.BlnEsDocEle AndAlso
                    MobjObjetoWin.BlnAfectaFrasRegEFac Then
                MnuReprocesarEDoc.Visibility = Visibility.Visible
            End If
        End If
    End Sub

    Private Sub SVerifiqueEstadoEDoc()
        Dim lstrMens As String, lblnRegistrar = False
        Dim lenuEstadoEFac As EnuEstadoEDoc = MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro
        If lenuEstadoEFac < EnuEstadoEDoc.EnuEnviada Then
            If lenuEstadoEFac = EnuEstadoEDoc.EnuErrorFtp Then
                lstrMens = "La Nota no fue aceptada por MisFacturas. " &
                        "Ya esta corregida?"
                lblnRegistrar = MsgBox(lstrMens, vbYesNo, "Procesar Nota?") = vbYes
            ElseIf lenuEstadoEFac = EnuEstadoEDoc.EnuInvalida Then
                lstrMens = "La Nota fue rechazada por la DIAN. " &
                        "Ya esta corregida?"
                lblnRegistrar = MsgBox(lstrMens, vbYesNo, "Procesar Nota?") = vbYes
            ElseIf lenuEstadoEFac = EnuEstadoEDoc.EnuNoReg OrElse
                        lenuEstadoEFac = EnuEstadoEDoc.EnuEnProceso OrElse
                        lenuEstadoEFac = EnuEstadoEDoc.EnuRegi Then
                MnuReprocesarEDoc.Visibility = Visibility.Visible
                lstrMens = "Se debe ejecutar el menú 'EFactura -> 
                            Procesar Documentos Electrónicos'"
                MsgBox(lstrMens, vbOKOnly, "Procesar Documentos")
            End If
            If lblnRegistrar Then
                MobjObjetoWin.SHabiliteProcesarEFac()
            End If
        ElseIf lenuEstadoEFac = EnuEstadoEDoc.EnuRechazada Then
            lstrMens = "Esta Nota fue rechazada por el Cliente." & vbCrLf &
                "En consecuencia se debe elaborar una nota crédito " & vbCrLf &
                "por intereses de mora a " & "la factura " &
                MobjObjetoWin.ObjFacturaAfectada.StrIdObjeto & "!"
            MsgBox(lstrMens, vbOKOnly, "Nota rechazada!")
        ElseIf lenuEstadoEFac <> EnuEstadoEDoc.EnuEnviada Then
            lstrMens = "Por favor informe este estado a soporte en Optimusoft!"
            MsgBox(lstrMens, vbOKOnly, "Estado actual!")
        End If
    End Sub
#End Region

#Region "Eventos en la Ventana"
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                Dim lmnuOpcion As MenuItem = lelmElemento
                If lmnuOpcion.Name = "MnuEnviarPorCorreo" Then
                    SEnvieEmailNotaDb()
                ElseIf lmnuOpcion.Name = "MnuEstadoProvEfac" Then
                    Dim lblnV1 = (MobjObjetoWin.ObjVerEFacEnt.ObjValorPro = EnuVerEFac.EnuV1)
                    Mouse.OverrideCursor = Cursors.Wait
                    SRefresqueWin()
                    SMuestreEstadoEFac(MobjObjetoWin.ObjCUDocStr.ObjValorPro,
                            EnuTipoDocOri.EnuNotaDb, MobjObjetoWin.StrNumeroNotaDb, lblnV1,
                            MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro, lstrMens)
                    Mouse.OverrideCursor = Cursors.Arrow
                ElseIf lmnuOpcion.Name = "MnuReprocesarEDoc" Then
                    If MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                            EnuEstadoEDoc.EnuNoEDoc AndAlso
                            MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                            EnuEstadoEDoc.EnuNoReg AndAlso
                            MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                             EnuEstadoEDoc.EnuInvalida Then
                        If MobjObjetoWin.ObjFecha_NotaDbDtm.ObjValorPro > Today.AddDays(-30) Then
                            If MsgBox("Realmente desea reprocesar esta Nota?", vbYesNo,
                                    "Reprocesar Nota?") = vbYes Then
                                MobjObjetoWin.SPrepareParaReprocesarEfac()
                                SProceseEFac(lstrMens)
                            End If
                        Else
                            lstrMens = "Solo es posible reprocesar un documento si su fecha está " &
                                    "dentro de los 30 dias anteriores al día de hoy!"
                        End If
                    End If
                ElseIf lmnuOpcion.Name = "MnuImportarNsDb" Then
                    SImporteNotasDb()
                ElseIf lmnuOpcion.Name = "MnuImprimirNotasMes" Then
                    Dim lstrIdNotasDb = ClsOrionCop.FstrIdUltimasNotasDb()
                    If Not String.IsNullOrEmpty(lstrIdNotasDb) Then
                        SImprimaNotasDb(lstrIdNotasDb, False)
                    Else
                        lstrMens = "No hay Notas de Intereses de Mora en el presente Mes!"
                    End If
                ElseIf lmnuOpcion.Name = "MnuMostrarLista" Then
                    If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                        SMuestreNotas(True)
                    End If
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
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    End If
                Else
                    GobjPanDat.SControleProcesoObj(False, True)
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
    End Sub

    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            ltxtTextBox.SelectAll()
        End If
    End Sub

    Private Sub OnCboCambio(sender As Object, e As RoutedEventArgs)
        If Not MblnPoblandoCombo Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is ComboBox AndAlso Not HblnSeEstaCerrando Then
                If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                    If lelmElemento.Name = "cboPref" Then
                        SInicialiceNota()
                        SMuestreDatos()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Txt_KeyDown(sender As Object, e As KeyEventArgs) Handles txtIdNota.KeyDown
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            If e.Key = Key.Return OrElse e.Key = Key.Tab Then
                SAbraNotaDb()
            End If
        End If
    End Sub

    Private Sub Dgr_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
                dgrItemsNotaDb.MouseRightButtonUp, dgrNovedades.MouseRightButtonUp
        Dim ldgrSeleccionada As DataGrid
        If TypeOf sender Is DataGrid Then
            ldgrSeleccionada = sender
            Dim ldrvFactura As DataRowView = ldgrSeleccionada.SelectedItem
            If Not IsNothing(ldrvFactura) Then
                Dim lstrNroFact As String
                If ldgrSeleccionada.Name = "dgrItemsNotaDb" Then
                    lstrNroFact = ldrvFactura("NroFact")
                Else
                    lstrNroFact = ldrvFactura("NroFac")
                End If
                Dim lstrPrefFac = ClsPanorama.FstrPrefijoDcto(lstrNroFact)
                Dim lentIdFact = ClsPanorama.FentIdDcto(lstrNroFact)
                SAbraFactura(lstrPrefFac, lentIdFact)
            End If
        End If
    End Sub

    Private Sub DgrNotas_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) _
            Handles dgrNotas.MouseDoubleClick
        Dim ldrvFilaActual As DataRowView
        If TypeOf sender Is DataGrid Then
            Dim ldgrActual As DataGrid = sender
            Dim lstrPrefijo As String
            Dim lentIdNota As Integer
            ldrvFilaActual = ldgrActual.SelectedItem
            If Not IsNothing(ldrvFilaActual) AndAlso ldgrActual.SelectedIndex >= 0 Then
                lstrPrefijo = ldrvFilaActual("PrefijoNotaDb")
                lentIdNota = ldrvFilaActual("IdNotaDb")
                cboPref.SelectedItem = lstrPrefijo
                txtIdNota.Text = lentIdNota
                SAbraNotaDb()
            End If
        End If
        SMuestreNotas(False)
    End Sub

    Private Sub DgrNotas_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) _
            Handles dgrNotas.MouseRightButtonUp
        Dim ldrvFilaActual As DataRowView
        If TypeOf sender Is DataGrid Then
            Dim ldgrActual As DataGrid = sender
            Dim lstrPrefijo As String
            Dim lentIdNota As Integer
            ldrvFilaActual = ldgrActual.SelectedItem
            If Not IsNothing(ldrvFilaActual) AndAlso ldgrActual.SelectedIndex >= 0 Then
                lstrPrefijo = ldrvFilaActual("PrefijoNotaDb")
                lentIdNota = ldrvFilaActual("IdNotaDb")
                cboPref.SelectedItem = lstrPrefijo
                txtIdNota.Text = lentIdNota
                SAbraNotaDb()
            End If
        End If
        SMuestreNotas(False)
    End Sub

    Private Sub OnRatonUp(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TabItem Then
            SEstablezcaDataContext()
        End If
    End Sub
#End Region
End Class
