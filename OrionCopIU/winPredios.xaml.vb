Class WinPredios
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuArea = 0
        enuFactorPond
        enuComentario
        enuEMail
        enuIdFichaCat
        enuIdMatriInmob
        enuIdPredAgru
        enuIdPredio
        enuNomPredio
        enuIdRegMercan
        enuRefPago
        enuIdSectorPredio
        enuIdTerAdmini
        enuIdTerArrenda
        enuIdTerRepLeg
        enuIdTipoDestFac
        enuNombreComer
        enuValorSerId
        enuEstadoDeuda
        EnuFactPorSer
        EnuPropietarios
    End Enum
    Private Enum EnuTipoClienteTer As Byte
        None = 0
        enuArrendatario
        enuRepLegalArre
        enuAdministrador
    End Enum
#End Region
    ' Variables
    Private WithEvents MwinVentana As ClsFormInterface = Nothing
    Private ReadOnly MenuTamanoIcono As EnuTamanoIconos
    Private MenuTipoClienteCreando As EnuTipoClienteTer = EnuTipoClienteTer.enuAdministrador
    Private MobjObjetoWin As ClsPredio = Nothing
    Private MobjPropSel As ClsPropietario
    Private MdtbProp As DataTable
    '
    Private MnuImportarPredio As MenuItemPan = Nothing
    Private MnuImportarPropietarios As MenuItemPan = Nothing
    Private MnuCalcularCP As MenuItemPan = Nothing
    Private MnuVerificarIntegridadPredios As MenuItemPan = Nothing
    Private MnuAbrirPredioConC As MenuItem = Nothing
    Private MentTabSeleccionado As Integer = 0
    Private MblnEsPredioAgru As Boolean = False
    Private MstrIdPropSel As String = String.Empty
    ' Manejo Ubicacion
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomPre
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuPredio
        MenuTamanoIcono = GenuTamanoIcono
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicionesCtlsRestringidos()
        Dim lcolControlesLlave As New Collection From {
            txtIdPredio
        }
        SCargueForma(EnuElementosAdicionalesDef.enuBuscar,
                21, lcolControlesLlave, txtNombrePredio, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SMuestrePropietarios()
        lblNotaProp.Visibility = Visibility.Collapsed
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
        If IsNothing(ObjObjetoWin) Then
            ObjObjetoWin = ClsOrionCop.FobjNuevoPredio(EnuModoInstanciaObjDef.EnuNavegable)
            If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                ObjObjetoWin.SVayaAlPrimero()
            End If
        End If
        MobjObjetoWin = ObjObjetoWin
        MdtbProp = MobjObjetoWin.FdtbPropietarios
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        chkFacturarSer.IsChecked = False
        chkNoConsItemsFac.IsChecked = True
        StcValidaControl(EnuValidEntrada.enuArea) = lblArea
        StcValidaControl(EnuValidEntrada.enuFactorPond) = lblFactorPond
        StcValidaControl(EnuValidEntrada.enuComentario) = lblComentarios
        StcValidaControl(EnuValidEntrada.enuEMail) = lblEMail
        StcValidaControl(EnuValidEntrada.enuIdFichaCat) = lblFichaCatastral
        StcValidaControl(EnuValidEntrada.enuIdMatriInmob) = lblMatInmobiliaria
        StcValidaControl(EnuValidEntrada.enuIdPredAgru) = lblPredioAgrupador
        StcValidaControl(EnuValidEntrada.enuIdPredio) = lblIdPredio
        StcValidaControl(EnuValidEntrada.enuNomPredio) = lblNombrePredio
        StcValidaControl(EnuValidEntrada.enuIdRegMercan) = lblMatriculaMercantil
        StcValidaControl(EnuValidEntrada.enuRefPago) = lblRefPago
        StcValidaControl(EnuValidEntrada.enuIdSectorPredio) = lblSector
        StcValidaControl(EnuValidEntrada.enuIdTerAdmini) = lblIdAdministrador
        StcValidaControl(EnuValidEntrada.enuIdTerArrenda) = lblIdArrendatario
        StcValidaControl(EnuValidEntrada.enuIdTerRepLeg) = lblIdRepLegalArrendatario
        StcValidaControl(EnuValidEntrada.enuIdTipoDestFac) = lblDestFact
        StcValidaControl(EnuValidEntrada.enuNombreComer) = lblNombreComercial
        StcValidaControl(EnuValidEntrada.enuValorSerId) = lblValorServicioID
        StcValidaControl(EnuValidEntrada.enuEstadoDeuda) = lblEstadoActualDeuda
        StcValidaControl(EnuValidEntrada.EnuFactPorSer) = chkFacturarSer
        StcValidaControl(EnuValidEntrada.EnuPropietarios) = lblPropietrios
        '
        SVisibiliceCtrls()
        SPuebleComboBoxes()
        tbiPropietarios.IsSelected = True
        HbttAceptar.TabIndex = 80
        HbttCancelar.TabIndex = 81
    End Sub
    Protected Overrides Sub SMuestreDatos()
        HblnMostrandoDatos = True
        MblnEsPredioAgru = MobjObjetoWin.ObjIdPredioStr.ToString().ToUpper =
                MobjObjetoWin.ObjIdPredioAgrupadorStr.ToString().ToUpper
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                SLevanteEveNoti("No hay Predios para ser mostrados!", "", 0,
                        EnuSeveridadNot.EnuInformacion)
                txtIdPredio.IsEnabled = False
            End If
        End If
        SHabiliteBttsProp()
        With MobjObjetoWin
            txtIdPredio.Text = .ObjIdPredioStr.ObjValorPro
            txtNombrePredio.Text = .ObjNombrePredioStr.ObjValorPro
            txtIdPredioAgr.Text = .ObjIdPredioAgrupadorStr.ObjValorPro
            txtFichaCatastral.Text = .ObjIdFichaCatastralStr.ToString
            txtMatInmobiliaria.Text = .ObjIdMatriculaInmobiliariaStr.ToString
            txtAliasCont.Text = .ObjAliasContStr.ToString
            txtRefPago.Text = .ObjReferenciaPagoStr.ToString
            txtEMailAdi.Text = .ObjEmailAdiStr.ObjValorPro
        End With
        SMuestreInfTab()
        SValide()
        Title = My.Resources.FichaPredio
        If Not String.IsNullOrEmpty(txtIdPredio.Text) Then
            Title &= txtIdPredio.Text
        End If
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If txtIdPredio.Focus Then
                txtIdPredio.SelectAll()
            End If
        End If
        HblnMostrandoDatos = False
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.CenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(EnuValidEntrada.enuArea) = .ObjAreaPredioDec.BlnEsValido
                StcValidValido(EnuValidEntrada.enuFactorPond) = .ObjFactorPonderaCPDbl.BlnEsValido
                StcValidValido(EnuValidEntrada.enuComentario) = .ObjComentarioStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuEMail) = .ObjEmailAdiStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuIdFichaCat) = .ObjIdFichaCatastralStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuIdMatriInmob) = .ObjIdMatriculaInmobiliariaStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuIdPredAgru) = .ObjIdPredioAgrupadorStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuIdPredio) = .ObjIdPredioStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuNomPredio) = .ObjNombrePredioStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuIdRegMercan) = .ObjIdRegistroMercantilStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuRefPago) = .ObjReferenciaPagoStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuIdSectorPredio) = .ObjIdSector_PredioShr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuIdTerAdmini) = .ObjIdClienteAdministradorDbl.BlnEsValido
                StcValidValido(EnuValidEntrada.enuIdTerArrenda) = .ObjIdClienteArrendatarioDbl.BlnEsValido
                StcValidValido(EnuValidEntrada.enuIdTerRepLeg) = .ObjIdClienteRepLegArrendatariodbl.BlnEsValido
                StcValidValido(EnuValidEntrada.enuIdTipoDestFac) = .ObjIdTipoDestinatarioFacturaByt.BlnEsValido
                StcValidValido(EnuValidEntrada.enuNombreComer) = .ObjNombreComercialStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuValorSerId) = .ObjValorServicioIdDec.BlnEsValido
                StcValidValido(EnuValidEntrada.enuEstadoDeuda) = .ObjIdEstadoDeuda_PredioByt.BlnEsValido
                StcValidValido(EnuValidEntrada.EnuFactPorSer) = .ObjFacturarPorServicio_PreBln.BlnEsValido
                StcValidValido(EnuValidEntrada.EnuPropietarios) = FblnPropietariosOk(True) AndAlso
                        FblnPrediosIntegros()
            End With
        End If
        SHabiliteBotonesTlb()
        If FblnEstanTodosBien() Then
            If MobjObjetoWin.ColPropietarios.Count = 0 Then
                If Not ClsCentroUtilOriCop.FblnHayPropietarios Then
                    SLevanteEveNoti("Debe crear o importar los Propietarios!", "", 0,
                                    EnuSeveridadNot.EnuInformacion)
                Else
                    SLevanteEveNoti("Debe crear o importar los Propietarios del Predio!", "", 0,
                                    EnuSeveridadNot.EnuCamInsatis)
                End If
                HbttAceptar.IsEnabled = False
                lblPropietrios.Style = FindResource("RecLblInvalido")
            Else
                HbttAceptar.IsEnabled = True
                lblPropietrios.Style = FindResource("RecCtlValido")
            End If
        End If
        If MobjObjetoWin.BlnExiste AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.CenuConsultando Then
            If GobjParametros.EnuEstadoInstalacion = EnuEstadoInstalacion.Todos Then
                SHabiliteMenuItem(FblnHabilitarMenuPan(3), MnuCalcularCP)
            End If
        Else
            SHabiliteMenuItem(False, MnuCalcularCP)
        End If
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            If Not String.IsNullOrEmpty(txtAliasCont.Text) Then
                .ObjAliasContStr.ObjValorPro = txtAliasCont.Text
            End If
            .ObjFacturarPorServicio_PreBln.ObjValorPro = chkFacturarSer.IsChecked
            .ObjNoConsolidarItemsFacBln.ObjValorPro = chkNoConsItemsFac.IsChecked
            .ObjIdFichaCatastralStr.ObjValorPro = txtFichaCatastral.Text
            .ObjIdMatriculaInmobiliariaStr.ObjValorPro = txtMatInmobiliaria.Text
            .ObjIdPredioAgrupadorStr.ObjValorPro = txtIdPredioAgr.Text
            .ObjAreaPredioDec.ObjValorPro = txtArea.Text
            .ObjFactorPonderaCPDbl.ObjValorPro = FdblTasa(txtFactorPond.Text)
            .ObjIdSector_PredioShr.ObjValorPro = cboSectores.SelectedIndex
            .ObjComentarioStr.ObjValorPro = txtComentarios.Text
            .ObjIdClienteArrendatarioDbl.ObjValorPro = txtIdArrendatario.Text
            .ObjIdClienteRepLegArrendatariodbl.ObjValorPro = txtIdRepLegalArrendatario.Text
            .ObjIdClienteAdministradorDbl.ObjValorPro = txtIdAdministrador.Text
            .ObjNombreComercialStr.ObjValorPro = txtNombreComercial.Text
            .ObjIdRegistroMercantilStr.ObjValorPro = txtMatriculaMercantil.Text
            .ObjReferenciaPagoStr.ObjValorPro = txtRefPago.Text
            .ObjIdTipoDestinatarioFacturaByt.ObjValorPro = cboDestFact.SelectedIndex
            .ObjValorServicioIdDec.ObjValorPro = txtValorServicioId.Text
            .ObjEmailAdiStr.ObjValorPro = txtEMailAdi.Text
        End With
    End Sub
    Protected Overrides Sub SConfigureMenuesPropios()
        ' Adicionar menú Importar
        MnuImportarPredio = FmnuiMenuItemPan("MnuImportarPredio", "_Importar Predios",
                1, "")
        MnuImportarPropietarios = FmnuiMenuItemPan("MnuImportarPropietarios",
                "_Importar Propietarios", 1, "")
        MnuVerificarIntegridadPredios = FmnuiMenuItemPan("MnuIntegridadPredios",
                    "_Verificar Integridad Predios", 2, "")
        MnuCalcularCP = FmnuiMenuItemPan("MnuCalcularCP",
                "_Calcular Coeficientes de Propiedad", 3, "")
        Dim lsepSeparador As New Separator
        HmnuAcciones.Items.Insert(7, MnuImportarPredio)
        HmnuAcciones.Items.Insert(8, MnuImportarPropietarios)
        HmnuAcciones.Items.Insert(9, MnuCalcularCP)
        HmnuAcciones.Items.Insert(10, MnuVerificarIntegridadPredios)
        HmnuAcciones.Items.Insert(11, lsepSeparador)
        ' Adicionar Menu Contextual
        SAsigneMenuContextual()
        ' Adicionar Menues Reportes
        Dim lmnuReportes As MenuItem = FmnuiMenuItem("MnuReportes", "R_eportes", "RecMnuItemPriInf")
        MenuVen.Items.Insert(1, lmnuReportes)
        Dim lmnuItem = FmnuiMenuItem("MnuRepPrediosSector", "Predios por _Sector", "RecMnuItemSec")
        lmnuItem.ToolTip = "Genera un Reporte de los Predios agrupador por Sector en orden alfabético."
        lmnuReportes.Items.Add(lmnuItem)
        lmnuItem = FmnuiMenuItem("MnuRepPrediosCliente", "Predios por _Propietario", "RecMnuItemSec")
        lmnuItem.ToolTip = "Genera un Reporte de los Propietarios con sus respectivos Predios."
        lmnuReportes.Items.Add(lmnuItem)
        lmnuItem = FmnuiMenuItem("MnuRepPropiPorCP", "Propi_etarios por Coeficiente de Propiedad",
                "RecMnuItemSec")
        lmnuItem.ToolTip = "Genera un Reporte de los propietarios y el total de sus coeficientes " &
                "de propiedad"
        lmnuReportes.Items.Add(lmnuItem)
        lmnuItem = FmnuiMenuItem("MnuRepCuotasAdminProp", "Cuotas _Administración por Propietario",
                "RecMnuItemSec")
        lmnuItem.ToolTip = "Genera un Reporte de los Propietarios con la Cuota de Administración de cada Predio."
        lmnuReportes.Items.Add(lmnuItem)
        lmnuItem = FmnuiMenuItem("MnuRepDirTf", "Directorio _Telefónico de Propietarios", "RecMnuItemSec")
        lmnuItem.ToolTip = "Genera un Directorio Telefónico de los Propietarios."
        lmnuReportes.Items.Add(lmnuItem)
        Dim lsep As New Separator
        lmnuReportes.Items.Add(lsep)
        lmnuItem = FmnuiMenuItem("MnuRepCarteraPredAgr", "Ca_rtera por Predio Agrupador",
                "RecMnuItemSec")
        lmnuItem.ToolTip = "Muestra el valor de las Cuentas por Cobrar por Predio Agrupador a hoy."
        lmnuReportes.Items.Add(lmnuItem)
        lmnuItem = FmnuiMenuItem("MnuRepCarteraPredios", "Car_tera por Predio",
                "RecMnuItemSec")
        lmnuItem.ToolTip = "Muestra el valor de las Cuentas por Cobrar por Predio a hoy."
        Dim lsep1 As New Separator
        lmnuItem = FmnuiMenuItem("MnuRepPazYSalvo", "Pa_z y Salvo del Predio",
                "RecMnuItemSec")
        lmnuReportes.Items.Add(lsep1)
        lmnuReportes.Items.Add(lmnuItem)
    End Sub
    Private Sub SAsigneMenuContextual()
        Dim lmnuMenuContextual As ContextMenu = FindResource("RecMnuPrediosAgruMC")
        lsbPrediosAgrupados.ContextMenu = lmnuMenuContextual
        For Each lobjObjetoMenu As Object In lmnuMenuContextual.Items
            If TypeOf lobjObjetoMenu Is MenuItem Then
                Dim lmnuItemMenuCont As MenuItem = lobjObjetoMenu
                If lmnuItemMenuCont.Name = "MnuAbrirPredioC" Then
                    MnuAbrirPredioConC = lmnuItemMenuCont
                End If
            End If
        Next
    End Sub
#End Region
#Region "Invalida otros metodos de la clase base"
    Protected Overrides Sub SHabiliteMenues()
        MyBase.SHabiliteMenues()
        Dim lblnNoPermitido = Not (GobjParametros.EnuEstadoInstalacion AndAlso
                EnuEstadoInstalacion.Predios)
        If lblnNoPermitido Then
            SHabiliteMenuItemPan(False, MnuVerificarIntegridadPredios)
            SHabiliteMenuItemPan(False, MnuCalcularCP)
            SHabiliteMenuItemPan(False, MnuImportarPropietarios)
        Else
            If MobjObjetoWin.FblnEstaVacioOrigenDatos Then
                SHabiliteMenuItemPan(False, MnuImportarPropietarios)
            Else
                SHabiliteMenuItemPan(True, MnuImportarPropietarios)
            End If
        End If
    End Sub
    Protected Overrides Sub SCree()
        tbiPredio.IsSelected = True
        txtIdPredio.IsEnabled = True
        txtIdPredio.Focus()
        MyBase.SCree()
        bttEncontrarAdmin.Visibility = Visibility.Visible
        bttEncontrarArren.Visibility = Visibility.Visible
        bttEncontrarRepL.Visibility = Visibility.Visible
    End Sub
    Protected Overrides Sub SModifique()
        SMuestreTodosTabs()
        MyBase.SModifique()
        SHabiliteBttsProp()
        bttEncontrarAdmin.Visibility = Visibility.Visible
        bttEncontrarArren.Visibility = Visibility.Visible
        bttEncontrarRepL.Visibility = Visibility.Visible
    End Sub
    Protected Overrides Sub SCancele()
        MyBase.SCancele()
        SLevanteEveOk()
        SHabiliteBttsProp()
    End Sub
    Protected Overrides Sub SGuarde()
        Dim lstrMens As String, lblnPredIteg As Boolean = True
        Dim lblnCambioArea = MobjObjetoWin.ObjAreaPredioDec.BlnCambio
        If MobjObjetoWin.FblnPropietariosCambiaron Then
            lblnPredIteg = FblnPrediosIntegros()
        End If
        If Not lblnPredIteg Then
            HbttAceptar.IsEnabled = False
            lblPropietrios.Style = FindResource("RecLblInvalido")
            lstrMens = "Hay problemas de integridad en los predios." &
                    " No es posible guardar los cambios!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        Else
            MyBase.SGuarde()
            If lblnCambioArea Then
                lstrMens = "Calculando Coeficientes de Propiedad!"
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                SCalculeCPCopr()
            End If
        End If
        SHabiliteBttsProp()
    End Sub
    Protected Overrides Sub SEstablezcaWinConsultando()
        MyBase.SEstablezcaWinConsultando()
        MdtbProp = Nothing
        bttEncontrarAdmin.Visibility = Visibility.Collapsed
        bttEncontrarArren.Visibility = Visibility.Collapsed
        bttEncontrarRepL.Visibility = Visibility.Collapsed
    End Sub
#End Region
#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        MyBase.SBuscar()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.EnuNavegable Then
                    txtIdPredio.Text = StrResultadoBusqueda
                    SAbraPredio()
                End If
            End If
        End If
    End Sub
    ''' <summary>
    ''' Invalida la funcion "fblnDefinioBusqueda" de la clase base.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        SDefineBusquedaNombreCliente()
        SDefineBusquedaPredioAgr()
        Return True
    End Function
    Private Function FblnDefinioBusquedaCliente() As Boolean
        Dim lstrCamposMostrar As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                         ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdClienteDbl.SstrNombreCampoBd
        Dim lstrTabla As String = ClsCliente.SstrNombreTabla
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " & lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
        Return True
    End Function
    Private Sub SDefineBusquedaNombreCliente()
        Dim lstrTablaPri = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec = ClsPropietario.SstrNombreTabla
        Dim lstrCampSelePri = {ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampSeleSec = {ClsIdCliente_PropDbl.SstrNombreCampoBd,
                ClsNombreCompleto_PropStr.SstrNombreCampoBd}
        Dim LstrCampRelPri = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampRelSec = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdPredio_PropStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompleto_PropStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdPredioStr.SstrNombreCampoBd
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Nombre Propietario", lstrTablaPri, lstrTablaSec,
                lstrCampSelePri, lstrCampSeleSec, LstrCampRelPri, lstrCampRelSec,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, True)
    End Sub
    Private Sub SDefineBusquedaPredioAgr()
        Dim lstrTablaPri = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec = ClsPropietario.SstrNombreTabla
        Dim lstrCampSelePri As String() = {"DISTINCT " & ClsIdPredioAgrupadorStr.SstrNombreCampoBd}
        Dim lstrCampSeleSec = {ClsIdCliente_PropDbl.SstrNombreCampoBd,
                ClsNombreCompleto_PropStr.SstrNombreCampoBd}
        Dim LstrCampRelPri = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampRelSec = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdPredio_PropStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador", lstrTablaPri, lstrTablaSec,
                lstrCampSelePri, lstrCampSeleSec, LstrCampRelPri, lstrCampRelSec,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, True)
    End Sub
    Private Sub SBusqueCliente(aenuTipoTer As EnuTipoClienteTer)
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            StrResultadoBusqueda = String.Empty
            Cursor = Cursors.Wait
            If IsNothing(HwinBusqueda) Then
                HwinBusqueda = New WinBusqueda With {
                    .WinPadre = Me
                }
            End If
            If FblnDefinioBusquedaCliente() Then
                HwinBusqueda.ShowDialog()
            End If
            HwinBusqueda = Nothing
            Cursor = Cursors.Arrow
            If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                Select Case aenuTipoTer
                    Case EnuTipoClienteTer.enuArrendatario
                        txtIdArrendatario.Text = StrResultadoBusqueda
                        SRegistreTextBox(txtIdArrendatario)
                    Case EnuTipoClienteTer.enuRepLegalArre
                        txtIdRepLegalArrendatario.Text = StrResultadoBusqueda
                        SRegistreTextBox(txtIdRepLegalArrendatario)
                    Case EnuTipoClienteTer.enuAdministrador
                        txtIdAdministrador.Text = StrResultadoBusqueda
                        SRegistreTextBox(txtIdAdministrador)
                End Select
                SMuestreDatos()
            End If
        End If
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SAdicioneCtlsRestingidos()
        SAdicioneControlRestringido(dgrFacturas)
        SAdicioneControlRestringido(dgrPropietarios)
    End Sub
    Private Sub SAbraCliente(ablnPropietario As Boolean)
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
            Dim ldblIdCliente As Double
            If MobjObjetoWin.ObjIdTipoDestinatarioFacturaByt.ObjValorPro =
                    EnuDestinatarioFacturaDef.EnuPropietario OrElse ablnPropietario Then
                If MobjPropSel IsNot Nothing Then
                    ldblIdCliente = MobjPropSel.ObjIdCliente_PropDbl.ObjValorPro
                End If
            Else
                ldblIdCliente = MobjObjetoWin.ObjIdClienteArrendatarioDbl.ObjValorPro
            End If
            lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente})
            If lobjCliente.BlnExiste Then
                Dim lwinPropietario As New WinClientes With {
                    .WinPadre = Me,
                    .ObjObjetoWin = lobjCliente,
                    .BlnVentanaAux = True
                }
                lwinPropietario.ShowDialog()
            End If
        End If
    End Sub
    Private Sub SMuestreInfTab()
        Select Case True
            Case tbiPredio.IsSelected
                SMuestrePredio()
            Case tbiPropietarios.IsSelected
                SMuestrePropietarios()
            Case tbiArrendatario.IsSelected
                SMuestreArrendatario()
            Case tbiFacturas.IsSelected
                SMuestreFacturas()
        End Select
    End Sub
    Private Sub SMuestreTodosTabs()
        SMuestrePredio()
        SMuestrePropietarios()
        SMuestreArrendatario()
    End Sub
    Private Sub SMuestrePredio()
        If Not IsNothing(MobjObjetoWin) Then
            With MobjObjetoWin
                txtIdPredioAgr.Text = .ObjIdPredioAgrupadorStr.ObjValorPro
                txtArea.Text = .ObjAreaPredioDec.ToString
                txtFactorPond.Text = Format(.ObjFactorPonderaCPDbl.ObjValorPro, "p")
                chkFacturarSer.IsChecked = .ObjFacturarPorServicio_PreBln.ObjValorPro
                chkNoConsItemsFac.IsChecked = .ObjNoConsolidarItemsFacBln.ObjValorPro
                txtCoefPro.Content = Math.Round(.ObjCoeficientePropiedadDec.ObjValorPro * 100, 6
                        ).ToString & "%"
                txtValorServicioId.Text = Format(.ObjValorServicioIdDec.ObjValorPro, "c")
                cboSectores.SelectedIndex = .ObjIdSector_PredioShr.ObjValorPro
                cboEstadoActualDeuda.SelectedIndex = .ObjIdEstadoDeuda_PredioByt.ObjValorPro
                If MobjObjetoWin.BlnExiste Then
                    txtEstadoSugerido.Content = ClsOrionCop.FstrNombreDatoConstanteOri(
                            EnuGrupoConstantesOriDef.EnuEstadoDeuda, .FenuEstadoSugeridoDeuda)
                Else
                    txtEstadoSugerido.Content = String.Empty
                End If
                txtComentarios.Text = .ObjComentarioStr.ObjValorPro
                cboDestFact.SelectedIndex = .ObjIdTipoDestinatarioFacturaByt.ObjValorPro
                SPueblePrediosAgrupados()
            End With
        End If
    End Sub
    Private Sub SMuestrePropietarios()
        SActualiceTblPropietarios()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            MdtbProp = MobjObjetoWin.FdtbPropietarios
        End If
        grdPropietarios.DataContext = MdtbProp
        dgrPropietarios.SelectedIndex = 0
        SEstablezcaPropSel()
    End Sub
    Private Sub SMuestreArrendatario()
        If Not IsNothing(MobjObjetoWin) Then
            With MobjObjetoWin
                txtIdArrendatario.Text = .ObjIdClienteArrendatarioDbl.ToString
                If MobjObjetoWin.ObjArrendatario IsNot Nothing Then
                    txtNombreArrendatario.Content = .ObjArrendatario.ObjNombreCompletoStr.ObjValorPro
                Else
                    txtNombreArrendatario.Content = String.Empty
                End If
                txtIdRepLegalArrendatario.Text = .ObjIdClienteRepLegArrendatariodbl.ToString
                txtNombreRepLegalArrendatario.Content = .ObjIdClienteRepLegArrendatariodbl.StrNombreRepLegalArrendatario
                txtIdAdministrador.Text = .ObjIdClienteAdministradorDbl.ToString
                txtNomAdministrador.Content = .ObjIdClienteAdministradorDbl.StrNombreAdministrador
                txtNombreComercial.Text = .ObjNombreComercialStr.ObjValorPro
                txtMatriculaMercantil.Text = .ObjIdRegistroMercantilStr.ObjValorPro
            End With
        End If
    End Sub
    Private Sub SMuestreFacturas()
        If Not IsNothing(MobjObjetoWin) Then
            Mouse.OverrideCursor = Cursors.Wait
            Dim ldtbFacturas = Nothing
            Select Case True
                Case rdbVencidas.IsChecked
                    ldtbFacturas = MobjObjetoWin.FdtbFacturasPredio(EnuEstadoFacturaDef.EnuVencida)
                Case rdbTodas.IsChecked
                    ldtbFacturas = MobjObjetoWin.FdtbFacturasPredio(EnuEstadoFacturaDef.EnuNormal)
                Case rdbCanceladas.IsChecked
                    ldtbFacturas = MobjObjetoWin.FdtbFacturasPredio(EnuEstadoFacturaDef.EnuCancelada)
                Case rdbAnuladas.IsChecked
                    ldtbFacturas = MobjObjetoWin.FdtbFacturasPredio(EnuEstadoFacturaDef.EnuAnulada)
            End Select
            dgrFacturas.DataContext = ldtbFacturas
            Mouse.OverrideCursor = Cursors.Arrow
        End If
    End Sub
    Private Sub SVisibiliceCtrls()
        If GobjParametros.ObjServicioIdActivoBln.ObjValorPro Then
            lblValorServicioID.Visibility = Visibility.Visible
            txtValorServicioId.Visibility = Visibility.Visible
        Else
            lblValorServicioID.Visibility = Visibility.Hidden
            txtValorServicioId.Visibility = Visibility.Hidden
        End If
    End Sub
    Private Sub SAbraCuenta()
        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
        lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, MobjObjetoWin.ObjIdPredioStr.ObjValorPro})
        Dim lwinCuenta As New WinCuentaPredios With {
            .WinPadre = Me,
            .ObjObjetoWin = lobjPredio,
            .BlnVentanaAux = True
        }
        lwinCuenta.ShowDialog()
    End Sub
    Private Sub SAbraPredio()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                If txtIdPredio.Text <> MobjObjetoWin.ObjIdPredioStr.ToString() Then
                    Dim lobjVlrLlave = {GshrIdCarpeta, GshrIdCentroUtil,
                            txtIdPredio.Text}
                    MobjObjetoWin.SAbra(lobjVlrLlave)
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
                If Not lblnNoHayError Then
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
    End Sub
    Private Sub SRegistreTextBox(atxtTextBox As TextBox)
        With MobjObjetoWin
            Select Case atxtTextBox.Name
                Case "txtIdPredio"
                    .ObjIdPredioStr.ObjValorPro = txtIdPredio.Text
                Case "txtNombrePredio"
                    .ObjNombrePredioStr.ObjValorPro = txtNombrePredio.Text
                Case "txtFichaCatastral"
                    .ObjIdFichaCatastralStr.ObjValorPro = txtFichaCatastral.Text
                Case "txtMatInmobiliaria"
                    .ObjIdMatriculaInmobiliariaStr.ObjValorPro = txtMatInmobiliaria.Text
                Case "txtIdPredioAgr"
                    .ObjIdPredioAgrupadorStr.ObjValorPro = txtIdPredioAgr.Text
                    If EnuOperacionEnWin = EnuOperacionEnVentana.CenuModificando AndAlso
                            .ObjIdPredioStr.ObjValorPro = .ObjIdPredioAgrupadorStr.ObjValorPro Then
                        MblnEsPredioAgru = True
                    Else
                        MblnEsPredioAgru = False
                    End If
                    SHabiliteBttsProp()
                Case "txtArea"
                    .ObjAreaPredioDec.ObjValorPro = txtArea.Text
                Case "txtFactorPond"
                    .ObjFactorPonderaCPDbl.ObjValorPro = FdblTasa(txtFactorPond.Text)
                Case "txtAliasCont"
                    .ObjAliasContStr.ObjValorPro = txtAliasCont.Text
                Case "txtRefPago"
                    .ObjReferenciaPagoStr.ObjValorPro = txtRefPago.Text
                Case "txtComentarios"
                    .ObjComentarioStr.ObjValorPro = txtComentarios.Text
                Case "txtIdRepLegalArrendatario"
                    .ObjIdClienteRepLegArrendatariodbl.ObjValorPro = txtIdRepLegalArrendatario.Text
                    If Not .ObjIdClienteRepLegArrendatariodbl.BlnExisteCliente Then
                        MenuTipoClienteCreando = EnuTipoClienteTer.enuRepLegalArre
                        SCrearCliente(txtIdRepLegalArrendatario.Text)
                    End If
                Case "txtIdArrendatario"
                    .ObjIdClienteArrendatarioDbl.ObjValorPro = txtIdArrendatario.Text
                Case "txtIdAdministrador"
                    .ObjIdClienteAdministradorDbl.ObjValorPro = txtIdAdministrador.Text
                    If Not .ObjIdClienteAdministradorDbl.BlnExisteCliente Then
                        MenuTipoClienteCreando = EnuTipoClienteTer.enuAdministrador
                        SCrearCliente(txtIdAdministrador.Text)
                    End If
                Case "txtNombreComercial"
                    .ObjNombreComercialStr.ObjValorPro = txtNombreComercial.Text
                Case "txtMatriculaMercantil"
                    .ObjIdRegistroMercantilStr.ObjValorPro = txtMatriculaMercantil.Text
                Case "txtValorServicioId"
                    .ObjValorServicioIdDec.ObjValorPro = txtValorServicioId.Text
                Case "txtEMailAdi"
                    .ObjEmailAdiStr.ObjValorPro = txtEMailAdi.Text
            End Select
        End With
    End Sub
    Private Sub SPueblePrediosAgrupados()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            lsbPrediosAgrupados.Items.Clear()
            If MobjObjetoWin.BlnExiste Then
                Dim lstrPrediosAgrupados() As String = MobjObjetoWin.StrPrediosAgrupados
                If Not IsNothing(lstrPrediosAgrupados) AndAlso lstrPrediosAgrupados.Length > 0 Then
                    For Each lstrPredio As String In lstrPrediosAgrupados
                        lsbPrediosAgrupados.Items.Add(lstrPredio)
                    Next
                End If
            End If
        End If
    End Sub
    Private Sub SPuebleComboBoxes()
        ' Poblar combo Destinatario Factura
        cboDestFact.Items.Clear()
        Dim ldrwDataRow = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuDestinatarioFactura)
        SPuebleComboBox(ldrwDataRow, cboDestFact)
        ldrwDataRow = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuEstadoDeuda)
        SPuebleComboBox(ldrwDataRow, cboEstadoActualDeuda)
        ' Poblar combo Sectores
        cboSectores.Items.Clear()
        Dim lcolSectores As Collection = GobjParametros.ColSectores
        cboSectores.Items.Add("<Ninguno>")
        For Each lobjSector As ClsSector In lcolSectores
            cboSectores.Items.Add(lobjSector.ObjNombreSectorStr.ObjValorPro)
        Next
    End Sub
    Private Sub SAbraImportar(ablnPredio As Boolean)
        Dim lobjAImportar As ClsCBObjetoPan
        lobjAImportar = If(ablnPredio, MobjObjetoWin,
                New ClsPropietario(EnuModoInstanciaObjDef.EnuUnico))
        Dim lwinImportar As New WinImportar(lobjAImportar) With {
            .WinPadre = Me
        }
        lwinImportar.ShowDialog()
        If ablnPredio Then
            ClsOrionCop.SActualiceTotalAreaCopr()
        End If
        SRefrescarClic()
    End Sub
    Private Sub SCrearCliente(astrIdCliente As String)
        If IsNumeric(astrIdCliente) Then
            If Not String.IsNullOrEmpty(astrIdCliente) AndAlso astrIdCliente <> "0" Then
                If MsgBox("El Cliente ingresado no existe. Desea crearlo ahora?", MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                        "Crear Cliente ?") = MsgBoxResult.Yes Then
                    Dim lobjCliente = ClsOrionCop.FobjCliente(EnuModoInstanciaObjDef.EnuNavegable)
                    lobjCliente.SCreeObj({GshrIdCarpeta, GshrIdCentroUtil, astrIdCliente})
                    lobjCliente.ObjIdClienteDbl.ObjValorPro = astrIdCliente
                    MwinVentana = New WinClientes With {
                        .ObjObjetoWin = lobjCliente,
                        .EnuOperacionEnWin = EnuOperacionEnWin.CenuCreando,
                        .WinPadre = Me
                    }
                    MwinVentana.ShowDialog()
                    Select Case MenuTipoClienteCreando
                        Case EnuTipoClienteTer.enuAdministrador
                            MobjObjetoWin.ObjIdClienteAdministradorDbl.ObjValorPro = txtIdAdministrador.Text
                        Case EnuTipoClienteTer.enuArrendatario
                            MobjObjetoWin.ObjIdClienteArrendatarioDbl.ObjValorPro = txtIdArrendatario.Text
                        Case EnuTipoClienteTer.enuRepLegalArre
                            MobjObjetoWin.ObjIdClienteRepLegArrendatariodbl.ObjValorPro = txtIdRepLegalArrendatario.Text
                    End Select
                    SMuestreDatos()
                End If
            End If
        Else
            SLevanteEveNoti("La identificación del Cliente debe ser númerica!", "", 0,
                    EnuSeveridadNot.EnuDatoInvalido)
        End If
    End Sub
    Private Sub SGenereRep(astrNomMenu As String)
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            Mouse.OverrideCursor = Cursors.Wait
            Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO)
            Select Case astrNomMenu
                Case "MnuRepPrediosSector"
                    lobjRep.EnuReporte = EnuReporteDef.enuPrediosSector
                Case "MnuRepPrediosCliente"
                    lobjRep.EnuReporte = EnuReporteDef.enuPrediosPropietario
                Case "MnuRepPropiPorCP"
                    lobjRep.EnuReporte = EnuReporteDef.enuPropietariosXCP
                Case "MnuRepCuotasAdminProp"
                    Dim lshrIdAno As Short = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
                    Dim lobjParametrosRep As New ClsParametrosReportesDocs("", lshrIdAno, 0)
                    lobjRep.ObjParRepDocs = lobjParametrosRep
                    lobjRep.EnuReporte = EnuReporteDef.enuCuotasAdminPropi
                Case "MnuRepCarteraPredAgr"
                    lobjRep.EnuReporte = EnuReporteDef.enuCarteraPorPredioAgr
                Case "MnuRepCarteraPredios"
                    lobjRep.EnuReporte = EnuReporteDef.enuCarteraPorPredio
                Case "MnuRepDirTf"
                    lobjRep.EnuReporte = EnuReporteDef.enuDirTf
            End Select
            lobjRep.SGenereReporte()
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
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub
    Private Sub SGenerePazYSalvo()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Dim lblnProcese = False
            GobjPanDat.SControleProcesoObj(True)
            If FblnPuedeGenPazYSalvo() Then
                Mouse.OverrideCursor = Cursors.Wait
                lblnProcese = MobjObjetoWin.FblnEstaPazYSalvo(lstrMens)
            Else
                lstrMens = "Para generar un Paz y Salvo, no deben haber procesos pendientes " &
                        "de llevarse a cabo!"
            End If
            If lblnProcese Then
                Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO)
                Dim lobjPredio As ClsPredio
                If MobjObjetoWin.ObjIdPredioStr.ObjValorPro =
                        MobjObjetoWin.ObjIdPredioAgrupadorStr.ObjValorPro Then
                    lobjPredio = MobjObjetoWin
                Else
                    lobjPredio = MobjObjetoWin.ObjPredioAgrupador
                End If
                lobjRep.SGenerePazYSalvo(lobjPredio)
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ValorArgumentoInvalidoException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentoInvalidoPanException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                If String.IsNullOrEmpty(lstrMens) Then
                    lstrMens = "Paz y Salvo generado exitosamente!"
                End If
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                GobjPanDat.SControleProcesoObj(False)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                GobjPanDat.SControleProcesoObj(False, True)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub
    Private Function FblnPuedeGenPazYSalvo() As Boolean
        Dim lblnPuede As Boolean = Not (ClsOrionCop.FblnHayItemsPorFacturar OrElse
                ClsOrionCop.FblnHayPrefacturas OrElse ClsOrionCop.FblnHacerCierreMes OrElse
                ClsOrionCop.FblnDebeCausarInt)
        If lblnPuede Then
            With GobjParametros
                lblnPuede = Not (.ObjAnoActual.FblnDebeCalcularCuotas OrElse
                    .ObjAnoActual.FblnDebeAjustarCuotasAdmin OrElse
                    ClsOrionCop.FblnHayItemsPorFacturar OrElse
                    ClsOrionCop.FblnCrearAno)
            End With
        End If
        Return lblnPuede
    End Function
    Private Sub SAdicionesCtlsRestringidos()
        SAdicioneControlRestringido(tbcPredios)
        SAdicioneControlRestringido(bttAbrirCliente)
        SAdicioneControlRestringido(bttAbrirCuenta)
        SAdicioneControlRestringido(bttEncontrarAdmin)
        SAdicioneControlRestringido(bttEncontrarArren)
        SAdicioneControlRestringido(bttEncontrarRepL)
        SAdicioneControlRestringido(lsbPrediosAgrupados)
        SAdicioneControlRestringido(rdbAnuladas)
        SAdicioneControlRestringido(rdbCanceladas)
        SAdicioneControlRestringido(rdbTodas)
        SAdicioneControlRestringido(rdbVencidas)
    End Sub
    ''' <summary>
    ''' Calcula el total de la base de calculo para el coeficiente de propiedad y calcula el
    ''' coeficiente de propiedad de cada uno de los predios. Actualiza los datos en la BD
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SCalculeCP()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            lstrMens = "Calculando Coeficientes de Propiedad!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            SCalculeCPCopr()
            lblnNoHayError = True
        Catch ex As ProveedorBdPanException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString()
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString()
        Catch ex As ArgumentNullException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString()
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString()
        Finally
            If lblnNoHayError Then
                MobjObjetoWin.SRefresqueObj()
                SMuestreDatos()
                lstrMens = "Proceso finalizado exitosamente."
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Private Sub SCalculeCPCopr()
        Mouse.OverrideCursor = Cursors.Wait
        ClsOrionCop.SActualiceTotalAreaCopr()
        Mouse.OverrideCursor = Cursors.Arrow
        SLevanteEveNoti(String.Empty, String.Empty, 0, EnuSeveridadNot.EnuOk)
    End Sub
    ''' <summary>
    ''' Verifica que el Propietario de los Predios Agrupados sea el mismo Propietario del Predio Agrupador.
    ''' </summary>
    ''' <remarks></remarks>
    Private Function FblnPrediosIntegros() As Boolean
        Mouse.OverrideCursor = Cursors.Wait
        Dim lstrPrediosConProb = ClsOrionCop.FstrPrediosDifieren()
        Dim lstrMens As String, lblnPrediosIntegros As Boolean
        If Not IsNothing(lstrPrediosConProb) AndAlso lstrPrediosConProb.Length > 0 Then
            lstrMens = "Predios con problemas de integridad con el predio agrupador: " &
                    vbCrLf
            For Each lstrPredio As String In lstrPrediosConProb
                lstrMens &= lstrPredio & vbCrLf
            Next
            lstrMens = lstrMens.Substring(0, lstrMens.Length - 2)
            MsgBox(lstrMens, vbOKOnly + MsgBoxStyle.Exclamation, "Error")
            lblnPrediosIntegros = False
        Else
            lstrMens = "No hay problemas de Integridad en los Predios!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            lblnPrediosIntegros = True
        End If
        Mouse.OverrideCursor = Cursors.Arrow
        Return lblnPrediosIntegros
    End Function
#End Region
#Region "Propietarios"
    Private Sub SHabiliteBttsProp()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            If MblnEsPredioAgru Then
                Dim lblnHayPropietarios = Not String.IsNullOrEmpty(MstrIdPropSel)
                bttVincularProp.IsEnabled = True
                bttModificarProp.IsEnabled = lblnHayPropietarios
                bttDesvincularProp.IsEnabled = lblnHayPropietarios
                lblNotaProp.Visibility = Visibility.Collapsed
            Else
                bttVincularProp.IsEnabled = False
                bttModificarProp.IsEnabled = False
                bttDesvincularProp.IsEnabled = False
                lblNotaProp.Visibility = Visibility.Visible
            End If
        Else
            bttVincularProp.IsEnabled = False
            bttModificarProp.IsEnabled = False
            bttDesvincularProp.IsEnabled = False
            lblNotaProp.Visibility = Visibility.Collapsed
        End If
    End Sub
    Private Sub SEstablezcaPropSel()
        Dim ldrvFilaActual As DataRowView
        ldrvFilaActual = dgrPropietarios.SelectedItem
        MobjPropSel = Nothing
        If Not IsNothing(ldrvFilaActual) Then
            MstrIdPropSel = ldrvFilaActual("IdTerceroPropietario")
            If MobjObjetoWin.BlnExiste AndAlso Not String.IsNullOrEmpty(MstrIdPropSel) Then
                If MobjObjetoWin.ColPropietarios.Count > 0 Then
                    If MobjObjetoWin.ColPropietarios.Contains(MstrIdPropSel) Then
                        MobjPropSel = MobjObjetoWin.ColPropietarios(MstrIdPropSel.ToString())
                    Else
                        MobjPropSel = MobjObjetoWin.ColPropietarios(1)
                    End If
                Else
                    MobjPropSel = Nothing
                End If
            End If
        Else
            MstrIdPropSel = String.Empty
        End If
    End Sub
    Private Sub SVincularProp()
        MobjPropSel = MobjObjetoWin.FobjNewPropietario()
        If MblnEsPredioAgru Then
            Dim lwinProp As New WinPropietario(MobjPropSel, True) With {
                .WinPadre = Me
            }
            lwinProp.Show()
        End If
        SValide()
        SHabiliteBttsProp()
    End Sub
    Private Sub SModifiqueProp()
        If MobjPropSel IsNot Nothing Then
            SLevanteEveOk()
            Dim lwinProp As New WinPropietario(MobjPropSel, False) With {
                .WinPadre = Me
            }
            lwinProp.Show()
        End If
        SHabiliteBttsProp()
    End Sub
    Private Sub SDesvinculeProp()
        If Not String.IsNullOrEmpty(MstrIdPropSel) Then
            MobjObjetoWin.SDesvinculeProp(MstrIdPropSel)
            SMuestrePropietarios()
            SValide()
            FblnPropietariosOk(True)
        End If
        SHabiliteBttsProp()
    End Sub
    Friend Sub SAcepteProp(ablnVinculando As Boolean)
        If ablnVinculando Then
            MobjObjetoWin.SAdicioneNewProp(MobjPropSel)
        End If
        SLevanteEveOk()
        SMuestrePropietarios()
        MobjObjetoWin.ObjIdPredioAgrupadorStr.SValide()
        SValide()
        FblnPropietariosOk(True)
        SHabiliteBttsProp()
    End Sub
    Private Function FblnEsValidoProp() As Boolean
        Dim lblnEsValido = MobjPropSel IsNot Nothing
        If lblnEsValido Then
            lblnEsValido = MobjPropSel.ObjIdCliente_PropDbl.BlnEsValido AndAlso
                    MobjPropSel.ObjPorcentajePartiDbl.BlnEsValido
        End If
        Return lblnEsValido
    End Function
    Private Sub SActualiceTblPropietarios()
        If MdtbProp Is Nothing Then
            MdtbProp = MobjObjetoWin.FdtbPropietarios
        ElseIf EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            MdtbProp.Rows.Clear()
            For Each lobjProp As ClsPropietario In MobjObjetoWin.ColPropietarios
                Dim ldrwNewFila As DataRow = MdtbProp.NewRow
                ldrwNewFila(StrCampoCarpeta) = GshrIdCarpeta
                ldrwNewFila(StrCampoCentroUtil) = GshrIdCentroUtil
                ldrwNewFila(ClsIdPredio_PropStr.SstrNombreCampoBd) =
                        lobjProp.ObjIdPredio_PropStr.ObjValorPro
                ldrwNewFila(ClsIdCliente_PropDbl.SstrNombreCampoBd) =
                        lobjProp.ObjIdCliente_PropDbl.ObjValorPro
                ldrwNewFila(ClsNombreCompleto_PropStr.SstrNombreCampoBd) =
                        lobjProp.ObjNombreCompleto_PropStr.ObjValorPro
                ldrwNewFila(ClsPorcentajePartiDbl.SstrNombreCampoBd) =
                        lobjProp.ObjPorcentajePartiDbl.ObjValorPro
                MdtbProp.Rows.Add(ldrwNewFila)
            Next
        End If
    End Sub
    Private Function FblnPropietariosOk(ablnNotifica As Boolean)
        Dim ldblPorcientoProp As Double, ldblTotProProp = 0.0, lblnOk As Boolean
        Dim lstrMens = String.Empty
        If MdtbProp IsNot Nothing Then
            If MdtbProp.Rows.Count = 0 Then
                lblnOk = False
                lstrMens = "El predio debe tener al menos un Propietario vinculado!"
            Else
                For Each ldrwProp As DataRow In MdtbProp.Rows
                    ldblPorcientoProp = ClsPanorama.FobjValorCampo(ldrwProp(
                        ClsPorcentajePartiDbl.SstrNombreCampoBd), EnuTipoValor.EnuDouble)
                    ldblTotProProp += ldblPorcientoProp
                Next
                lblnOk = Math.Round(ldblTotProProp, 4) = 1
                If ablnNotifica AndAlso Not lblnOk Then
                    lstrMens = "El porcentaje de propiedad total debe ser 100%"
                End If
            End If
        Else
            lblnOk = False
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, "", 0, EnuSeveridadNot.EnuInformacion)
        End If
        Return lblnOk
    End Function
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            Select Case lelmElemento.Name
                Case HbttTercero.Name
                    SAbraCliente(True)
                Case "bttAbrirCliente"
                    SAbraCliente(False)
                Case "bttAbrirCuenta"
                    SAbraCuenta()
                Case "bttVincularProp"
                    SVincularProp()
                Case "bttModificarProp"
                    SModifiqueProp()
                Case "bttDesvincularProp"
                    SDesvinculeProp()
                Case "bttEncontrarArren"
                    SBusqueCliente(EnuTipoClienteTer.enuArrendatario)
                Case "bttEncontrarRepL"
                    SBusqueCliente(EnuTipoClienteTer.enuRepLegalArre)
                Case "bttEncontrarAdmin"
                    SBusqueCliente(EnuTipoClienteTer.enuAdministrador)
            End Select
        End If
    End Sub
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            Select Case lelmElemento.Name
                Case "MnuImportarPredio"
                    SAbraImportar(True)
                Case "MnuImportarPropietarios"
                    SAbraImportar(False)
                Case "MnuIntegridadPredios"
                    Dim lblnNoUsado = FblnPrediosIntegros()
                Case "MnuRepPrediosSector"
                    SGenereRep("MnuRepPrediosSector")
                Case "MnuRepPrediosCliente"
                    SGenereRep("MnuRepPrediosCliente")
                Case "MnuRepPropiPorCP"
                    SGenereRep("MnuRepPropiPorCP")
                Case "MnuRepCuotasAdminProp"
                    SGenereRep("MnuRepCuotasAdminProp")
                Case "MnuRepCarteraPredAgr"
                    SGenereRep("MnuRepCarteraPredAgr")
                Case "MnuRepCarteraPredios"
                    SGenereRep("MnuRepCarteraPredios")
                Case "MnuRepDirTf"
                    SGenereRep("MnuRepDirTf")
                Case "MnuRepPazYSalvo"
                    SGenerePazYSalvo()
                Case "MnuCalcularCP"
                    SCalculeCP()
            End Select
        End If
    End Sub
    Private Sub MnuContextual_Click(sender As Object, e As RoutedEventArgs)
        If sender.Equals(MnuAbrirPredioConC) Then
            Dim lstrMens As String
            If IsNothing(lsbPrediosAgrupados.SelectedItem) OrElse
                    String.IsNullOrEmpty(lsbPrediosAgrupados.SelectedItem.ToString) Then
                lstrMens = "Debe seleccionar un Predio!"
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            Else
                If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                    Dim lstrPredio = lsbPrediosAgrupados.SelectedItem.ToString
                    txtIdPredio.Text = lstrPredio
                    SAbraPredio()
                Else
                    lstrMens = "Solo puede cambiar de predio cuando se está consultando!"
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            End If
        End If
    End Sub
    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            ltxtTextBox.SelectAll()
        End If
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        If Not HblnSeEstaCerrando AndAlso Not HblnMostrandoDatos Then
            Dim lstrNombreTextBox = String.Empty
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is TextBox OrElse TypeOf lelmElemento Is CheckBox Then
                If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                    Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError As Boolean
                    Try
                        If TypeOf lelmElemento Is TextBox Then
                            lstrNombreTextBox = lelmElemento.Name
                            SRegistreTextBox(lelmElemento)
                        End If
                        With MobjObjetoWin
                            Select Case lelmElemento.Name
                                Case "chkFacturarSer"
                                    .ObjFacturarPorServicio_PreBln.ObjValorPro = chkFacturarSer.IsChecked
                                Case "chkNoConsItemsFac"
                                    .ObjNoConsolidarItemsFacBln.ObjValorPro = chkNoConsItemsFac.IsChecked
                            End Select
                        End With
                        lblnNoHayError = True
                    Catch ex As PanLException
                        lstrMens = ex.Message
                        lstrMensEx = ex.ToString()
                    Catch ex As PanDatException
                        lstrMens = ex.Message
                        lstrMensEx = ex.ToString()
                    Catch ex As Exception
                        lstrMens = ex.Message
                        lstrMensEx = ex.ToString()
                    Finally
                        If Not lstrNombreTextBox = "txtPorcientoPropNew" Then
                            SMuestreDatos()
                        End If
                        If lblnNoHayError Then
                            If Not String.IsNullOrEmpty(lstrMens) Then
                                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuAdvertencia)
                            End If
                        Else
                            SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                        End If
                    End Try
                End If
            End If
        End If
    End Sub
    Private Sub OnCboCambio(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is ComboBox Then
            If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando AndAlso
                    Not HblnMostrandoDatos Then
                Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError As Boolean
                Try
                    With MobjObjetoWin
                        Select Case lelmElemento.Name
                            Case "cboSectores"
                                .ObjIdSector_PredioShr.ObjValorPro = cboSectores.SelectedIndex
                            Case "cboDestFact"
                                .ObjIdTipoDestinatarioFacturaByt.ObjValorPro = cboDestFact.SelectedIndex
                            Case "cboEstadoActualDeuda"
                                .ObjIdEstadoDeuda_PredioByt.ObjValorPro = cboEstadoActualDeuda.SelectedIndex
                        End Select
                    End With
                    lblnNoHayError = True
                Catch ex As PanLException
                    lstrMens = ex.Message
                    lstrMensEx = ex.ToString()
                Catch ex As PanDatException
                    lstrMens = ex.Message
                    lstrMensEx = ex.ToString()
                Catch ex As Exception
                    lstrMens = ex.Message
                    lstrMensEx = ex.ToString()
                Finally
                    SMuestreDatos()
                    If Not lblnNoHayError Then
                        SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                    End If
                End Try
            End If
        End If
    End Sub
    Private Sub TbcPredios_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles tbcPredios.SelectionChanged
        If TypeOf sender Is TabControl Then
            If MentTabSeleccionado <> tbcPredios.SelectedIndex Then
                MentTabSeleccionado = tbcPredios.SelectedIndex
                SMuestreInfTab()
            End If
        End If
    End Sub
    Private Sub Txt_KeyDown(sender As Object, e As KeyEventArgs) Handles txtIdPredio.KeyDown
        If e.Key = Key.Return OrElse e.Key = Key.Tab Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando AndAlso
                    MobjObjetoWin.ObjIdPredioStr.ToString() <> txtIdPredio.Text Then
                SAbraPredio()
            End If
        End If
    End Sub
    Private Sub Chk_Click(sender As Object, e As RoutedEventArgs) Handles chkFacturarSer.Click,
            chkNoConsItemsFac.Click
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando AndAlso
                Not HblnMostrandoDatos Then
            MobjObjetoWin.ObjFacturarPorServicio_PreBln.ObjValorPro = chkFacturarSer.IsChecked
            MobjObjetoWin.ObjNoConsolidarItemsFacBln.ObjValorPro = chkNoConsItemsFac.IsChecked
        End If
        SMuestreDatos()
    End Sub
    Private Sub Dgr_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) _
            Handles dgrPropietarios.SelectionChanged
        If TypeOf sender Is DataGrid Then
            SEstablezcaPropSel()
        End If
    End Sub
    Private Sub Rdb_Click(sender As Object, e As RoutedEventArgs) Handles rdbVencidas.Click, rdbAnuladas.Click,
            rdbCanceladas.Click, rdbTodas.Click
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            SMuestreFacturas()
        End If
    End Sub
    Private Sub Dgr_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
                dgrFacturas.MouseRightButtonUp
        Dim ldrvFilaActual As DataRowView
        If TypeOf sender Is DataGrid Then
            Dim ldgrActual As DataGrid = sender
            Dim lstrNroFact As String, lstrPrefijo As String, lentIdFact As Integer
            ldrvFilaActual = ldgrActual.SelectedItem
            If Not IsNothing(ldrvFilaActual) Then
                lstrNroFact = ldrvFilaActual("NroFact")
                lstrPrefijo = ClsPanorama.FstrPrefijoDcto(lstrNroFact)
                lentIdFact = ClsPanorama.FentIdDcto(lstrNroFact)
                SAbraFactura(lstrPrefijo, lentIdFact)
            End If
        End If
    End Sub
    Protected Overrides Sub EwinClosed(sender As Object, e As EventArgs)
        GenuTamanoIcono = MenuTamanoIcono
        If WinPadre IsNot Nothing Then
            If WinPadre.Visibility <> Visibility.Visible Then
                WinPadre.Visibility = Visibility.Visible
                WinPadre.SRefresqueWin()
            End If
        End If
    End Sub
#End Region
End Class