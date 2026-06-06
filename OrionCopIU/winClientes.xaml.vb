Public Class WinClientes
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuIdCliente = 0
        enuRegimenVtas
        enuEstadoDeuda
        enuComentario
        enuRecibeCorreos
        enuMedioPago
        enuCodPostal
        enuRetFte
        enuRetIca
        enuGranCont
        enuAutoret
        enuRetIva
        enuRST
        enuFacPorSer
        ' Tercero
        enuApellido1
        enuApellido2
        enuDirec1
        enuDirec2
        enuPagWeb
        enuEmail
        enuTelMovil
        enuTelMovil2
        enuTel1
        enuTel2
        enuTipoIdent
        enuTipoTercero
        enuIdTercero
        enuNom1
        enuNom2
        enuRazonSocial
        enuPaisDir
        enuDptoDir
        enuCiudadDir
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsCliente = Nothing
    Private WithEvents MwinVentana As ClsFormInterface = Nothing
    Private MblnPoblandoCbo As Boolean = False
    Private MblnIngresoTel As Boolean = False
    Private MnuImportarClientes As MenuItemPan = Nothing
    Private MnuReasignarId As MenuItem = Nothing
    ' Manejo Ubicacion
    Private MobjUbicacion As ClsUbicacion = Nothing
    '
    Private MnuAbrirPredioConC As MenuItem = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomClientes
#End Region

#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuCliente
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicioneControlRestringido(bttAbrirTer)
        SAdicioneControlRestringido(lsbPrediosDelCliente)
        SAdicioneControlRestringido(bttAbrirCuenta)
        Dim lcolControlesLlave As New Collection From {
            txtNroDoc
        }
        SCargueForma(EnuElementosAdicionalesDef.enuBuscar, 33, lcolControlesLlave, cboTipoDoc, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        txtNroDoc.Focus()
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
            ObjObjetoWin = ClsOrionCop.FobjCliente(EnuModoInstanciaObjDef.enuNavegable)
            If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                ObjObjetoWin.SVayaAlPrimero()
            End If
        End If
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub

    Protected Overrides Sub SInicialiceControles()
        chkFacturarSer.IsChecked = False
        StcValidaControl(EnuValidEntrada.enuIdCliente) = lblNroDoc
        StcValidaControl(EnuValidEntrada.enuRegimenVtas) = lblRegimenVentas
        StcValidaControl(EnuValidEntrada.enuEstadoDeuda) = lblEstadoActualDeuda
        StcValidaControl(EnuValidEntrada.enuComentario) = lblComentario
        StcValidaControl(EnuValidEntrada.enuRecibeCorreos) = chkRecibeDocsEmail
        StcValidaControl(EnuValidEntrada.enuMedioPago) = lblMedPagoPreferido
        StcValidaControl(EnuValidEntrada.enuCodPostal) = lblCodPostal
        StcValidaControl(EnuValidEntrada.enuRetFte) = chkAgenteRetFte
        StcValidaControl(EnuValidEntrada.enuRetIca) = chkAgenteRetIca
        StcValidaControl(EnuValidEntrada.enuGranCont) = chkGranCont
        StcValidaControl(EnuValidEntrada.enuAutoret) = chkAutoRet
        StcValidaControl(EnuValidEntrada.enuRetIva) = chkRetieneIva
        StcValidaControl(EnuValidEntrada.enuRST) = chkRegSimpleTri
        StcValidaControl(EnuValidEntrada.enuFacPorSer) = chkFacturarSer
        'Tercero
        StcValidaControl(EnuValidEntrada.enuApellido1) = lblApellido1
        StcValidaControl(EnuValidEntrada.enuApellido2) = lblApellido2
        StcValidaControl(EnuValidEntrada.enuDirec1) = lblDireccion1
        StcValidaControl(EnuValidEntrada.enuDirec2) = lblDireccion2
        StcValidaControl(EnuValidEntrada.enuPagWeb) = lblPaginaWeb
        StcValidaControl(EnuValidEntrada.enuEmail) = lblEmail
        StcValidaControl(EnuValidEntrada.enuTelMovil) = lblMovil
        StcValidaControl(EnuValidEntrada.enuTelMovil2) = lblMovil2
        StcValidaControl(EnuValidEntrada.enuTel1) = lblTelefono1
        StcValidaControl(EnuValidEntrada.enuTel2) = lblTelefono2
        StcValidaControl(EnuValidEntrada.enuTipoIdent) = lblTipoDoc
        StcValidaControl(EnuValidEntrada.enuTipoTercero) = lblTipoTercero
        StcValidaControl(EnuValidEntrada.enuIdTercero) = lblNroDoc
        StcValidaControl(EnuValidEntrada.enuNom1) = lblNombre1
        StcValidaControl(EnuValidEntrada.enuNom2) = lblNombre2
        StcValidaControl(EnuValidEntrada.enuRazonSocial) = lblRazonSocial
        StcValidaControl(EnuValidEntrada.enuPaisDir) = lblPaisDir
        StcValidaControl(EnuValidEntrada.enuDptoDir) = lblDptoDir
        StcValidaControl(EnuValidEntrada.enuCiudadDir) = lblCiudadDir
        ' Visibilizar boton terceros
        If GstrIdUsuario = GCSTRUSUARIOU Then
            bttAbrirTer.Visibility = Visibility.Visible
        End If
        '
        SPuebleComboBoxes()
        SPuebleCboPaisDir()
        HbttAceptar.TabIndex = 100
        HbttCancelar.TabIndex = 101
    End Sub

    Protected Overrides Sub SMuestreDatos()
        HblnMostrandoDatos = True
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            SLevanteEveNoti("No hay Clientes para ser mostrados!", "", 0,
                        EnuSeveridadNot.EnuInformacion)
            txtNroDoc.IsEnabled = False
            'Esconder Empresa
            grbDatosEmpresas.Visibility = Visibility.Collapsed
            lblDigitoVerificacion.Visibility = Visibility.Hidden
            txtDigVer.Visibility = Visibility.Hidden
        End If
        With MobjObjetoWin
            cboRegimenVentas.SelectedIndex = .ObjIdRegimenVentasByt.ObjValorPro
            cboEstadoActualDeuda.SelectedIndex = .ObjIdEstadoDeudaByt.ObjValorPro
            cboMedPagoPref.SelectedIndex = .ObjIdMedioPagoClienteByt.ObjValorPro
            chkFacturarSer.IsChecked = .ObjFactPorServicio_CliBln.ObjValorPro
            chkAgenteRetFte.IsChecked = .ObjEsAgenteReteFteBln.ObjValorPro
            chkAutoRet.IsChecked = .ObjEsAutorretenedorBln.ObjValorPro
            chkGranCont.IsChecked = .ObjEsGranContrBln.ObjValorPro
            chkRegSimpleTri.IsChecked = .ObjEsRegimenSimpleTBln.ObjValorPro
            chkRetieneIva.IsChecked = .ObjRetieneIvaBln.ObjValorPro
            chkAgenteRetIca.IsChecked = .ObjRetieneIcaBln.ObjValorPro
            txtCodPostal.Text = .ObjCodigoPostalStr.ObjValorPro
            txtEmail.Text = .ObjEmailStr.ObjValorPro
            chkRecibeDocsEmail.IsChecked = .ObjRecibeDocsPorEmailBln.ObjValorPro
            txtFechaIngreso.Content = .ObjFechaIngresoDtm.ToString
            txtNombreCompleto.Content = .ObjNombreCompletoStr.ObjValorPro
            txtComentario.Text = .ObjComentario_ClienteStr.ObjValorPro
            txtEstadoSugeridoDeuda.Content = ClsOrionCop.FstrNombreDatoConstanteOri(EnuGrupoConstantesOriDef.EnuEstadoDeuda,
                    MobjObjetoWin.FenuEstadoSugeridoDeuda)
        End With
        SMuestreDatosTer()
        SMuestreUbicacionDir()
        SPueblePrediosDelCliente()
        Title = My.Resources.FichaCliente
        If Not String.IsNullOrEmpty(txtNroDoc.Text) Then
            Title &= txtNroDoc.Text
            Title &= " " & MobjObjetoWin.ObjNombreCompletoStr.ObjValorPro
        End If
        SValide()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If txtNroDoc.Focus Then
                txtNroDoc.SelectAll()
            End If
        End If
        HblnMostrandoDatos = False
    End Sub

    Private Sub SMuestreDatosTer()
        With MobjObjetoWin
            txtApellido1.Text = .ObjApellidoPrimeroStr.ObjValorPro
            txtApellido2.Text = .ObjApellidoSegundoStr.ObjValorPro
            txtDireccion1.Text = .ObjDireccionUnoStr.ObjValorPro
            txtDireccion2.Text = .ObjDireccionDosStr.ObjValorPro
            txtPaginaWeb.Text = .ObjPaginaWebStr.ObjValorPro
            txtEmail.Text = .ObjEmailStr.ObjValorPro
            txtFechaIngreso.Content = .ObjFechaCreacionDtm.ObjValorPro
            cboTipoDoc.SelectedIndex = .ObjTipoDocIdentidadByt.ObjValorPro
            cboTipoTercero.SelectedIndex = .ObjTipoTerceroByt.ObjValorPro
            txtNroDoc.Text = Format(.ObjIdTerceroDbl.ObjValorPro, GCSTRFMTIDTERCERO)
            txtDigVer.Text = .ObjIdTerceroDbl.SbyDigitoVerificacion
            txtNombre1.Text = .ObjNombrePrimeroStr.ObjValorPro
            txtNombre2.Text = .ObjNombreSegundoStr.ObjValorPro
            txtRazonSocial.Text = .ObjRazonSocialStr.ObjValorPro
            txtTelefono1.Text = .ObjTelefonoUnoStr.ObjValorPro
            txtTelefono2.Text = .ObjTelefonoDosStr.ObjValorPro
            txtMovil.Text = .ObjCelularStr.ObjValorPro
            txtMovil2.Text = .ObjCelular2Str.ObjValorPro
        End With
        SUbiqueControles()
        SDeshabiliteControles()
    End Sub

    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.CenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(aintIndice:=EnuValidEntrada.enuIdCliente) = .ObjIdClienteDbl.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuRegimenVtas) = .ObjIdRegimenVentasByt.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuComentario) = .ObjComentario_ClienteStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuRecibeCorreos) = .ObjRecibeDocsPorEmailBln.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuEstadoDeuda) = .ObjIdEstadoDeudaByt.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuMedioPago) = .ObjIdMedioPagoClienteByt.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuRetFte) = .ObjEsAgenteReteFteBln.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuRetIca) = .ObjRetieneIcaBln.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuGranCont) = .ObjEsGranContrBln.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuAutoret) = .ObjEsAutorretenedorBln.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuRetIva) = .ObjRetieneIvaBln.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuRST) = .ObjEsRegimenSimpleTBln.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuFacPorSer) = .ObjFactPorServicio_CliBln.BlnEsValido
                'Tercero
                StcValidValido(aintIndice:=EnuValidEntrada.enuApellido1) = .ObjApellidoPrimeroStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuApellido2) = .ObjApellidoSegundoStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuDirec1) = .ObjDireccionUnoStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuDirec2) = .ObjDireccionDosStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuPagWeb) = .ObjPaginaWebStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuEmail) = .FblnEsValidoEmail
                StcValidValido(aintIndice:=EnuValidEntrada.enuTelMovil) = .ObjCelularStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuTelMovil2) = .ObjCelular2Str.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuTel1) = .ObjTelefonoUnoStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuTel2) = .ObjTelefonoDosStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuTipoIdent) = .ObjTipoDocIdentidadByt.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuTipoTercero) = .ObjTipoTerceroByt.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuIdTercero) = .ObjTerceroCliente.ObjIdTerceroDbl.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuNom1) = .ObjNombrePrimeroStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuNom2) = .ObjNombreSegundoStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuRazonSocial) = .ObjRazonSocialStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuPaisDir) = .ObjPaisDirStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuDptoDir) = .ObjDepartamentoDirByt.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuCiudadDir) = .ObjCiudadDirShr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuCodPostal) = .ObjCodigoPostalStr.BlnEsValido
            End With
        End If
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
        If MblnIngresoTel Then
            If Not (MobjObjetoWin.ObjTelefonoUnoStr.BlnEsValido AndAlso
                    MobjObjetoWin.ObjTelefonoDosStr.BlnEsValido AndAlso
                    MobjObjetoWin.ObjCelularStr.BlnEsValido AndAlso
                    MobjObjetoWin.ObjCelular2Str.BlnEsValido) Then
                Dim lstrMens = "El campo Teléfono solo admite números. Debe tener al menos 10 " &
                        "caracteres o puede dejarlo vacío.!"
                SLevanteEveNoti(lstrMens, "Revise los números de teléfono ingresados.", 0,
                            EnuSeveridadNot.EnuInformacion)
            End If
            MblnIngresoTel = False
        End If
    End Sub

    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjIdCarpetaClienteShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtilClienteShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdClienteDbl.ObjValorPro = txtNroDoc.Text
            .ObjFactPorServicio_CliBln.ObjValorPro = chkFacturarSer.IsChecked
            .ObjIdRegimenVentasByt.ObjValorPro = cboRegimenVentas.SelectedIndex
            .ObjIdEstadoDeudaByt.ObjValorPro = cboEstadoActualDeuda.SelectedIndex
            .ObjEsAgenteReteFteBln.ObjValorPro = chkAgenteRetFte.IsChecked
            .ObjEsAutorretenedorBln.ObjValorPro = chkAutoRet.IsChecked
            .ObjEsGranContrBln.ObjValorPro = chkGranCont.IsChecked
            .ObjEsRegimenSimpleTBln.ObjValorPro = chkRegSimpleTri.IsChecked
            .ObjRetieneIvaBln.ObjValorPro = chkRetieneIva.IsChecked
            .ObjRetieneIcaBln.ObjValorPro = chkAgenteRetIca.IsChecked
            .ObjCodigoPostalStr.ObjValorPro = txtCodPostal.Text
            .ObjComentario_ClienteStr.ObjValorPro = txtComentario.Text
            .ObjRecibeDocsPorEmailBln.ObjValorPro = chkRecibeDocsEmail.IsChecked
            .ObjDireccionDosStr.ObjValorPro = txtDireccion2.Text
            .ObjDireccionUnoStr.ObjValorPro = txtDireccion1.Text
            .ObjPaginaWebStr.ObjValorPro = txtPaginaWeb.Text
            .ObjEmailStr.ObjValorPro = txtEmail.Text
            .ObjPaisDirStr.ObjValorPro = FstrIdPaisActual()
            .ObjDepartamentoDirByt.ObjValorPro = FBytIdDptoActual()
            .ObjCiudadDirShr.ObjValorPro = FshrIdCiudadActual()
            .ObjIdTerceroDbl.ObjValorPro = txtNroDoc.Text
            .ObjIdTerceroDbl.SbyDigitoVerificacion = txtDigVer.Text
            .ObjNombrePrimeroStr.ObjValorPro = txtNombre1.Text
            .ObjNombreSegundoStr.ObjValorPro = txtNombre2.Text
            .ObjApellidoPrimeroStr.ObjValorPro = txtApellido1.Text
            .ObjRazonSocialStr.ObjValorPro = txtRazonSocial.Text
            .ObjApellidoSegundoStr.ObjValorPro = txtApellido2.Text
            .ObjCelularStr.ObjValorPro = txtMovil.Text
            .ObjCelular2Str.ObjValorPro = txtMovil2.Text
            .ObjTelefonoUnoStr.ObjValorPro = txtTelefono1.Text
            .ObjTelefonoDosStr.ObjValorPro = txtTelefono2.Text
            .ObjTipoDocIdentidadByt.ObjValorPro = cboTipoDoc.SelectedIndex
            .ObjTipoTerceroByt.ObjValorPro = cboTipoTercero.SelectedIndex
            .ObjIdMedioPagoClienteByt.ObjValorPro = cboMedPagoPref.SelectedIndex
        End With
    End Sub

    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        ' Adicionar menú Importar
        MnuImportarClientes = FmnuiMenuItemPan("MnuImportarClientes", "_Importar Clientes", 1, "")
        Dim lsepSeparador As New Separator
        Dim lmnuCuentaCliente As MenuItem = FmnuiMenuItem("MnuCuentaCliente",
                "Consultar Cuenta del Cliente", "RecMnuItemSec")
        ' Adicionar menú Reasignar Id tercero
        MnuReasignarId = FmnuiMenuItem("MnuReasignarId", "Reasi_gnar Id. Tercero",
                "RecMnuItemSec")
        HmnuAcciones.Items.Insert(7, lmnuCuentaCliente)
        HmnuAcciones.Items.Insert(8, MnuImportarClientes)
        HmnuAcciones.Items.Insert(9, MnuReasignarId)
        HmnuAcciones.Items.Insert(10, lsepSeparador)
        ' Adicionar Menues Reportes
        Dim lmnuReportes As MenuItem = FmnuiMenuItem("MnuReportes", "R_eportes", "RecMnuItemPriInf")
        Dim lmnuConsPago = FmnuiMenuItem("MnuConstPago", "Constancia Valores Facturados", "RecMnuItemSec")
        lmnuConsPago.ToolTip = "Genera Constancia de los Pagos Efectuados en el año."
        lmnuReportes.Items.Add(lmnuConsPago)
        MenuVen.Items.Insert(1, lmnuReportes)
        Dim lstrAnos = FstrUltimosAnos()
        Dim lmnuItem As MenuItem
        Dim lstrNombreMenuI As String
        For Each lstrAno As String In lstrAnos
            lstrNombreMenuI = "MnuAno" & lstrAno
            lmnuItem = FmnuiMenuItem(lstrNombreMenuI, lstrAno, "RecMnuItemSec")
            lmnuConsPago.Items.Add(lmnuItem)
        Next
        'Reporte directorio clientes
        lmnuItem = FmnuiMenuItem("MnuRepDirClientes", "Directorio de Clientes", "RecMnuItemSec")
        lmnuItem.ToolTip = "Genera un Listado de Clientes con dirección, teléfonos y correo."
        lmnuReportes.Items.Add(lmnuItem)
        ' Reporte Facturas vivas con int. mora
        Dim lsepSepa2 As New Separator
        lmnuReportes.Items.Add(lsepSepa2)
        lmnuItem = FmnuiMenuItem("MnuRepFacturas", "Facturas por Pagar discriminando Int. Mora", "RecMnuItemSec")
        lmnuItem.ToolTip = "Genera un Listado de la Facturas que debe el Cliente discriminado los Intereses de Mora."
        lmnuReportes.Items.Add(lmnuItem)
        ' Adicionar Menu Contextual
        Dim lmnuMenuContextual As ContextMenu = FindResource("RecMnuPrediosClienteMC")
        lsbPrediosDelCliente.ContextMenu = lmnuMenuContextual
        For Each lobjObjetoMenu As Object In lmnuMenuContextual.Items
            If TypeOf lobjObjetoMenu Is MenuItem Then
                If lobjObjetoMenu.Name = "MnuAbrirPredioC" Then
                    MnuAbrirPredioConC = lobjObjetoMenu
                End If
            End If
        Next
    End Sub
#End Region

#Region "Invalida otros metodos de la clase base"
    Protected Overrides Sub SCree()
        MyBase.SCree()
        tbiId.IsSelected = True
        txtNroDoc.IsEnabled = True
        txtNroDoc.Focus()
    End Sub

    Protected Overrides Sub SModifique()
        MyBase.SModifique()
        SMuestreDatos()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            SRegistre()
            SValide()
        End If
        SDeshabiliteControles()
        cboTipoDoc.Style = FindResource("RecCtlNoHabilitado")
    End Sub

    Protected Overrides Sub SGuarde()
        MyBase.SGuarde()
        SRefresqueBotones()
    End Sub

    Protected Overrides Sub SCancele()
        MyBase.SCancele()
        SRefresqueWin()
    End Sub

    Protected Overrides Sub SEstablezcaWinConsultando()
        MyBase.SEstablezcaWinConsultando()
        cboTipoDoc.Style = FindResource("RecCtlNoHabilitado")
        bttAbrirCuenta.IsEnabled = True
        bttAbrirTer.IsEnabled = True
        txtNroDoc.Focus()
    End Sub
#End Region

#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        MyBase.SBuscar()
        If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
            If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable Then
                txtNroDoc.Text = StrResultadoBusqueda
                SAbraCliente()
                SMuestreDatos()
            End If
        End If
    End Sub

    ''' <summary>
    ''' Invalida la funcion "fblnDefinioBusqueda" de la clase base.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        SDefineBusquedaNombreCompleto()
        SDefineBusquedaPredioAgr_Prop()
        SDefineBusquedaPredioAgr_Arren()
        SDefineBusquedaPrimerApell()
        Return True
    End Function

    Private Sub SDefineBusquedaNombreCompleto()
        Dim lstrCamposMostrar As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdClienteDbl.SstrNombreCampoBd
        Dim lstrTabla As String = ClsCliente.SstrNombreTabla
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " & lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Nombre Completo", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub

    Private Sub SDefineBusquedaPrimerApell()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsTercero.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {"CONCAT(S." & ClsApellidoPrimeroStr.SstrNombreCampoBd & ", " &
                "' '" & ", S." & ClsNombrePrimeroStr.SstrNombreCampoBd & ")" & " AS ApellidoNombre"}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdTerceroDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = "ApellidoNombre"
        Dim lstrCampoRetornar As String = ClsIdClienteDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND S." &
                ClsNombrePrimeroStr.SstrNombreCampoBd & " <> ''"
        HwinBusqueda.SDefinaBusqueda("Primer Apellido", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, True)
    End Sub

    Private Sub SDefineBusquedaPredioAgr_Prop()
        Dim lstrTablaPri = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec = ClsPropietario.SstrNombreTabla
        Dim lstrCamSelTablaPri As String() = {"DISTINCT " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd}
        Dim lstrCampSelTablaSec As String() = {ClsIdCliente_PropDbl.SstrNombreCampoBd,
                ClsNombreCompleto_PropStr.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdCliente_PropDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = "P." & StrCampoCarpeta &
                " = " & GshrIdCarpeta & " AND P." &
                StrCampoCentroUtil & " = " &
                GshrIdCentroUtil & " AND " & "P." & ClsIdPredioStr.SstrNombreCampoBd &
                " = " & ClsIdPredioAgrupadorStr.SstrNombreCampoBd & " AND " &
                lstrCampoBusqueda & " <> ''"
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador - Propietario", lstrTablaPri,
                lstrTablaSec, lstrCamSelTablaPri, lstrCampSelTablaSec, lstrCampRelPri,
                lstrCampRelSec, lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, False)
    End Sub

    Private Sub SDefineBusquedaPredioAgr_Arren()
        Dim lstrTablaSec As String = ClsPredio.SstrNombreTabla
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrCamposTabSec = {ClsIdPredioAgrupadorStr.SstrNombreCampoBd,
                                 ClsIdClienteArrendatarioDbl.SstrNombreCampoBd}
        Dim lstrCamposTabPri As String() = {"DISTINCT " & ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdClienteArrendatarioDbl.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdClienteArrendatarioDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta &
                " AND P." & StrCampoCentroUtil & " = " & GshrIdCentroUtil
        lstrFiltro &= " AND " & ClsIdClienteArrendatarioDbl.SstrNombreCampoBd & " > 0 AND " &
                lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador - Arrendatario", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, False)
    End Sub
#End Region

#Region "Procedimientos Propios"
    Private Sub SPuebleComboBoxes()
        MblnPoblandoCbo = True
        Dim ldrwDataRow = ClsAdministrador.FdrwConstantesPan(EnuGrupoConstantesPanDef.enuTipoDocIdentidad)
        SPuebleComboBox(ldrwDataRow, cboTipoDoc)
        ldrwDataRow = ClsAdministrador.FdrwConstantesPan(EnuGrupoConstantesPanDef.enuTipoTercero)
        SPuebleComboBox(ldrwDataRow, cboTipoTercero)
        Dim ldrwRegimenIva = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuRegimenIva)
        SPuebleComboBox(ldrwRegimenIva, cboRegimenVentas)
        Dim ldrwEstadoDeuda = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuEstadoDeuda)
        SPuebleComboBox(ldrwEstadoDeuda, cboEstadoActualDeuda)
        ' Poblar combo Medios de Pago
        ldrwDataRow = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuMediosPago)
        SPuebleComboBox(ldrwDataRow, cboMedPagoPref)
        MblnPoblandoCbo = False
    End Sub
    Private Sub SPueblePrediosDelCliente()
        lsbPrediosDelCliente.Items.Clear()
        Dim lstrPrediosDelCliente() As String = MobjObjetoWin.FstrPrediosDelCliente
        If Not IsNothing(lstrPrediosDelCliente) AndAlso lstrPrediosDelCliente.Length > 0 Then
            For Each lstrPredio As String In lstrPrediosDelCliente
                lsbPrediosDelCliente.Items.Add(lstrPredio)
            Next
        End If
    End Sub
    Private Sub SAbraCuenta()
        Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
        lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, MobjObjetoWin.ObjIdClienteDbl.ObjValorPro})
        Dim lwinCuenta As New WinCuentaClientes With {
            .WinPadre = Me,
            .ObjObjetoWin = lobjCliente,
            .BlnVentanaAux = True
        }
        lwinCuenta.ShowDialog()
    End Sub
    Private Sub SAbraTercero()
        MwinVentana = New WinTerceros() With {
        .ObjObjetoWin = MobjObjetoWin.ObjTerceroCliente,
        .WinPadre = Me,
        .BlnVentanaAux = True
    }
        MwinVentana.ShowDialog()
    End Sub
    Private Sub SAbraCliente()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                If txtNroDoc.Text <> MobjObjetoWin.ObjIdClienteDbl.ToString Then
                    Dim lobjVlrLlave = {GshrIdCarpeta, GshrIdCentroUtil,
                            txtNroDoc.Text}
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
    Private Sub SAbraPredio(astrIdPredio As String)
        If Not String.IsNullOrEmpty(astrIdPredio) Then
            Dim lobjPredio As ClsPredio = ClsOrionCop.FobjNuevoPredio(EnuModoInstanciaObjDef.enuUnico)
            lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, astrIdPredio})
            If lobjPredio.BlnExiste Then
                Dim lwinPredio As New WinPredios() With {
                        .WinPadre = Me,
                        .ObjObjetoWin = lobjPredio,
                        .BlnVentanaAux = True
                    }
                lwinPredio.ShowDialog()
            End If
        End If
    End Sub
    Private Sub SAbraImportar()
        Dim lwinImportar As New WinImportar(MobjObjetoWin) With {
            .WinPadre = Me
        }
        lwinImportar.ShowDialog()
        lwinImportar.BlnVentanaAux = True
        SRefrescarClic()
    End Sub
    Private Sub SGenereRepVlrsFacturados(astrIdAno As String)
        Dim lstrFecIni As String, lstrFecFin As String
        GobjPanDat.SControleProcesoObj(True)
        Mouse.OverrideCursor = Cursors.Wait
        Dim lobjAno As ClsAno = GobjParametros.ColAnos(astrIdAno)
        If lobjAno.FblnEsAnoActual Then
            lstrFecFin = ClsPanoramaDat.FstrFechaNormalizada(Today)
        Else
            lstrFecFin = ClsPanoramaDat.FstrFechaNormalizada(lobjAno.DtmFechaFinAno)
        End If
        lstrFecIni = ClsPanoramaDat.FstrFechaNormalizada(lobjAno.DtmFechaInicioAno)
        Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
            .EnuReporte = EnuReporteDef.enuValoresFacturados,
            .StrFechaDesde = lstrFecIni,
            .StrFechaHasta = lstrFecFin,
            .DblIdCliente = MobjObjetoWin.ObjIdClienteDbl.ObjValorPro
        }
        lobjRep.SGenereReporte()
        GobjPanDat.SControleProcesoObj(False)
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Shared Sub SGenereDirClientes()
        GobjPanDat.SControleProcesoObj(True)
        Mouse.OverrideCursor = Cursors.Wait
        Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
            .EnuReporte = EnuReporteDef.enuDirClientes
            }
        lobjRep.SGenereReporte()
        GobjPanDat.SControleProcesoObj(False)
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Shared Function FstrUltimosAnos() As String()
        Dim lobjAno As ClsAno
        Dim lstrUltAnos() As String = Array.Empty(Of String)
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
    Private Sub SGenereRepFacVivas()
        Dim lstrMens = String.Empty
        GobjPanDat.SControleProcesoObj(True)
        Mouse.OverrideCursor = Cursors.Wait
        If MobjObjetoWin.FblnTieneDeuda(False) Then
            Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
            .DblIdCliente = MobjObjetoWin.ObjIdClienteDbl.ObjValorPro,
            .EnuReporte = EnuReporteDef.enuFacVivas
            }
            lobjRep.SGenereReporte()
        Else
            lstrMens = "Este Cliente no tiene Deuda en mora!"
        End If
        GobjPanDat.SControleProcesoObj(False)
        Mouse.OverrideCursor = Cursors.Arrow
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, "", 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SUbiqueControles()
        If MobjObjetoWin.ObjTipoDocIdentidadByt.ObjValorPro = EnuTipoDocIdDef.enuNit OrElse
                MobjObjetoWin.ObjTipoDocIdentidadByt.ObjValorPro = EnuTipoDocIdDef.enuNuip Then
            'Mostrar Empresa
            grbDatosEmpresas.Visibility = Visibility.Visible
            lblDigitoVerificacion.Visibility = Visibility.Visible
            txtDigVer.Visibility = Visibility.Visible
            ' Esconder Persona
            grbDatosPersonas.Visibility = Visibility.Collapsed
        Else
            'Esconder Empresa
            grbDatosEmpresas.Visibility = Visibility.Collapsed
            lblDigitoVerificacion.Visibility = Visibility.Hidden
            txtDigVer.Visibility = Visibility.Hidden
            'Mostrar Persona
            grbDatosPersonas.Visibility = Visibility.Visible
        End If
    End Sub
    Private Sub SDeshabiliteControles()
        Dim lblnHabilite As Boolean
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            Dim lstrEstiloHabilitado As String
            Dim lstrEstiloNoHabilitado As String
            lblnHabilite = MobjObjetoWin.ObjTipoDocIdentidadByt.ObjValorPro =
                    EnuTipoDocIdDef.enuNit OrElse MobjObjetoWin.ObjTipoDocIdentidadByt.ObjValorPro =
                    EnuTipoDocIdDef.enuNuip
            If lblnHabilite Then
                lstrEstiloHabilitado = "RecCtlHabilitado"
                lstrEstiloNoHabilitado = "RecCtlNoHabilitado"
            Else
                lstrEstiloHabilitado = "RecCtlNoHabilitado"
                lstrEstiloNoHabilitado = "RecCtlHabilitado"
            End If
            txtRazonSocial.Style = FindResource(lstrEstiloHabilitado)
            txtDigVer.Style = FindResource(lstrEstiloHabilitado)
            txtApellido1.Style = FindResource(lstrEstiloNoHabilitado)
            txtApellido2.Style = FindResource(lstrEstiloNoHabilitado)
            txtNombre1.Style = FindResource(lstrEstiloNoHabilitado)
            txtNombre2.Style = FindResource(lstrEstiloNoHabilitado)
            bttAbrirTer.IsEnabled = False
            bttAbrirCuenta.IsEnabled = False
        End If
    End Sub
#End Region

#Region "Registrar entradas"
    Private Sub SRegistreEntrada(aobjControl As Control)
        With MobjObjetoWin
            Select Case aobjControl.Name
                Case "cboTipoTercero"
                    .ObjTipoTerceroByt.ObjValorPro = cboTipoTercero.SelectedIndex
                Case "txtDireccion2"
                    .ObjDireccionDosStr.ObjValorPro = txtDireccion2.Text
                Case "txtDireccion1"
                    .ObjDireccionUnoStr.ObjValorPro = txtDireccion1.Text
                Case "txtPaginaWeb"
                    .ObjPaginaWebStr.ObjValorPro = txtPaginaWeb.Text
                Case "txtEmail"
                    .ObjEmailStr.ObjValorPro = txtEmail.Text
                    .ObjRecibeDocsPorEmailBln.SValide()
                Case "cboMedPagoPref"
                    .ObjIdMedioPagoClienteByt.ObjValorPro = cboMedPagoPref.SelectedIndex
                Case "txtNombre1"
                    .ObjNombrePrimeroStr.ObjValorPro = txtNombre1.Text
                Case "txtNombre2"
                    .ObjNombreSegundoStr.ObjValorPro = txtNombre2.Text
                Case "txtApellido1"
                    .ObjApellidoPrimeroStr.ObjValorPro = txtApellido1.Text
                Case "txtRazonSocial"
                    .ObjRazonSocialStr.ObjValorPro = txtRazonSocial.Text
                Case "txtApellido2"
                    .ObjApellidoSegundoStr.ObjValorPro = txtApellido2.Text
                Case "txtTelefono2"
                    .ObjTelefonoDosStr.ObjValorPro = txtTelefono2.Text
                    MblnIngresoTel = True
                Case "txtMovil"
                    .ObjCelularStr.ObjValorPro = txtMovil.Text
                    MblnIngresoTel = True
                Case "txtMovil2"
                    .ObjCelular2Str.ObjValorPro = txtMovil2.Text
                    MblnIngresoTel = True
                Case "txtTelefono1"
                    .ObjTelefonoUnoStr.ObjValorPro = txtTelefono1.Text
                    MblnIngresoTel = True
            End Select
        End With
    End Sub

    Private Sub SRegistreTipoDoc()
        With MobjObjetoWin
            .ObjTipoDocIdentidadByt.ObjValorPro = cboTipoDoc.SelectedIndex
            If cboTipoDoc.SelectedIndex = EnuTipoDocIdDef.enuNit OrElse
                    cboTipoDoc.SelectedIndex = EnuTipoDocIdDef.enuNuip Then
                .ObjApellidoPrimeroStr.ObjValorPro = String.Empty
                .ObjApellidoSegundoStr.ObjValorPro = String.Empty
                .ObjNombrePrimeroStr.ObjValorPro = String.Empty
                .ObjNombreSegundoStr.ObjValorPro = String.Empty
            End If
            SUbiqueControles()
            .ObjIdTerceroDbl.SbyDigitoVerificacion = txtDigVer.Text
            SValide()
        End With
    End Sub

    Private Sub SRegistreTercero()
        With MobjObjetoWin
            cboTipoDoc.SelectedIndex = .ObjTipoDocIdentidadByt.ObjValorPro
            cboTipoTercero.SelectedIndex = .ObjTipoTerceroByt.ObjValorPro
            txtDigVer.Text = .ObjIdTerceroDbl.SbyDigitoVerificacion
            SValide()
            txtNroDoc.Style = FindResource("RecCtlNoHabilitado")
            txtDigVer.Style = FindResource("RecCtlNoHabilitado")
            cboTipoDoc.Style = FindResource("RecCtlNoHabilitado")
            cboTipoTercero.Style = FindResource("RecCtlNoHabilitado")
            txtRazonSocial.Style = FindResource("RecCtlNoHabilitado")
            txtApellido1.Style = FindResource("RecCtlNoHabilitado")
            txtApellido2.Style = FindResource("RecCtlNoHabilitado")
            txtNombre1.Style = FindResource("RecCtlNoHabilitado")
            txtNombre2.Style = FindResource("RecCtlNoHabilitado")
        End With
    End Sub
#End Region

#Region "Manejo Ubicación"
    Private Sub SPuebleCboPaisDir()
        MblnPoblandoCbo = True
        GobjPanDat.SControleProcesoObj(True)
        If IsNothing(MobjUbicacion) Then
            MobjUbicacion = New ClsUbicacion
        End If
        cboPaisDir.Items.Clear()
        cboPaisDir.Items.Add("<Ninguno>")
        Dim lstrPaises As String() = MobjUbicacion.FstrPaises
        If lstrPaises.Length > 0 Then
            For Each lstrPais In lstrPaises
                cboPaisDir.Items.Add(lstrPais)
            Next
        End If
        cboPaisDir.SelectedIndex = 0
        GobjPanDat.SControleProcesoObj(False)
        MblnPoblandoCbo = False
    End Sub
    Private Sub SPuebleCboDptoDir()
        MblnPoblandoCbo = True
        cboDptoDir.Items.Clear()
        cboCiudadDir.Items.Clear()
        cboDptoDir.Items.Add("<Ninguno>")
        cboCiudadDir.Items.Add("<Ninguno>")
        cboCiudadDir.SelectedIndex = 0
        If MobjUbicacion IsNot Nothing Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                Dim lstrDptos As String() = MobjUbicacion.FstrDptos(FstrIdPaisActual())
                If lstrDptos IsNot Nothing Then
                    If lstrDptos.Length > 0 Then
                        For Each lstrDpto As String In lstrDptos
                            cboDptoDir.Items.Add(lstrDpto)
                        Next
                    End If
                End If
                cboDptoDir.SelectedIndex = 0
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
                If Not lblnNoHayError Then
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
                MblnPoblandoCbo = False
            End Try
        End If
    End Sub
    Private Sub SPuebleCboCiudadDir()
        MblnPoblandoCbo = True
        cboCiudadDir.Items.Clear()
        If MobjUbicacion IsNot Nothing Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                Dim lstrCiudades As String() = MobjUbicacion.FstrCiudades(FstrIdPaisActual(),
                        FBytIdDptoActual())
                cboCiudadDir.Items.Add("<Ninguno>")
                If lstrCiudades.Length > 0 Then
                    For Each lstrCiudad As String In lstrCiudades
                        cboCiudadDir.Items.Add(lstrCiudad)
                    Next
                End If
                cboCiudadDir.SelectedIndex = 0
                MblnPoblandoCbo = False
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
                If Not lblnNoHayError Then
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        Else
            MblnPoblandoCbo = False
        End If
    End Sub
    Private Sub SMuestreUbicacionDir()
        Dim lstrNomDep As String, lstrNomCiu As String
        Dim lstrIdPaisDir = MobjObjetoWin.ObjPaisDirStr.ToString()
        Dim lentIdDpto As Integer = MobjObjetoWin.ObjDepartamentoDirByt.ObjValorPro
        Dim lentIdCiu As Integer = MobjObjetoWin.ObjCiudadDirShr.ObjValorPro
        Dim lstrNomPais As String = MobjUbicacion.StrNombrePais(lstrIdPaisDir)
        If String.IsNullOrEmpty(lstrNomPais) Then
            lstrNomPais = "<Ninguno>"
        End If
        lstrNomDep = MobjUbicacion.StrNombreDpto(lstrIdPaisDir, lentIdDpto)
        lstrNomCiu = MobjUbicacion.StrNombreCiudad(lstrIdPaisDir, lentIdDpto, lentIdCiu)
        cboPaisDir.SelectedItem = lstrNomPais
        SPuebleCboDptoDir()
        cboDptoDir.SelectedItem = lstrNomDep
        SPuebleCboCiudadDir()
        cboCiudadDir.SelectedItem = lstrNomCiu
    End Sub
    Private Function FstrIdPaisActual() As String
        Dim lstrIdPais As String, lstrNombrePais = GCSTRNINGUNO
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            If Not IsNothing(cboPaisDir.SelectedItem) Then
                lstrNombrePais = cboPaisDir.SelectedItem
            End If
        Else
            lstrNombrePais = cboPaisDir.SelectedItem
        End If
        lstrIdPais = MobjUbicacion.StrIdPais(lstrNombrePais)
        Return lstrIdPais
    End Function
    Private Function FBytIdDptoActual() As Byte
        Dim lbytIdDptoDir As Byte, lstrNomDpto = String.Empty
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            If Not IsNothing(cboDptoDir.SelectedItem) Then
                lstrNomDpto = cboDptoDir.SelectedItem
            End If
        Else
            lstrNomDpto = cboDptoDir.SelectedItem
        End If
        lbytIdDptoDir = MobjUbicacion.BytIdDpto(FstrIdPaisActual(), lstrNomDpto)
        Return lbytIdDptoDir
    End Function
    Private Function FshrIdCiudadActual() As Short
        Dim lshrIdCiu As Short, lstrNomCiu = String.Empty
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            If Not IsNothing(cboCiudadDir.SelectedItem) Then
                lstrNomCiu = cboCiudadDir.SelectedItem
            End If
        Else
            lstrNomCiu = cboCiudadDir.SelectedItem
        End If
        lshrIdCiu = MobjUbicacion.ShrIdCiudad(FstrIdPaisActual(), FBytIdDptoActual(),
                    lstrNomCiu)
        Return lshrIdCiu
    End Function
#End Region

#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            If lelmElemento.Name = "bttAbrirTer" OrElse lelmElemento.Name = "bttTercero" Then
                SAbraTercero()
            ElseIf lelmElemento.Name = "bttAbrirCuenta" Then
                SAbraCuenta()
            End If
        End If
    End Sub

    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Dim lmnuItemMenu As MenuItem
        If TypeOf lelmElemento Is MenuItem Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                lmnuItemMenu = lelmElemento
                Dim lstrNombreMenu = lmnuItemMenu.Name
                Select Case lstrNombreMenu
                    Case "MnuImportarClientes"
                        SAbraImportar()
                    Case "MnuRepDirClientes"
                        SGenereDirClientes()
                    Case "MnuRepFacturas"
                        SGenereRepFacVivas()
                    Case "MnuCuentaCliente"
                        SAbraCuenta()
                    Case "MnuReasignarId"
                        Dim lobjWin As New WinReasignaIdTercero(
                                MobjObjetoWin.ObjTerceroCliente) With {
                        .WinPadre = Me
                        }
                        lobjWin.ShowDialog()
                        SRefrescarClic()
                    Case Else
                        If lmnuItemMenu.Name.StartsWith("MnuAno") Then
                            SGenereRepVlrsFacturados(Right(lmnuItemMenu.Name, 4))
                            If String.IsNullOrEmpty(lstrMens) Then
                                lstrMens = "Reporte generado exitosamente!"
                            End If
                        End If
                End Select
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

    Private Sub MnuContextual_Click(sender As Object, e As RoutedEventArgs)
        If sender.Equals(MnuAbrirPredioConC) Then
            If IsNothing(lsbPrediosDelCliente.SelectedItem) OrElse
                    String.IsNullOrEmpty(lsbPrediosDelCliente.SelectedItem.ToString) Then
                Dim lstrMens = "Debe seleccionar un Predio!"
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            Else
                Dim lstrPredio = lsbPrediosDelCliente.SelectedItem.ToString
                SAbraPredio(lstrPredio)
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
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox OrElse TypeOf lelmElemento Is ComboBox OrElse
                TypeOf lelmElemento Is DatePicker AndAlso Not HblnSeEstaCerrando Then
            GobjPanDat.SControleProcesoObj(True)
            If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando AndAlso
                    Not HblnMostrandoDatos Then
                SRegistreEntrada(lelmElemento)
                Select Case lelmElemento.Name
                    Case "txtNroDoc"
                        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
                            MobjObjetoWin.ObjIdClienteDbl.ObjValorPro = txtNroDoc.Text
                            If MobjObjetoWin.ObjTerceroCliente.BlnExiste Then
                                SRegistreTercero()
                            Else
                                MobjObjetoWin.ObjTerceroCliente.SVacie()
                                MobjObjetoWin.ObjIdClienteDbl.ObjValorPro = txtNroDoc.Text
                            End If
                            SMuestreDatos()
                        End If
                    Case "txtDigVer"
                        MobjObjetoWin.ObjIdTerceroDbl.SbyDigitoVerificacion = txtDigVer.Text
                    Case "txtComentario"
                        MobjObjetoWin.ObjComentario_ClienteStr.ObjValorPro = txtComentario.Text
                    Case "txtCodPostal"
                        MobjObjetoWin.ObjCodigoPostalStr.ObjValorPro = txtCodPostal.Text
                End Select
            End If
            SValide()
            GobjPanDat.SControleProcesoObj(False)
        End If
    End Sub

    Private Sub Chk_Click(sender As Object, e As RoutedEventArgs) Handles chkAgenteRetFte.Click,
            chkRetieneIva.Click, chkAgenteRetIca.Click, chkFacturarSer.Click, chkRecibeDocsEmail.Click,
            chkAutoRet.Click, chkGranCont.Click, chkRegSimpleTri.Click
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando AndAlso
                Not HblnMostrandoDatos Then
            If TypeOf lelmElemento Is CheckBox Then
                Dim lchkCB As CheckBox = lelmElemento
                Select Case lelmElemento.Name
                    Case "chkAutoRet"
                        MobjObjetoWin.ObjEsAutorretenedorBln.ObjValorPro = lchkCB.IsChecked
                    Case "chkGranCont"
                        MobjObjetoWin.ObjEsGranContrBln.ObjValorPro = lchkCB.IsChecked
                    Case "chkRegSimpleTri"
                        MobjObjetoWin.ObjEsRegimenSimpleTBln.ObjValorPro = lchkCB.IsChecked
                    Case "chkAgenteRetFte"
                        MobjObjetoWin.ObjEsAgenteReteFteBln.ObjValorPro = lchkCB.IsChecked
                    Case "chkRetieneIva"
                        MobjObjetoWin.ObjRetieneIvaBln.ObjValorPro = lchkCB.IsChecked
                    Case "chkAgenteRetIca"
                        MobjObjetoWin.ObjRetieneIcaBln.ObjValorPro = lchkCB.IsChecked
                    Case "chkFacturarSer"
                        MobjObjetoWin.ObjFactPorServicio_CliBln.ObjValorPro = lchkCB.IsChecked
                    Case "chkRecibeDocsEmail"
                        MobjObjetoWin.ObjRecibeDocsPorEmailBln.ObjValorPro = chkRecibeDocsEmail.IsChecked
                End Select
                SMuestreDatos()
            End If
        End If
    End Sub

    Private Sub OnCboCambio(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            If TypeOf lelmElemento Is ComboBox Then
                If Not MblnPoblandoCbo Then
                    GobjPanDat.SControleProcesoObj(True)
                    Select Case lelmElemento.Name
                        Case "cboPaisDir"
                            MobjObjetoWin.ObjPaisDirStr.ObjValorPro = FstrIdPaisActual()
                            MobjObjetoWin.ObjDepartamentoDirByt.ObjValorPro = FBytIdDptoActual()
                            MobjObjetoWin.ObjCiudadDirShr.ObjValorPro = FshrIdCiudadActual()
                            SPuebleCboDptoDir()
                            SMuestreUbicacionDir()
                        Case "cboDptoDir"
                            MobjObjetoWin.ObjDepartamentoDirByt.ObjValorPro = FBytIdDptoActual()
                            MobjObjetoWin.ObjCiudadDirShr.ObjValorPro = FshrIdCiudadActual()
                            SPuebleCboCiudadDir()
                        Case "cboCiudadDir"
                            MobjObjetoWin.ObjCiudadDirShr.ObjValorPro = FshrIdCiudadActual()
                        Case "cboMedPagoPref"
                            MobjObjetoWin.ObjIdMedioPagoClienteByt.ObjValorPro = cboMedPagoPref.SelectedIndex
                        Case "cboRegimenVentas"
                            MobjObjetoWin.ObjIdRegimenVentasByt.ObjValorPro = cboRegimenVentas.SelectedIndex
                        Case "cboEstadoActualDeuda"
                            MobjObjetoWin.ObjIdEstadoDeudaByt.ObjValorPro = cboEstadoActualDeuda.SelectedIndex
                        Case "cboTipoDoc"
                            SRegistreTipoDoc()
                            SDeshabiliteControles()
                    End Select
                    GobjPanDat.SControleProcesoObj(False)
                End If
            End If
        End If
    End Sub

    Private Sub Txt_KeyDown(sender As Object, e As KeyEventArgs) Handles txtNroDoc.KeyDown
        If e.Key = Key.Return OrElse e.Key = Key.Tab Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando AndAlso
                    MobjObjetoWin.ObjIdClienteDbl.ToString() <> txtNroDoc.Text Then
                SAbraCliente()
            End If
        End If
    End Sub

    Private Sub Lbl_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles _
            lblOutlook.MouseDoubleClick, lblGmail.MouseDoubleClick, lblHotmail.MouseDoubleClick,
            lblYahoo.MouseDoubleClick
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            If TypeOf lelmElemento Is Label Then
                Select Case lelmElemento.Name
                    Case "lblOutlook"
                        txtEmail.Text &= lblOutlook.Content
                    Case "lblGmail"
                        txtEmail.Text &= lblGmail.Content
                    Case "lblHotmail"
                        txtEmail.Text &= lblHotmail.Content
                    Case "lblYahoo"
                        txtEmail.Text &= lblYahoo.Content
                End Select
            End If
        End If
    End Sub

    Protected Sub EwinVentanaClosed(sender As Object, e As EventArgs) Handles MwinVentana.Closed
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            MobjObjetoWin.ObjIdClienteDbl.ObjValorPro = txtNroDoc.Text
            SMuestreDatos()
        End If
    End Sub

    Private Sub TxtTel_TextChanged(sender As Object, e As TextChangedEventArgs) Handles _
                txtMovil.TextChanged, txtMovil2.TextChanged, txtTelefono1.TextChanged,
                txtTelefono2.TextChanged
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            If lelmElemento.Name.StartsWith("txtTel") OrElse
                    lelmElemento.Name.StartsWith("txtMov") Then
                Dim ltxtTextBox As TextBox = lelmElemento
                Dim lstrTexto As String = ltxtTextBox.Text
                Dim lstrTextoNumeros As String = String.Empty
                For Each lchrCaracter As Char In lstrTexto
                    If Char.IsDigit(lchrCaracter) Then
                        lstrTextoNumeros &= lchrCaracter
                    End If
                Next
                If lstrTexto <> lstrTextoNumeros Then
                    Dim lstrMens = "Sólo se permiten números en este campo!"
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    ltxtTextBox.Text = lstrTextoNumeros
                    ltxtTextBox.SelectionStart = ltxtTextBox.Text.Length
                End If
            End If
        End If
    End Sub
#End Region
End Class