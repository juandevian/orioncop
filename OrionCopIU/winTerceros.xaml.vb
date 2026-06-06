Imports Microsoft.Win32
Imports System.IO
Public Class WinTerceros
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuApellido1 = 0
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
        enuNomDpto
        enuNomCiu
        enuCodPostal
    End Enum

    Private Enum EnuIAFIdAccionFormaDef As Integer
        enuIAFReasignarId = 1
        enuIAFFotos
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsTercero = Nothing
    Private WithEvents MwinVentana As WinReasignaIdTercero = Nothing
    ' Manejo Ubicacion
    Private MobjUbicacion As Ubicacion.ClsUbicacion
    Private mblnPoblandoCbo As Boolean = False
    ' Opciones menu Fotos
    Private MmnuImportar As MenuItemPan = Nothing
    Private MmnuReasignarId As MenuItemPan = Nothing
    Private MmnuAsociarFoto As MenuItemPan = Nothing
    Private MmnuSuprimirFoto As MenuItemPan = Nothing
    Private MmnuRotarFoto As MenuItem = Nothing
    Private MblnFotoImportada As Boolean = False
    Private MimgFoto As System.Drawing.Image = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomTercer
#End Region
#Region "Constructor"
    Public Sub New(aobjAplicacionL As Object)
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuTerceros
        HobjAplicacionL = aobjAplicacionL
    End Sub
#End Region
#Region "Invalida metodos de la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolControlesLlave As New Collection From {
            txtNroDoc
        }
        SCargueForma(EnuElementosAdicionalesDef.None, 22,
                lcolControlesLlave, cboTipoDoc, False)
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
            ObjObjetoWin = New ClsTercero(EnuModoInstanciaObjDef.enuNavegable)
            If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                ObjObjetoWin.SVayaAlPrimero()
            End If
        End If
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
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
        StcValidaControl(EnuValidEntrada.enuNomDpto) = lblDptoDir
        StcValidaControl(EnuValidEntrada.enuNomCiu) = lblCiudadDir
        StcValidaControl(EnuValidEntrada.enuCodPostal) = lblCodPostal
        '
        SPuebleComboBoxes()
        SPuebleCboPaisDir()
        SSeleccioneControles()
        HbttAceptar.TabIndex = 100
        HbttCancelar.TabIndex = 101
    End Sub
    Protected Overrides Sub SMuestreDatos()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
            SLevanteEveNoti("No hay Terceros para ser mostradas!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
            txtNroDoc.IsEnabled = False
        End If
        With MobjObjetoWin
            txtApellido1.Text = .ObjApellidoPrimeroStr.ObjValorPro
            txtApellido2.Text = .ObjApellidoSegundoStr.ObjValorPro
            txtDireccion1.Text = .ObjDireccionUnoStr.ObjValorPro
            txtDireccion2.Text = .ObjDireccionDosStr.ObjValorPro
            txtPaginaWeb.Text = .ObjPaginaWebStr.ObjValorPro
            txtEmail.Text = .ObjEmailStr.ObjValorPro
            txtCodPostal.Text = .ObjCodigoPostalStr.ToString
            txtFechaIngreso.Content = .ObjFechaCreacionDtm.ObjValorPro
            cboTipoDoc.SelectedIndex = .ObjTipoDocIdentidadByt.ObjValorPro
            cboTipoTercero.SelectedIndex = .ObjTipoTerceroByt.ObjValorPro
            txtNroDoc.Text = Format(.ObjIdTerceroDbl.ObjValorPro, GCSTRFMTIDTERCERO)
            If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
                txtDigVer.Text = .ObjIdTerceroDbl.SbyDigitoVerificacion
                If chkMostrarFoto.IsChecked Then
                    SCargueFoto()
                End If
            End If
            lblNombreCompleto.Content = .FstrNombreCompleto()
            txtNombre1.Text = .ObjNombrePrimeroStr.ObjValorPro
            txtNombre2.Text = .ObjNombreSegundoStr.ObjValorPro
            txtRazonSocial.Text = .ObjRazonSocialStr.ObjValorPro
            txtTelefono1.Text = .ObjTelefonoUnoStr.ObjValorPro
            txtTelefono2.Text = .ObjTelefonoDosStr.ObjValorPro
            txtMovil.Text = .ObjCelularStr.ObjValorPro
            txtMovil2.Text = .ObjCelular2Str.ObjValorPro
        End With
        SMuestreUbicaciones()
        SUbiqueControles()
        SDeshabiliteControles()
        SDeshabiliteMenuesFoto()
        Title = My.Resources.Tercero & My.Resources.DosPuntosEspacio
        If Not String.IsNullOrEmpty(txtNroDoc.Text) Then
            Title &= txtNroDoc.Text
            If txtDigVer.Text <> "-1" Then
                Title &= My.Resources.GuionSeparador & txtDigVer.Text
            End If
            Title &= " " & MobjObjetoWin.FstrNombreCompleto()
        End If
        SValide()
        If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
            txtNroDoc.SelectAll()
        End If
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentanaDef.cenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(aintIndice:=EnuValidEntrada.enuApellido1) = .ObjApellidoPrimeroStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuApellido2) = .ObjApellidoSegundoStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuDirec1) = .ObjDireccionUnoStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuDirec2) = .ObjDireccionDosStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuPagWeb) = .ObjPaginaWebStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuEmail) = .ObjEmailStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuTelMovil) = .ObjCelularStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuTelMovil2) = .ObjCelular2Str.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuTel1) = .ObjTelefonoUnoStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuTel2) = .ObjTelefonoDosStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuTipoIdent) = .ObjTipoDocIdentidadByt.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuTipoTercero) = .ObjTipoTerceroByt.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuIdTercero) = .ObjIdTerceroDbl.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuNom1) = .ObjNombrePrimeroStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuNom2) = .ObjNombreSegundoStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuRazonSocial) = .ObjRazonSocialStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuPaisDir) = .ObjPaisDirStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuDptoDir) = .ObjDepartamentoDirByt.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuCiudadDir) = .ObjCiudadDirShr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuNomDpto) = .ObjNombreDptoDirStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuNomCiu) = .ObjNombreCiuDirStr.BlnEsValido
                StcValidValido(aintIndice:=EnuValidEntrada.enuCodPostal) = .ObjCodigoPostalStr.BlnEsValido
            End With
        End If
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjDireccionDosStr.ObjValorPro = txtDireccion2.Text
            .ObjDireccionUnoStr.ObjValorPro = txtDireccion1.Text
            .ObjPaginaWebStr.ObjValorPro = txtPaginaWeb.Text
            .ObjCodigoPostalStr.ObjValorPro = txtCodPostal.Text
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
            .ObjNombreDptoDirStr.ObjValorPro = txtDptoDir.Text
            .ObjNombreCiuDirStr.ObjValorPro = txtCiudadDir.Text
        End With
        SValide()
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        Dim lmnuOpcionesFoto As MenuItem = FindResource("RecMnuOpcFoto")
        HmnuMiMenu.Items.Insert(1, lmnuOpcionesFoto)
        MmnuAsociarFoto = FmnuiMenuItemPan("MnuAsociarFoto", "A_sociar Fotografía", 3, "")
        lmnuOpcionesFoto.Items.Insert(0, MmnuAsociarFoto)
        MmnuSuprimirFoto = FmnuiMenuItemPan("MnuSuprimirFoto", "S_uprimir Fotografía", 4, "")
        lmnuOpcionesFoto.Items.Insert(1, MmnuSuprimirFoto)
        MmnuRotarFoto = FmnuiMenuItem("MnuRotarFoto", "_Rotar Fotografía", "RecMnuItemSec")
        lmnuOpcionesFoto.Items.Insert(3, MmnuRotarFoto)
        ' Adicionar menú Importar
        MmnuImportar = FmnuiMenuItemPan("MnuImportar", "_Importar Terceros", 1, "")
        HmnuAcciones.Items.Insert(7, MmnuImportar)
        ' Adicionar menú Reasignar Id tercero
        MmnuReasignarId = FmnuiMenuItemPan("MnuReasignarId", "Reasi_gnar Id. Tercero", 2, "")
        HmnuAcciones.Items.Insert(8, MmnuReasignarId)
        ' Adicionar Separador
        HmnuAcciones.Items.Insert(9, New Separator)
    End Sub
#End Region
#Region "Invalida otros metodos de la clase base"
    Protected Overrides Sub SCree()
        MyBase.SCree()
        SSeleccioneControles()
        txtNroDoc.IsEnabled = True
        txtNroDoc.Focus()
    End Sub
    Protected Overrides Sub SModifique()
        MyBase.SModifique()
        SMuestreDatos()
        SSeleccioneControles()
        SDeshabiliteControles()
        SDeshabiliteMenuesFoto()
        cboTipoDoc.Style = FindResource("RecCtlNoHabilitado")
    End Sub
    Protected Overrides Sub SGuarde()
        SRegistre()
        If FblnEstanTodosBien() Then
            If MblnFotoImportada Then
                MobjObjetoWin.SAdicioneImagen(MimgFoto)
                MblnFotoImportada = False
            End If
            If ObjObjetoWin.BlnTengoCambios Then
                MobjObjetoWin.SActualice(True)
            End If
            If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuCreando Then
                Dim lstrMens = FstrNombreDoc() & " fue creado exitosamente!"
                SFinaliceOperacion()
                SCrearClic()
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            Else
                SFinaliceOperacion()
            End If
            If Not (IsNothing(WinPadre) OrElse WinPadre.EnuIdVentana = EnuIdVentanaDef.enuParametrizacion) Then
                If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
                    SCerrarClic()
                End If
            End If
        End If
    End Sub
    Protected Overrides Sub SFinaliceOperacion()
        MyBase.SFinaliceOperacion()
        SSeleccioneControles()
    End Sub
    Protected Overrides Sub SHabiliteMenues()
        MyBase.SHabiliteMenues()
        Dim lblnHabilite As Boolean
        If MmnuAsociarFoto.IsEnabled Then
            lblnHabilite = GobjPanorama.ObjUsuarioActual.FblnTienePermiso(EnuIdClasesPanDef.enuTercero, EnuIdAccionDef.enuCrear)
            SHabiliteMenuItemPan(lblnHabilite, MmnuAsociarFoto)
        End If
        If MmnuReasignarId.IsEnabled Then
            lblnHabilite = GobjPanorama.ObjUsuarioActual.FblnTienePermiso(EnuIdClasesPanDef.enuTercero, EnuIdAccionDef.enuModificar)
            SHabiliteMenuItemPan(lblnHabilite, MmnuReasignarId)
        End If
        If MmnuImportar.IsEnabled Then
            lblnHabilite = GobjPanorama.ObjUsuarioActual.FblnTienePermiso(EnuIdClasesPanDef.enuTercero, EnuIdAccionDef.enuCrear)
            SHabiliteMenuItemPan(lblnHabilite, MmnuImportar)
        End If
        MnuAsociarFotoC.IsEnabled = MmnuAsociarFoto.IsEnabled
        MnuSuprimirFotoC.IsEnabled = MmnuSuprimirFoto.IsEnabled
    End Sub
#End Region
#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        MyBase.SBuscar()
        If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
            If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable Then
                txtNroDoc.Text = StrResultadoBusqueda
                SAbraTercero()
            End If
        End If
    End Sub
    ''' <summary>
    ''' Invalida la funcion "fblnDefinioBusqueda" de la clase base.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        SDefineBusquedaRazonSocial()
        SDefineBusquedaPrimerApellido()
        SDefineBusquedaPrimerNombre()
        Return True
    End Function
    Private Sub SDefineBusquedaRazonSocial()
        Dim lstrTabla As String = ClsTercero.SstrNombreTabla
        Dim lstrCamposMostrar As String() = {ClsIdTerceroDbl.SstrNombreCampoBd,
                                             ClsRazonSocialStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsRazonSocialStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdTerceroDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Razon Social", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
    Private Sub SDefineBusquedaPrimerApellido()
        Dim lstrTabla As String = ClsTercero.SstrNombreTabla
        Dim lstrCamposMostrar As String() = {ClsIdTerceroDbl.SstrNombreCampoBd,
                                             ClsNombrePrimeroStr.SstrNombreCampoBd,
                                             ClsNombreSegundoStr.SstrNombreCampoBd,
                                             ClsApellidoPrimeroStr.SstrNombreCampoBd,
                                             ClsApellidoSegundoStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsApellidoPrimeroStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdTerceroDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Primer Apellido", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
    Private Sub SDefineBusquedaPrimerNombre()
        Dim lstrTabla As String = ClsTercero.SstrNombreTabla
        Dim lstrCamposMostrar As String() = {ClsIdTerceroDbl.SstrNombreCampoBd,
                                             ClsNombrePrimeroStr.SstrNombreCampoBd,
                                             ClsNombreSegundoStr.SstrNombreCampoBd,
                                             ClsApellidoPrimeroStr.SstrNombreCampoBd,
                                             ClsApellidoSegundoStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombrePrimeroStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdTerceroDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Primer Nombre", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SPuebleComboBoxes()
        Dim ldrwDataRow = ClsAdministrador.FdrwConstantesPan(EnuGrupoConstantesPanDef.enuTipoDocIdentidad)
        SPuebleComboBox(ldrwDataRow, cboTipoDoc)
        ldrwDataRow = ClsAdministrador.FdrwConstantesPan(EnuGrupoConstantesPanDef.enuTipoTercero)
        SPuebleComboBox(ldrwDataRow, cboTipoTercero)
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
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
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
        End If
    End Sub
    Private Sub SAbraTercero()
        If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                If txtNroDoc.Text <> MobjObjetoWin.ObjIdTerceroDbl.ToString() Then
                    Dim lobjVlrLlave = {txtNroDoc.Text}
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
    Private Sub SRegistreEntrada(aobjControl As Control)
        SRegistreEntrada_Adi(aobjControl)
        With MobjObjetoWin
            Select Case aobjControl.Name
                Case "cboTipoDoc"
                    SRegistreTipoDoc()
                Case "cboTipoTercero"
                    .ObjTipoTerceroByt.ObjValorPro = cboTipoTercero.SelectedIndex
                Case "txtNroDoc"
                    .ObjIdTerceroDbl.ObjValorPro = txtNroDoc.Text
                Case "txtRazonSocial"
                    .ObjRazonSocialStr.ObjValorPro = txtRazonSocial.Text.Trim
                Case "txtDireccion1"
                    .ObjDireccionUnoStr.ObjValorPro = txtDireccion1.Text.Trim
                Case "txtDireccion2"
                    .ObjDireccionDosStr.ObjValorPro = txtDireccion2.Text.Trim
                Case "txtPaginaWeb"
                    .ObjPaginaWebStr.ObjValorPro = txtPaginaWeb.Text.Trim
                Case "txtCodPostal"
                    .ObjCodigoPostalStr.ObjValorPro = txtCodPostal.Text.Trim
                Case "txtEmail"
                    .ObjEmailStr.ObjValorPro = txtEmail.Text.Trim
            End Select
        End With
        If aobjControl.Name <> "txtNroDoc" Then
            SMuestreDatos()
        End If
    End Sub
    Private Sub SRegistreEntrada_Adi(aobjControl As Control)
        With MobjObjetoWin
            Select Case aobjControl.Name
                Case "txtCiudadDir"
                    .ObjNombreCiuDirStr.ObjValorPro = txtCiudadDir.Text
                Case "txtDptoDir"
                    MobjObjetoWin.ObjNombreDptoDirStr.ObjValorPro = txtDptoDir.Text
                Case "txtDigVer"
                    .ObjIdTerceroDbl.SbyDigitoVerificacion = txtDigVer.Text
                Case "txtNombre1"
                    .ObjNombrePrimeroStr.ObjValorPro = txtNombre1.Text
                Case "txtNombre2"
                    .ObjNombreSegundoStr.ObjValorPro = txtNombre2.Text
                Case "txtApellido1"
                    .ObjApellidoPrimeroStr.ObjValorPro = txtApellido1.Text
                Case "txtApellido2"
                    .ObjApellidoSegundoStr.ObjValorPro = txtApellido2.Text
                Case "txtTelefono2"
                    .ObjTelefonoDosStr.ObjValorPro = txtTelefono2.Text
                Case "txtMovil"
                    .ObjCelularStr.ObjValorPro = txtMovil.Text
                Case "txtMovil2"
                    .ObjCelular2Str.ObjValorPro = txtMovil2.Text
                Case "txtTelefono1"
                    .ObjTelefonoUnoStr.ObjValorPro = txtTelefono1.Text
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
            Else
                .ObjRazonSocialStr.ObjValorPro = String.Empty
            End If
        End With
    End Sub
    Private Sub SEjecuteAccionCheckBox(achkCheckBox As CheckBox)
        Dim lstrNombre = achkCheckBox.Name
        If lstrNombre = "chkMostrarFoto" Then
            If chkMostrarFoto.IsChecked Then
                If IsNothing(imgFoto.Source) Then
                    imgFoto.Visibility = Visibility.Visible
                    SCargueFoto()
                Else
                    imgFoto.Visibility = Visibility.Visible
                End If
            Else
                imgFoto.Visibility = Visibility.Hidden
            End If
            SDeshabiliteMenuesFoto()
        End If
    End Sub
    Private Sub SAbraImportar()
        Dim lwinImportar As New WinImportar(MobjObjetoWin) With {
            .WinPadre = Me
        }
        lwinImportar.ShowDialog()
        SRefrescarClic()
    End Sub
#Region "Manejo Fotos"
    Private Sub SDeshabiliteMenuesFoto()
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
            SHabiliteMnuAsociarFoto(True)
            SHabiliteMnuSuprimirFoto(False)
            SHabiliteMenuItemPan(False, MmnuRotarFoto)
            MnuRotarFotoC.IsEnabled = False
        Else
            chkMostrarFoto.IsEnabled = True
            SHabiliteMnuAsociarFoto(False)
            If Not IsNothing(imgFoto.Source) Then
                Dim lblnHabilite As Boolean = chkMostrarFoto.IsChecked
                SHabiliteMnuSuprimirFoto(lblnHabilite)
                SHabiliteMenuItemPan(lblnHabilite, MmnuRotarFoto)
                MnuRotarFotoC.IsEnabled = MmnuRotarFoto.IsEnabled
            Else
                SHabiliteMnuSuprimirFoto(False)
                SHabiliteMenuItemPan(False, MmnuRotarFoto)
                MnuRotarFotoC.IsEnabled = MmnuRotarFoto.IsEnabled
            End If
        End If
    End Sub
    Private Sub SHabiliteMnuAsociarFoto(ablnHabilte As Boolean)
        If ablnHabilte Then
            If FblnHabilitarMenuPan(MmnuAsociarFoto.EntIdAccion) Then
                SHabiliteMenuItemPan(True, MmnuAsociarFoto)
            End If
        Else
            If MmnuAsociarFoto.IsEnabled Then
                SHabiliteMenuItemPan(False, MmnuAsociarFoto)
            End If
        End If
        MnuAsociarFotoC.IsEnabled = MmnuAsociarFoto.IsEnabled
    End Sub
    Private Sub SHabiliteMnuSuprimirFoto(ablnHabilte As Boolean)
        If ablnHabilte Then
            If FblnHabilitarMenuPan(MmnuSuprimirFoto.EntIdAccion) Then
                SHabiliteMenuItemPan(True, MmnuSuprimirFoto)
            End If
        Else
            If MmnuSuprimirFoto.IsEnabled Then
                SHabiliteMenuItemPan(False, MmnuSuprimirFoto)
            End If
        End If
        MnuSuprimirFotoC.IsEnabled = MmnuSuprimirFoto.IsEnabled
    End Sub
    Private Sub SImporteFoto()
        Dim lofdFoto As New OpenFileDialog With {
            .DefaultExt = ".jpg",
            .Filter = My.Resources.TipoArchivoImagen
        }
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Dim lblnOk As Boolean = lofdFoto.ShowDialog
            If lblnOk Then
                ' Obtener el objeto Image de la Foto
                Dim lstmFoto As Stream = lofdFoto.OpenFile
                Dim llngTamano As Long = lstmFoto.Length
                If llngTamano > 70000 Then
                    lstrMens = "El tamaño máximo permitido de la foto es de 68K. " &
                            "Debe importar una foto más pequeña!"
                End If
                MimgFoto = Bitmap.FromStream(lstmFoto)
                ' Asigno la imagen al control imgFoto de la ventana
                Dim lstrTray As String = lofdFoto.FileName
                Dim lbimFoto As New BitmapImage
                lbimFoto.BeginInit()
                lbimFoto.UriSource = New Uri(lstrTray)
                lbimFoto.DecodePixelWidth = 135
                lbimFoto.EndInit()
                imgFoto.Source = lbimFoto
                lblnNoHayError = True
            End If
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            MblnFotoImportada = lblnNoHayError
        End Try
    End Sub
    Private Sub SCargueFoto()
        Dim lmstImagenGuardada As MemoryStream = MobjObjetoWin.FmstImagenAnterior
        If Not IsNothing(lmstImagenGuardada) Then
            Dim lbimFoto As New BitmapImage
            lbimFoto.BeginInit()
            lbimFoto.StreamSource = lmstImagenGuardada
            lbimFoto.EndInit()
            imgFoto.Source = lbimFoto
            lblFechaFoto.Content = "Fecha Foto: " & Format(MobjObjetoWin.DtmFechaFoto, GCSTRFMTFECHASIMPLE)
        Else
            imgFoto.Source = Nothing
        End If
    End Sub
    Private Sub SSuprimaFoto()
        Dim lstrMens = String.Empty
        Dim lblnSuprimio = MobjObjetoWin.FblnSuprimioFoto()
        SCargueFoto()
        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
    End Sub
#End Region
#Region "Manejo Ubicación"
    Private Sub SPuebleCboPaisDir()
        If IsNothing(MobjUbicacion) Then
            MobjUbicacion = New Ubicacion.ClsUbicacion
        End If
        mblnPoblandoCbo = True
        cboPaisDir.Items.Clear()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Dim lstrPaises As String() = MobjUbicacion.FstrPaises
            cboPaisDir.Items.Add("<Ninguno>")
            If lstrPaises.Length > 0 Then
                For Each lstrPais In lstrPaises
                    cboPaisDir.Items.Add(lstrPais)
                Next
            End If
            cboPaisDir.SelectedIndex = 0
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
            mblnPoblandoCbo = False
        End Try
    End Sub
    Private Sub SPuebleCboDptoDir()
        mblnPoblandoCbo = True
        cboDptoDir.Items.Clear()
        cboCiudadDir.Items.Clear()
        cboDptoDir.Items.Add("<Ninguno>")
        cboCiudadDir.Items.Add("<Ninguno>")
        cboCiudadDir.SelectedIndex = 0
        If FstrIdPaisActual() = "CO" Then
            If Not IsNothing(MobjUbicacion) Then
                Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                Try
                    Dim lstrDptos As String() = MobjUbicacion.FstrDptos(FstrIdPaisActual())
                    If Not IsNothing(lstrDptos) Then
                        If lstrDptos.Length > 0 Then
                            For Each lstrDpto As String In lstrDptos
                                cboDptoDir.Items.Add(lstrDpto)
                            Next
                        End If
                    End If
                    cboDptoDir.SelectedIndex = 0
                    mblnPoblandoCbo = False
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
            End If
        Else
            mblnPoblandoCbo = False
        End If
    End Sub
    Private Sub SPuebleCboCiudadDir()
        mblnPoblandoCbo = True
        cboCiudadDir.Items.Clear()
        If Not IsNothing(MobjUbicacion) Then
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
                mblnPoblandoCbo = False
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
            mblnPoblandoCbo = False
        End If
    End Sub
    Private Sub SMuestreUbicaciones()
        If tbiLocaliza.IsSelected Then
            SMuestreUbicacionDir()
        End If
    End Sub
    Private Sub SMuestreUbicacionDir()
        Dim lstrNomDep As String, lstrNomCiu As String
        Dim lstrIdPaisDir = MobjObjetoWin.ObjPaisDirStr.ObjValorPro
        Dim lentIdDpto As Integer = MobjObjetoWin.ObjDepartamentoDirByt.ObjValorPro
        Dim lentIdCiu As Integer = MobjObjetoWin.ObjCiudadDirShr.ObjValorPro
        Dim lstrNomPais As String = MobjUbicacion.StrNombrePais(lstrIdPaisDir)
        If lstrIdPaisDir = "CO" Then
            lstrNomDep = MobjUbicacion.StrNombreDpto(lstrIdPaisDir, lentIdDpto)
            lstrNomCiu = MobjUbicacion.StrNombreCiudad(lstrIdPaisDir, lentIdDpto, lentIdCiu)
        Else
            lstrNomDep = MobjObjetoWin.ObjNombreDptoDirStr.ObjValorPro
            lstrNomCiu = MobjObjetoWin.ObjNombreCiuDirStr.ObjValorPro
        End If
        If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
            txtPaisDir.Text = lstrNomPais
            txtDptoDir.Text = lstrNomDep
            txtCiudadDir.Text = lstrNomCiu
        Else
            cboPaisDir.SelectedItem = lstrNomPais
            If lstrIdPaisDir <> "CO" Then
                txtDptoDir.Text = MobjObjetoWin.ObjNombreDptoDirStr.ObjValorPro
                txtCiudadDir.Text = MobjObjetoWin.ObjNombreCiuDirStr.ObjValorPro
            Else
                txtDptoDir.Text = String.Empty
                txtCiudadDir.Text = String.Empty
                cboDptoDir.SelectedItem = lstrNomDep
                cboCiudadDir.SelectedItem = lstrNomCiu
            End If
        End If
    End Sub
    Private Sub SSeleccioneControles()
        If EnuOperacionEnWin = OriWin.EnuOperacionEnVentanaDef.cenuConsultando Then
            SSeleccioneCtrlsConsulta()
        Else
            SSeleccioneCtrlsModiYCrear()
        End If
    End Sub
    Private Sub SSeleccioneCtrlsConsulta()
        cboPaisDir.Visibility = Visibility.Collapsed
        cboDptoDir.Visibility = Visibility.Collapsed
        cboCiudadDir.Visibility = Visibility.Collapsed
        '
        txtPaisDir.Visibility = Visibility.Visible
        txtDptoDir.Visibility = Visibility.Visible
        txtCiudadDir.Visibility = Visibility.Visible
    End Sub
    Private Sub SSeleccioneCtrlsModiYCrear()
        cboPaisDir.Visibility = Visibility.Visible
        txtPaisDir.Visibility = Visibility.Collapsed
        If MobjObjetoWin.ObjPaisDirStr.ObjValorPro = "CO" OrElse FstrIdPaisActual() = "CO" Then
            txtDptoDir.Visibility = Visibility.Collapsed
            txtCiudadDir.Visibility = Visibility.Collapsed
            cboDptoDir.Visibility = Visibility.Visible
            cboCiudadDir.Visibility = Visibility.Visible
        Else
            txtDptoDir.Visibility = Visibility.Visible
            txtCiudadDir.Visibility = Visibility.Visible
            txtDptoDir.Text = String.Empty
            txtCiudadDir.Text = String.Empty
            cboDptoDir.Visibility = Visibility.Collapsed
            cboCiudadDir.Visibility = Visibility.Collapsed
        End If
    End Sub
    Private Function FstrIdPaisActual() As String
        Dim lstrIdPais As String, lstrNombrePais As String
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
            If Not IsNothing(cboPaisDir.SelectedItem) AndAlso cboPaisDir.SelectedIndex > 0 Then
                lstrNombrePais = cboPaisDir.SelectedItem
            Else
                lstrNombrePais = txtPaisDir.Text
            End If
        Else
            lstrNombrePais = txtPaisDir.Text
        End If
        lstrIdPais = MobjUbicacion.StrIdPais(lstrNombrePais)
        Return lstrIdPais
    End Function
    Private Function FBytIdDptoActual() As Byte
        Dim lbytIdDptoDir As Byte, lstrNomDpto As String
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
            If Not IsNothing(cboDptoDir.SelectedItem) AndAlso cboDptoDir.SelectedIndex > 0 Then
                lstrNomDpto = cboDptoDir.SelectedItem
            Else
                lstrNomDpto = txtDptoDir.Text
            End If
        Else
            lstrNomDpto = txtDptoDir.Text
        End If
        lbytIdDptoDir = MobjUbicacion.BytIdDpto(FstrIdPaisActual(), lstrNomDpto)
        Return lbytIdDptoDir
    End Function
    Private Function FshrIdCiudadActual() As Short
        Dim lshrIdCiu As Short, lstrNomCiu As String
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
            If Not IsNothing(cboCiudadDir.SelectedItem) AndAlso cboCiudadDir.SelectedIndex > 0 Then
                lstrNomCiu = cboCiudadDir.SelectedItem
            Else
                lstrNomCiu = txtCiudadDir.Text
            End If
        Else
            lstrNomCiu = txtCiudadDir.Text
        End If
        lshrIdCiu = MobjUbicacion.ShrIdCiudad(FstrIdPaisActual(), FBytIdDptoActual(),
                    lstrNomCiu)
        Return lshrIdCiu
    End Function
#End Region
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is CheckBox Then
            SEjecuteAccionCheckBox(lelmElemento)
        End If
    End Sub
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Dim lstrNombreAccion As String = lelmElemento.Name.Substring(3)
        Select Case lstrNombreAccion
            Case Is = "Importar"
                SAbraImportar()
            Case Is = "ReasignarId"
                Dim lobjWin As New WinReasignaIdTercero(MobjObjetoWin) With {
                        .WinPadre = Me
                    }
                lobjWin.ShowDialog()
                SRefrescarClic()
            Case Is = "AsociarFoto"
                SImporteFoto()
            Case "RotarFoto"
                SCargueFoto()
            Case "SuprimirFoto"
                SSuprimaFoto()
        End Select
    End Sub
    Private Sub EmnuiClick(sender As Object, e As RoutedEventArgs) Handles MnuAsociarFotoC.Click, MnuRotarFotoC.Click,
            MnuSuprimirFotoC.Click
        Select Case True
            Case sender.Equals(MnuAsociarFotoC)
                SImporteFoto()
            Case sender.Equals(MnuRotarFotoC)
                SCargueFoto()
            Case sender.Equals(MnuSuprimirFotoC)
                SSuprimaFoto()
        End Select
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        If Not HblnSeEstaCerrando Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is ComboBox OrElse TypeOf lelmElemento Is DatePicker OrElse
                    TypeOf lelmElemento Is TextBox Then
                If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
                    Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                    Try
                        SRegistreEntrada(lelmElemento)
                        SMuestreDatos()
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
    Private Sub OnCboCambio(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is ComboBox AndAlso Not mblnPoblandoCbo Then
            Select Case lelmElemento.Name
                Case "cboPaisDir"
                    MobjObjetoWin.ObjPaisDirStr.ObjValorPro = FstrIdPaisActual()
                    SSeleccioneControles()
                    SPuebleCboDptoDir()
                    SMuestreUbicaciones()
                Case "cboDptoDir"
                    MobjObjetoWin.ObjDepartamentoDirByt.ObjValorPro = FBytIdDptoActual()
                    SPuebleCboCiudadDir()
                Case "cboCiudadDir"
                    MobjObjetoWin.ObjCiudadDirShr.ObjValorPro = FshrIdCiudadActual()
            End Select
        End If
    End Sub
    Private Sub Txt_KeyDown(sender As Object, e As KeyEventArgs) Handles txtNroDoc.KeyDown
        If e.Key = Key.Return OrElse e.Key = Key.Tab Then
            If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando AndAlso
                    MobjObjetoWin.ObjIdTerceroDbl.ToString() <> txtNroDoc.Text Then
                SAbraTercero()
            End If
        End If
    End Sub
    Private Sub TbcTerceros_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) _
            Handles tbcTerceros.SelectionChanged
        If TypeOf e.Source Is TabControl Then
            If Not IsNothing(MobjObjetoWin) AndAlso Not mblnPoblandoCbo Then
                SMuestreUbicaciones()
            End If
        End If
    End Sub
#End Region
End Class
