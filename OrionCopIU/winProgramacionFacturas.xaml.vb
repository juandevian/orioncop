Imports System.ComponentModel
Public Class WinProgramacionFacturas
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        EnuDestinatario = 0
        EnuAno
        EnuServicio
        EnuPeriodoIni
        EnuCantPeriodos
        EnuVlrPeriodo
        EnuLectAnterior
        EnuLectActual
        EnuValorUnidad
    End Enum
#End Region
    ' Variables
    Private MblnDejoUltimoControl As Boolean = False
    Private MobjObjetoWin As ClsPredio = Nothing
    Private MobjCliente As ClsCliente = Nothing
    Private MobjItemProgramaFact As ClsItemProgramaFact = Nothing
    Private McolServicios As Collection = Nothing
    Private MblnPoblandoCombo As Boolean = False
    Private MshrIdServicio As Short = 0
    Private MnuImportarItems As MenuItemPan = Nothing
    Private MnuImportarCobroConsumos As MenuItemPan = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomProFact
    Private ReadOnly MstrRutaArchivo As String = IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "OPTIMUSOFT",
            "ori-cc-servicios", "ori-cc-servicios.exe")
#End Region

#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuProgramaFact
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolControlesLlave As New Collection From {
            rdbPredio,
            rdbCliente,
            txtIdDestinatario
        }
        SAdicioneControlRestringido(txtServicio)
        SAdicioneControlRestringido(txtPeriodoIni)
        SAdicioneControlRestringido(txtCantPeri)
        SAdicioneControlRestringido(txtVlrPeri)
        SAdicioneControlRestringido(txtOrigen)
        SAdicioneControlRestringido(cboAno)
        SAdicioneControlRestringido(dgrItemsProgramaFact)
        SCargueForma(EnuElementosAdicionalesDef.None, 9,
                lcolControlesLlave, txtVlrPeriNuevo, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
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
            ObjObjetoWin = ClsOrionCop.FobjNuevoPredio(EnuModoInstanciaObjDef.EnuNavegable)
        End If
        MobjObjetoWin = ObjObjetoWin
        If Not MobjObjetoWin.FblnEstaVacioOrigenDatos Then
            MobjObjetoWin.SVayaAlPrimero()
            MobjItemProgramaFact = MobjObjetoWin.ObjItemProgramaFact(0, 0)
        Else
            MobjItemProgramaFact = New ClsItemProgramaFact
        End If
        ObjHijoObjWin = MobjItemProgramaFact
        MobjCliente = ClsOrionCop.FobjCliente(EnuModoInstanciaObjDef.EnuNavegable)
        If Not MobjCliente.FblnEstaVacioOrigenDatos Then
            MobjCliente.SVayaAlPrimero()
        End If
        Dim lobjItemProFac As New ClsItemProgramaFact
        EnuTipoPermisoObjWin = lobjItemProFac.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuDestinatario) = lblIdDestinatario
        StcValidaControl(EnuValidEntrada.enuAno) = lblAno
        StcValidaControl(EnuValidEntrada.enuServicio) = lblServicio
        StcValidaControl(EnuValidEntrada.enuPeriodoIni) = lblPeriodoIni
        StcValidaControl(EnuValidEntrada.enuCantPeriodos) = lblCantidadPer
        StcValidaControl(EnuValidEntrada.EnuVlrPeriodo) = lblValorPerio
        StcValidaControl(EnuValidEntrada.EnuLectActual) = lblLecAActual
        StcValidaControl(EnuValidEntrada.EnuLectAnterior) = lblLecAnterior
        StcValidaControl(EnuValidEntrada.EnuValorUnidad) = lblVlrUnidad
        '
        SDeshabiliteControlesActuales()
        rdbPredio.IsChecked = True
        grdProgramaFact.DataContext = MobjObjetoWin.DtbItemsProgFac()
        SPuebleComboAnos()
        SPuebleCombosPeriodo()
        SPuebleComboServicios(False)
        '
        HbttAceptar.TabIndex = 13
        HbttCancelar.TabIndex = 14
    End Sub
    Protected Overrides Sub SMuestreDatos()
        HblnMostrandoDatos = True
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                SLevanteEveNoti("No hay Programación para ser mostrada!", "", 0,
                        EnuSeveridadNot.EnuInformacion)
                txtIdDestinatario.IsEnabled = False
            End If
        End If
        If rdbPredio.IsChecked Then
            txtIdDestinatario.Text = MobjObjetoWin.ObjIdPredioStr.ObjValorPro
            Dim lobjProp As ClsPropietario = MobjObjetoWin.ColPropietarios(1)
            Dim lstrNombrePro = lobjProp.ObjCliente.ObjNombreCompletoStr.ToString()
            txtNombre.Content = lstrNombrePro
            If MobjObjetoWin.ColPropietarios.Count > 1 Then
                txtNombre.Content &= " y otros."
            End If
        Else
            txtIdDestinatario.Text = MobjCliente.ObjIdClienteDbl.ObjValorPro
            txtNombre.Content = MobjCliente.ObjNombreCompletoStr.ObjValorPro
        End If
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If MobjCliente.BlnExiste Then
                SVinculeControles()
            End If
        End If
        txtNombre.ToolTip = txtNombre.Content
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            If Not IsNothing(MobjItemProgramaFact) Then
                With MobjItemProgramaFact
                    Dim lstrKey = .ObjIdAno_ItemProgramaFactShr.ToString & "," &
                            .ObjIdServicio_ItemProgramaFactShr.ToString
                    cboServicioNuevo.SelectedIndex = FentSelectedIndex(lstrKey)
                    cboAnoPeriodoNuevo.SelectedItem = Strings.Left(.ObjPeriodoIni_ItemProgStr.ToString, 4)
                    cboMesPeriodoNuevo.SelectedItem = Strings.Right(.ObjPeriodoIni_ItemProgStr.ToString, 2)
                    txtCantPeriNuevo.Text = .ObjCantidadPeriodosShr.ObjValorPro
                    txtVlrPeriNuevo.Text = Format(.ObjValorPeriodo_ItemProgramaFactDec.ObjValorPro, "c")
                End With
                dgrItemsProgramaFact.Visibility = Visibility.Hidden
            End If
        End If
        SMuestreComplemento()
        SValide()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            txtIdDestinatario.SelectAll()
        End If
        HblnMostrandoDatos = False
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            SInicialiceValido()
        Else
            With MobjItemProgramaFact
                If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                    If MobjItemProgramaFact IsNot Nothing Then
                        StcValidValido(EnuValidEntrada.EnuCantPeriodos) =
                                .ObjCantidadPeriodosShr.BlnEsValido
                        StcValidValido(EnuValidEntrada.EnuAno) =
                                .ObjIdAno_ItemProgramaFactShr.BlnEsValido
                        StcValidValido(EnuValidEntrada.EnuPeriodoIni) =
                                .ObjPeriodoIni_ItemProgStr.BlnEsValido
                        StcValidValido(EnuValidEntrada.EnuServicio) =
                                .ObjIdServicio_ItemProgramaFactShr.BlnEsValido
                        StcValidValido(EnuValidEntrada.EnuVlrPeriodo) =
                                .ObjValorPeriodo_ItemProgramaFactDec.BlnEsValido
                        If .BlnEsServicioConsumo Then
                            StcValidValido(EnuValidEntrada.EnuLectAnterior) =
                                    .ObjLecturaAnterior_ItemProgramaFactDec.BlnEsValido
                            StcValidValido(EnuValidEntrada.EnuLectActual) =
                                    .ObjLecturaActual_ItemProgramaFactDec.BlnEsValido
                            StcValidValido(EnuValidEntrada.EnuValorUnidad) =
                                    .ObjValorUnitario_ItemProgramaFactDec.BlnEsValido
                        End If
                    End If
                    Else
                    StcValidValido(EnuValidEntrada.enuCantPeriodos) = True
                    StcValidValido(EnuValidEntrada.enuAno) = True
                    StcValidValido(EnuValidEntrada.enuPeriodoIni) = True
                    StcValidValido(EnuValidEntrada.enuServicio) = True
                    StcValidValido(EnuValidEntrada.enuVlrPeriodo) = True
                End If
                If rdbPredio.IsEnabled Then
                    StcValidValido(EnuValidEntrada.enuDestinatario) =
                            MobjObjetoWin.ObjIdPredioStr.BlnEsValido
                Else
                    StcValidValido(EnuValidEntrada.enuDestinatario) =
                            MobjCliente.ObjIdClienteDbl.BlnEsValido
                End If
            End With
        End If
        '
        SHabiliteBotonesTlb()
        SHabiliteModificar()
        SHabiliteSuprimir()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjItemProgramaFact
            .ObjIdAno_ItemProgramaFactShr.ObjValorPro = FshrIdAnoActual()
            .ObjCantidadPeriodosShr.ObjValorPro = txtCantPeriNuevo.Text
            .ObjPeriodoIni_ItemProgStr.ObjValorPro = cboAnoPeriodoNuevo.SelectedItem &
                    cboMesPeriodoNuevo.SelectedItem
            .ObjIdCarpeta_ItemProgramaFactShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_ItemProgramaFactShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdServicio_ItemProgramaFactShr.ObjValorPro = FshrIdServicioNuevo()
            .ObjOrigen_ItemProgramaFacByt.ObjValorPro = EnuOrigenItemProgramaFactDef.EnuUsuario
            If rdbPredio.IsChecked Then
                .ObjIdPredio_ItemProgramaFactStr.ObjValorPro = txtIdDestinatario.Text
                .ObjIdCliente_ItemProgramaFactDbl.ObjValorPro = 0
            Else
                .ObjIdPredio_ItemProgramaFactStr.ObjValorPro = String.Empty
                .ObjIdCliente_ItemProgramaFactDbl.ObjValorPro = txtIdDestinatario.Text
            End If
        End With
        SValide()
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        Dim lsepSeparador As New Separator
        ' Adicionar menú Importar
        MnuImportarItems = FmnuiMenuItemPan("MnuImportarItems",
            "_Importar Cobros Períodicos", 1, "")
        HmnuAcciones.Items.Insert(7, MnuImportarItems)

        ' Adicionar menú Importar Cobros por Consumo
        If My.Computer.FileSystem.FileExists(MstrRutaArchivo) Then
            ' Crear el nuevo ítem de menú
            MnuImportarCobroConsumos = FmnuiMenuItemPan("MnuImportarCobroConsumos",
                "_Importar Cobros por Consumo", 2, "")
            MnuImportarCobroConsumos.ToolTip = "Permite importar masivamente ítems de cobro por consumo desde archivos Excel"
            ' Insertar el nuevo ítem debajo de "Importar Cobros Períodicos"
            HmnuAcciones.Items.Insert(8, MnuImportarCobroConsumos)
            HmnuAcciones.Items.Insert(9, lsepSeparador)
        Else
            HmnuAcciones.Items.Insert(8, lsepSeparador)
        End If
        ' Adicionar Menues Reportes
        Dim lmnuReportes As MenuItem = FmnuiMenuItem("MnuReportes", "Reportes", "RecMnuItemPriInf")
        MenuVen.Items.Insert(1, lmnuReportes)
        Dim lmnuItem = FmnuiMenuItem("MnuItemsProgramaFact", "Cobros programados", "RecMnuItemSec")
        lmnuItem.ToolTip = "Genera un Reporte de los cobros programados para ser generados " &
                "automáticamente."
        lmnuReportes.Items.Add(lmnuItem)
        HmnuSuprimir.Visibility = Visibility.Visible
        HbttSuprimir.Visibility = Visibility.Visible
    End Sub
#End Region

#Region "Procedimientos invalidantes"
    ''' <summary>
    ''' Sub que prepara a la ventana y a su objeto para crear un nuevo objeto. Invalida el Sub
    ''' "SCree" de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SCree()
        If Not ClsOrionCop.FblnHayPrefacturas Then
            If rdbPredio.IsChecked Then
                MobjItemProgramaFact = MobjObjetoWin.ObjNewItemProgramaFact
            Else
                MobjItemProgramaFact = MobjCliente.ObjNewItemProgramaFact
            End If
            ObjHijoObjWin = MobjItemProgramaFact
            EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando
            SMuestreDatos()
            SRegistre()
            SCreeItemPrograma()
        Else
            SLevanteEveNoti("No es posible modificar la Programación de Facturas mientras " &
                    " hayan pre-Facturas generadas!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    ''' <summary>
    ''' Prepara la ventana y su objeto para modificar el objeto. Invalida la función "SModifique"
    ''' de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SModifique()
        If MobjItemProgramaFact.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            EnuOperacionEnWin = EnuOperacionEnVentana.CenuModificando
            MobjItemProgramaFact.SModifique()
            If Not IsNothing(HbttCancelar) Then
                HbttCancelar.Content = My.Resources.Cancelar
            End If
            SHabiliteWin(True)
            txtVlrPeriNuevo.IsEnabled = Not MobjItemProgramaFact.BlnEsServicioConsumo
            txtCantPeriNuevo.IsEnabled = Not MobjItemProgramaFact.BlnEsServicioConsumo
            lblCantidadPer.IsEnabled = Not MobjItemProgramaFact.BlnEsServicioConsumo
            SMuestreDatos()
            SRegistre()
            Else
                Throw New ErrorInesperadoPanLException("Estado inesperado del objeto!")
        End If
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            SModifiqueItemPrograma()
        End If
    End Sub
    ''' <summary>
    ''' . Invalida el procedimiento "SGuarde" de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SGuarde()
        SRegistre()
        SValide()
        If FblnEstanTodosBien() Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
                If rdbPredio.IsChecked Then
                    MobjObjetoWin.SAgregueItemProgramaFact(MobjItemProgramaFact)
                Else
                    MobjCliente.SAgregueItemProgramaFact(MobjItemProgramaFact)
                End If
            ElseIf EnuOperacionEnWin = EnuOperacionEnVentana.CenuModificando Then
                MobjItemProgramaFact.SActualice(True)
            End If
        End If
        SFinaliceOperacion()
    End Sub
    Protected Overrides Sub SSuprima()
        Dim lstrMens = String.Empty, lblnNoHayError As Boolean
        Dim lstrMensEx = String.Empty, lblnSuprimio = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                Dim lblnEsSuprimible = MobjItemProgramaFact.FblnEsSuprimible()
                If lblnEsSuprimible Then
                    If MsgBox("Esta seguro de suprimir el presente " &
                            MobjItemProgramaFact.StrNombreClase & "?",
                            MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Supresión") =
                            MsgBoxResult.Yes Then
                        lblnSuprimio = MobjItemProgramaFact.FblnSuprimio()
                        If lblnSuprimio Then
                            MobjItemProgramaFact = Nothing
                        End If
                        SFinaliceOperacion()
                    End If
                Else
                    lstrMens = "El Item del Programa de facturación seleccionado no es suprimible!"
                End If
            End If
            If lblnSuprimio Then
                lstrMens = "El Item del Programa de Facturas fue suprimido exitosamente!"
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
                End If
                GobjPanDat.SControleProcesoObj(False)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuError)
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Protected Overrides Sub SFinaliceOperacion()
        Dim lstrMens As String = "Los datos del objeto " & ObjObjetoWin.StrNombreClase &
                " han cambiado!" & vbCrLf & "Desea guardar los cambios?"
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            With MobjItemProgramaFact
                If MobjItemProgramaFact IsNot Nothing Then
                    If .EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando Then
                        If FblnEstanTodosBien() AndAlso .BlnTengoCambios Then
                            If MsgBox(lstrMens, vbYesNo, "Aceptar Cambios") = vbYes Then
                                .SActualice(True)
                            Else
                                .SNormaliceEstado(True)
                            End If
                        Else
                            .SNormaliceEstado(True)
                        End If
                    End If
                End If
                EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
                dgrItemsProgramaFact.Visibility = Visibility.Visible
            End With
            HbttCancelar.Content = My.Resources.BtnCerrar
            SHabiliteWin(False)
        End If
        dgrItemsProgramaFact.IsEnabled = True
        SRefrescarClic()
        SVisibiliceControlesNuevos(False)
        If cboAno.Visibility = Visibility.Hidden Then
            cboAno.Visibility = Visibility.Visible
        End If
        rdbPredio.Focus()
    End Sub
    Protected Overrides Sub SRefresqueWin()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            Mouse.OverrideCursor = Cursors.Wait
            MshrIdServicio = 0
            SMuestreDatos()
            SLevanteEveOk()
            Mouse.OverrideCursor = Cursors.Arrow
        End If
    End Sub
#End Region

#Region "Invalida otros metodos se la clase base"
    Protected Overrides Sub SHabiliteMenues()
        MyBase.SHabiliteMenues()
    End Sub
#End Region

#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        MyBase.SBuscar()
        If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
            txtIdDestinatario.Text = StrResultadoBusqueda
            If rdbPredio.IsChecked Then
                SAbraPredio()
            Else
                SAbraCliente()
            End If
        End If
    End Sub
    ''' <summary>
    ''' Invalida la funcion "fblnDefinioBusqueda" de la clase base.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        If rdbPredio.IsChecked Then
            SDefineBusquedaPredio()
        Else
            SDefineBusquedaCliente()
        End If
        Return True
    End Function
    Private Sub SDefineBusquedaPredio()
        Dim lstrTablaPri As String = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec As String = ClsPropietario.SstrNombreTabla
        Dim lstrCampTablaPri As String() = {ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampTablaSec As String() = {ClsNombreCompleto_PropStr.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdPredio_PropStr.SstrNombreCampoBd}
        Dim lstrCampBusqueda = ClsNombreCompleto_PropStr.SstrNombreCampoBd
        Dim lstrFiltro As String = "P." & StrCampoCarpeta &
                " = " & GshrIdCarpeta & " AND P." &
                StrCampoCentroUtil & " = " &
                GshrIdCentroUtil & " AND " & lstrCampBusqueda & "<> ''"
        Dim lstrCampoRetornar As String = ClsIdPredioStr.SstrNombreCampoBd
        HwinBusqueda.SDefinaBusqueda("Nombre Propioetario", lstrTablaPri, lstrTablaSec,
                lstrCampTablaPri, lstrCampTablaSec, lstrCampRelPri, lstrCampRelSec,
                lstrCampBusqueda, lstrCampoRetornar, lstrFiltro, True)
    End Sub
    Private Sub SDefineBusquedaCliente()
        Dim lstrTabla As String = ClsCliente.SstrNombreTabla
        Dim lstrCamposMostrar As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                                 ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdClienteDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " & lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
#End Region

#Region "Procedimientos Propios"
    Private Sub SDeshabiliteControlesActuales()
        Dim lstyEstiloNoHabilitado As Style = FindResource("RecCtlNoHabilitado")
        txtServicio.Style = lstyEstiloNoHabilitado
        txtPeriodoIni.Style = lstyEstiloNoHabilitado
        txtCantPeri.Style = lstyEstiloNoHabilitado
        txtVlrPeri.Style = lstyEstiloNoHabilitado
        txtIdDestinatario.Style = FindResource("RecCtlHabilitado")
    End Sub
    Private Sub SHabiliteControlesNuevos(ablnHabilite As Boolean)
        Dim lstyEstiloHabilitado As Style = FindResource("RecCtlHabilitado")
        Dim lstyEstiloNoHabilitado As Style = FindResource("RecCtlNoHabilitado")
        SVisibiliceControlesNuevos(ablnHabilite)
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            cboServicioNuevo.Style = lstyEstiloHabilitado
            cboAnoPeriodoNuevo.Style = lstyEstiloHabilitado
            cboMesPeriodoNuevo.Style = lstyEstiloHabilitado
            txtCantPeriNuevo.Style = lstyEstiloHabilitado
            txtVlrPeriNuevo.Style = lstyEstiloHabilitado
            txtIdDestinatario.Style = lstyEstiloNoHabilitado
            rdbPredio.Style = lstyEstiloNoHabilitado
            rdbCliente.Style = lstyEstiloNoHabilitado
        End If
    End Sub
    Private Sub SVisibiliceControlesNuevos(ablnHabilite As Boolean)
        Dim lvisVisibilidadNuevos As Visibility
        Dim lvisVisibilidadActual As Visibility
        If ablnHabilite Then
            lvisVisibilidadNuevos = Visibility.Visible
            lvisVisibilidadActual = Visibility.Hidden
        Else
            lvisVisibilidadNuevos = Visibility.Hidden
            lvisVisibilidadActual = Visibility.Visible
        End If
        txtServicio.Visibility = lvisVisibilidadActual
        cboServicioNuevo.Visibility = lvisVisibilidadNuevos
        txtPeriodoIni.Visibility = lvisVisibilidadActual
        cboAnoPeriodoNuevo.Visibility = lvisVisibilidadNuevos
        cboMesPeriodoNuevo.Visibility = lvisVisibilidadNuevos
        lblAnoPeriodo.Visibility = lvisVisibilidadNuevos
        lblMesPeriodo.Visibility = lvisVisibilidadNuevos
        txtCantPeri.Visibility = lvisVisibilidadActual
        txtCantPeriNuevo.Visibility = lvisVisibilidadNuevos
        txtVlrPeri.Visibility = lvisVisibilidadActual
        txtVlrPeriNuevo.Visibility = lvisVisibilidadNuevos
    End Sub
    Private Sub SHabiliteSuprimir()
        Dim lstrRecursoBtt As String
        Dim lstrRecursoMnu As String
        Dim lblnHabilitarSuprimir = False
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If dgrItemsProgramaFact.Items.Count > 0 Then
                If Not IsNothing(MobjItemProgramaFact) AndAlso MobjItemProgramaFact.BlnExiste Then
                    lblnHabilitarSuprimir = Not MobjItemProgramaFact.ObjServicio_ItemProgramaFact.ObjGeneraProgramBln.ObjValorPro
                End If
            End If
        End If
        If lblnHabilitarSuprimir Then
            lstrRecursoBtt = "RecBttHabilitado"
            lstrRecursoMnu = "RecMnuItemSecHab"
        Else
            lstrRecursoBtt = "RecBttDesHabilitado"
            lstrRecursoMnu = "RecMnuItemSecDes"
        End If
        HbttSuprimir.Style = FindResource(lstrRecursoBtt)
        HmnuSuprimir.Style = FindResource(lstrRecursoMnu)
    End Sub
    Private Sub SHabiliteModificar()
        Dim lblnHabiliteModificar = (Not IsNothing(MobjItemProgramaFact)) AndAlso
                MobjItemProgramaFact.BlnExiste AndAlso
                CType(MobjItemProgramaFact.EnuPermisosObj And EnuPermisosDef.EnuModificar, Boolean)
        SHabiliteBotonTlb(lblnHabiliteModificar, HbttModificar)
        SHabiliteMenuItem(lblnHabiliteModificar, HmnuModificar)
    End Sub
    Private Sub SCreeItemPrograma()
        If Not IsNothing(HbttCancelar) Then
            HbttCancelar.Content = My.Resources.Cancelar
        End If
        SHabiliteWin(True)
        cboAno.Visibility = Visibility.Hidden
        SValide()
        SHabiliteControlesNuevos(True)
        dgrItemsProgramaFact.IsEnabled = False
        SPuebleComboServicios(True)
        cboServicioNuevo.Focus()
    End Sub
    Private Sub SModifiqueItemPrograma()
        If Not IsNothing(MobjItemProgramaFact) AndAlso MobjItemProgramaFact.BlnExiste Then
            SHabiliteControlesNuevos(True)
            cboServicioNuevo.Style = FindResource("RecCtlNoHabilitado")
            cboAno.Style = FindResource("RecCtlNoHabilitado")
            dgrItemsProgramaFact.IsEnabled = False
            SVerifiqueItemCalculado(False)
            txtVlrPeriNuevo.Focus()
        End If
    End Sub
    Private Sub SVerifiqueItemCalculado(ablnComparar As Boolean)
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuModificando Then
            Dim lblnMostrar = True
            If ablnComparar Then
                lblnMostrar = MobjItemProgramaFact.ObjValorPeriodo_ItemProgramaFactDec.BlnCambio
            End If
            If lblnMostrar Then
                Dim lobjServicio = MobjItemProgramaFact.ObjServicio_ItemProgramaFact
                If lobjServicio.ObjGeneraProgramBln.ObjValorPro Then
                    If lobjServicio.BlnEsCuotaAdministracion Then
                        cboAnoPeriodoNuevo.Style = FindResource("RecCtlNoHabilitado")
                        cboMesPeriodoNuevo.Style = FindResource("RecCtlNoHabilitado")
                        txtCantPeriNuevo.Style = FindResource("RecCtlNoHabilitado")
                    End If
                    SMensajeAdvertencia()
                End If
            End If
        End If
    End Sub
    Private Sub SMensajeAdvertencia()
        Dim lstrMens = "El Valor de este Item del Programa de Facturación fue calculado " &
                "por el Sistema." & vbCrLf & "Cualquier modificación hace que la " &
                "Integridad del Servicio se pierda!"
        MsgBox(lstrMens, MsgBoxStyle.Exclamation, "Advertencia")
        lstrMens = lstrMens.Replace(vbCrLf, " ")
        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuAdvertencia)
    End Sub
    Private Sub SAdecueAPredio()
        ObjObjetoWin = MobjObjetoWin
        lblIdDestinatario.Content = My.Resources.IdPredio
        cboAno.Style = FindResource("RecCtlHabilitado")
    End Sub
    Private Sub SAdecueACliente()
        ObjObjetoWin = MobjCliente
        lblIdDestinatario.Content = My.Resources.IdCliente
        cboAno.Style = FindResource("RecCtlNoHabilitado")
    End Sub
    Private Sub SPuebleComboAnos()
        MblnPoblandoCombo = True
        cboAno.Items.Clear()
        If rdbPredio.IsChecked Then
            Dim lcolAnos As Collection = GobjParametros.ColAnos
            For Each lobjAno As ClsAno In lcolAnos
                cboAno.Items.Add(lobjAno.ObjIdAnoShr.ObjValorPro)
            Next
        End If
        cboAno.Items.Add(0)
        MblnPoblandoCombo = False
        If rdbPredio.IsChecked Then
            cboAno.SelectedItem = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
        Else
            cboAno.SelectedItem = 0
        End If
    End Sub
    Private Sub SPuebleCombosPeriodo()
        Dim lshrIdAno As Short = Date.Today.Year
        Dim lstrIdAno = lshrIdAno.ToString
        cboAnoPeriodoNuevo.Items.Clear()
        cboAnoPeriodoNuevo.Items.Add(lstrIdAno)
        For i As Integer = 0 To 2
            lshrIdAno += 1
            cboAnoPeriodoNuevo.Items.Add(lshrIdAno.ToString)
        Next i
        For i As Integer = 1 To 12
            cboMesPeriodoNuevo.Items.Add(Format(i, "0#"))
        Next
        cboAnoPeriodoNuevo.SelectedIndex = 0
        cboMesPeriodoNuevo.SelectedItem = Format(Date.Today.Month, "0#")
    End Sub
    Private Sub SPuebleComboServicios(ablnRefresque As Boolean)
        MblnPoblandoCombo = True
        SPuebleColServicios(ablnRefresque)
        cboServicioNuevo.Items.Clear()
        cboServicioNuevo.Items.Add(My.Resources.Ninguno)
        For Each lobjServicio As ClsServicio In McolServicios
            If lobjServicio.ObjEstaActivoServicioBln.ObjValorPro Then
                cboServicioNuevo.Items.Add(lobjServicio.ObjConceptoServicioStr.ObjValorPro)
            End If
        Next
        cboServicioNuevo.SelectedItem = My.Resources.Ninguno
        MblnPoblandoCombo = False
    End Sub
    Private Sub SPuebleColServicios(ablnRefresque As Boolean)
        If ablnRefresque Then
            McolServicios = Nothing
        End If
        If IsNothing(McolServicios) Then
            McolServicios = New Collection
            Dim lcolServicios As Collection = GobjParametros.ColServiciosPer
            For Each lobjServicio As ClsServicio In lcolServicios
                If lobjServicio.ObjEsFactProgramableBln.ObjValorPro Then
                    Dim lstrKey = lobjServicio.ObjIdAno_ServicioShr.ToString & "," &
                            lobjServicio.ObjIdServicioShr.ToString
                    McolServicios.Add(lobjServicio, lstrKey)
                End If
            Next
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                If rdbPredio.IsChecked Then
                    Dim lshrIdAno As Short = CType(cboAno.SelectedItem, Short)
                    Dim lobjAno As ClsAno = GobjParametros.ColAnos(lshrIdAno.ToString)
                    Dim lcolServiciosAno As Collection = lobjAno.ColServiciosAno
                    For Each lobjServicio As ClsServicio In lcolServiciosAno
                        Dim lstrKey = lobjServicio.ObjIdAno_ServicioShr.ToString & "," &
                                lobjServicio.ObjIdServicioShr.ToString
                        McolServicios.Add(lobjServicio, lstrKey)
                    Next
                End If
            End If
        End If
    End Sub
    Private Sub SVinculeControles()
        If Not MblnPoblandoCombo Then
            Dim lshrIdAno As Short = cboAno.SelectedItem
            Dim ldtbItemsProgFact As DataTable
            If rdbPredio.IsChecked Then
                ldtbItemsProgFact = MobjObjetoWin.DtbItemsProgFac(lshrIdAno)
            Else
                ldtbItemsProgFact = MobjCliente.DtbItemsProgFac()
            End If
            grdProgramaFact.DataContext = ldtbItemsProgFact
            SOrdeneDataGrid(dgrItemsProgramaFact, dgrItemsProgramaFact.Columns(0), "IdAno",
                      ListSortDirection.Ascending)
        End If
    End Sub
    Private Function FshrIdAnoActual() As Short
        Dim lshrIdAno As Short = 0
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuCreando Then
            Dim ldrvItemProgFac As DataRowView = dgrItemsProgramaFact.SelectedItem
            If Not IsNothing(ldrvItemProgFac) Then
                lshrIdAno = ldrvItemProgFac("IdAno")
            End If
        End If
        Return lshrIdAno
    End Function
    Private Function FshrIdServicioNuevo() As Short
        Dim lshrIdServicio As Short = 0
        Dim lstrNombreServicio As String = cboServicioNuevo.SelectedItem
        For Each lobjServicio As ClsServicio In McolServicios
            If lobjServicio.ObjConceptoServicioStr.ObjValorPro = lstrNombreServicio Then
                lshrIdServicio = lobjServicio.ObjIdServicioShr.ObjValorPro
                Exit For
            End If
        Next
        Return lshrIdServicio
    End Function
    Private Function FentSelectedIndex(astrKey As String) As Integer
        Dim lentIndiceCombo = 0
        Dim lstrNombreServicio = String.Empty
        Dim lobjServicio As ClsServicio = Nothing
        If McolServicios.Contains(astrKey) Then
            lobjServicio = McolServicios(astrKey)
        End If
        If Not IsNothing(lobjServicio) Then
            lstrNombreServicio = lobjServicio.ObjConceptoServicioStr.ObjValorPro
        End If
        For i = 0 To cboServicioNuevo.Items.Count - 1
            If cboServicioNuevo.Items(i).ToString = lstrNombreServicio Then
                lentIndiceCombo = i
                Exit For
            End If
        Next
        Return lentIndiceCombo
    End Function
    Private Sub SActualiceColorModo()
        If txtOrigen.Text.StartsWith("1") OrElse String.IsNullOrEmpty(txtOrigen.Text) Then
            txtOrigen.Style = FindResource("RecTxtOrigenSistema")
        Else
            txtOrigen.Style = FindResource("RecTxtOrigenUsuario")
        End If
    End Sub
    Private Sub SAbraPredio()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                Dim lobjVlrLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil,
                        txtIdDestinatario.Text}
                MobjObjetoWin.SAbra(lobjVlrLlave)
                If Not MobjObjetoWin.BlnExiste Then
                    MobjObjetoWin.SRefresqueObj()
                    lstrMens = "El predio ingresado no existe!"
                End If
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
                If lblnNoHayError Then
                    If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                        If Not String.IsNullOrEmpty(lstrMens) Then
                            SLevanteEveNoti(lstrMens, String.Empty, 0,
                                EnuSeveridadNot.EnuInformacion)
                        Else
                            SRefrescarClic()
                            txtIdDestinatario.SelectAll()
                        End If
                    End If
                Else
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
    End Sub
    Private Sub SAbraCliente()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                Dim lobjVlrLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, txtIdDestinatario.Text}
                MobjCliente.SAbra(lobjVlrLlave)
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
                If lblnNoHayError Then
                    If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                        MobjCliente.SRefresqueObj()
                        SRefrescarClic()
                    End If
                Else
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
    End Sub
    Private Sub SAbraImportar()
        Dim lstrMens = String.Empty, lstrMensErr = String.Empty, lblnNoHayError = False
        If Not ClsOrionCop.FblnHayPrefacturas Then
            If IsNothing(MobjItemProgramaFact) Then
                MobjItemProgramaFact = MobjObjetoWin.ObjNewItemProgramaFact
            End If
            Dim lwinImportar As New WinImportar(MobjItemProgramaFact) With {
                .WinPadre = Me,
                .BlnExigeRequeridos = False
            }
            GblnOK = True
            lwinImportar.ShowDialog()
            If GblnOK Then
                Try
                    GobjPanDat.SControleProcesoObj(True)
                    GobjPanDat.SInicialiceTransaccion()
                    GblnImportando = True
                    GblnImportando = False
                    SRefrescarClic()
                    GobjParametros.SRefresqueObj()
                    lblnNoHayError = True
                Catch ex As ErrorInesperadoPanLException
                    lstrMens = ex.Message
                    lstrMensErr = ex.ToString()
                Catch ex As Exception
                    lstrMens = ex.Message
                    lstrMensErr = ex.ToString()
                Finally
                    If lblnNoHayError Then
                        GobjPanDat.SConfirmeTransaccion()
                        GobjPanDat.SControleProcesoObj(False)
                        If Not String.IsNullOrEmpty(lstrMens) Then
                            SLevanteEveNoti(lstrMens, lstrMensErr, 0,
                                    EnuSeveridadNot.EnuInformacion)
                        End If
                    Else
                        GobjPanDat.SAborteTransaccion()
                        GobjPanDat.SControleProcesoObj(False, True)
                        SLevanteEveNoti(lstrMens, lstrMensErr, 0,
                                EnuSeveridadNot.EnuExcep)
                    End If
                End Try
            End If
        Else
            lstrMens = "No es posible modificar la Programación de Facturas mientras " &
                    "haya pre-Facturas generadas!"
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, lstrMensErr, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SAbraImportarCobroConsumos()
        Try
            ' Verificar si el archivo existe
            If My.Computer.FileSystem.FileExists(MstrRutaArchivo) Then
                ' Abrir el archivo ejecutable
                Process.Start(MstrRutaArchivo)
            Else
                ' Mostrar un mensaje si el archivo no existe
                MsgBox("El archivo especificado no se encuentra en la ruta: " & MstrRutaArchivo, MsgBoxStyle.Exclamation, "Archivo no encontrado")
            End If
        Catch ex As Exception
            ' Manejar cualquier error al intentar abrir el archivo
            MsgBox("Ocurrió un error al intentar abrir el archivo: " & ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub SRefresqueObjetoWin()
        Dim lobjItemProgramaFact As ClsItemProgramaFact
        Dim lshrIdAno = FshrIdAnoActual()
        Dim lshrIdServicio = 0
        If Not IsNothing(dgrItemsProgramaFact.SelectedItem) Then
            lshrIdServicio = MshrIdServicio
        End If
        If lshrIdServicio <> 0 Then
            If rdbPredio.IsChecked Then
                lobjItemProgramaFact = MobjObjetoWin.ObjItemProgramaFact(lshrIdAno, lshrIdServicio)
            Else
                lobjItemProgramaFact = MobjCliente.ObjItemProgramaFact(lshrIdServicio)
            End If
        Else
            If rdbPredio.IsChecked Then
                lobjItemProgramaFact = MobjObjetoWin.ObjItemProgramaFact(0, 0)
            Else
                lobjItemProgramaFact = MobjCliente.ObjItemProgramaFact(0)
            End If
        End If
        MobjItemProgramaFact = lobjItemProgramaFact
        ObjHijoObjWin = MobjObjetoWin
    End Sub
    Private Sub SGenereRepItemsProgFact()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
                .EnuReporte = EnuReporteDef.enuItemsProgramaFact
                }
            lobjRep.SGenereReporte()
            lblnNoHayError = True
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Private Sub SMuestreComplemento()
        If MobjItemProgramaFact IsNot Nothing Then
            txtTarifaIva.Content = Format(MobjItemProgramaFact.DblTarifaIva_Servicio, "p")
            txtSaldo.Content = Format(MobjItemProgramaFact.ObjSaldo_ItemProgramaFactDec.ObjValorPro,
                    "c")
        Else
            txtTarifaIva.Content = String.Empty
            txtSaldo.Content = String.Empty
        End If
    End Sub
    Private Sub SMuestreConsumo()
        If MobjItemProgramaFact IsNot Nothing Then
            With MobjItemProgramaFact
                If .ObjLecturaActual_ItemProgramaFactDec.ObjValorPro > 0 Then
                    grbConsumo.Visibility = Visibility.Visible
                    txtLecAnterior.Text = .ObjLecturaAnterior_ItemProgramaFactDec.ObjValorPro
                    txtLecActual.Text = .ObjLecturaActual_ItemProgramaFactDec.ObjValorPro
                    txtVlrUnidad.Text = Format(.ObjValorUnitario_ItemProgramaFactDec.ObjValorPro,
                            "c")
                Else
                    grbConsumo.Visibility = Visibility.Collapsed
                    txtLecAnterior.Text = String.Empty
                    txtLecActual.Text = String.Empty
                    txtVlrUnidad.Text = String.Empty
                End If
            End With
        Else
            grbConsumo.Visibility = Visibility.Collapsed
            txtLecAnterior.Text = String.Empty
            txtLecActual.Text = String.Empty
            txtVlrUnidad.Text = String.Empty
        End If
    End Sub
#End Region

#Region "Eventos en la Ventana"
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            Select Case lelmElemento.Name
                Case "MnuItemsProgramaFact"
                    SGenereRepItemsProgFact()
                Case "MnuImportarItems"
                    SAbraImportar()
                Case "MnuImportarCobroConsumos"
                    SAbraImportarCobroConsumos()
            End Select
        End If
    End Sub

    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            ltxtTextBox.SelectAll()
        ElseIf lelmElemento.Equals(HbttCancelar) Then
            If MblnDejoUltimoControl Then
                HbttAceptar.Focus()
            End If
        End If
    End Sub

    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        If Not HblnSeEstaCerrando AndAlso Not HblnMostrandoDatos Then
            Dim lblnNoHayError = False
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is TextBox Then
                MblnDejoUltimoControl = False
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    GobjPanDat.SControleProcesoObj(True)
                    With MobjItemProgramaFact
                        Select Case lelmElemento.Name
                            Case "txtVlrPeriNuevo"
                                .ObjValorPeriodo_ItemProgramaFactDec.ObjValorPro = txtVlrPeriNuevo.Text
                                SVerifiqueItemCalculado(True)
                            Case "txtCantPeriNuevo"
                                .ObjCantidadPeriodosShr.ObjValorPro = txtCantPeriNuevo.Text
                                MblnDejoUltimoControl = True
                            Case "txtLecAnterior"
                                .ObjLecturaAnterior_ItemProgramaFactDec.ObjValorPro =
                                        txtLecAnterior.Text
                            Case "txtLecActual"
                                .ObjLecturaActual_ItemProgramaFactDec.ObjValorPro =
                                        txtLecActual.Text
                            Case "txtVlrUnidad"
                                .ObjValorUnitario_ItemProgramaFactDec.ObjValorPro =
                                        txtVlrUnidad.Text
                        End Select
                        .SActualiceSaldoActual()
                    End With
                    SMuestreDatos()
                    GobjPanDat.SControleProcesoObj(False)
                Else
                    If lelmElemento.Equals(txtIdDestinatario) Then
                        If rdbPredio.IsChecked Then
                            SAbraPredio()
                        Else
                            SAbraCliente()
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Txt_TextChanged(sender As Object, e As TextChangedEventArgs) Handles _
            txtIdDestinatario.TextChanged, txtOrigen.TextChanged
        Select Case True
            Case sender.Equals(txtOrigen)
                SActualiceColorModo()
            Case Else
                SRefresqueObjetoWin()
        End Select
    End Sub

    Private Sub CboAno_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cboAno.SelectionChanged
        If Not MblnPoblandoCombo Then
            SVinculeControles()
        End If
    End Sub

    Private Sub Cbo_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cboServicioNuevo.SelectionChanged,
            cboAnoPeriodoNuevo.SelectionChanged, cboMesPeriodoNuevo.SelectionChanged
        If Not MblnPoblandoCombo Then
            Dim lblnNoHayError = False
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                If TypeOf lelmElemento Is ComboBox Then
                    GobjPanDat.SControleProcesoObj(True)
                    With MobjItemProgramaFact
                        Select Case lelmElemento.Name
                            Case "cboServicioNuevo"
                                .ObjIdServicio_ItemProgramaFactShr.ObjValorPro =
                                        FshrIdServicioNuevo()
                                txtTarifaIva.Content = Format(.DblTarifaIva_Servicio, "p")
                            Case "cboAnoPeriodoNuevo", "cboMesPeriodoNuevo"
                                Dim lstrPeriodoIni As String = cboAnoPeriodoNuevo.SelectedItem &
                                        cboMesPeriodoNuevo.SelectedItem
                                .ObjPeriodoIni_ItemProgStr.ObjValorPro = lstrPeriodoIni
                                .SActualiceSaldoActual()
                        End Select
                    End With
                    SMuestreDatos()
                    GobjPanDat.SControleProcesoObj(False)
                End If
            End If
        End If
    End Sub

    Private Sub DgrItemsProgramaFact_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles dgrItemsProgramaFact.SelectionChanged
        Dim ldrvItemProgFac As DataRowView = dgrItemsProgramaFact.SelectedItem
        If Not IsNothing(ldrvItemProgFac) Then
            MshrIdServicio = ldrvItemProgFac("IdServicio")
        End If
        SRefresqueObjetoWin()
        SMuestreComplemento()
        SMuestreConsumo()
        SValide()
    End Sub

    Private Sub ClsFormInterface_KeyDow(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Return Then
            If txtIdDestinatario.IsFocused Then
                If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                    If rdbPredio.IsChecked Then
                        SAbraPredio()
                    Else
                        SAbraCliente()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub RdbPredio_Click(sender As Object, e As RoutedEventArgs) Handles rdbPredio.Click,
            rdbCliente.Click
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is RadioButton Then
            If rdbPredio.IsChecked Then
                SAdecueAPredio()
            Else
                SAdecueACliente()
            End If
            SPuebleComboAnos()
            SPuebleComboServicios(True)
            If dgrItemsProgramaFact.Items.Count > 0 Then
                dgrItemsProgramaFact.SelectedIndex = 0
            End If
            SMuestreDatos()
        End If
    End Sub

    Private Sub Txt_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles _
            txtIdDestinatario.MouseDoubleClick
        If Not String.IsNullOrEmpty(txtIdDestinatario.Text) Then
            If rdbPredio.IsChecked Then
                SAbraWinPredio(txtIdDestinatario.Text)
            Else
                SAbraWinCliente(txtIdDestinatario.Text)
            End If
        End If

    End Sub
#End Region
End Class
