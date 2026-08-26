Public Class WinParametrizacion
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuIdAccionWin As Integer
        None = 0
        EnuTablasAuxiliares
        EnuAbrirSctores
        EnuModContrib
        EnuTerceros
        EnuUbicaciones
        EnuCopiasSeg
        EnuCarpeta
        EnuCtasCont
        EnuCenUtil
        EnuOpcionesCenUtil
        EnuAbrirDocs
        EnuTasasInt
        EnuAgruSer
        EnuAbrirSer
        EnuAno
        EnuAbrirAno
        EnuCrearAno
        EnuSupimirAno
        EnuAbrirBancos
    End Enum

#End Region
    ' Variables
    Private WithEvents MwinVentana As ClsFormInterface = Nothing
    ' Menu acciones de objeto (Se adiciona al menu principal dependiendo del Nodo del Arbol seleccionado)
    Private MnuAccionesDeObjeto As Controls.MenuItem = Nothing
    '
    Private MnuTablasAuxiliares As MenuItemPan = Nothing
    Private MnuSectores As MenuItemPan = Nothing
    Private MnuModulos As MenuItemPan = Nothing
    Private MnuTerceros As MenuItemPan = Nothing
    Private MnuUbicaciones As MenuItemPan = Nothing
    Private MnuCuentaOrigenEmail As MenuItemPan = Nothing
    Private MnuOpcionesBK As MenuItemPan = Nothing
    '
    Private MnuAbrirCtaCon As MenuItemPan = Nothing
    Private MnuAbrirAno As MenuItem = Nothing
    Private MnuAbrirAnoC As MenuItem = Nothing
    Private MnuOpcionesCentroUtil As MenuItem = Nothing
    Private MnuAbrirConfigC As MenuItem = Nothing
    Private MnuAbrirCtaConC As MenuItem = Nothing
    Private MnuAbrirServicioPer As MenuItem = Nothing
    Private MnuAbrirServicioAno As MenuItem = Nothing
    Private MnuAbrirServicioPerC As MenuItem = Nothing
    Private MnuAbrirServicioAnoC As MenuItem = Nothing
    Private MnuAbrirTasasMora As MenuItem = Nothing
    Private MnuAbrirTasasMoraC As MenuItem = Nothing
    Private MnuAbrirBancos As MenuItem = Nothing
    Private MnuAbrirBancosC As MenuItem = Nothing
    Private MnuAbrirDocumentos As MenuItem = Nothing
    Private MnuAbrirDocumentosC As MenuItem = Nothing

    Private MnuCrearAno As MenuItemPan = Nothing 'Ok
    Private MnuCrearAnoC As MenuItem = Nothing
    Private MnuAbrirServicioAnoCN As MenuItem = Nothing

    Private MnuSuprimirAno As MenuItem = Nothing
    Private MnuSuprimirAnoC As MenuItem = Nothing

    Private WithEvents MtviPlanContable As TreeViewItem = Nothing
    Private MtviCarpeta As TreeViewItem = Nothing
    Private MtviCentroUtilOriCop As TreeViewItem = Nothing
    Private MtviOrionCop As TreeViewItem = Nothing
    Private MtviServiciosPermanentes As TreeViewItem = Nothing
    Private MtviTasasMora As TreeViewItem = Nothing
    Private MtviBancos As TreeViewItem = Nothing
    Private MtviDocumentos As TreeViewItem = Nothing

    Private MblnModificandoArbol As Boolean = False
    Private MblnNodosAnoModificados As Boolean = False
    Private MblnRefrescando As Boolean = False
    Private MblnCerrando As Boolean = False
    '
    Private MfrmAyuda As FrmAyuda = Nothing
    '
    Private ReadOnly MobjCarpeta As ClsCarpeta = GobjPanorama.ObjCarpetaActual
    Private ReadOnly MobjCentroUtil As ClsCentroUtilidad = MobjCarpeta.ObjCentroUtilidadActual
    Private MobjServicio As ClsServicio = Nothing
    Private MobjAno As ClsAno = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomPara
    ' 
    Private Property MdblnIncrementoCA As Double = Nothing
    Private Property MenuTipoCalculoCuota As EnuTipoBaseCalculo = Nothing
    Friend Property BlnInicioFacMesActual As Boolean = Nothing
#End Region

#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuParametrizacion
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            SCargueForma(EnuElementosAdicionalesDef.None, 0, Nothing, Nothing, True)
            SPuebleBarraEstado(HcolLabelsBarraEstado)
            MtviOrionCop.IsSelected = True
            If EnuEstadoAyuda <> EnuEstadoAyudaDef.EnuOff AndAlso
                    Not GobjParametros.ObjNoMostrarAyudaBln.ObjValorPro Then
                SAbraWinAyuda()
            End If
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentNullException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                SLevanteEveNoti("", "", 0, EnuSeveridadNot.EnuOk)
                GobjPanDat.SControleProcesoObj(False)
            Else
                GblnOK = False
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            HblnCargandoForma = False
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
        ObjObjetoWin = GobjParametros
    End Sub
    Protected Overrides Sub SInicialiceControles()
        SInicialiceNodos()
    End Sub
    Protected Overrides Sub SMuestreDatos()
        trvOrionCop.Focus()
    End Sub
    Protected Overrides Sub SValide()
        '
    End Sub
    Protected Overrides Sub SRegistre()
        '
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin".
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        Dim lsepTabAux_1 As New Separator
        Dim lsepTabAux_2 As New Separator
        HmnuCerrar = MnuCerrar
        HmnuRefrescar = MnuRefrescar
        SDefinaMenusWin()
        MenuVen.Items.Insert(1, MnuTablasAuxiliares)
        With MnuTablasAuxiliares
            .Items.Add(MnuSectores)
            .Items.Add(MnuModulos)
            .Items.Add(lsepTabAux_1)
            .Items.Add(MnuTerceros)
            .Items.Add(MnuCuentaOrigenEmail)
            .Items.Add(MnuUbicaciones)
            .Items.Add(lsepTabAux_2)
            .Items.Add(MnuOpcionesBK)
        End With
        SAsigneMenuesContextuales()
    End Sub
#End Region

#Region "Procedimientos invalidantes"
    Protected Overrides Sub SHabiliteMenues()
        MyBase.SHabiliteMenues()
        SHabiliteMenuPan(MnuAbrirCtaCon)
        SHabiliteMenuPan(MnuOpcionesCentroUtil)
        SHabiliteMenuPan(MnuAbrirBancos)
        SHabiliteMenuPan(MnuAbrirDocumentos)
        SHabiliteMenuPan(MnuAbrirTasasMora)
        SHabiliteMenuPan(MnuAbrirServicioPer)
        SHabiliteMenuPan(MnuAbrirServicioAno)
        SHabiliteMenuPan(MnuAbrirAno)
        SHabiliteMenuPan(MnuCrearAno)
        SHabiliteMenuPan(MnuSuprimirAno)
        ' Menus contextuales
        MnuAbrirCtaConC.IsEnabled = MnuAbrirCtaCon.IsEnabled
        MnuAbrirConfigC.IsEnabled = MnuOpcionesCentroUtil.IsEnabled
        MnuAbrirBancosC.IsEnabled = MnuAbrirBancos.IsEnabled
        MnuAbrirDocumentosC.IsEnabled = MnuAbrirDocumentos.IsEnabled
        MnuAbrirTasasMoraC.IsEnabled = MnuAbrirTasasMora.IsEnabled
        MnuAbrirServicioPerC.IsEnabled = MnuAbrirServicioPer.IsEnabled
        MnuAbrirServicioAnoC.IsEnabled = MnuAbrirServicioAno.IsEnabled
        MnuAbrirServicioAnoCN.IsEnabled = MnuAbrirServicioAno.IsEnabled
        MnuAbrirAnoC.IsEnabled = MnuAbrirAno.IsEnabled
        MnuCrearAnoC.IsEnabled = MnuCrearAno.IsEnabled
        MnuSuprimirAnoC.IsEnabled = MnuSuprimirAno.IsEnabled
        If GstrIdUsuario = GCSTRUSUARIOU Then
            MnuTerceros.Visibility = Visibility.Visible
        Else
            MnuTerceros.Visibility = Visibility.Collapsed
        End If
    End Sub

    Protected Overrides Sub SRefresqueWin()
        MblnRefrescando = True
        SContraigaTodo(MtviOrionCop)
        SContraigaRama(MtviOrionCop)
        MtviServiciosPermanentes.Items.Clear()
        SQuiteNodosAnos()
        MobjCarpeta.SRefresqueObj()
        MobjCentroUtil.SRefresqueObj()
        MobjServicio = Nothing
        MobjAno = Nothing
        MblnNodosAnoModificados = False
        GobjParametros.SRefresqueObj()
        SExpandaTodo(MtviOrionCop)
        MblnRefrescando = False
        SVerifiqueInstalacion()
        SHabiliteMenues()
        MtviCentroUtilOriCop.IsSelected = True
        Dim lblnEnfocado = Focus()
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        If EnuEstadoAyuda = EnuEstadoAyudaDef.EnuFrmOn Then
            SRefresqueAyuda()
        End If
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub

    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        SDefineCuentaCont()
        Return True
    End Function

    Protected Overrides Sub SCerrarClic()
        If GobjParametros.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            GobjParametros.SNormaliceEstado(False)
        End If
        MblnCerrando = True
        If MfrmAyuda IsNot Nothing Then
            MfrmAyuda.SCierre()
            MfrmAyuda = Nothing
        End If
        MyBase.SCerrarClic()
    End Sub
#End Region

#Region "Procedimientos Propios"
#Region "Menues"
    Private Sub SDefinaMenusWin()
        MnuAccionesDeObjeto = FmnuiMenuItem("MnuAccionesDeObjeto", "", "RecMnuItemPri")
        MnuTablasAuxiliares = FmnuiMenuItemPan("MnuTablasAuxiliares", "_Tablas Auxiliares",
                EnuIdAccionWin.EnuTablasAuxiliares, "", True)
        MnuSectores = FmnuiMenuItemPan("MnuSectores", "Sec_tores", EnuIdAccionWin.EnuAbrirSctores, "")
        MnuModulos = FmnuiMenuItemPan("MnuModulos", "_Módulos de Contribución",
                EnuIdAccionWin.EnuModContrib, "")
        MnuTerceros = FmnuiMenuItemPan("MnuTerceros", "_Terceros", EnuIdAccionWin.EnuTerceros,
                "", "terceros.png")
        MnuCuentaOrigenEmail = FmnuiMenuItemPan("MnuCuentaOrigenEmail", "_Cuenta Origen Correo",
                EnuIdAccionWin.EnuTerceros, "", "mensaje.png")
        MnuUbicaciones = FmnuiMenuItemPan("MnuUbicaciones", "_Ubicaciones",
                EnuIdAccionWin.EnuUbicaciones, "")
        MnuOpcionesBK = FmnuiMenuItemPan("MnuOpcionesBK", "_Opciones Copias de Seguridad",
                EnuIdAccionWin.EnuCopiasSeg, "")
        '
        MnuAbrirCtaCon = FmnuiMenuItemPan("MnuAbrirCtaCon", "_Abrir Cuentas Contabilidad",
                EnuIdAccionWin.EnuCtasCont, "", "cuentasContabilidad.png")
        MnuOpcionesCentroUtil = FmnuiMenuItemPan("MnuOpcionesCentroUtil",
                "_Abrir Opciones de la Copropiedad", EnuIdAccionWin.EnuOpcionesCenUtil,
                "", "config.png")
        MnuAbrirBancos = FmnuiMenuItemPan("MnuAbrirBancos", "_Abrir Cuentas Bancarias",
                EnuIdAccionWin.EnuAbrirBancos, "", "CtaBanco.png")
        MnuAbrirDocumentos = FmnuiMenuItemPan("MnuAbrirDocumentos", "_Abrir Documentos",
                EnuIdAccionWin.EnuAbrirDocs, "", "Documentos.png")
        MnuAbrirTasasMora = FmnuiMenuItemPan("MnuAbrirTasasMora", "_Abrir Tasas de Mora",
                EnuIdAccionWin.EnuTasasInt, "", "CondicionesMora.png")
        MnuAbrirServicioPer = FmnuiMenuItemPan("MnuAbrirServicioPer", "_Abrir Servicio Permanente",
                EnuIdAccionWin.EnuAbrirSer, "", "Servicios.png")
        MnuAbrirServicioAno = FmnuiMenuItemPan("MnuAbrirServicioAno", "_Abrir Servicio Anual",
                EnuIdAccionWin.EnuAbrirSer, "", "Servicios.png")
        MnuAbrirAno = FmnuiMenuItemPan("MnuAbrirAno", "_Abrir Año", EnuIdAccionWin.EnuAbrirAno,
                "", "ano.png")
        MnuCrearAno = FmnuiMenuItemPan("MnuCrearAno", "_Crear Año", EnuIdAccionWin.EnuCrearAno,
                "", "ano.png")
        MnuSuprimirAno = FmnuiMenuItemPan("MnuSuprimirAno", "_Suprimir Año",
                EnuIdAccionWin.EnuSupimirAno, "", "ano.png")
    End Sub
#End Region
#Region "Manejo Arbol"
    Private Shared Function FstrTrayectoria(atviNodo As TreeViewItem)
        Dim lstrTextoTvi As String
        Dim ltviPadre As TreeViewItem = Nothing
        lstrTextoTvi = atviNodo.Tag
        Dim lstrTray = lstrTextoTvi
        If atviNodo.Name <> "tviNodoRaiz" Then
            ltviPadre = atviNodo.Parent
        End If
        Do While lstrTextoTvi <> GobjPanorama.ObjAppActual.StrNombreCompleto
            lstrTextoTvi = ltviPadre.Tag
            lstrTray = lstrTextoTvi & "/" & lstrTray
            If lstrTextoTvi <> GobjPanorama.ObjAppActual.StrNombreCompleto Then
                ltviPadre = ltviPadre.Parent
            End If
        Loop
        Return lstrTray
    End Function
    Private Sub SExpandaTodo(atviNodo As TreeViewItem)
        Static ltviNodoIni As TreeViewItem = Nothing
        If IsNothing(ltviNodoIni) Then
            ltviNodoIni = atviNodo
        End If
        MblnModificandoArbol = True
        atviNodo.IsExpanded = True
        If atviNodo.Items.Count > 0 Then
            For Each ltviNodo As TreeViewItem In atviNodo.Items
                SExpandaTodo(ltviNodo)
            Next
        End If
        If ltviNodoIni.Name = atviNodo.Name Then
            ltviNodoIni = Nothing
            MblnModificandoArbol = False
        End If
        SRefresqueNodosAno()
    End Sub
    Private Sub SExpandaRama(atviNodo As TreeViewItem)
        MblnModificandoArbol = True
        atviNodo.IsExpanded = True
        MblnModificandoArbol = False
    End Sub
    Private Sub SContraigaTodo(atviNodo As TreeViewItem)
        Static ltviNodoIni As TreeViewItem = Nothing
        If IsNothing(ltviNodoIni) Then
            ltviNodoIni = atviNodo
        End If
        MblnModificandoArbol = True
        atviNodo.IsExpanded = False
        If atviNodo.Items.Count > 0 Then
            For Each ltviNodo As TreeViewItem In atviNodo.Items
                SContraigaTodo(ltviNodo)
            Next
        End If
        If ltviNodoIni.Name = atviNodo.Name Then
            ltviNodoIni = Nothing
            MblnModificandoArbol = False
        End If
    End Sub
    Private Sub SContraigaRama(atviNodo As TreeViewItem)
        MblnModificandoArbol = True
        atviNodo.IsExpanded = False
        MblnModificandoArbol = False
    End Sub
#End Region
#Region "Manejo de Objetos"
    Private Sub SDefineCuentaCont()
        Dim lstrCamposMostrar As String() = {ClsIdCuentaContStr.SstrNombreCampoBd,
                                                 ClsNombreCuentaStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCuentaStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdCuentaContStr.SstrNombreCampoBd
        Dim lstrTabla As String = ClsCuentaContabilidad.SstrNombreTabla
        Dim lstrFiltro As String = ClsIdCarpetaCuentaShr.SstrNombreCampoBd & " = " & GshrIdCarpeta &
                " AND " & lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Nombre Cuenta", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
    ''' <summary>
    ''' Crea un nuevo año.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SCreeAno()
        Dim lblnNoHayError = False
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty
        Dim lobjAno As ClsAno = Nothing
        Try
            Dim ltviOri As TreeViewItem = trvOrionCop.SelectedItem
            Dim ldblInc = 0.0, lenuTipoCalCuo = EnuTipoBaseCalculo.None
            Dim lblnIniMesAct = False, lblnProcesar As Boolean = True
            GblnOK = True
            If GobjParametros.ColAnos.Count > 0 Then
                Dim lobjUltAno As ClsAno = GobjParametros.ColAnos(
                        GobjParametros.ColAnos.Count)
                Dim lstrIdAno = (lobjUltAno.ObjIdAnoShr.ObjValorPro + 1).ToString
                lstrMens = "Creando el Ano " & lstrIdAno & "!"
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                ldblInc = FdblIncrementoCA(lenuTipoCalCuo, lblnProcesar)
            Else
                lstrMens = "Creando el Año inicial!"
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                lblnIniMesAct = FblnInicioMesAct()
            End If
            If lblnProcesar Then
                Mouse.OverrideCursor = Cursors.Wait
                If lenuTipoCalCuo = EnuTipoBaseCalculo.EnuImportadas Then
                    lenuTipoCalCuo = EnuTipoBaseCalculo.None
                End If
                lobjAno = GobjParametros.SCreeAno(lblnIniMesAct, ldblInc, lenuTipoCalCuo)
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
                MblnNodosAnoModificados = False
                SRefresqueNodosAno()
                If lobjAno IsNot Nothing Then
                    If GobjParametros.ObjParametrizacionOkBln.ObjValorPro Then
                        lstrMens = "Se creó el año " & lobjAno.ObjIdAnoShr.ToString
                        If Not lobjAno.FblnDebeCalcularCuotas Then
                            If GobjParametros.ColAnos.Count > 1 Then
                                lstrMens &= ", se calcularon los Módulos y " &
                                        "se generaron las Cuotas de Administración!"
                            Else
                                lstrMens &= "!"
                            End If
                        End If
                    Else
                        Dim lshrIdAno As Short = lobjAno.ObjIdAnoShr.ObjValorPro
                        lstrMens = If(GobjParametros.ColAnos.Count > 1,
                                "Se craron los Años " & lshrIdAno.ToString & " y " &
                                (lshrIdAno + 1).ToString, "Se creó el Año " & lshrIdAno.ToString)
                    End If
                Else
                    lstrMens = "Proceso cancelado por el Usuario!"
                End If
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub
    Private Function FdblIncrementoCA(ByRef aenuTipoCalCuota As EnuTipoBaseCalculo,
            ByRef ablnProcesar As Boolean) As Double
        Dim ldblInc = 0.0
        Dim lwinInc As New WinIncrementoCA()
        Dim NoUsado = lwinInc.ShowDialog()
        ablnProcesar = lwinInc.BlnProcesar
        If ablnProcesar Then
            ldblInc = lwinInc.DblIncCA
            aenuTipoCalCuota = lwinInc.EnuTipoCalculoCuotaAno
        End If
        Return ldblInc
    End Function
    Private Function FblnInicioMesAct() As Boolean
        Dim lwinMesIniFac As New WinMesInicioFact()
        Dim NoUsado = lwinMesIniFac.ShowDialog()
        Return lwinMesIniFac.BlnInicioFacMesActual
    End Function
    Private Sub SAbraAno()
        Dim ltviOri As TreeViewItem = trvOrionCop.SelectedItem
        If ltviOri.Name = My.Resources.Ano Then
            Dim lstrIdAno As String = MobjAno.ObjIdAnoShr.ToString
            MwinVentana = New WinAnos(MobjAno) With {
            .WinPadre = Me
            }
            MwinVentana.ShowDialog()
            MwinVentana = Nothing
        Else
            Dim lstrMens = "Primero debe seleccionar el Nodo correspondiente al Año que quiere abrir!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SSuprimaAno()
        Dim ltviOri As TreeViewItem = trvOrionCop.SelectedItem
        Dim lstrMens = String.Empty
        If ltviOri.Name = My.Resources.Ano Then
            If MsgBox("Realmente desea suprimir el Año seleccionado ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Dim lstrAno = MobjAno.ObjIdAnoShr.ToString, lblnSuprimioAno = False
                MobjAno.SSuprima(lstrMens)
                If String.IsNullOrEmpty(lstrMens) Then
                    MobjAno = Nothing
                    MblnNodosAnoModificados = False
                    SRefresqueNodosAno()
                    lstrMens = "El Año " & lstrAno & " fue suprimido exitosamente!"
                End If
            End If
        Else
            lstrMens = "Primero debe seleccionar el Nodo correspondiente al Año que quiere Suprimir!"
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SAbraServicio()
        Dim ltviOri As TreeViewItem = trvOrionCop.SelectedItem
        If ltviOri.Equals(MtviServiciosPermanentes) OrElse
                    ltviOri.Parent.Equals(MtviServiciosPermanentes) OrElse
                    ltviOri.Name = My.Resources.NombreNodoServicioAno OrElse
                    ltviOri.Name = My.Resources.Ano Then
            Dim ltviPadre As TreeViewItem = ltviOri.Parent
            Dim lstrIdAno = ltviPadre.Tag.ToString
            Dim lshrIdServicio As Short
            If IsNothing(MobjServicio) Then
                lshrIdServicio = 0
            Else
                lshrIdServicio = MobjServicio.ObjIdServicioShr.ObjValorPro
            End If
            If ltviOri.Equals(MtviServiciosPermanentes) OrElse
                    ltviOri.Parent.Equals(MtviServiciosPermanentes) Then
                MwinVentana = New WinServicios(lshrIdServicio, Nothing)
            ElseIf ltviOri.Name = My.Resources.NombreNodoServicioAno OrElse
                    ltviOri.Name = My.Resources.Ano Then
                MwinVentana = New WinServicios(lshrIdServicio, MobjAno)
            End If
            MwinVentana.WinPadre = Me
            MwinVentana.ShowDialog()
            MwinVentana = Nothing
        Else
            Dim lstrMens = "Primero debe seleccionar el Nodo correspondiente al Servicio " &
                    "que quiere abrir!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SVerifiqueInstalacion()
        GobjParametros.SVerifiqueApp(True, True)
    End Sub
    Private Sub SAbraWinAyuda()
        Dim lstrTituloAyuda = String.Empty
        Dim lstrMensAyuda = FstrMensajeAyuda(lstrTituloAyuda)
        If String.IsNullOrEmpty(lstrMensAyuda) Then
            If MfrmAyuda IsNot Nothing Then
                MfrmAyuda.Close()
                MfrmAyuda = Nothing
            End If
            EnuEstadoAyuda = EnuEstadoAyudaDef.EnuOff
            SLevanteEveNoti("No hay ayuda para ser mostrada!", "", 0, EnuSeveridadNot.EnuInformacion)
        Else
            If MfrmAyuda Is Nothing Then
                MfrmAyuda = New FrmAyuda(Me) With {
                    .StrTitulo = lstrTituloAyuda,
                    .StrMensaje = lstrMensAyuda
                    }
            Else
                MfrmAyuda.StrTitulo = lstrTituloAyuda
                MfrmAyuda.StrMensaje = lstrMensAyuda
            End If
            MfrmAyuda.Show()
            ClsOrionCop.SiempreEncima(MfrmAyuda.Handle.ToInt32)
            EnuEstadoAyuda = EnuEstadoAyudaDef.EnuFrmOn
        End If
    End Sub
    Private Sub SRefresqueAyuda()
        If MfrmAyuda IsNot Nothing Then
            Dim lstrTituloAyuda = String.Empty
            Dim lstrMensAyuda = FstrMensajeAyuda(lstrTituloAyuda)
            MfrmAyuda.StrTitulo = lstrTituloAyuda
            MfrmAyuda.StrMensaje = lstrMensAyuda
        End If
    End Sub
    Friend Sub SCerroAyuda()
        MfrmAyuda = Nothing
        EnuEstadoAyuda = EnuEstadoAyudaDef.EnuOff
    End Sub
#End Region
#Region "Manejo de nodos"
    Private Sub SInicialiceNodos()
        Dim lstrTextoTvi = String.Empty
        Dim lstrTrayIcono = String.Empty
        ' Nodo Raiz (Aplicacion)
        lstrTextoTvi = GobjPanorama.ObjAppActual.StrNombreCompleto & My.Resources.App
        lstrTrayIcono = "RecImagenes/aplicacion.png"
        MtviOrionCop = FtviTviPan(lstrTextoTvi, lstrTrayIcono)
        MtviOrionCop.Tag = GobjPanorama.ObjAppActual.StrNombreCompleto
        MtviOrionCop.Name = "tviNodoRaiz"
        trvOrionCop.Items.Add(MtviOrionCop)
        ' Nodo Carpeta
        lstrTextoTvi = "Carpeta: " & MobjCarpeta.ObjNombreStr.ToString
        lstrTrayIcono = "RecImagenes/carpeta.png"
        MtviCarpeta = FtviTviPan(lstrTextoTvi, lstrTrayIcono)
        MtviCarpeta.Tag = MobjCarpeta.ObjNombreStr.ToString
        MtviCarpeta.Name = "tviCarpeta"
        MtviCarpeta.ToolTip = My.Resources.TTSeleccionarAccion
        MtviOrionCop.Items.Add(MtviCarpeta)
        ' Nodo Cuentas Contabilidad
        lstrTextoTvi = My.Resources.CtasConta
        lstrTrayIcono = "RecImagenes/CuentasContabilidad.png"
        MtviPlanContable = FtviTviPan(lstrTextoTvi, lstrTrayIcono)
        MtviPlanContable.Tag = My.Resources.CtasConta
        MtviPlanContable.Name = "tviPanContable"
        MtviPlanContable.ContextMenu = FindResource("RecMnuCarpetaMC")
        MtviPlanContable.ToolTip = My.Resources.TTPlanContable
        MtviCarpeta.Items.Add(MtviPlanContable)
        ' Nodo Copropiedad
        lstrTextoTvi = "Copropiedad: " & MobjCentroUtil.ObjNombreCentroUtilStr.ToString
        lstrTrayIcono = "RecImagenes/CentroUtilidad.png"
        MtviCentroUtilOriCop = FtviTviPan(lstrTextoTvi, lstrTrayIcono)
        MtviCentroUtilOriCop.Tag = MobjCentroUtil.ObjNombreCentroUtilStr.ToString
        MtviCentroUtilOriCop.Name = "tviCenUtilidad"
        MtviCentroUtilOriCop.ContextMenu = FindResource("RecMnuCentroUtilMC")
        MtviCentroUtilOriCop.ToolTip = My.Resources.TTSeleccionarAccion
        MtviCarpeta.Items.Add(MtviCentroUtilOriCop)
        ' Nodo Cuentas Bancarias
        lstrTextoTvi = My.Resources.CtaBanc
        lstrTrayIcono = "RecImagenes/Ctabanco.png"
        MtviBancos = FtviTviPan(lstrTextoTvi, lstrTrayIcono)
        MtviBancos.Tag = My.Resources.CtaBanc
        MtviBancos.Name = "mtviBancos"
        MtviBancos.ContextMenu = FindResource("RecMnuBancosMC")
        MtviCentroUtilOriCop.Items.Add(MtviBancos)
        ' Nodo Documentos
        lstrTextoTvi = My.Resources.Docs
        lstrTrayIcono = "RecImagenes/Documentos.png"
        MtviDocumentos = FtviTviPan(lstrTextoTvi, lstrTrayIcono)
        MtviDocumentos.Tag = My.Resources.Docs
        MtviDocumentos.Name = "mtviDocumentos"
        MtviDocumentos.ContextMenu = FindResource("RecMnuDocumentosMC")
        MtviCentroUtilOriCop.Items.Add(MtviDocumentos)
        ' Nodo Tasas de mora
        lstrTextoTvi = My.Resources.TasasMora
        lstrTrayIcono = "RecImagenes/CondicionesMora.png"
        MtviTasasMora = FtviTviPan(lstrTextoTvi, lstrTrayIcono)
        MtviTasasMora.Tag = My.Resources.TasasMora
        MtviTasasMora.Name = "tviTasasMora"
        MtviTasasMora.ContextMenu = FindResource("RecMnuTasasMoraMC")
        MtviCentroUtilOriCop.Items.Add(MtviTasasMora)
        ' Nodo ServiciosPermanentes
        lstrTextoTvi = My.Resources.ServiciosPerm
        lstrTrayIcono = "RecImagenes/Servicios.png"
        MtviServiciosPermanentes = FtviTviPan(lstrTextoTvi, lstrTrayIcono)
        MtviServiciosPermanentes.Tag = My.Resources.ServiciosPerm
        MtviServiciosPermanentes.Name = "tviServiciosPermanentes"
        MtviServiciosPermanentes.ContextMenu = FindResource("RecMnuServicioPerMC")
        MtviCentroUtilOriCop.Items.Add(MtviServiciosPermanentes)
    End Sub
    Private Sub SRefresqueNodosAno()
        If Not MblnNodosAnoModificados Then
            Dim lcolAnos As Collection = GobjParametros.ColAnos
            Dim ltviAno As TreeViewItem = Nothing
            Dim lstrTrayIcono = "RecImagenes/ano.png"
            MblnModificandoArbol = True
            SQuiteNodosAnos()
            For Each lobjAno As ClsAno In lcolAnos
                Dim lstrTextoTvi = String.Empty
                With lobjAno
                    lstrTextoTvi = .ObjIdAnoShr.ToString
                    ltviAno = FtviTviPan(lstrTextoTvi, lstrTrayIcono)
                    ltviAno.Name = My.Resources.Ano
                    ltviAno.Tag = .ObjIdAnoShr.ObjValorPro
                    ltviAno.ContextMenu = FindResource("RecMnuAnoMC")
                    ltviAno.ToolTip = My.Resources.TTSeleccionarAccion
                    MtviCentroUtilOriCop.Items.Add(ltviAno)
                End With
            Next
            MblnModificandoArbol = False
            MblnNodosAnoModificados = True
        End If
    End Sub
    Private Sub SQuiteNodosAnos()
        Dim ltviNodoAno As TreeViewItem = Nothing
        MblnModificandoArbol = True
        If MtviCentroUtilOriCop.Items.Count > 0 Then
            For i = MtviCentroUtilOriCop.Items.Count - 1 To 0 Step -1
                ltviNodoAno = MtviCentroUtilOriCop.Items(i)
                If ltviNodoAno.Name = My.Resources.Ano Then
                    MtviCentroUtilOriCop.Items.Remove(ltviNodoAno)
                End If
            Next
        End If
        MblnModificandoArbol = False
    End Sub
    Private Sub SPuebleNodoServiciosPer()
        Dim lcolServPer As Collection = GobjParametros.ColServiciosPer
        MtviServiciosPermanentes.Items.Clear()
        For Each lobjServPer As ClsServicio In lcolServPer
            With lobjServPer
                Dim lstrTextoTvi = .ObjIdServicioShr.ToString & " - " &
                        .ObjConceptoServicioStr.ObjValorPro
                Dim ltviSerPer = FtviTviPan(lstrTextoTvi, String.Empty)
                ltviSerPer.Tag = .ObjIdServicioShr.ObjValorPro
                ltviSerPer.ContextMenu = FindResource("RecMnuServicioPerMC")
                ltviSerPer.ToolTip = My.Resources.TTServicio
                MtviServiciosPermanentes.Items.Add(ltviSerPer)
            End With
        Next
        MobjServicio = Nothing
    End Sub
    Private Sub SPuebleNodoAno(ashrIdAno As Short)
        Dim ltviOri As TreeViewItem = trvOrionCop.SelectedItem
        If ltviOri.Tag <> ashrIdAno Then
            Throw New ErrorInesperadoPanLException("El argumento pasado no coincide con el nodo seleccionado")
        End If
        Dim lcolServAno As Collection = MobjAno.ColServiciosAno
        ltviOri.Items.Clear()
        For Each lobjServAno As ClsServicio In lcolServAno
            With lobjServAno
                Dim lstrTextoTvi = .ObjIdServicioShr.ToString & " - " &
                        .ObjNombreServicioStr.ObjValorPro
                Dim ltviSerAno = FtviTviPan(lstrTextoTvi, String.Empty)
                ltviSerAno.Tag = .ObjIdServicioShr.ObjValorPro
                ltviSerAno.Name = My.Resources.NombreNodoServicioAno
                ltviSerAno.ContextMenu = FindResource("RecMnuServicioAnoMC")
                ltviSerAno.ToolTip = My.Resources.TTServicio
                ltviOri.Items.Add(ltviSerAno)
            End With
        Next
        MobjServicio = Nothing
    End Sub
#End Region
#Region "Manejo de Menues"
    Private Sub SCargueMenues()
        Dim lblnInsertarAcciones = True
        Dim ltviOri As TreeViewItem = trvOrionCop.SelectedItem
        If MenuVen.Items.Count > 4 Then
            Dim lmnuItem As MenuItem = MenuVen.Items(1)
            MenuVen.Items.Remove(lmnuItem)
        End If
        MnuAccionesDeObjeto.Items.Clear()
        Select Case True
            Case ltviOri.Equals(MtviPlanContable)
                SPuebleMenuPlanContable()
            Case ltviOri.Equals(MtviCentroUtilOriCop)
                SPuebleMenuCentroUtil()
            Case ltviOri.Equals(MtviBancos)
                SPuebleMenuBancos()
            Case ltviOri.Equals(MtviTasasMora)
                SPuebleMenuTasasMora()
            Case ltviOri.Equals(MtviDocumentos)
                SPuebleMenuDocumentos()
            Case ltviOri.Parent.Equals(MtviServiciosPermanentes), ltviOri.Equals(MtviServiciosPermanentes)
                SPuebleMenuServicioPer()
            Case ltviOri.Name = My.Resources.Ano
                SPuebleMenuAno()
            Case Else
                lblnInsertarAcciones = False
        End Select
        If lblnInsertarAcciones Then
            MenuVen.Items.Insert(1, MnuAccionesDeObjeto)
        End If
    End Sub
    Private Sub SPuebleMenuPlanContable()
        MnuAccionesDeObjeto.Header = "_Acciones del Plan Contable"
        MnuAccionesDeObjeto.Items.Add(MnuAbrirCtaCon)
    End Sub
    Private Sub SPuebleMenuCentroUtil()
        Dim lsepSepaCenUtil_0 As New Separator
        Dim lsepSepaCenUtil_1 As New Separator
        MnuAccionesDeObjeto.Header = "_Acciones Copropiedad"
        With MnuAccionesDeObjeto.Items
            .Add(MnuOpcionesCentroUtil)
            .Add(lsepSepaCenUtil_0)
            .Add(MnuCrearAno)
            .Add(lsepSepaCenUtil_1)
        End With
    End Sub
    Private Sub SPuebleMenuTasasMora()
        MnuAccionesDeObjeto.Header = "_Acciones de Tasas de Mora"
        MnuAccionesDeObjeto.Items.Add(MnuAbrirTasasMora)
    End Sub
    Private Sub SPuebleMenuBancos()
        MnuAccionesDeObjeto.Header = "_Acciones de Documentos"
        MnuAccionesDeObjeto.Items.Add(MnuAbrirBancos)
    End Sub
    Private Sub SPuebleMenuDocumentos()
        MnuAccionesDeObjeto.Header = "_Acciones de Documentos"
        MnuAccionesDeObjeto.Items.Add(MnuAbrirDocumentos)
    End Sub
    Private Sub SPuebleMenuServicioPer()
        MnuAccionesDeObjeto.Header = "Acciones de Servicio Permanente"
        MnuAccionesDeObjeto.Items.Add(MnuAbrirServicioPer)
    End Sub
    Private Sub SPuebleMenuAno()
        Dim lsepAno As New Separator
        MnuAccionesDeObjeto.Header = "Acciones de Añ_o"
        With MnuAccionesDeObjeto.Items
            .Add(MnuAbrirAno)
            .Add(MnuSuprimirAno)
            .Add(lsepAno)
            .Add(MnuAbrirServicioAno)
        End With
    End Sub
    Private Sub SAsigneMenuesContextuales()
        SAsigneMenuContexCarpeta()
        SAsigneMenuContexCenUtiliOriCop()
        SAsigneMenuContexTasasMora()
        SAsigneMenuContexDocumento()
        SAsigneMenuContexBanco()
        SAsigneMenuServicioPer()
        SAsigneMenuServicioAno()
        SAsigneMenuAno()
    End Sub
    Private Sub SAsigneMenuContexCarpeta()
        Dim lmnuMenuContextual As ContextMenu = FindResource("RecMnuCarpetaMC")
        For Each lobjObjetoMenu As Object In lmnuMenuContextual.Items
            If TypeOf lobjObjetoMenu Is MenuItem Then
                Dim lmnuItemMenuCont As MenuItem = lobjObjetoMenu
                If lmnuItemMenuCont.Name = "MnuAbrirCtaConC" Then
                    MnuAbrirCtaConC = lmnuItemMenuCont
                End If
            End If
        Next
    End Sub
    Private Sub SAsigneMenuContexCenUtiliOriCop()
        Dim lmnuItemMenuCont As MenuItem
        Dim lmnuMenuContextual As ContextMenu = FindResource("RecMnuCentroUtilMC")
        For Each lobjObjetoMenu As Object In lmnuMenuContextual.Items
            If TypeOf lobjObjetoMenu Is MenuItem Then
                lmnuItemMenuCont = lobjObjetoMenu
                Select Case lmnuItemMenuCont.Name
                    Case Is = "MnuOpcionesCentroUtilC"
                        MnuAbrirConfigC = lmnuItemMenuCont
                    Case Is = "MnuCrearAnoC"
                        MnuCrearAnoC = lmnuItemMenuCont
                    Case Is = "MnuCrearAnoC"
                End Select
            End If
        Next
    End Sub
    Private Sub SAsigneMenuContexTasasMora()
        Dim lmnuMenuContextual As ContextMenu
        lmnuMenuContextual = FindResource("RecMnuTasasMoraMC")
        For Each lobjObjetoMenu As Object In lmnuMenuContextual.Items
            If TypeOf lobjObjetoMenu Is MenuItem Then
                Dim lmnuItemMenuCont As MenuItem = lobjObjetoMenu
                If lmnuItemMenuCont.Name = "MnuAbrirTasaMoraC" Then
                    MnuAbrirTasasMoraC = lmnuItemMenuCont
                End If
            End If
        Next
    End Sub
    Private Sub SAsigneMenuContexBanco()
        Dim lmnuMenuContextual As ContextMenu
        lmnuMenuContextual = FindResource("RecMnuBancosMC")
        For Each lobjObjetoMenu As Object In lmnuMenuContextual.Items
            If TypeOf lobjObjetoMenu Is MenuItem Then
                Dim lmnuItemMenuCont As MenuItem = lobjObjetoMenu
                If lmnuItemMenuCont.Name = "MnuAbrirBancosC" Then
                    MnuAbrirBancosC = lmnuItemMenuCont
                End If
            End If
        Next
    End Sub
    Private Sub SAsigneMenuContexDocumento()
        Dim lmnuMenuContextual As ContextMenu
        lmnuMenuContextual = FindResource("RecMnuDocumentosMC")
        For Each lobjObjetoMenu As Object In lmnuMenuContextual.Items
            If TypeOf lobjObjetoMenu Is MenuItem Then
                Dim lmnuItemMenuCont As MenuItem = lobjObjetoMenu
                If lmnuItemMenuCont.Name = "MnuAbrirDocumentosC" Then
                    MnuAbrirDocumentosC = lmnuItemMenuCont
                End If
            End If
        Next
    End Sub
    Private Sub SAsigneMenuServicioPer()
        Dim lmnuMenuContextual As ContextMenu
        lmnuMenuContextual = FindResource("RecMnuServicioPerMC")
        For Each lobjObjetoMenu As Object In lmnuMenuContextual.Items
            If TypeOf lobjObjetoMenu Is MenuItem Then
                Dim lmnuItemMenuCont As MenuItem = lobjObjetoMenu
                If lmnuItemMenuCont.Name = "MnuAbrirServicioPerC" Then
                    MnuAbrirServicioPerC = lmnuItemMenuCont
                End If
            End If
        Next
    End Sub
    Private Sub SAsigneMenuServicioAno()
        Dim lmnuMenuContextual As ContextMenu
        lmnuMenuContextual = FindResource("RecMnuServicioAnoMC")
        For Each lobjObjetoMenu As Object In lmnuMenuContextual.Items
            If TypeOf lobjObjetoMenu Is MenuItem Then
                Dim lmnuItemMenuCont As MenuItem = lobjObjetoMenu
                If lmnuItemMenuCont.Name = "MnuAbrirServicioAnoC" Then
                    MnuAbrirServicioAnoC = lmnuItemMenuCont
                End If
            End If
        Next
    End Sub
    Private Sub SAsigneMenuAno()
        Dim lmnuMenuContextual As ContextMenu
        lmnuMenuContextual = FindResource("RecMnuAnoMC")
        For Each lobjObjetoMenu In lmnuMenuContextual.Items
            If TypeOf lobjObjetoMenu Is MenuItem Then
                Dim lmnuItemMenuCont As MenuItem = lobjObjetoMenu
                Select Case lmnuItemMenuCont.Name
                    Case Is = "MnuAbrirAnoC"
                        MnuAbrirAnoC = lmnuItemMenuCont
                    Case Is = "MnuSuprimirAnoC"
                        MnuSuprimirAnoC = lmnuItemMenuCont
                    Case Is = "MnuAbrirServicioAnoCN"
                        MnuAbrirServicioAnoCN = lmnuItemMenuCont
                End Select
            End If
        Next
    End Sub
#End Region
#End Region

#Region "Mostrar Objetos"
    Private Sub SMuestreServicioPer(ashrIdServicio As Short)
        lvOrionCop.Items.Clear()
        Dim lstrKey = "0" & "," & ashrIdServicio.ToString
        MobjServicio = GobjParametros.ColServiciosPer(lstrKey)
        If Not IsNothing(MobjServicio) AndAlso MobjServicio.BlnExiste Then
            SMuestreServicio()
        End If
    End Sub
    Private Sub SMuestreServicioAno(ashrIdServicio As Short)
        lvOrionCop.Items.Clear()
        Dim lstrKey = MobjAno.ObjIdAnoShr.ToString & "," & ashrIdServicio.ToString
        MobjServicio = MobjAno.ColServiciosAno(lstrKey)
        If Not IsNothing(MobjServicio) AndAlso MobjServicio.BlnExiste Then
            SMuestreServicio()
        End If
    End Sub
    Private Sub SMuestreServicio()
        lvOrionCop.Items.Clear()
        With MobjServicio
            lvOrionCop.Items.Add(New ClsViewItemPar("Id. Servicio", .ObjIdServicioShr.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar("Nombre del Servicio",
                    .ObjConceptoServicioStr.ObjValorPro))
            If .ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual Then
                lvOrionCop.Items.Add(New ClsViewItemPar("Tipo de Servicio", "Anual"))
            ElseIf .ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuPermanente Then
                lvOrionCop.Items.Add(New ClsViewItemPar("Tipo de Servicio", "Permanente"))
            End If
            Dim lstrCausaMora = If(.FblnCausaMora, "Si", "No")
            lvOrionCop.Items.Add(New ClsViewItemPar("Causa Intereses de Mora ?", lstrCausaMora))
            If .ObjEsFactProgramableBln.ObjValorPro Then
                lvOrionCop.Items.Add(New ClsViewItemPar("La facturación es Programable", "Si"))
            Else
                lvOrionCop.Items.Add(New ClsViewItemPar("La facturación es Programable", "No"))
            End If
            If .ObjGeneraProgramBln.ObjValorPro Then
                lvOrionCop.Items.Add(New ClsViewItemPar("Período de inicio de la Facturación",
                            .ObjPeriodoInicioStr.ObjValorPro))
                lvOrionCop.Items.Add(New ClsViewItemPar("Cantidad de Períodos a Facturar",
                            .ObjCantPeriodos_ServicioShr.ToString))
                If .ObjEstaGenaradaProgramBln.ObjValorPro Then
                    lvOrionCop.Items.Add(New ClsViewItemPar("Está generada la Programación de las Facturas ?", "Si"))
                Else
                    lvOrionCop.Items.Add(New ClsViewItemPar("Está generada la Programación de las Facturas ?", "No"))
                End If
                If .BlnEsCuotaAdministracion Then
                    If Not .ObjEsAjusteBln.ObjValorPro Then
                        If .ObjEstaAjustadoBln.ObjValorPro Then
                            lvOrionCop.Items.Add(New ClsViewItemPar("Está generado el Retroactivo ?", "Si"))
                        Else
                            lvOrionCop.Items.Add(New ClsViewItemPar("Está generado el Retroactivo ?", "No"))
                        End If
                    Else
                        lvOrionCop.Items.Add(New ClsViewItemPar("Es Ajuste de las Cuotas de Administración ?", "Si"))
                    End If
                End If
            End If
            lvOrionCop.Items.Add(New ClsViewItemPar(" ", " "))
            If .ObjEsAjusteBln.ObjValorPro Then
                lvOrionCop.Items.Add(New ClsViewItemPar("Valor del Servicio", Format(.DecValorAjuste, "c")))
            Else
                lvOrionCop.Items.Add(New ClsViewItemPar("Valor del Servicio", Format(.FdecValor, "c")))
            End If
            If .ObjGeneraProgramBln.ObjValorPro Then
                Dim lshrIdAno As Short = MobjServicio.ObjIdAno_ServicioShr.ObjValorPro
                Dim lshrIdServicio As Short = MobjServicio.ObjIdServicioShr.ObjValorPro
                lvOrionCop.Items.Add(New ClsViewItemPar("Valor Calculado para la cantidad de Períodos",
                        Format(ClsOrionCop.FdecValorTotalCalculadoServicio(lshrIdAno, lshrIdServicio),
                        "c")))
            End If
            lvOrionCop.Items.Add(New ClsViewItemPar(" ", " "))
            lvOrionCop.Items.Add(New ClsViewItemPar("Código Cuenta Débito",
                    .ObjCodigoCuentaDbStr.ObjValorPro & " - " & .ObjCodigoCuentaDbStr.StrNombreCuenta))
            lvOrionCop.Items.Add(New ClsViewItemPar("Código Cuenta Crédito",
                    .ObjCodigoCuentaCrStr.ObjValorPro & " - " & .ObjCodigoCuentaCrStr.StrNombreCuenta))
            lvOrionCop.Items.Add(New ClsViewItemPar("Tipo Tercero Cuenta Crédito",
                    .ObjIdTipoTerCtaCrSerByt.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar("Tercero Cuenta Crédito",
                    .ObjIdTerceroCtaCrDbl.ToString & " - " & .ObjIdTerceroCtaCrDbl.StrNombreTercero))
            lvOrionCop.Items.Add(New ClsViewItemPar("Código Cuenta Intereses de Mora",
                    .ObjCodigoCuentaMoraStr.ObjValorPro & " - " & .ObjCodigoCuentaMoraStr.StrNombreCuenta))
            lvOrionCop.Items.Add(New ClsViewItemPar("Código Cuenta Devolución",
                    .ObjCodigoCuentaDevStr.ObjValorPro & " - " & .ObjCodigoCuentaDevStr.StrNombreCuenta))
            lvOrionCop.Items.Add(New ClsViewItemPar(" ", " "))
            lvOrionCop.Items.Add(New ClsViewItemPar("Código Cuenta IVA",
                    .ObjCodigoCuentaIvaStr.ObjValorPro & " - " & .ObjCodigoCuentaIvaStr.StrNombreCuenta))
            lvOrionCop.Items.Add(New ClsViewItemPar("Tarifa del IVA", .ObjTarifaIvaDbl.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar("Tarifa Retención el Fuente", .ObjTarifaRetFteDbl.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar("Base ReteFuente mínima", .ObjBaseMinimaReteFuenteDec.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar("Tarifa Retención Industria y Comercio", .ObjTarifaRetIcaDbl.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar("Base ReteIca mínima", .ObjBaseMinimaReteIcaDec.ToString))
        End With
    End Sub
    Private Sub SMuestreAno(ashrIdAno As Short)
        lvOrionCop.Items.Clear()
        MobjAno = GobjParametros.ColAnos(ashrIdAno.ToString)
        With MobjAno
            lvOrionCop.Items.Add(New ClsViewItemPar("Año", .ObjIdAnoShr.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar("Está cerrado el Año ?",
                    ClsPanorama.FstrBuleanoToString(.ObjEstaCerradoAnoBln.ObjValorPro)))
            lvOrionCop.Items.Add(New ClsViewItemPar("Presupuesto de Ingresos del Año",
                    Format(.ObjValorPres_AnoDec.ObjValorPro, "c")))
            lvOrionCop.Items.Add(New ClsViewItemPar("Presupuesto Calculado", Format(.FdecValorCuotaAdminAno, "c")))
            lvOrionCop.Items.Add(New ClsViewItemPar(" ", " "))
            If .FblnEstaGenCuota Then
                lvOrionCop.Items.Add(New ClsViewItemPar("Está generada la Programación de las Facturas ?", "Si"))
            Else
                lvOrionCop.Items.Add(New ClsViewItemPar("Está generada la Programación de las Facturas ?", "No"))
            End If
            If .FblnEstaAjustada Then
                lvOrionCop.Items.Add(New ClsViewItemPar("Está generado el Retroactivo ?", "Si"))
            Else
                lvOrionCop.Items.Add(New ClsViewItemPar("Está generado el Retroactivo  ?", "No"))
            End If
        End With
    End Sub
    Private Sub SMuestreApp()
        lvOrionCop.Items.Clear()
        With GobjPanorama.ObjAppActual
            lvOrionCop.Items.Add(New ClsViewItemPar("Nombre Aplicación", .StrNombreCompleto))
            lvOrionCop.Items.Add(New ClsViewItemPar("Versión Instalada", .ObjVersionStr.ObjValorPro))
            lvOrionCop.Items.Add(New ClsViewItemPar("Licencia", .StrNombreLicencia))
            lvOrionCop.Items.Add(New ClsViewItemPar("Fecha Instalación", .ObjFechaInstalacionDtm.ObjValorPro.ToLongDateString))
            lvOrionCop.Items.Add(New ClsViewItemPar("Fecha Actualización", .ObjFechaActualizacionDtm.ObjValorPro.ToLongDateString))
        End With
    End Sub
    Private Sub SMuestreCarpeta()
        lvOrionCop.Items.Clear()
        With MobjCarpeta
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.IdCarpeta, .ObjIdCarpetaShr.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.NombreCarpeta, .ObjNombreStr.ToString))
        End With
    End Sub
    Private Sub SMuestreCtasCont()
        gvcPropiedad.Header = "Cuenta Contable"
        gvcValor.Header = "Nombre"
        lvOrionCop.Items.Clear()
        Dim lstrIdCtaCon As String, lstrNomCtaCon As String
        For Each lobjCtaCon As ClsCuentaContabilidad In GobjPanorama.ObjCarpetaActual.FcolCuentasCont
            lstrIdCtaCon = lobjCtaCon.ObjIdCuentaContStr.ObjValorPro
            lstrNomCtaCon = lobjCtaCon.ObjNombreCuentaStr.ObjValorPro
            lvOrionCop.Items.Add(New ClsViewItemPar(lstrIdCtaCon, lstrNomCtaCon))
        Next
    End Sub
    Private Sub SMuestreTasas()
        gvcPropiedad.Header = "Fecha Desde"
        gvcValor.Header = "Tasa"
        lvOrionCop.Items.Clear()
        Dim ldtmFecha As Date, ldblTasa As Double
        Dim lstrFecha As String, lstrTasa As String
        Dim ldtbTasasMora = GobjParametros.FdtbTasasMora
        For Each ldrwTasa As DataRow In ldtbTasasMora.Rows
            ldtmFecha = ClsPanorama.FobjValorCampo(ldrwTasa(
                    ClsFechaDesdeTasaMoraDtm.SstrNombreCampoBd), EnuTipoValor.enuDate)
            lstrFecha = Format(ldtmFecha, GCSTRFMTFECHASIMPLE)
            ldblTasa = ClsPanorama.FobjValorCampo(ldrwTasa(ClsTasaMoraDbl.SstrNombreCampoBd),
                    EnuTipoValor.enuDouble)
            lstrTasa = Format(ldblTasa, "#0.000%")
            lvOrionCop.Items.Add(New ClsViewItemPar(lstrFecha, lstrTasa))
        Next
    End Sub
    Private Sub SMuestreBancos()
        gvcPropiedad.Header = "Nombre Banco"
        gvcValor.Header = "Número de la cuenta"
        lvOrionCop.Items.Clear()
        Dim lstrBanco As String
        Dim lstrInf As String
        Dim lcolCuentasBanco = GobjParametros.FcolCuentasBanco
        For Each lobjCtaBanco As ClsCuentaBanco In lcolCuentasBanco
            lstrBanco = lobjCtaBanco.ObjNombreBancoStr.ToString()
            lstrInf = lobjCtaBanco.ObjNumeroCuentaStr.ToString()
            lvOrionCop.Items.Add(New ClsViewItemPar(lstrBanco, lstrInf))
        Next
    End Sub
    Private Sub SMuestreDocs()
        gvcPropiedad.Header = "Documento"
        gvcValor.Header = "Doc.Con / Prefijo/ Nro. Inicial"
        lvOrionCop.Items.Clear()
        Dim lstrDoc As String, lstrDocCon As String, lstrPref As String, lstrNumIni As String
        Dim lstrInf As String
        For Each lobjDocCon As ClsDocumento In GobjParametros.ColDocumentos
            lstrDoc = lobjDocCon.ObjNombre_DocStr.ToString
            lstrDocCon = lobjDocCon.ObjTipoDocumentoStr.ToString
            lstrPref = lobjDocCon.ObjPrefijo_DocStr.ToString
            lstrNumIni = lobjDocCon.ObjNumeroInicial_DocEnt.ToString
            If String.IsNullOrEmpty(lstrDocCon) Then lstrDocCon = " "
            If String.IsNullOrEmpty(lstrPref) Then lstrPref = " "
            If String.IsNullOrEmpty(lstrNumIni) Then lstrNumIni = "0"
            lstrInf = lstrDocCon & " / " & lstrPref & " / " & lstrNumIni.ToString
            lvOrionCop.Items.Add(New ClsViewItemPar(lstrDoc, lstrInf))
        Next
    End Sub
    Private Sub SMuestreCenUtilidad()
        lvOrionCop.Items.Clear()
        With MobjCentroUtil
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.IdCenUtil, .ObjIdCentroUtilShr.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.NomCenUti, .ObjNombreCentroUtilStr.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.IdTerCenUti, .ObjIdTerceroCentroUtilDbl.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.NomTerCenUti, .ObjTerceroCentroUtilidad.FstrNombreCompleto()))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.IdRepLegal, .ObjIdTerceroRepLegalDbl.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.NomRepLeg, .ObjIdTerceroRepLegalDbl.StrNombreRepLegal))
        End With
        SMuestreConfiguracion()
    End Sub
    Private Sub SAbraConfiguracion()
        MwinVentana = New WinCentroUtilidadOriCop With {
            .WinPadre = Me
        }
        If Not GobjParametros.FblnEstanTodosOk Then
            If ClsOrionCop.FblnExisteRegCenUtil Then
                MwinVentana.EnuOperacionEnWin = EnuOperacionEnVentana.CenuModificando
            Else
                MwinVentana.EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando
            End If
        End If
        MwinVentana.ShowDialog()
        MwinVentana = Nothing
        SMuestreCenUtilidad()
        SVerifiqueInstalacion()
    End Sub
    Private Sub SMuestreConfiguracion()
        With GobjParametros
            lvOrionCop.Items.Add(New ClsViewItemPar("", ""))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.TotalAreaCoprop,
                    Format(.ObjTotalAreaCopropDec.ObjValorPro, "###,##0.00")))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.TotalAreaPond,
                    Format(.ObjTotalAreaPondDec.ObjValorPro, "###,##0.00")))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.DiasPersuasivo,
                    .ObjDiasParaPersuasivoShr.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.DiasPreju,
                    .ObjDiasParaPrejuridicoShr.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.DiasJuri,
                    .ObjDiasParaJuridicoShr.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.DiasPerdida,
                    .ObjDiasParaPerdidaShr.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.BaseRedCP,
                    .ObjBaseRedondeoCPByt.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.BaseRedIM,
                    .ObjBaseRedondeoIntMoraDbl.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.BaseRedGral,
                    .ObjBaseRedondeoGeneralDbl.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.PlazoFacManual,
                    .ObjPlazoDefectoFacManualShr.ToString))
            lvOrionCop.Items.Add(New ClsViewItemPar("", ""))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.IdCtaCaja,
                    .ObjIdCtaCajaStr.ToString & " - " & .ObjIdCtaCajaStr.StrNombreCuenta))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.IdCtaInMoDb,
                    .ObjIdCtaIntMoraDbStr.ToString & " - " & .ObjIdCtaIntMoraDbStr.StrNombreCuenta))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.IdCtaAntRec,
                    .ObjIdCtaAnticiposRecibidosStr.ToString & " - " &
                    .ObjIdCtaAnticiposRecibidosStr.StrNombreCuenta))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.IdCtaIngPorIden,
                    .ObjIdCtaIngPorIdentificarStr.ToString & " - " &
                    .ObjIdCtaIngPorIdentificarStr.StrNombreCuenta))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.IdCtaDctoPP,
                    .ObjIdCtaDescuentosPPStr.ToString & " - " &
                    .ObjIdCtaDescuentosPPStr.StrNombreCuenta))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.IdCtaRetFte,
                    .ObjIdCtaReteFuenteStr.ToString & " - " &
                    .ObjIdCtaReteFuenteStr.StrNombreCuenta))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.IdCtaRetIva,
                    .ObjIdCtaReteIvaStr.ToString & " - " & .ObjIdCtaReteIvaStr.StrNombreCuenta))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.IdCtaRetIca,
                    .ObjIdCtaReteIcaStr.ToString & " - " & .ObjIdCtaReteIcaStr.StrNombreCuenta))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.IdCtaImptosAsumidos,
                    .ObjIdCtaImptosAsumidosStr.ToString & " - " &
                    .ObjIdCtaImptosAsumidosStr.StrNombreCuenta))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.AppCont,
                    .ObjIdAppContableByt.StrNombreApp))
            lvOrionCop.Items.Add(New ClsViewItemPar(My.Resources.TerCtaCajaBancos,
                    .ObjTipoTerceroCajaByt.ToString))
        End With
    End Sub
#End Region

#Region "Eventos en la Ventana"
#Region "Eventos Menues y Botones"
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError As Boolean
        If TypeOf lelmElemento Is MenuItem Then
            Try
                MwinVentana = Nothing
                If Not FblnEjecutoMenu(lelmElemento, lstrMens) Then
                    Select Case lelmElemento.Name
                    ' Tablas Auxiliares
                        Case "MnuSectores"
                            MwinVentana = New WinSectores()
                        Case "MnuModulos"
                            MwinVentana = New WinModulosContribucion
                        Case "MnuTerceros"
                            Dim lobjOrionCop = New ClsOrionCop(GCOBJREGISTRO, False)
                            MwinVentana = New WinTerceros()
                        Case "MnuCuentaOrigenEmail"
                            MwinVentana = New WinCuentaCorreoOrigen()
                        Case "MnuUbicaciones"
                            MwinVentana = New WinUbicacion()
                        Case "MnuOpcionesBK"
                            MwinVentana = New WinOpcionesCopiasSeg()
                        ' Manues adicionales
                        Case "MnuAbrirCtaCon"
                            MwinVentana = New WinCuentasContabilidad()
                        Case "MnuOpcionesCentroUtil"
                            SAbraConfiguracion()
                        ' Cuentas bancarias
                        Case "MnuAbrirBancos"
                            MwinVentana = New WinCuentasBanco
                        ' Tasas de Mora
                        Case "MnuAbrirTasasMora"
                            MwinVentana = New WinTasasMora()
                        Case "MnuAbrirDocumentos"
                            MwinVentana = New WinDocumentos
                        ' Años
                        Case "MnuAbrirAno"
                            SAbraAno()
                        Case "MnuCrearAno"
                            SCreeAno()
                        Case "MnuSuprimirAno"
                            SSuprimaAno()
                    End Select
                End If
                If Not IsNothing(MwinVentana) Then
                    MwinVentana.WinPadre = Me
                    MwinVentana.ShowDialog()
                    MwinVentana = Nothing
                    SRefresqueWin()
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
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuInformacion)
                    End If
                Else
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
    End Sub
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            Dim lbttAyuda As Button = lelmElemento
            If lbttAyuda.Name = "bttAbrirWinAyuda" Then
                SAbraWinAyuda()
                If EnuEstadoAyuda = EnuEstadoAyudaDef.EnuOn Then
                    bttAbrirWinAyuda.IsEnabled = False
                End If
            End If
            End If
    End Sub
    Private Function FblnEjecutoMenu(amnuMenuItem As MenuItem, ByRef astrMens As String) As Boolean
        Dim lblnEjecutoMenu = True
        Select Case amnuMenuItem.Name
            ' Manejo del arbol
            Case "MnuExpandirTodo"
                SExpandaTodo(MtviOrionCop)
            Case "MnuExpandirRama"
                If Not IsNothing(trvOrionCop.SelectedItem) Then
                    SExpandaRama(trvOrionCop.SelectedItem)
                End If
            Case "MnuContraerTodo"
                SContraigaTodo(MtviOrionCop)
            Case "MnuContraerRama"
                If Not IsNothing(trvOrionCop.SelectedItem) Then
                    SContraigaRama(trvOrionCop.SelectedItem)
                End If
                ' Servicios
            Case "MnuAbrirServicioPer", "MnuAbrirServicioAno"
                If GobjParametros.ColAnos.Count = 0 Then
                    astrMens = "Antes de abrir los servicios es necesario haber " &
                                    "creado previamente un año!"
                Else
                    SAbraServicio()
                End If
            Case Else
                lblnEjecutoMenu = False
        End Select
        Return lblnEjecutoMenu
    End Function
    Private Sub MnuContextual_Click(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError As Boolean
        If TypeOf lelmElemento Is MenuItem Then
            Try
                Select Case True
                    ' Manues adicionales
                    Case sender.Equals(MnuAbrirCtaConC)
                        MwinVentana = New WinCuentasContabilidad()
                    Case sender.Equals(MnuAbrirConfigC)
                        SAbraConfiguracion()
                        ' Cuentas bancarias
                    Case sender.Equals(MnuAbrirBancosC)
                        MwinVentana = New WinCuentasBanco()
                    ' Tasas de Mora
                    Case sender.Equals(MnuAbrirTasasMoraC)
                        MwinVentana = New WinTasasMora()
                    ' Documentos
                    Case sender.Equals(MnuAbrirDocumentosC)
                        MwinVentana = New WinDocumentos
                    ' Años
                    Case sender.Equals(MnuAbrirAnoC)
                        SAbraAno()
                    Case sender.Equals(MnuCrearAnoC)
                        SCreeAno()
                    Case sender.Equals(MnuSuprimirAnoC)
                        SSuprimaAno()
                    ' Servicios
                    Case sender.Equals(MnuAbrirServicioPerC), sender.Equals(MnuAbrirServicioAnoC)
                        If GobjParametros.ColAnos.Count = 0 Then
                            lstrMens = "Antes de abrir los servicios es necesario haber " &
                                    "creado previamente un año!"
                        Else
                            SAbraServicio()
                        End If
                    Case sender.Equals(MnuAbrirServicioAnoCN)
                        If MobjAno Is Nothing Then
                            lstrMens = "Antes de abrir los servicios es necesario haber " &
                                    "creado previamente un año!"
                        Else
                            SAbraServicio()
                        End If
                End Select
                If Not IsNothing(MwinVentana) Then
                    MwinVentana.WinPadre = Me
                    MwinVentana.ShowDialog()
                    MwinVentana = Nothing
                Else
                    SRefresqueWin()
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
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuInformacion)
                    End If
                Else
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
    End Sub
#End Region
#Region "Evento de los controles"
    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            ltxtTextBox.SelectAll()
        End If
    End Sub
    Private Sub OnCambioSeleccion(sender As Object, e As RoutedPropertyChangedEventArgs(Of System.Object))
        Dim lelmElemento As FrameworkElement = Nothing
        If Not IsNothing(e) Then
            lelmElemento = CType(e.NewValue, TreeViewItem)
        ElseIf TypeOf sender Is TreeViewItem Then
            lelmElemento = sender
        End If
        gvcPropiedad.Header = "Propiedad"
        gvcValor.Header = "Valor"
        If Not IsNothing(lelmElemento) Then
            If Not (MblnModificandoArbol OrElse MblnRefrescando) Then
                Dim ltviOri As TreeViewItem = lelmElemento
                SExpandaRama(ltviOri)
                txtTrayectoria.Text = FstrTrayectoria(ltviOri)
                lvOrionCop.Items.Clear()
                Select Case True
                    Case ltviOri.Equals(MtviOrionCop)
                        SMuestreApp()
                    Case ltviOri.Equals(MtviCarpeta)
                        SMuestreCarpeta()
                    Case ltviOri.Equals(MtviPlanContable)
                        SMuestreCtasCont()
                    Case ltviOri.Equals(MtviCentroUtilOriCop)
                        SMuestreCenUtilidad()
                        SRefresqueNodosAno()
                    Case ltviOri.Equals(MtviBancos)
                        SMuestreBancos()
                    Case ltviOri.Equals(MtviTasasMora)
                        SMuestreTasas()
                    Case ltviOri.Equals(MtviDocumentos)
                        SMuestreDocs()
                    Case ltviOri.Equals(MtviServiciosPermanentes)
                        SPuebleNodoServiciosPer()
                    Case ltviOri.Parent.Equals(MtviServiciosPermanentes)
                        SMuestreServicioPer(ltviOri.Tag)
                    Case ltviOri.Name = My.Resources.Ano
                        SMuestreAno(ltviOri.Tag)
                        SPuebleNodoAno(ltviOri.Tag)
                    Case ltviOri.Name = My.Resources.NombreNodoServicioAno
                        Dim ltviAnoPadre As TreeViewItem = ltviOri.Parent
                        MobjAno = GobjParametros.ColAnos(ltviAnoPadre.Tag.ToString)
                        SMuestreServicioAno(ltviOri.Tag)
                End Select
                SCargueMenues()
                SLevanteEveNoti("", "", 0, EnuSeveridadNot.EnuOk)
                SVerifiqueInstalacion()
            End If
        End If
    End Sub
    Private Sub ClsFormInterface_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyboardDevice.Modifiers = ModifierKeys.Control Then
            If e.Key = Key.E Then
                SExpandaTodo(MtviOrionCop)
            End If
            If e.Key = Key.C Then
                SContraigaTodo(MtviOrionCop)
            End If
        End If
    End Sub
    Private Sub GridSplitter_MouseEnter(sender As Object, e As MouseEventArgs)
        Mouse.OverrideCursor = Cursors.ScrollWE
    End Sub
    Private Sub GridSplitter_MouseLeave(sender As Object, e As MouseEventArgs)
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Sub ClsFormInterface_Closing(sender As Object, e As CancelEventArgs)
        Dim lwinVentanaPro As MWOrionCop = WinPadre
        lwinVentanaPro.SRestablezcaWinAyuda()
    End Sub
#End Region
#End Region
End Class