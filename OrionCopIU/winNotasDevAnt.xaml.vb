Imports System.Windows.Controls
Imports System.Data
Public Class WinNotasDevAnt
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuFechaDev
        enuIdCliente
        enuIdPreAgru
        enuIdAnticipo
        enuValor
        enuComen
    End Enum
#End Region
    ' Variables
    Private MblnDejoUltimoControl As Boolean = False
    Private MobjObjetoWin As ClsNotaDevAnt = Nothing
    Private MblnPoblandoCbo As Boolean = False
    Private MobjCliente As ClsCliente = Nothing
    Private MdtbAnticipos As DataTable = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomNotDevAnt
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuNotaDevAnt
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolControlesLlave As New Collection From {
            cboPref,
            txtIdNota
        }
        SAdicioneControlRestringido(bttEncontrarCliente)
        SAdicioneControlRestringido(dgrNovedades)
        SCargueForma(EnuElementosAdicionalesDef.enuImprimir, 6,
                lcolControlesLlave, dtpFechaDevAnt, False)
        dgrNovedades.Visibility = Visibility.Visible
        lblValorDev.Visibility = Visibility.Hidden
        txtValorDev.Visibility = Visibility.Hidden
        lblValor.Visibility = Visibility.Visible
        txtValorR.Visibility = Visibility.Visible
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
            Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuNotaReintegroAnt)
            ObjObjetoWin = New ClsNotaDevAnt(lstrPref)
            If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                ObjObjetoWin.SVayaAlUltimo()
            End If
        End If
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        If IsNothing(ObjObjetoWin) Then
            SCerrarClic()
            Exit Sub
        End If
        StcValidaControl(EnuValidEntradaDef.enuIdCliente) = lblIdCliente
        StcValidaControl(EnuValidEntradaDef.enuIdPreAgru) = lblIdPredioAgr
        StcValidaControl(EnuValidEntradaDef.enuFechaDev) = lblFechaDevAnt
        StcValidaControl(EnuValidEntradaDef.enuIdAnticipo) = lblIdAntSel
        StcValidaControl(EnuValidEntradaDef.enuValor) = lblValorDev
        StcValidaControl(EnuValidEntradaDef.enuComen) = lblComentario
        SPuebleComboBoxes()
        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub
    Protected Overrides Sub SMuestreDatos()
        HblnMostrandoDatos = True
        If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos AndAlso
                EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            SLevanteEveNoti("No hay Notas para ser mostradas!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
            cboPref.IsEnabled = False
            txtIdNota.IsEnabled = False
        Else
            With MobjObjetoWin
                cboPref.SelectedItem = .ObjPrefijo_NotaDevAntStr.ToString()
                txtIdNota.Text = .ObjIdNotaDevAntEnt.ToString()
                dtpFechaDevAnt.SelectedDate = .ObjFecha_NotaDevAntDtm.ObjValorPro
                txtIdCliente.Text = .ObjIdCliente_NotaDevAntDbl.ToString
                txtNombreCliente.Content = .ObjClienteNota.ObjNombreCompletoStr.ToString()
                If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                    Dim lstrIdPredioAgr = String.Empty
                    If String.IsNullOrEmpty(.ObjIdPredioAgrupador_NotaDevAntStr.ToString) Then
                        lstrIdPredioAgr = GCSTRSINPA
                    Else
                        lstrIdPredioAgr = .ObjIdPredioAgrupador_NotaDevAntStr.ObjValorPro
                    End If
                    txtPredioAgr.Text = lstrIdPredioAgr
                End If
                txtValorDev.Text = Format(.ObjValor_NotaDevAntDec.ObjValorPro, "c")
                txtValorR.Content = Format(.ObjValor_NotaDevAntDec.ObjValorPro, "c")
                txtIdAntSeleccionado.Content = .ObjIdAnticipo_NotaDevAntEnt.ObjValorPro
                txtComentario.Text = .ObjComentario_NotaDevAntStr.ObjValorPro
            End With
        End If
        SMuestreUsuarios()
        SMuestreEstado()
        Title = My.Resources.FichaNDev
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            Title &= "Nuevo " & My.Resources.De & txtNombreCliente.Content
        Else
            Title &= txtIdNota.Text & My.Resources.De & txtNombreCliente.Content
        End If
        SValide()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            SEstablezcaDataContext()
            txtIdNota.Focus()
        End If
        HblnMostrandoDatos = False
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.cenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(EnuValidEntradaDef.enuIdCliente) = .ObjIdCliente_NotaDevAntDbl.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuIdPreAgru) = .ObjIdPredioAgrupador_NotaDevAntStr.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuFechaDev) = .ObjFecha_NotaDevAntDtm.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuValor) = .ObjValor_NotaDevAntDec.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuIdAnticipo) = .ObjIdAnticipo_NotaDevAntEnt.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuComen) = .ObjComentario_NotaDevAntStr.BlnEsValido
            End With
        End If
        '
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjIdNotaDevAntEnt.ObjValorPro = 0
            .ObjIdCliente_NotaDevAntDbl.ObjValorPro = txtIdCliente.Text
            .ObjFecha_NotaDevAntDtm.ObjValorPro = dtpFechaDevAnt.SelectedDate
            .ObjIdAnticipo_NotaDevAntEnt.ObjValorPro = txtIdAntSeleccionado.Content
            .ObjComentario_NotaDevAntStr.ObjValorPro = txtComentario.Text
            If cboPredioAgru.SelectedItem = GCSTRSINPA Then
                .ObjIdPredioAgrupador_NotaDevAntStr.ObjValorPro = String.Empty
            Else
                .ObjIdPredioAgrupador_NotaDevAntStr.ObjValorPro = cboPredioAgru.SelectedItem
            End If
            .ObjValor_NotaDevAntDec.ObjValorPro = txtValorDev.Text
        End With
        SValide()
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        HmnuImprimir = FmnuiMenuItem("MnuImprimir", "Im_primir", "RecMnuItemSec")
        Dim lentPosicion = HmnuAcciones.Items.Count - 2
        Dim lsepSeparad As New Separator
        HmnuAcciones.Items.Insert(lentPosicion, HmnuImprimir)
        HmnuAcciones.Items.Insert(lentPosicion, lsepSeparad)
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    Protected Overrides Sub SCree()
        MyBase.SCree()
        bttEncontrarCliente.Visibility = Visibility.Visible
        cboPredioAgru.Visibility = Visibility.Visible
        dgrAnticipos.Visibility = Visibility.Visible
        lblValorDev.Visibility = Visibility.Visible
        txtValorDev.Visibility = Visibility.Visible
        dgrNovedades.Visibility = Visibility.Hidden
        txtPredioAgr.Visibility = Visibility.Hidden
        lblValor.Visibility = Visibility.Hidden
        txtValorR.Visibility = Visibility.Hidden
        txtIdNota.IsEnabled = False
        With GobjParametros
            If .ObjExigeFechaHoyDocsBln.ObjValorPro Then
                MobjObjetoWin.ObjFecha_NotaDevAntDtm.ObjValorPro = Date.Today
                dtpFechaDevAnt.Style = FindResource("RecCtlNoHabilitado")
            Else
                If .ObjAnoActual.StrIdPeriodoActual < ClsOrionCop.FstrPeriodoDeFecha(Date.Today) Then
                    MobjObjetoWin.ObjFecha_NotaDevAntDtm.ObjValorPro =
                                .ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                End If
            End If
        End With
        If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
            bttEncontrarCliente.Focus()
        Else
            dtpFechaDevAnt.Focus()
        End If
    End Sub
    Protected Overrides Sub SGuarde()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Dim lblnGuardo As Boolean
        Try
            GobjPanDat.SControleProcesoObj(True)
            If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                SRegistre()
                SValide()
                lblnGuardo = FblnGravo()
                If lblnGuardo Then
                    If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
                        SFinaliceOperacion()
                        lstrMens = "Desea imprimir la Nota?"
                        If MsgBox(lstrMens, vbYesNo + MsgBoxStyle.Question, "Imprimir Nota?") = vbYes Then
                            SImprima()
                        End If
                        lstrMens = FstrNombreDoc()
                        If lstrMens.StartsWith("El") Then
                            lstrMens &= " fue creado exitosamente!"
                        Else
                            lstrMens &= " fue creada exitosamente!"
                        End If
                    Else
                        SFinaliceOperacion()
                    End If
                End If
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
                GobjPanDat.SControleProcesoObj(False)
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
                SFinaliceOperacion()
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Protected Overrides Function SAnule() As Boolean
        Dim lblnAnulo = MyBase.SAnule()
        If lblnAnulo Then
            SImprima()
        End If
        Return lblnAnulo
    End Function
    Protected Overrides Sub SEstablezcaWinConsultando()
        MyBase.SEstablezcaWinConsultando()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            bttEncontrarCliente.Visibility = Visibility.Hidden
            cboPredioAgru.Visibility = Visibility.Hidden
            dgrAnticipos.Visibility = Visibility.Hidden
            dgrNovedades.Visibility = Visibility.Visible
            txtPredioAgr.Visibility = Visibility.Visible
            lblValorDev.Visibility = Visibility.Hidden
            txtValorDev.Visibility = Visibility.Hidden
            lblValor.Visibility = Visibility.Visible
            txtValorR.Visibility = Visibility.Visible
            SMuestreDatos()
            SEstablezcaDataContext()
        End If
    End Sub
    Protected Overrides Sub SImprima()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Mouse.OverrideCursor = Cursors.Wait
            If MobjObjetoWin.BlnExiste Then
                SLevanteEveNoti("Imprimiendo", String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                Dim lstrPrefNota As String = MobjObjetoWin.ObjPrefijo_NotaDevAntStr.ObjValorPro
                Dim lentIdNotaPrimera = MobjObjetoWin.ObjIdNotaDevAntEnt.ObjValorPro
                Dim lentIdNotaUltima = lentIdNotaPrimera
                Dim lobjParaFact As New ClsParametrosReportesDocs(lstrPrefNota,
                            lentIdNotaPrimera, lentIdNotaUltima)
                Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
                    .ObjParRepDocs = lobjParaFact,
                    .EnuReporte = EnuReporteDef.enuNotaDevAnt
                    }
                lobjRep.SGenereReporte()
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
            If Not lblnNoHayError Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            Else
                SLevanteEveNoti(String.Empty, String.Empty, 0, EnuSeveridadNot.EnuOk)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub
#End Region
#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        MyBase.SBuscar()
        If BlnBusquedaOk Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
                If Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                    txtIdCliente.Text = StrResultadoBusqueda
                    SRegistreCliente()
                End If
            Else
                If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable Then
                    If BlnBusquedaOk AndAlso StrResutadosBusqueda.Length > 0 Then
                        cboPref.SelectedItem = StrResutadosBusqueda(0)
                        txtIdNota.Text = StrResutadosBusqueda(1)
                        SAbraNota()
                    End If
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
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            If txtIdCliente.Focus Then
                SDefineBusquedaPredioAgr_Prop()
                SDefineBusquedaPredioAgr_Arren()
                SDefineBusquedaCliente()
            End If
        Else
            SDefineNombreCliente()
            SDefinePredioAgr()
        End If
        Return True
    End Function
    Private Sub SDefineBusquedaCliente()
        Dim lstrTabla = ClsCliente.SstrNombreTabla
        Dim lstrCamposMostrar = {ClsIdClienteDbl.SstrNombreCampoBd, ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdClienteDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " & lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
    Private Sub SDefineBusquedaPredioAgr_Prop()
        Dim lstrTablaPri = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec = ClsPropietario.SstrNombreTabla
        Dim lstrCamSelTablaPri As String() = {"DISTINCT " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd}
        Dim lstrCampSelTablaSec As String() = {ClsIdCliente_PropDbl.SstrNombreCampoBd,
                ClsNombreCompleto_PropStr.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta,
                StrCampoCentroutil,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {StrCampoCarpeta,
                StrCampoCentroutil,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdCliente_PropDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = "P." & StrCampoCarpeta &
                " = " & GshrIdCarpeta & " AND P." &
                StrCampoCentroutil & " = " &
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
    Private Sub SDefineNombreCliente()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsNotaDevAnt.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsIdPredioAgrupador_NotaDevAntStr.SstrNombreCampoBd,
                                            ClsFecha_NotaDevAntDtm.SstrNombreCampoBd,
                                            ClsIdNotaDevAntEnt.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdCliente_NotaDevAntDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdNotaDevAntEnt.SstrNombreCampoBd
        Dim lstrFiltro As String = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta &
                " AND P." & StrCampoCentroUtil & " = " & GshrIdCentroUtil
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, False)
    End Sub
    Private Sub SDefinePredioAgr()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsNotaDevAnt.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsIdPredioAgrupador_NotaDevAntStr.SstrNombreCampoBd,
                                            ClsFecha_NotaDevAntDtm.SstrNombreCampoBd,
                                            ClsPrefijo_NotaDevAntStr.SstrNombreCampoBd,
                                            ClsIdNotaDevAntEnt.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdCliente_NotaDevAntDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsIdPredioAgrupador_NotaDevAntStr.SstrNombreCampoBd
        Dim lstrCamposRetornar As String() = {ClsPrefijo_NotaDevAntStr.SstrNombreCampoBd,
                                            ClsIdNotaDevAntEnt.SstrNombreCampoBd}
        Dim lstrFiltro = "S." & StrCampoCarpeta & " = " &
                GshrIdCarpeta.ToString & " AND S." & StrCampoCentroUtil &
                " = " & GshrIdCentroUtil.ToString
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador", lstrTablaPri, lstrTablaSec,
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
            If MobjObjetoWin.ObjPrefijo_NotaDevAntStr.ToString() <> lstrPref Then
                ObjObjetoWin = New ClsNotaDevAnt(lstrPref)
                MobjObjetoWin = ObjObjetoWin
                If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                    MobjObjetoWin.SVayaAlUltimo()
                End If
            End If
        End If
    End Sub
    Private Sub SMuestreUsuarios()
        With MobjObjetoWin
            If MobjObjetoWin.BlnExiste Then
                If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                    txtUsuarioGenero.Visibility = Visibility.Visible
                    lblUsuarioGenero.Visibility = Visibility.Visible
                    txtUsuarioGenero.Content = MobjObjetoWin.ObjIdUsuario_NotaDevAntStr.ObjValorPro
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
            Else
                lblUsuarioAnulo.Visibility = Visibility.Collapsed
                txtUsuarioAnulo.Visibility = Visibility.Collapsed
            End If
        End With
    End Sub
    Private Sub SMuestreEstado()
        If MobjObjetoWin.ObjAnuladoBln.ObjValorPro Then
            txtEstado.Style = FindResource("RecDocAnulado")
        Else
            txtEstado.Style = FindResource("RecDocNormal")
        End If
    End Sub
    Private Sub SPuebleComboBoxes()
        MblnPoblandoCbo = True
        Dim ldrwConst = ClsOrionCop.FdrwPrefDoc(EnuTipoDocOri.EnuNotaDevAnt)
        SPuebleComboBox(ldrwConst, cboPref)
        MblnPoblandoCbo = False
    End Sub
    Private Sub SPuebleCboPredAgru()
        MblnPoblandoCbo = True
        cboPredioAgru.Items.Clear()
        cboPredioAgru.Items.Add(My.Resources.Ninguno)
        If Not IsNothing(MobjCliente) Then
            Dim ldrwPrediosAgr As DataRow() = MobjCliente.FdrwPrediosAgruAnt()
            For Each ldrwPreAgr As DataRow In ldrwPrediosAgr
                If String.IsNullOrEmpty(ldrwPreAgr(0)) Then
                    cboPredioAgru.Items.Add(GCSTRSINPA)
                Else
                    cboPredioAgru.Items.Add(ldrwPreAgr(0))
                End If
            Next
        End If
        If cboPredioAgru.Items.Count = 1 Then
            SLevanteEveNoti("El Cliente seleccionado no tiene Anticipos!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
        End If
        MblnPoblandoCbo = False
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            cboPredioAgru.SelectedIndex = 0
            MobjObjetoWin.ObjIdPredioAgrupador_NotaDevAntStr.ObjValorPro = cboPredioAgru.SelectedItem
        Else
            cboPredioAgru.SelectedIndex = 0
            Dim lreaArgumento As New RoutedEventArgs With {
                .RoutedEvent = ComboBox.SelectionChangedEvent,
                .Source = cboPredioAgru
            }
            OnCboCambio(cboPredioAgru, lreaArgumento)
        End If
    End Sub
    Private Sub SRegistreCliente()
        With MobjObjetoWin
            Dim ldblIdCliente As Double = .ObjIdCliente_NotaDevAntDbl.ObjValorPro
            .ObjIdCliente_NotaDevAntDbl.ObjValorPro = txtIdCliente.Text
            If .ObjIdCliente_NotaDevAntDbl.BlnEsValido Then
                MobjCliente = .ObjClienteNota
                txtNombreCliente.Content = .ObjClienteNota.ObjNombreCompletoStr.ToString
                SPuebleCboPredAgru()
                cboPredioAgru.Focus()
            End If
        End With
    End Sub
    Private Sub SRegistrePredAgr()
        With MobjObjetoWin
            If cboPredioAgru.Items.Count = 1 Then
                SLevanteEveNoti("El Cliente seleccionado no tiene Anticipos!", "", 0,
                        EnuSeveridadNot.EnuInformacion)
            Else
                If cboPredioAgru.SelectedItem = GCSTRSINPA Then
                    .ObjIdPredioAgrupador_NotaDevAntStr.ObjValorPro = String.Empty
                Else
                    .ObjIdPredioAgrupador_NotaDevAntStr.ObjValorPro = cboPredioAgru.SelectedItem
                End If
                If .ObjIdPredioAgrupador_NotaDevAntStr.BlnEsValido Then
                    Dim lstrIdPredioAgr As String = .ObjIdPredioAgrupador_NotaDevAntStr.ObjValorPro
                    MdtbAnticipos = MobjCliente.DtbAnticiposVivos(lstrIdPredioAgr)
                Else
                    MdtbAnticipos = Nothing
                End If
            End If
        End With
    End Sub
    Private Sub SEstablezcaDataContext()
        If Me.EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            dgrAnticipos.DataContext = MdtbAnticipos
        Else
            dgrNovedades.DataContext = MobjObjetoWin.DtbNovedades
        End If
    End Sub
    Private Sub SAbraNota()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                If txtIdNota.Text <> MobjObjetoWin.ObjIdNotaDevAntEnt.ToString() Then
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
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            If lelmElemento.Name = "bttEncontrarCliente" Then
                If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
                    StrResultadoBusqueda = String.Empty
                    SBuscar()
                End If
            End If
        End If
    End Sub
    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Select Case True
            Case TypeOf lelmElemento Is TextBox
                Dim ltxtTextBox As TextBox = lelmElemento
                ltxtTextBox.SelectAll()
            Case TypeOf lelmElemento Is Button
                If lelmElemento.Equals(HbttCancelar) AndAlso MblnDejoUltimoControl Then
                    If FblnEstanTodosBien() Then
                        HbttAceptar.IsEnabled = True
                        HbttAceptar.Focus()
                        MblnDejoUltimoControl = False
                    End If
                End If
        End Select
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        If Not HblnSeEstaCerrando AndAlso Not HblnMostrandoDatos Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is TextBox OrElse TypeOf lelmElemento Is DatePicker Then
                MblnDejoUltimoControl = False
                If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
                    With MobjObjetoWin
                        Select Case lelmElemento.Name
                            Case "txtIdCliente"
                                SRegistreCliente()
                            Case "dtpFechaDevAnt"
                                .ObjFecha_NotaDevAntDtm.ObjValorPro = dtpFechaDevAnt.SelectedDate
                            Case "txtValorDev"
                                .ObjValor_NotaDevAntDec.ObjValorPro = txtValorDev.Text
                            Case "txtComentario"
                                .ObjComentario_NotaDevAntStr.ObjValorPro = txtComentario.Text
                                MblnDejoUltimoControl = True
                        End Select
                    End With
                    SMuestreDatos()
                End If
            End If
        End If
    End Sub
    Private Sub OnCboCambio(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If Not MblnPoblandoCbo AndAlso TypeOf lelmElemento Is ComboBox AndAlso Not HblnSeEstaCerrando Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando AndAlso
                    Not HblnMostrandoDatos Then
                If TypeOf lelmElemento Is ComboBox Then
                    If lelmElemento.Name = "cboPredioAgru" Then
                        SRegistrePredAgr()
                        dgrAnticipos.DataContext = MdtbAnticipos
                    End If
                    SMuestreDatos()
                End If
            ElseIf EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                If lelmElemento.Name = "cboPref" Then
                    SInicialiceNota()
                    SMuestreDatos()
                End If
            End If
        End If
    End Sub
    Private Sub Txt_KeyDown(sender As Object, e As KeyEventArgs) Handles txtIdNota.KeyDown
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            If e.Key = Key.Return OrElse e.Key = Key.Tab Then
                SAbraNota()
            End If
        End If
    End Sub
    Private Sub DgrAnticipos_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles dgrAnticipos.SelectionChanged
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            Dim ldrvAnticipo As DataRowView = dgrAnticipos.SelectedItem
            Dim lentIdAnticipo = 0
            If Not IsNothing(ldrvAnticipo) Then
                lentIdAnticipo = ldrvAnticipo(ClsIdAnticipoEnt.SstrNombreCampoBd)
            End If
            txtIdAntSeleccionado.Content = lentIdAnticipo
            MobjObjetoWin.ObjIdAnticipo_NotaDevAntEnt.ObjValorPro = lentIdAnticipo
            SValide()
        End If
    End Sub
    Private Sub Dgr_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
                 dgrNovedades.MouseRightButtonUp
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If lelmElemento.Name = "dgrNovedades" Then
            SAbraAnticipo(CType(txtIdAntSeleccionado.Content, Integer))
        End If
    End Sub
    Private Sub Ctl_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles _
            txtIdAntSeleccionado.MouseDoubleClick
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If lelmElemento.Name = "txtIdAntSeleccionado" Then
            SAbraAnticipo(CType(txtIdAntSeleccionado.Content, Integer))
        End If
    End Sub
#End Region
End Class
