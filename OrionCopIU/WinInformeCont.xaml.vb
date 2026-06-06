Public Class WinInformeCont
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuFecIni
        enuFecFin
        enuFecRad
        enuIdFacIni
        enuIdFacFin
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsInformeCont = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomInfCon
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuInformeCon
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolControlesLlave As New Collection From {
            txtIdInformeCont
        }
        SCargueForma(EnuElementosAdicionalesDef.None, 5, lcolControlesLlave,
                Nothing, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        txtNomDoc.Content = ClsInformeCont.StrNombreInfContingencia
        txtIdInformeCont.SelectAll()
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
        If IsNothing(MobjObjetoWin) Then
            MobjObjetoWin = New ClsInformeCont(EnuModoInstanciaObjDef.enuNavegable)
            If Not MobjObjetoWin.FblnEstaVacioOrigenDatos Then
                MobjObjetoWin.SVayaAlUltimo()
            End If
        End If
        ObjObjetoWin = MobjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuIdFacFin) = lblFacFinCon
        StcValidaControl(EnuValidEntrada.enuIdFacIni) = lblFacIniCon
        StcValidaControl(EnuValidEntrada.enuFecRad) = lblFecRadCon
        StcValidaControl(EnuValidEntrada.enuFecFin) = lblFecFinCon
        StcValidaControl(EnuValidEntrada.enuFecIni) = lblFecIniCon
        '
        txtPrefFacFin.Content = GobjParametros.ObjPrefijoFactContStr.ObjValorNuevo
        txtPrefFacIni.Content = GobjParametros.ObjPrefijoFactContStr.ObjValorNuevo
        HbttAceptar.TabIndex = 20
        HbttCancelar.TabIndex = 21
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos Then
            SLevanteEveNoti("No hay Informes de Contingencia para ser mostrados!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
            txtIdInformeCont.IsEnabled = False
        End If
        With MobjObjetoWin
            txtIdInformeCont.Text = .ObjIdInformeContEnt.ObjValorPro
            dtpFechaIni.SelectedDate = .ObjFechaInicioContDtm.ObjValorPro
            txtHoraIni.Text = .ObjFechaInicioContDtm.EntHoraFecIni
            txtMinutosIni.Text = .ObjFechaInicioContDtm.EntMinFecIni
            dtpFechaFin.SelectedDate = .ObjFechaFinContDtm.ObjValorPro
            txtHoraFin.Text = .ObjFechaFinContDtm.EntHoraFecFin
            txtMinutosFin.Text = .ObjFechaFinContDtm.EntMinFecFin
            dtpFechaRad.SelectedDate = .ObjFechaRadicoDtm.ObjValorPro
            txtHoraRad.Text = .ObjFechaRadicoDtm.EntHoraFecRad
            txtMinutosRad.Text = .ObjFechaRadicoDtm.EntMinFecRad
            txtPrefFacIni.Content = .ObjPrefFactContStr.ObjValorPro
            txtIdFacIniCon.Text = .ObjIdFactContIniEnt.ObjValorPro
            txtPrefFacFin.Content = .ObjPrefFactContStr.ObjValorPro
            txtIdFacFinCon.Text = .ObjIdFactContFinEnt.ObjValorPro
            txtComentario.Text = .ObjComentario_InfContStr.ObjValorPro
        End With
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            If txtIdInformeCont.Focus Then
                txtIdInformeCont.SelectAll()
            End If
        End If
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.cenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(EnuValidEntrada.enuIdFacFin) = .ObjIdFactContFinEnt.BlnEsValido
                StcValidValido(EnuValidEntrada.enuIdFacIni) = .ObjIdFactContIniEnt.BlnEsValido
                StcValidValido(EnuValidEntrada.enuFecRad) = .ObjFechaRadicoDtm.BlnEsValido
                StcValidValido(EnuValidEntrada.enuFecFin) = .ObjFechaFinContDtm.BlnEsValido
                StcValidValido(EnuValidEntrada.enuFecIni) = .ObjFechaInicioContDtm.BlnEsValido
            End With
        End If
        '
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjFechaFinContDtm.ObjValorPro = dtpFechaFin.SelectedDate
            .ObjFechaFinContDtm.EntHoraFecFin = txtHoraFin.Text
            .ObjFechaFinContDtm.EntMinFecFin = txtMinutosFin.Text
            .ObjFechaInicioContDtm.ObjValorPro = dtpFechaIni.SelectedDate
            .ObjFechaInicioContDtm.EntHoraFecIni = txtHoraIni.Text
            .ObjFechaInicioContDtm.EntMinFecIni = txtMinutosIni.Text
            .ObjFechaRadicoDtm.ObjValorPro = dtpFechaRad.SelectedDate
            .ObjFechaRadicoDtm.EntHoraFecRad = txtHoraRad.Text
            .ObjFechaRadicoDtm.EntMinFecRad = txtMinutosRad.Text
            .ObjIdFactContFinEnt.ObjValorPro = txtIdFacFinCon.Text
            .ObjIdFactContIniEnt.ObjValorPro = txtIdFacIniCon.Text
            .ObjComentario_InfContStr.ObjValorPro = txtComentario.Text
        End With
        SValide()
    End Sub
    Protected Overrides Sub SConfigureMenuesPropios()
        '
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    Protected Overrides Sub SCree()
        MyBase.SCree()
        txtIdInformeCont.IsEnabled = False
        MobjObjetoWin.ObjPrefFactContStr.ObjValorPro = GobjParametros.ObjPrefijoFactContStr.ObjValorPro
        SMuestreDatos()
        dtpFechaIni.Focus()
    End Sub
    Protected Overrides Sub SModifique()
        Dim lblnPuedeModi = MobjObjetoWin.ObjFechaRadicoDtm.ObjValorPro > GCDTMFECHANULA
        If lblnPuedeModi Then
            MyBase.SModifique()
        Else
            SLevanteEveNoti("Este Informe de Contingencia ya está cerrado!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SAbraInformeCont()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                If txtIdInformeCont.Text <> MobjObjetoWin.ObjIdInformeContEnt.ToString() Then
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, txtIdInformeCont.Text}
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
    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            ltxtTextBox.SelectAll()
        End If
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                Select Case lelmElemento.Name
                    Case "txtHoraIni"
                        MobjObjetoWin.ObjFechaInicioContDtm.EntHoraFecIni = txtHoraIni.Text
                    Case "txtHoraFin"
                        MobjObjetoWin.ObjFechaFinContDtm.EntHoraFecFin = txtHoraFin.Text
                    Case "txtMinutosIni"
                        MobjObjetoWin.ObjFechaInicioContDtm.EntMinFecIni = txtMinutosIni.Text
                    Case "txtMinutosFin"
                        MobjObjetoWin.ObjFechaFinContDtm.EntMinFecFin = txtMinutosFin.Text
                    Case "txtHoraRad"
                        MobjObjetoWin.ObjFechaRadicoDtm.EntHoraFecRad = txtHoraRad.Text
                    Case "txtMinutosRad"
                        MobjObjetoWin.ObjFechaRadicoDtm.EntMinFecRad = txtMinutosRad.Text
                    Case "txtIdFacIniCon"
                        MobjObjetoWin.ObjIdFactContIniEnt.ObjValorPro = txtIdFacIniCon.Text
                    Case "txtIdFacFinCon"
                        MobjObjetoWin.ObjIdFactContFinEnt.ObjValorPro = txtIdFacFinCon.Text
                    Case "txtComentario"
                        MobjObjetoWin.ObjComentario_InfContStr.ObjValorPro = txtComentario.Text
                End Select
                SMuestreDatos()
            End If
        End If
    End Sub
    Private Sub Txt_KeyDown(sender As Object, e As KeyEventArgs) Handles txtIdInformeCont.KeyDown
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            If e.Key = Key.Return OrElse e.Key = Key.Tab Then
                SAbraInformeCont()
            End If
        End If
    End Sub
    Private Sub Dtp_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles _
            dtpFechaIni.SelectedDateChanged, dtpFechaFin.SelectedDateChanged, dtpFechaRad.SelectedDateChanged
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is DatePicker Then
            If lelmElemento.Name = "dtpFechaIni" Then
                MobjObjetoWin.ObjFechaInicioContDtm.ObjValorPro = dtpFechaIni.SelectedDate
            ElseIf lelmElemento.Name = "dtpFechaFin" Then
                MobjObjetoWin.ObjFechaFinContDtm.ObjValorPro = dtpFechaFin.SelectedDate
            ElseIf lelmElemento.Name = "dtpFechaRad" Then
                MobjObjetoWin.ObjFechaRadicoDtm.ObjValorPro = dtpFechaRad.SelectedDate
            End If
        End If
        SMuestreDatos()
        SValide()
    End Sub
#End Region
End Class
