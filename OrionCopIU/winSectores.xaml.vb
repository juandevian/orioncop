Imports System.ComponentModel
Public Class WinSectores
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuNombreSector
        enuDsctoPP
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsSector = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomSec
    Private MenuTipoDsctoPP As EnuTipoDsctoPP = EnuTipoDsctoPP.None
#End Region

#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuSectores
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolCamposLlave As New Collection From {
            txtIdSectorNuevo
        }
        SAdicioneControlRestringido(dgrSectores)
        SCargueForma(EnuElementosAdicionalesDef.None, 2,
                lcolCamposLlave, txtNombreSectorNuevo, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        BlnSiempreCreando = True
        dgrSectores.Focus()
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
            ObjObjetoWin = New ClsSector(EnuModoInstanciaObjDef.enuNavegable)
        End If
        MobjObjetoWin = ObjObjetoWin
        MobjObjetoWin.SVayaAlPrimero()
        If GobjParametros.ObjAnoActual IsNot Nothing Then
            MenuTipoDsctoPP = GobjParametros.ObjAnoActual.ObjTipoDsctoPPByt.ObjValorPro
        End If
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuNombreSector) = lblNombreSector
        StcValidaControl(EnuValidEntrada.enuDsctoPP) = lblDsctoPP
        '
        If Not GobjParametros.ObjParametrizacionOkBln.ObjValorPro Then
            txtDsctoPP.Visibility = Visibility.Hidden
            txtDsctoPPNuevo.Visibility = Visibility.Hidden
            lblDsctoPP.Visibility = Visibility.Hidden
        End If
        txtBaseSector.ToolTip = My.Resources.TTBaseSector
        If MenuTipoDsctoPP = EnuTipoDsctoPP.None Then
            txtDsctoPP.Visibility = Visibility.Collapsed
            txtDsctoPPNuevo.Visibility = Visibility.Collapsed
            lblDsctoPP.Visibility = Visibility.Collapsed
        End If
        SEstablezcaDataContext()
        '
        HbttAceptar.TabIndex = 15
        HbttCancelar.TabIndex = 16
        SModifiqueBarraHerramientas()
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If Not IsNothing(MobjObjetoWin) Then
            With MobjObjetoWin
                Dim lstrVlrDsctoPP As String
                If MenuTipoDsctoPP = EnuTipoDsctoPP.EnuProcentaje Then
                    lstrVlrDsctoPP = Format(.ObjDctoProntoPago_SecDbl.ObjValorPro, "p")
                Else
                    lstrVlrDsctoPP = Format(.ObjDctoProntoPago_SecDbl.ObjValorPro, "c")
                End If
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    txtIdSectorNuevo.Text = .ObjIdSectorShr.ToString
                    txtNombreSectorNuevo.Text = .ObjNombreSectorStr.ObjValorPro
                    txtDsctoPPNuevo.Text = lstrVlrDsctoPP
                Else
                    If GobjParametros.FdtbSectores().Rows.Count = 0 Then
                        SLevanteEveNoti("No hay Sectores para ser mostrados!", "", 0,
                                EnuSeveridadNot.EnuInformacion)
                    End If
                    txtDsctoPP.Text = lstrVlrDsctoPP
                End If
                txtArea.Content = Format(GobjParametros.ObjTotalAreaCopropDec.ObjValorPro,
                        "#,000.000")
                txtAreaPonderada.Content = Format(GobjParametros.ObjTotalAreaPondDec.ObjValorPro,
                        "#,000.000")
            End With
        End If
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        With MobjObjetoWin
            If GobjParametros.FdtbSectores().Rows.Count = 0 AndAlso EnuOperacionEnWin =
                    EnuOperacionEnVentana.CenuConsultando Then
                SInicialiceValido()
            Else
                StcValidValido(EnuValidEntrada.enuNombreSector) = .ObjNombreSectorStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuDsctoPP) = .ObjDctoProntoPago_SecDbl.BlnEsValido
            End If
        End With
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjIdSectorShr.ObjValorPro = txtIdSectorNuevo.Text
            If MenuTipoDsctoPP = EnuTipoDsctoPP.EnuProcentaje Then
                .ObjDctoProntoPago_SecDbl.ObjValorPro = FdblTasa(txtDsctoPPNuevo.Text)
            ElseIf MenuTipoDsctoPP = EnuTipoDsctoPP.EnuValorFijo Then
                .ObjDctoProntoPago_SecDbl.ObjValorPro = txtDsctoPPNuevo.Text
            Else
                .ObjDctoProntoPago_SecDbl.ObjValorPro = 0
            End If
            .ObjNombreSectorStr.ObjValorPro = txtNombreSectorNuevo.Text
        End With
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        '
    End Sub
#End Region

#Region "Procedimientos invalidantes"
    ''' <summary>
    ''' Sub que prepara a la ventana y a su objeto para crear un nuevo objeto. Invalida el Sub
    ''' "SCree" de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SCree()
        MyBase.SCree()
        SCreeSector()
    End Sub
    ''' <summary>
    ''' Prepara la ventana y su objeto para modificar el objeto. Invalida la función "SModifique"
    ''' de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SModifique()
        txtIdSectorNuevo.Text = txtIdSector.Text
        txtNombreSectorNuevo.Text = txtNombreSector.Text
        txtDsctoPPNuevo.Text = txtDsctoPP.Text
        MyBase.SModifique()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            SModifiqueSector()
        End If
        SMuestreDatos()
    End Sub
    Protected Overrides Sub SRefresqueWin()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            GobjPanDat.SControleProcesoObj(True)
            txtNombreSector.Text = String.Empty
            txtDsctoPP.Text = String.Empty
            MyBase.SRefresqueWin()
            SEstablezcaDataContext()
            GobjPanDat.SControleProcesoObj(False)
            Mouse.OverrideCursor = Cursors.Arrow
        End If
    End Sub
    Protected Overrides Sub SEstablezcaWinConsultando()
        MyBase.SEstablezcaWinConsultando()
        dgrSectores.IsEnabled = True
        txtIdSector.Visibility = Visibility.Visible
        txtIdSectorNuevo.Visibility = Visibility.Hidden
        txtNombreSector.Visibility = Visibility.Visible
        txtNombreSectorNuevo.Visibility = Visibility.Hidden
        txtDsctoPP.Visibility = Visibility.Visible
        txtDsctoPPNuevo.Visibility = Visibility.Hidden
        SEstablezcaDataContext()
        SMuestreDatos()
    End Sub
#End Region

#Region "Procedimientos Propios"
    Private Sub SModifiqueBarraHerramientas()
        HbttAlPrimero.Visibility = Visibility.Collapsed
        HbttAlAnterior.Visibility = Visibility.Collapsed
        HbttAlSiguiente.Visibility = Visibility.Collapsed
        HbttAlUltimo.Visibility = Visibility.Collapsed
        HbttBuscar.Visibility = Visibility.Collapsed
        Dim ltlbMiToolBar As ToolBar = Nothing
        For Each lobjObjeto As Object In PanelControl.Children
            If TypeOf lobjObjeto Is ToolBar Then
                ltlbMiToolBar = lobjObjeto
                Exit For
            End If
        Next
        If Not IsNothing(ltlbMiToolBar) Then
            For Each lobjObjeto In ltlbMiToolBar.Items
                If TypeOf (lobjObjeto) Is Separator Then
                    Dim lsepSeparador As Separator = lobjObjeto
                    If lsepSeparador.Name = "sepNavegar" Then
                        lsepSeparador.Visibility = Visibility.Collapsed
                        Exit For
                    End If
                End If
            Next
        End If
        HmnuNavegar.Visibility = Visibility.Collapsed
    End Sub
    Private Sub SHabiliteControlesNuevos(ablnHabilite As Boolean)
        SVisibiliceControlesNuevos(ablnHabilite)
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.cenuModificando Then
                txtIdSectorNuevo.Text = txtIdSector.Text
                txtNombreSectorNuevo.Text = txtNombreSector.Text
            End If
            txtIdSectorNuevo.Style = FindResource("RecCtlNoHabilitado")
            txtNombreSectorNuevo.Style = FindResource("RecCtlHabilitado")
            If GobjParametros.ObjAnoActual IsNot Nothing Then
                If GobjParametros.ObjAnoActual.ObjTipoIncentivoByt.ObjValorPro =
                        EnuTipoIncentivo.EnuDescuentoPP Then
                    txtDsctoPPNuevo.Style = FindResource("RecCtlHabilitado")
                Else
                    txtDsctoPPNuevo.Style = FindResource("RecCtlNoHabilitado")
                End If
            End If
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
        txtIdSector.Visibility = lvisVisibilidadActual
        txtNombreSector.Visibility = lvisVisibilidadActual
        txtDsctoPP.Visibility = lvisVisibilidadActual
        txtIdSectorNuevo.Visibility = lvisVisibilidadNuevos
        txtNombreSectorNuevo.Visibility = lvisVisibilidadNuevos
        If Not GobjParametros.ObjParametrizacionOkBln.ObjValorPro Then
            txtDsctoPPNuevo.Visibility = Visibility.Hidden
        Else
            txtDsctoPPNuevo.Visibility = lvisVisibilidadNuevos
        End If
    End Sub
    Private Sub SCreeSector()
        SHabiliteControlesNuevos(True)
        dgrSectores.IsEnabled = False
        MobjObjetoWin.ObjIdSectorShr.SValide()
        MobjObjetoWin.ObjDctoProntoPago_SecDbl.SValide()
        txtBaseSector.Content = "0.00"
        SMuestreDatos()
        txtNombreSectorNuevo.Focus()
    End Sub
    Private Sub SModifiqueSector()
        If Not IsNothing(MobjObjetoWin) AndAlso MobjObjetoWin.BlnExiste Then
            SHabiliteControlesNuevos(True)
            dgrSectores.IsEnabled = False
            txtNombreSectorNuevo.Focus()
        End If
    End Sub
    Private Sub SEstablezcaDataContext()
        grdSector.DataContext = GobjParametros.FdtbSectores()
        SOrdeneDataGrid(dgrSectores, dgrSectores.Columns(0), ClsIdSectorShr.SstrNombreCampoBd,
                ListSortDirection.Ascending)
        If dgrSectores.Items.Count > 0 Then
            dgrSectores.SelectedIndex = 0
        End If
    End Sub
#End Region

#Region "Eventos en la Ventana"
    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            ltxtTextBox.SelectAll()
        ElseIf lelmElemento.Equals(HbttCancelar) Then
            If FblnEstanTodosBien() Then
                HbttAceptar.Focus()
            End If
        End If
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        If Not HblnSeEstaCerrando Then
            Dim lelmelemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmelemento Is TextBox Then
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                    Try
                        With MobjObjetoWin
                            Select Case True
                                Case lelmelemento.Name = "txtIdSectorNuevo"
                                    .ObjIdSectorShr.ObjValorPro = txtIdSectorNuevo.Text
                                Case lelmelemento.Name = "txtNombreSectorNuevo"
                                    .ObjNombreSectorStr.ObjValorPro = txtNombreSectorNuevo.Text
                                Case lelmelemento.Name = "txtDsctoPPNuevo"
                                    If MenuTipoDsctoPP = EnuTipoDsctoPP.EnuProcentaje Then
                                        .ObjDctoProntoPago_SecDbl.ObjValorPro = FdblTasa(txtDsctoPPNuevo.Text)
                                    ElseIf MenuTipoDsctoPP = EnuTipoDsctoPP.EnuValorFijo Then
                                        .ObjDctoProntoPago_SecDbl.ObjValorPro = txtDsctoPPNuevo.Text
                                    Else
                                        lstrMens = "El Descuento por pronto Pago no esta activado!"
                                    End If
                            End Select
                        End With
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
                            If Not String.IsNullOrEmpty(lstrMens) Then
                                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuInformacion)
                            End If
                        Else
                            SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                        End If
                    End Try
                End If
            End If
        End If
    End Sub
    Private Sub TxtNombre_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtNombreSector.TextChanged
        Dim lobjvalorllave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, txtIdSector.Text}
        MobjObjetoWin.SAbra(lobjvalorllave)
        SMuestreDatos()
    End Sub
#End Region
End Class
