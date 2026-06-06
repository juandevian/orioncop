Public Class WinPropietario
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuIdProp
        enuPorPar
    End Enum
#End Region
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomProp
    Private ReadOnly MblnNuevo As Boolean = False
    Private ReadOnly MobjObjetoWin As ClsPropietario = Nothing
    Private ReadOnly MobjPredio As ClsPredio = Nothing
    Private MblnModificandoNuevo As Boolean = False
#End Region
#Region "Constructor"
    Public Sub New(aobjProp As ClsPropietario, ablnNuevo As Boolean)
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuPropietario
        MblnNuevo = ablnNuevo
        MobjObjetoWin = aobjProp
        MobjPredio = MobjObjetoWin.ObjPadre
        GblnOK = False
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolControlesLlave As New Collection From {
            txtIdPropietario
        }
        SCargueForma(EnuElementosAdicionalesDef.None,
                2, lcolControlesLlave, txtPorcientoProp, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        If MblnNuevo Then
            SCree()
        Else
            SModifique()
        End If
        SValide()
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
            ObjObjetoWin = MobjObjetoWin
        End If
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuIdProp) = lblIdPropietario
        StcValidaControl(EnuValidEntrada.enuPorPar) = lblPorcientoProp
        '
        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub
    Protected Overrides Sub SMuestreDatos()
        HblnMostrandoDatos = True
        With MobjObjetoWin
            txtIdPropietario.Text = MobjObjetoWin.ObjIdCliente_PropDbl.ObjValorPro
            txtNombrePropietario.Content = MobjObjetoWin.ObjCliente.ObjNombreCompletoStr.ObjValorPro
            txtPorcientoProp.Text = Format(MobjObjetoWin.ObjPorcentajePartiDbl.ObjValorPro, "p")
        End With
        SValide()
        Title = My.Resources.FichaProp
        If Not String.IsNullOrEmpty(txtNombrePropietario.Content) Then
            Title &= txtNombrePropietario.Content
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
                StcValidValido(EnuValidEntrada.enuIdProp) = .ObjIdCliente_PropDbl.BlnEsValido
                StcValidValido(EnuValidEntrada.enuPorPar) = .ObjPorcentajePartiDbl.BlnEsValido
            End With
        End If
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjIdCliente_PropDbl.ObjValorPro = txtIdPropietario.Text
            .ObjPorcentajePartiDbl.ObjValorPro = txtPorcientoProp.Text
            .ObjNombreCompleto_PropStr.ObjValorPro = txtNombrePropietario.Content
        End With
    End Sub
    Protected Overrides Sub SConfigureMenuesPropios()
        '
    End Sub
#End Region
#Region "Invalida otros metodos de la clase base"
    Protected Overrides Sub SCree()
        MyBase.SCree()
        txtIdPropietario.IsEnabled = True
        bttEncontrarProp.Visibility = Visibility.Visible
        Dim lblnNoUsado = bttEncontrarProp.Focus()
    End Sub
    Protected Overrides Sub SModifique()
        Dim lblnPropYaCreado = MobjObjetoWin.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando
        If Not lblnPropYaCreado Then
            MobjObjetoWin.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando
            MyBase.SModifique()
        Else
            MblnModificandoNuevo = True
            EnuOperacionEnWin = EnuOperacionEnVentana.CenuModificando
            MobjObjetoWin.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando
            MobjObjetoWin.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
            HbttCancelar.Content = My.Resources.Cancelar
            SHabiliteWin(True)
            Dim lblnNoUsado = txtPorcientoProp.Focus()
            SRegistre()
            SValide()
        End If
        txtIdPropietario.IsEnabled = False
        txtPorcientoProp.Focus()
    End Sub
    Protected Overrides Sub SCancele()
        MobjObjetoWin.ObjIdCliente_PropDbl.ObjValorPro = 0
        MobjObjetoWin.ObjPorcentajePartiDbl.ObjValorPro = 0
        SValide()
        MyBase.SCancele()
        GblnOK = False
        Close()
    End Sub
    Protected Overrides Sub SGuarde()
        GblnOK = True
        If MblnModificandoNuevo Then
            MobjObjetoWin.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando
            MobjObjetoWin.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando
        End If
        Close()
    End Sub
#End Region
#Region "Metodos"
    Private Sub SBusqueCliente()
        Cursor = Cursors.Wait
        If HwinBusqueda Is Nothing Then
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
            txtIdPropietario.Text = StrResultadoBusqueda
            SRegistrePropietario()
            txtPorcientoProp.Focus()
            SMuestreDatos()
        Else
            txtIdPropietario.Text = String.Empty
            txtIdPropietario.Focus()
        End If
    End Sub
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
    Private Sub SRegistrePropietario()
        SLevanteEveOk()
        Dim lstrMens = String.Empty
        If Not String.IsNullOrEmpty(txtIdPropietario.Text) Then
            If Not MobjPredio.FblnEstaPazYSalvo(lstrMens) Then
                MsgBox(lstrMens, vbOKOnly, "Advertencia")
            End If
            MobjObjetoWin.ObjIdCliente_PropDbl.ObjValorPro = txtIdPropietario.Text
            If Not MobjObjetoWin.BlnExisteCliente Then
                SCrearCliente(txtIdPropietario.Text)
            End If
            MobjObjetoWin.ObjNombreCompleto_PropStr.ObjValorPro =
                    MobjObjetoWin.ObjCliente.ObjNombreCompletoStr.ObjValorPro
            If MobjObjetoWin.ObjCliente.BlnExiste Then
                txtNombrePropietario.Content =
                        MobjObjetoWin.ObjCliente.ObjNombreCompletoStr.ObjValorPro
            End If
        End If
    End Sub
    Private Sub SCrearCliente(astrIdCliente As String)
        If IsNumeric(astrIdCliente) Then
            If Not String.IsNullOrEmpty(astrIdCliente) AndAlso astrIdCliente <> "0" Then
                If MsgBox("El Cliente ingresado no existe. Desea crearlo ahora?",
                        MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                        "Crear Cliente ?") = MsgBoxResult.Yes Then
                    Dim lobjCliente = ClsOrionCop.FobjCliente(EnuModoInstanciaObjDef.enuNavegable)
                    lobjCliente.SCreeObj({GshrIdCarpeta, GshrIdCentroUtil, astrIdCliente})
                    lobjCliente.ObjIdClienteDbl.ObjValorPro = astrIdCliente
                    Dim lwinVentana = New WinClientes With {
                        .ObjObjetoWin = lobjCliente,
                        .EnuOperacionEnWin = EnuOperacionEnWin.CenuCreando,
                        .WinPadre = Me
                    }
                    lwinVentana.ShowDialog()
                    If Not lobjCliente.BlnExiste Then
                        txtIdPropietario.Text = String.Empty
                        txtNombrePropietario.Content = String.Empty
                    Else
                        MobjObjetoWin.ObjIdCliente_PropDbl.ObjValorPro = astrIdCliente
                        txtNombrePropietario.Content =
                                MobjObjetoWin.ObjIdCliente_PropDbl.StrNombreCli
                    End If
                    SLevanteEveOk()
                End If
            End If
        Else
            SLevanteEveNoti("La identificación del Cliente debe ser númerica!", "", 0,
                    EnuSeveridadNot.EnuDatoInvalido)
        End If
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            ltxtTextBox.SelectAll()
        ElseIf TypeOf lelmElemento Is Button Then
            If lelmElemento.Name = "bttCancelar" AndAlso FblnEstanTodosBien() Then
                HbttAceptar.IsEnabled = True
                HbttAceptar.Focus()
            End If
        End If
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        If Not HblnSeEstaCerrando AndAlso Not HblnMostrandoDatos Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            Dim lstrNombreTextBox = lelmElemento.Name
            If TypeOf lelmElemento Is TextBox OrElse TypeOf lelmElemento Is CheckBox Then
                If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                    Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError As Boolean
                    Try
                        Select Case lstrNombreTextBox
                            Case "txtIdPropietario"
                                SRegistrePropietario()
                            Case "txtPorcientoProp"
                                MobjObjetoWin.ObjPorcentajePartiDbl.ObjValorPro =
                                        FdblTasa(txtPorcientoProp.Text)
                        End Select
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
    Private Sub BttEncontrarProp_Click(sender As Object, e As RoutedEventArgs) Handles _
            bttEncontrarProp.Click
        SBusqueCliente()
    End Sub
    Protected Overrides Sub EwinClosed(sender As Object, e As EventArgs)
        If WinPadre IsNot Nothing Then
            If WinPadre.Visibility <> Visibility.Visible Then
                WinPadre.Visibility = Visibility.Visible
                If GblnOK Then
                    Dim lwinPredio As WinPredios = WinPadre
                    lwinPredio.SAcepteProp(MblnNuevo)
                End If
            End If
        End If
    End Sub
#End Region
End Class
