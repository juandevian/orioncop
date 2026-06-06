Public Class WinCoeficientesPropPropietarios
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
#End Region
    ' Variables
    Private MnuExportar As MenuItem
    Private MdtbCoefiPropP As DataTable = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.CoefPropProp
#End Region
#Region "Constructor"
    Friend Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuPropietarioXCP
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicioneControlRestringido(dgrCoefPropP)
        SAdicioneControlRestringido(bttExportar)
        SCargueForma(EnuElementosAdicionalesDef.enuImprimir, 0, Nothing, Nothing, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        If Not FblnExisteExcel() Then
            bttExportar.Visibility = Visibility.Hidden
        End If
        dgrCoefPropP.IsEnabled = True
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
            ObjObjetoWin = GobjParametros
        End If
        EnuTipoPermisoObjWin = GobjParametros.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        SAdecueVentana()
        SVacie()
        HbttAceptar.TabIndex = 100
        bttExportar.TabIndex = 101
        HbttCancelar.TabIndex = 102
    End Sub
    Private Sub SVacie()
        '
    End Sub
    Protected Overrides Sub SRegistre()
        '
    End Sub
    Protected Overrides Sub SMuestreDatos()
        MdtbCoefiPropP = ClsOrionCop.FdtbCoeficientesPropPropietarios
        Mouse.OverrideCursor = Cursors.Wait
        dgrCoefPropP.DataContext = MdtbCoefiPropP
        txtCantReg.Content = Format(MdtbCoefiPropP.Rows.Count, "#,##0")
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Protected Overrides Sub SValide()
        '
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        Dim lentPosicion = HmnuAcciones.Items.Count - 2
        Dim lsepSeparad As New Separator
        MnuExportar = FmnuiMenuItem("MnuExportar", "Exportar a Excel", "RecMnuItemSec")
        HmnuAcciones.Items.Insert(lentPosicion, MnuExportar)
        HmnuAcciones.Items.Insert(lentPosicion, lsepSeparad)
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    Protected Overrides Sub SCancele()
        SCerrarClic()
    End Sub
    Protected Overrides Sub SCerrarClic()
        Close()
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SAdecueVentana()
        HbttModificar.Visibility = Visibility.Collapsed
        HbttRefrescar.Visibility = Visibility.Collapsed
        HbttCancelar.ToolTip = My.Resources.TTCierraVen
    End Sub
    Private Sub SExporteAExcel()
        SValide()
        If FblnEstanTodosBien() Then
            Mouse.OverrideCursor = Cursors.Wait
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                Dim lobjExp As New ClsExportToExcell, lblnExportoExcel As Boolean
                Dim lstrFecha = Today.Year.ToString & Format(Today.Month, "00") &
                        Format(Today.Day, "00")
                Dim lstrNomArchivo = lstrFecha & "_CoeficientesPropiedad"
                lstrNomArchivo = GstrTrayReportes & "\" & lstrNomArchivo
                lblnExportoExcel = lobjExp.FblnExportToExcel("Coeficientes de Propiedad",
                        lstrNomArchivo, dgrCoefPropP, MdtbCoefiPropP, FstrMapeoCols)
                If lblnExportoExcel Then
                    If MsgBox("Desea abrir el Archivo de Excel?", MsgBoxStyle.Question + vbYesNo,
                        "Abrir Archivo") = MsgBoxResult.Yes Then
                        lobjExp.SAbraExcel()
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
                Mouse.OverrideCursor = Cursors.Arrow
                If Not lblnNoHayError Then
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
    End Sub
    Private Function FstrMapeoCols() As String()
        Dim lstrMap(dgrCoefPropP.Columns.Count - 1) As String
        lstrMap(0) = "Id. Propietario=IdTerceroPropietario"
        lstrMap(1) = "Nombres=Nombres"
        lstrMap(2) = "Apellidos=Apellidos"
        lstrMap(3) = "Razón Social=RazonSocial"
        lstrMap(4) = "Teléfono=Telefono"
        lstrMap(5) = "Correo=Correo"
        lstrMap(6) = "Apartamento/Casa=Predio"
        lstrMap(7) = "Suma Coeficientes=TotalCoeficientes"
        Return lstrMap
    End Function
    Private Sub SAbraVentanaCliente(adblIdCliente As Double)
        If adblIdCliente > 0 Then
            Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
            Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, adblIdCliente}
            lobjCliente.SAbra(lobjValorLlave)
            Dim lwinVentana As New WinClientes With {
                .WinPadre = Me,
                .ObjObjetoWin = lobjCliente,
                .BlnVentanaAux = True
            }
            lwinVentana.ShowDialog()
        End If
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            Select Case lelmElemento.Name
                Case "bttExportar"
                    SExporteAExcel()
            End Select
        End If
    End Sub
    Private Sub Dgr_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
            dgrCoefPropP.MouseRightButtonUp
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If lelmElemento.Name = "dgrCoefPropP" Then
            If dgrCoefPropP.SelectedItem IsNot Nothing Then
                Dim ldrvCliente As DataRowView = dgrCoefPropP.SelectedItem
                Dim ldblIdCliente As Double = ldrvCliente("IdTerceroPropietario")
                SAbraVentanaCliente(ldblIdCliente)
            End If
        End If
    End Sub
#End Region
End Class
