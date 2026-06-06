Imports System.ComponentModel
Public Class WinModulosServicios
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuIdModlo
        enuValorCont
    End Enum
#End Region
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomModSer
    Private ReadOnly MblnNuevo As Boolean = False
    Private ReadOnly MobjObjetoWin As ClsModuloServicio = Nothing
    Private ReadOnly MobjServicio As ClsServicio = Nothing
    Private MblnPoblandoComboBox As Boolean = False
#End Region
#Region "Constructor"
    Friend Sub New(aobjModuloSer As ClsModuloServicio, ablnNuevo As Boolean)
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuModulosServicio
        MblnNuevo = ablnNuevo
        MobjObjetoWin = aobjModuloSer
        MobjServicio = MobjObjetoWin.ObjPadre
        GblnOK = False
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolControlesLlave As New Collection From {
            cboModuloServicio
        }
        SCargueForma(EnuElementosAdicionalesDef.None,
                2, lcolControlesLlave, txtValorCont, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        If MblnNuevo Then
            SCree()
        Else
            SModifique()
        End If
        SHabiliteValor
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
        SMuestreNota
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuIdModlo) = lblModuloServicio
        StcValidaControl(EnuValidEntrada.enuValorCont) = lblValorCont
        '
        SPuebleCboModulo
        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub
    Protected Overrides Sub SMuestreDatos()
        HblnMostrandoDatos = True
        With MobjObjetoWin
            cboModuloServicio.SelectedIndex = MobjObjetoWin.ObjIdModulo_ModuloServicioShr.ObjValorPro
            txtValorCont.Text = Format(MobjObjetoWin.ObjValorPres_ModuloServicioDec.ObjValorPro, "c")
        End With
        SValide()
        Title = My.Resources.NomModSer & My.Resources.DosPuntosEspacio
        Title &= cboModuloServicio.SelectedItem
        HblnMostrandoDatos = False
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.CenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(EnuValidEntrada.enuIdModlo) =
                        .ObjIdModulo_ModuloServicioShr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuValorCont) =
                        .ObjValorPres_ModuloServicioDec.BlnEsValido
            End With
        End If
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjIdModulo_ModuloServicioShr.ObjValorPro = cboModuloServicio.SelectedIndex
            .ObjValorPres_ModuloServicioDec.ObjValorPro = txtValorCont.Text
        End With
    End Sub
    Protected Overrides Sub SConfigureMenuesPropios()
        '
    End Sub
#End Region
#Region "Invalida otros metodos de la clase base"
    Protected Overrides Sub SCree()
        EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando
        SMuestreDatos()
        SFrmAdicione()
        txtNota.IsEnabled = False
        MobjObjetoWin.ObjValorPres_ModuloServicioDec.ObjValorPro = 0
        cboModuloServicio.IsEnabled = True
        cboModuloServicio.Focus()
    End Sub
    Protected Overrides Sub SModifique()
        EnuOperacionEnWin = EnuOperacionEnVentana.CenuModificando
        If Not IsNothing(HbttCancelar) Then
            HbttCancelar.Content = My.Resources.Cancelar
        End If
        SHabiliteWin(True)
        txtNota.IsEnabled = False
        SRegistre()
        cboModuloServicio.IsEnabled = False
        txtValorCont.Focus()
    End Sub
    Protected Overrides Sub SCancele()
        MobjObjetoWin.ObjIdModulo_ModuloServicioShr.ObjValorPro = 0
        MobjObjetoWin.ObjValorPres_ModuloServicioDec.ObjValorPro = 0
        SValide()
        MyBase.SCancele()
        GblnOK = False
        Close()
    End Sub
    Protected Overrides Sub SGuarde()
        GblnOK = True
        Close()
    End Sub
#End Region
#Region "Metodos"
    Private Sub SPuebleCboModulo()
        MblnPoblandoComboBox = True
        cboModuloServicio.Items.Clear()
        cboModuloServicio.Items.Add("Ninguno")
        For Each lobjModContribucion As ClsModuloContribucion In GobjParametros.ColModulos
            cboModuloServicio.Items.Add(lobjModContribucion.ObjNombreModuloStr.ObjValorPro)
        Next
        MblnPoblandoComboBox = False
    End Sub
    Private Sub SRegistreModuloServicio()
        SLevanteEveOk()
        If cboModuloServicio.SelectedIndex > 0 Then
            MobjObjetoWin.ObjIdModulo_ModuloServicioShr.ObjValorPro = cboModuloServicio.SelectedIndex
        End If
    End Sub
    Private Sub SHabiliteValor()
        Dim lobjServicio As ClsServicio = MobjServicio.ObjPadre
        If MobjServicio.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual Then
            If MobjServicio.ObjMiAno.ObjModuloPorServicioBln.ObjValorPro Then
                txtValorCont.Style = FindResource("RecCtlNoHabilitado")
                txtValorCont.Text = Format(0, "c")
            End If
        End If
    End Sub
    Private Sub SMuestreNota()
        Dim lstrNota = "Si este módulo corresponde a un servicio anual, y el año está " &
                "parametrizado para tener un solo módulo de contribucion por cada servicio, " &
                "el valor es calculado a partir del presupuesto anual, de lo contrario debe " &
                "ser ingresado aqui."
        txtNota.Text = lstrNota
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
                            Case "txtValorCont"
                                MobjObjetoWin.ObjValorPres_ModuloServicioDec.ObjValorPro =
                                        txtValorCont.Text
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
                        SMuestreDatos()
                        If Not lblnNoHayError Then
                            SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                        End If
                    End Try
                End If
            End If
        End If
    End Sub
    Private Sub OnCboCambio(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If Not MblnPoblandoComboBox AndAlso Not HblnMostrandoDatos Then
            If TypeOf lelmElemento Is ComboBox Then
                MobjObjetoWin.ObjIdModulo_ModuloServicioShr.ObjValorPro =
                        cboModuloServicio.SelectedIndex
                SMuestreDatos()
            End If
        End If
    End Sub
    Protected Overrides Sub EwinClosed(sender As Object, e As EventArgs)
        If WinPadre IsNot Nothing Then
            If WinPadre.Visibility <> Visibility.Visible Then
                WinPadre.Visibility = Visibility.Visible
                If GblnOK Then
                    Dim lstrMens = String.Empty
                    Dim lwinServicio As WinServicios = WinPadre
                    lwinServicio.SAcepteModuloSer(MblnNuevo, lstrMens)
                    GblnOK = String.IsNullOrEmpty(lstrMens)
                End If
            End If
        End If
    End Sub
#End Region
End Class
