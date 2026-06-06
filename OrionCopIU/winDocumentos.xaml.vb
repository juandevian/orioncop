Public Class WinDocumentos
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuIdDoc
        enuNomDoc
        enuPrefijo
        enuNroInicial
        enuTipoDoc
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsDocumento = Nothing
    '
    Private ReadOnly MstrNombreVentana As String = My.Resources.Documentos
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuDocumento
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolControlesLlave As New Collection
        SAdicioneControlRestringido(dgrDocumentos)
        lcolControlesLlave.Add(txtIdDocumento)
        SCargueForma(EnuElementosAdicionalesDef.None, 5,
                lcolControlesLlave, txtTipoDocumento, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SVinculeDocs()
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
            ObjObjetoWin = GobjParametros.ObjDocumento(EnuIdDocumentoDef.EnuFacturaVenta)
        End If
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuIdDoc) = lblIdDocumento
        StcValidaControl(EnuValidEntrada.enuNomDoc) = lblNombre
        StcValidaControl(EnuValidEntrada.enuNroInicial) = lblNroInicial
        StcValidaControl(EnuValidEntrada.enuPrefijo) = lblPrefijo
        StcValidaControl(EnuValidEntrada.enuTipoDoc) = lblTipoDocumento
        If GobjParametros.FblnHayDocumentosPorCrear Then
            HbttCrear.Visibility = Visibility.Visible
            HmnuCrear.Visibility = Visibility.Visible
        End If
        '
        HbttAceptar.TabIndex = 50
        HbttCancelar.TabIndex = 51
    End Sub
    Protected Overrides Sub SMuestreDatos()
        With MobjObjetoWin
            txtIdDocumento.Text = .ObjIdDocumentoEnt.ObjValorPro
            txtNombre.Content = .ObjNombre_DocStr.ToString
            txtPrefijo.Text = .ObjPrefijo_DocStr.ToString
            txtNroInicial.Text = .ObjNumeroInicial_DocEnt.ToString
            txtTipoDocumento.Text = .ObjTipoDocumentoStr.ObjValorPro
        End With
        Title = My.Resources.Docs
        SValide()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            txtIdDocumento.SelectAll()
        End If
    End Sub
    Protected Overrides Sub SValide()
        With MobjObjetoWin
            StcValidValido(EnuValidEntrada.enuIdDoc) = .ObjIdDocumentoEnt.BlnEsValido
            StcValidValido(EnuValidEntrada.enuNomDoc) = .ObjNombre_DocStr.BlnEsValido
            StcValidValido(EnuValidEntrada.enuNroInicial) = .ObjNumeroInicial_DocEnt.BlnEsValido
            StcValidValido(EnuValidEntrada.enuPrefijo) = .ObjPrefijo_DocStr.BlnEsValido
            StcValidValido(EnuValidEntrada.enuTipoDoc) = .ObjTipoDocumentoStr.BlnEsValido
        End With
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjIdDocumentoEnt.ObjValorPro = txtIdDocumento.Text
            .ObjNombre_DocStr.ObjValorPro = txtNombre.Content
            .ObjPrefijo_DocStr.ObjValorPro = txtPrefijo.Text
            .ObjNumeroInicial_DocEnt.ObjValorPro = txtNroInicial.Text
            .ObjTipoDocumentoStr.ObjValorPro = txtTipoDocumento.Text
        End With
        SValide()
    End Sub
    Protected Overrides Sub SConfigureMenuesPropios()
        '
    End Sub
#End Region
#Region "Invalida otros metodos se la clase base"
    Protected Overrides Sub SHabiliteMenues()
        MyBase.SHabiliteMenues()
    End Sub
    Protected Overrides Sub SCree()
        MobjObjetoWin = GobjParametros.FobjNuevoDocumento
        If Not IsNothing(MobjObjetoWin) Then
            ObjObjetoWin = MobjObjetoWin
            MyBase.SCree()
            txtIdDocumento.IsEnabled = False
        End If
    End Sub
    Protected Overrides Sub SFinaliceOperacion()
        MyBase.SFinaliceOperacion()
        SVinculeDocs()
        MobjObjetoWin = GobjParametros.ObjDocumento(EnuIdDocumentoDef.EnuFacturaVenta)
        ObjObjetoWin = MobjObjetoWin
        SMuestreDatos()
    End Sub
    Protected Overrides Sub SRefresqueWin()
        GobjParametros.SRefresqueObj()
        ObjObjetoWin = Nothing
        SInicialiceObjeto()
        SMuestreDatos()
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SAbraDocumento()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                If GobjParametros.ColDocumentos.Count > 0 Then
                    If GobjParametros.ColDocumentos.Contains(txtIdDocumento.Text) Then
                        MobjObjetoWin = GobjParametros.ColDocumentos(txtIdDocumento.Text)
                    Else
                        MobjObjetoWin = GobjParametros.ColDocumentos("1")
                        lstrMens = "El Documento ingresado no existe!"
                    End If
                End If
                dgrDocumentos.SelectedIndex = MobjObjetoWin.ObjIdDocumentoEnt.ObjValorPro - 1
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
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    End If
                Else
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
    End Sub
    Private Sub SVinculeDocs()
        Dim ldtbDocumentos = GobjParametros.FdtbDocumentos
        dgrDocumentos.DataContext = ldtbDocumentos
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
        If Not HblnSeEstaCerrando Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is TextBox Then
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    With MobjObjetoWin
                        Select Case lelmElemento.Name
                            Case "txtPrefijo"
                                .ObjPrefijo_DocStr.ObjValorPro = txtPrefijo.Text
                            Case "txtNroInicial"
                                .ObjNumeroInicial_DocEnt.ObjValorPro = txtNroInicial.Text
                            Case "txtTipoDocumento"
                                .ObjTipoDocumentoStr.ObjValorPro = txtTipoDocumento.Text
                        End Select
                    End With
                    SMuestreDatos()
                Else
                    If lelmElemento.Name = "txtIdCliente" Then
                        SAbraDocumento()
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub TxtIdDocumento_KeyDown(sender As Object, e As KeyEventArgs) Handles txtIdDocumento.KeyDown
        If e.Key = Key.Return Then
            If txtIdDocumento.Focus Then
                If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                    SAbraDocumento()
                End If
            End If
        End If
    End Sub
    Private Sub DgrDocumentos_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles dgrDocumentos.SelectionChanged
        Dim ldrvFilaActual As DataRowView
        If TypeOf sender Is DataGrid Then
            Dim ldgrActual As DataGrid = sender
            Dim lenuIdDocu As EnuIdDocumentoDef
            ldrvFilaActual = ldgrActual.SelectedItem
            If Not IsNothing(ldrvFilaActual) Then
                lenuIdDocu = CType(ldrvFilaActual("IdDocumento"), EnuIdDocumentoDef)
            Else
                lenuIdDocu = EnuIdDocumentoDef.EnuFacturaVenta
            End If
            MobjObjetoWin = GobjParametros.ObjDocumento(lenuIdDocu)
            ObjObjetoWin = MobjObjetoWin
            SMuestreDatos()
        End If
    End Sub
    Private Sub ClsFormInterface_KeyUp(sender As Object, e As KeyEventArgs)
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            Dim i As Integer
            If e.Key = Key.Down Then
                i = dgrDocumentos.SelectedIndex
                If i < dgrDocumentos.Items.Count - 1 Then
                    dgrDocumentos.SelectedIndex = i + 1
                End If
            ElseIf e.Key = Key.Up Then
                i = dgrDocumentos.SelectedIndex
                If i > 0 Then
                    dgrDocumentos.SelectedIndex = i - 1
                End If
            End If
        End If
    End Sub
#End Region
End Class
