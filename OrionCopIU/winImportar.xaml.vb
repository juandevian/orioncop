Imports System.Windows.Controls
Imports System.Data
Imports Microsoft.Win32
Public Class WinImportar
#Region "Definiciones"
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuRutaNombre = 0
        enuTablaOrigen
        enuListaCamposRelacionados
        enuValidos
    End Enum
#End Region
#Region "Delegados"
    Private Delegate Sub SdgtActualizaProgressBar(dp As DependencyProperty, value As Object)
    Private Delegate Sub SdgtActualizaLabel(dp As DependencyProperty, Content As Object)
    Private MdgtPgbActualiza As SdgtActualizaProgressBar = Nothing
    Private MdgtLblActualiza As SdgtActualizaLabel = Nothing
#End Region
    ' Variables
    Private WithEvents MobjObjetoWin As ClsImportar = Nothing
    Private ReadOnly MobjObjetoDestino As ClsCBObjetoPan = Nothing
    Private MdblCantObjetosProcesados As Double = 0D
    Private MdblCantObjetosAImportar As Double = 0D
    Private MblnCamposValidos As Boolean = False
    Private MstrCampoRelacionado As String = Nothing
    Private MintSelectedIndex As Integer = -1
    Private MblnPoblandoCbo As Boolean = False
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomImp
    Public Property BlnExigeRequeridos As Boolean = False
#End Region
#Region "Constructor"
    Public Sub New(aobjObjetoDestino As ClsCBObjetoPan)
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuImportar
        MobjObjetoDestino = aobjObjetoDestino
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicioneControlRestringido(lsbCamposRelacionados)
        SCargueForma(EnuElementosAdicionalesDef.None, 4,
                Nothing, txtOrigenDatos, False)
        If Not MobjObjetoWin.BlnHayDatosImportacion Then
            SCrearClic()
        Else
            SModificarClic()
        End If
        SPuebleListasCamposObjetoDestino()
        SPuebleListaCamposRelacionados()
        MblnCamposValidos = False
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
        MobjObjetoWin = New ClsImportar(MobjObjetoDestino)
        ObjObjetoWin = MobjObjetoWin
        MobjObjetoWin.ObjExigeRequeridosBln.ObjValorPro = BlnExigeRequeridos
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuRutaNombre) = lblOrigenDatos
        StcValidaControl(EnuValidEntrada.enuTablaOrigen) = lblTablasCargadas
        StcValidaControl(EnuValidEntrada.enuListaCamposRelacionados) = lblCamposRelacionados
        StcValidaControl(EnuValidEntrada.enuValidos) = lblValidar
        txtUltimaImportacion.Content = MobjObjetoWin.ObjArchivoOrigenStr.ObjValorPro
        txtUltimaImportacion.ToolTip = txtUltimaImportacion.Content
        chkExigeRequeridos.Visibility = Visibility.Visible
        chkExigeRequeridos.IsChecked = BlnExigeRequeridos
        '
        HbttAceptar.TabIndex = 29
        HbttCancelar.TabIndex = 30
    End Sub
    Private Sub SVacie()
        txtOrigenDatos.Text = String.Empty
        lsbCamposOrigen.Items.Clear()
        lsbCamposRelacionados.Items.Clear()
        txtUltimaImportacion.Content = String.Empty
        txtTabla.Content = String.Empty
        cboTablas.Items.Clear()
        With MobjObjetoWin
            .ObjArchivoOrigenStr.ObjValorPro = String.Empty
            .ObjTablaOrigenStr.ObjValorPro = String.Empty
            .ObjColumnasRelacionadasStr.ObjValorPro = Nothing
        End With
        chkExigeRequeridos.IsChecked = False
        SPuebleListasCamposObjetoDestino()
    End Sub
    Protected Overrides Sub SMuestreDatos()
        With MobjObjetoWin
            txtOrigenDatos.Text = .ObjArchivoOrigenStr.ObjValorPro
            txtOrigenDatos.ToolTip = txtOrigenDatos.Text
            txtTabla.Content = .ObjTablaOrigenStr.ObjValorPro
            chkExigeRequeridos.IsChecked = .ObjExigeRequeridosBln.ObjValorPro
        End With
        Title = My.Resources.ImportarObjeto
        Title &= MobjObjetoDestino.StrNombreClase
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        With MobjObjetoWin
            StcValidValido(EnuValidEntrada.enuRutaNombre) = .ObjArchivoOrigenStr.BlnEsValido
            StcValidValido(EnuValidEntrada.enuTablaOrigen) = .ObjTablaOrigenStr.BlnEsValido
            StcValidValido(EnuValidEntrada.enuListaCamposRelacionados) = .ObjColumnasRelacionadasStr.BlnEsValido
            StcValidValido(EnuValidEntrada.enuValidos) = MblnCamposValidos
        End With
        SHabiliteBotonesTlb()
        If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
            HbttModificar.Style = FindResource("RecBttHabilitado")
            HmnuModificar.Style = FindResource("RecMnuItemSecHab")
        End If
        bttValidar.IsEnabled = MobjObjetoWin.ObjArchivoOrigenStr.BlnEsValido AndAlso
                MobjObjetoWin.ObjTablaOrigenStr.BlnEsValido
        If FblnEstanTodosBien() Then
            HbttAceptar.Style = FindResource("RecBttAceCan")
        Else
            HbttAceptar.Style = FindResource("RecBttAceDesha")
        End If
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
        '
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    Protected Overrides Sub SCancele()
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
            HblnCancelando = True
            SFinaliceOperacion()
            HblnCancelando = False
            GblnOK = False
        End If
        SCerrarClic()
    End Sub
    Protected Overrides Sub SGuarde()
        Dim lblnNoHayError = False, lstrMens = String.Empty, lstrMensEx = String.Empty
        Try
            GobjPanDat.SControleProcesoObj(True)
            HbttCancelar.IsEnabled = False
            Dim lblnGuardo = FblnImporto()
            HbttCancelar.IsEnabled = True
            If lblnGuardo Then
                ObjObjetoWin.SActualice(True)
                SFinaliceOperacion()
                If FblnEstanTodosBien() Then
                    SLevanteEveNoti("La Importación terminó exitosamente!", "", 0,
                            EnuSeveridadNot.EnuInformacion)
                End If
                SMuestreInforme()
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
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SSeleccioneOrigen() 'Ok
        Dim lofdOrigenDatos As New OpenFileDialog With {
            .DefaultExt = ".xlsx",
            .Filter = My.Resources.TiposArchivosOrigen
        }
        Dim lblnOk As Boolean = lofdOrigenDatos.ShowDialog
        If lblnOk Then
            txtOrigenDatos.Text = lofdOrigenDatos.FileName
            SRegistreOrigenDatos()
            SMuestreDatos()
        End If
    End Sub
    Private Sub SRegistreOrigenDatos()
        Dim lblnLimpiar As Boolean = (MobjObjetoWin.ObjArchivoOrigenStr.ObjValorPro <> txtOrigenDatos.Text)
        MobjObjetoWin.ObjArchivoOrigenStr.ObjValorPro = txtOrigenDatos.Text
        If lblnLimpiar Then
            cboTablas.Items.Clear()
            SPuebleListasCamposObjetoDestino()
            lsbCamposOrigen.Items.Clear()
            lsbCamposRelacionados.Items.Clear()
        End If
    End Sub
    Private Sub SCargueOrigen()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            SPuebleComboTablas()
            lblnNoHayError = True
        Catch ex As ParametrosConexionBDPanException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ProveedorBdPanException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                SPuebleComboTablas()
            Else
                SFinaliceOperacion()
                SCancele()
                WinPadre.SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Private Sub SPuebleComboTablas()
        MblnPoblandoCbo = True
        cboTablas.Items.Clear()
        cboTablas.Items.Add("<Ninguna>")
        Dim lstrTablasOrigen As String() = MobjObjetoWin.FstrTablasOrigen
        If Not IsNothing(lstrTablasOrigen) Then
            For Each lstrNombreTabla As String In lstrTablasOrigen
                cboTablas.Items.Add(lstrNombreTabla)
            Next
        End If
        MblnPoblandoCbo = False
        cboTablas.SelectedIndex = 0
        SValide()
    End Sub
    Private Sub SPuebleListaCamposOrigen() 'Ok
        Dim lstrColumnasOrigen As String() = MobjObjetoWin.FstrColumnasOrigen
        If lstrColumnasOrigen IsNot Nothing Then
            lsbCamposOrigen.Items.Clear()
            For Each lstrNombreCampo As String In lstrColumnasOrigen
                lsbCamposOrigen.Items.Add(lstrNombreCampo.ToString.ToLower)
            Next
        End If
    End Sub
    Private Sub SPuebleListasCamposObjetoDestino() 'Ok
        Dim lstrArrayColumnasObjeto As String() = MobjObjetoWin.FstrColumnasObjetoDes
        Dim lstrArrayColumnasRequeridas As String() = MobjObjetoWin.StrColumnasRequeridas
        lsbCamposDestino.Items.Clear()
        lsbCamposRequeridos.Items.Clear()
        For Each lstrNombreCampo As String In lstrArrayColumnasObjeto
            lsbCamposDestino.Items.Add(lstrNombreCampo.ToString.ToLower)
        Next
        For Each lstrNombreCampo As String In lstrArrayColumnasRequeridas
            lsbCamposRequeridos.Items.Add(lstrNombreCampo.ToString.ToLower)
        Next
        SValide()
    End Sub
    Private Sub SPuebleListaCamposRelacionados()
        If IsArray(MobjObjetoWin.ObjColumnasRelacionadasStr.ObjValorPro) Then
            Dim lstrCamRel As String() = MobjObjetoWin.ObjColumnasRelacionadasStr.ObjValorPro
            If Not IsNothing(lstrCamRel) AndAlso lstrCamRel.Length > 0 Then
                lsbCamposRelacionados.Items.Clear()
                For Each lstrCampoRelacionado As String In lstrCamRel
                    lsbCamposRelacionados.Items.Add(lstrCampoRelacionado.ToString.ToLower)
                    Dim lstrNombreCampo As String()
                    lstrNombreCampo = lstrCampoRelacionado.Split("=")
                    If lsbCamposDestino.Items.Contains(lstrNombreCampo(0).ToString.ToLower) Then
                        lsbCamposDestino.Items.Remove(lstrNombreCampo(0).ToString.ToLower)
                    End If
                    If lsbCamposOrigen.Items.Contains(lstrNombreCampo(1).ToString.ToLower) Then
                        lsbCamposOrigen.Items.Remove(lstrNombreCampo(1).ToString.ToLower)
                    End If
                    If lsbCamposRequeridos.Items.Contains(lstrNombreCampo(0).ToString.ToLower) Then
                        lsbCamposRequeridos.Items.Remove(lstrNombreCampo(0).ToString.ToLower)
                    End If
                Next
            End If
        End If
        SValide()
    End Sub
    Private Sub SRelacioneCampos()
        If Not (lsbCamposDestino.SelectedIndex = -1 OrElse lsbCamposOrigen.SelectedIndex = -1) Then
            Dim lstrCampoDestino As String = lsbCamposDestino.SelectedItem.ToString
            Dim lstrCampoOrigen As String = lsbCamposOrigen.SelectedItem.ToString
            lsbCamposRelacionados.Items.Add(lstrCampoDestino & "=" & lstrCampoOrigen)
            lsbCamposDestino.Items.RemoveAt(lsbCamposDestino.SelectedIndex)
            lsbCamposOrigen.Items.RemoveAt(lsbCamposOrigen.SelectedIndex)
            If lsbCamposRequeridos.Items.Contains(lstrCampoDestino) Then
                lsbCamposRequeridos.Items.Remove(lstrCampoDestino)
            End If
        End If
        MobjObjetoWin.ObjColumnasRelacionadasStr.ObjValorPro = FstrColumnasRelacionadas()
    End Sub
    Private Sub SQuiteRelacionCampos()
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
            If Not IsNothing(MstrCampoRelacionado) Then
                Dim lstrCampos As String()
                lstrCampos = MstrCampoRelacionado.Split("=")
                lsbCamposDestino.Items.Add(lstrCampos(0))
                lsbCamposOrigen.Items.Add(lstrCampos(1))
                lsbCamposRelacionados.Items.RemoveAt(MintSelectedIndex)
            End If
            MstrCampoRelacionado = Nothing
            MintSelectedIndex = -1
        End If
        MobjObjetoWin.ObjColumnasRelacionadasStr.ObjValorPro = FstrColumnasRelacionadas()
    End Sub
    Private Function FstrColumnasRelacionadas() As String()
        Dim lstrColRel As String() = Nothing
        If lsbCamposRelacionados.Items.Count > 0 Then
            ReDim lstrColRel(lsbCamposRelacionados.Items.Count - 1)
            Dim i As Short = -1
            For Each lstrRelCol As String In lsbCamposRelacionados.Items
                i += 1
                lstrColRel(i) = lstrRelCol
            Next
        End If
        Return lstrColRel
    End Function
    Private Function FblnImporto() As Boolean
        Mouse.OverrideCursor = Cursors.Wait
        Dim ldtbDatosOrigen As DataTable = MobjObjetoWin.FdtbTablaOrigen
        MdblCantObjetosAImportar = ldtbDatosOrigen.Rows.Count
        pgbImportar.Minimum = 0
        pgbImportar.Maximum = MdblCantObjetosAImportar
        pgbImportar.Value = 0
        MdgtPgbActualiza = New SdgtActualizaProgressBar(AddressOf pgbImportar.SetValue)
        MdgtLblActualiza = New SdgtActualizaLabel(AddressOf lblResultado.SetValue)
        Dim lblnImporto = MobjObjetoWin.FblnImportoDatos(ldtbDatosOrigen)
        Mouse.OverrideCursor = Cursors.Arrow
        Return lblnImporto
    End Function
    Private Function FshrIdAno(adtbOrigen As DataTable, ByRef ashrIdServ As Short) As Short
        Dim lstrColOriAno = MobjObjetoWin.FstrColumnaOrigen(
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd)
        Dim lstrColOriServ = MobjObjetoWin.FstrColumnaOrigen(
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd)
        Dim ldrwOrigen = adtbOrigen.Rows(0)
        Dim lshrIdAno = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColOriAno),
                EnuTipoValorDef.enuShort)
        Dim lshrIdServ = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColOriServ),
                EnuTipoValorDef.enuShort)
        ashrIdServ = lshrIdServ
        Return lshrIdAno
    End Function
    Protected Overrides Sub SFinaliceOperacion()
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
            If Not IsNothing(HbttCancelar) Then
                HbttCancelar.Content = My.Resources.CerrarBtn
            End If
            EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando
            SMuestreDatos()
            MobjObjetoWin.SNormaliceEstado(True)
            SHabiliteWin(False)
            SHabiliteBotonesTlb()
            If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
                HbttModificar.Style = FindResource("RecBttHabilitado")
                HmnuModificar.Style = FindResource("RecMnuItemSecHab")
            End If
            Dim NoUsado = FblnEstanTodosBien()
        End If
    End Sub
    Private Sub SMuestreInforme()
        Dim lstrNombreArchivo = MobjObjetoWin.StrNombreArchivoResultados
        If My.Computer.FileSystem.FileExists(lstrNombreArchivo) Then
            Process.Start("notepad.exe", lstrNombreArchivo)
        End If
    End Sub
    Private Sub SValideTabla()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Dim lblnDatosValidos = False
        Try
            Mouse.OverrideCursor = Cursors.Wait
            lstrMens = "Validando Datos de origen!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            GobjPanDat.SControleProcesoObj(True)
            lblnDatosValidos = MobjObjetoWin.FblnSonValidosDatosTabla
            MblnCamposValidos = lblnDatosValidos
            SValide()
            lblnNoHayError = True
        Catch ex As ParametrosConexionBDPanException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ProveedorBdPanException
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
                GobjPanDat.SControleProcesoObj(False)
                If lblnDatosValidos Then
                    lstrMens = "Los datos de Origen son válidos"
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuInformacion)
                End If
            Else
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub
    Private Sub SNotifiqueIniVal()
        Dim lstrMens = "Validando Datos de origen!"
        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            Dim lstrNombreAccion As String = lelmElemento.Name.Substring(3)
            Select Case lstrNombreAccion
                Case "Examinar"
                    SSeleccioneOrigen()
                Case "Cargar"
                    SCargueOrigen()
                Case "Relacionar"
                    SRelacioneCampos()
                Case "DeshacerRel"
                    SQuiteRelacionCampos()
                Case "Validar"
                    SValideTabla()
            End Select
        End If
    End Sub
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
            If TypeOf lelmElemento Is TextBox OrElse TypeOf lelmElemento Is PasswordBox OrElse
                    TypeOf lelmElemento Is ComboBox OrElse TypeOf lelmElemento Is ListBox Then
                Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                Try
                    With MobjObjetoWin
                        Select Case True
                            Case lelmElemento.Equals(txtOrigenDatos)
                                SRegistreOrigenDatos()
                            Case lelmElemento.Equals(lsbCamposRelacionados)
                                .ObjColumnasRelacionadasStr.ObjValorPro = FstrColumnasRelacionadas()
                        End Select
                        SValide()
                    End With
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
    End Sub
    Private Sub OnCboCambio(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is ComboBox Then
            If lelmElemento.Name = "cboTablas" Then
                If Not MblnPoblandoCbo Then
                    If cboTablas.SelectedItem <> "<Ninguna>" AndAlso
                                Not String.IsNullOrEmpty(cboTablas.Text) Then
                        MobjObjetoWin.ObjTablaOrigenStr.ObjValorPro = cboTablas.SelectedItem
                        SPuebleListasCamposObjetoDestino()
                        SPuebleListaCamposOrigen()
                        SPuebleListaCamposRelacionados()
                        MblnCamposValidos = False
                        SValide()
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub Ctl_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lsbCamposRelacionados.SelectionChanged
        If lsbCamposRelacionados.SelectedIndex > -1 Then
            MstrCampoRelacionado = lsbCamposRelacionados.SelectedItem.ToString
            MintSelectedIndex = lsbCamposRelacionados.SelectedIndex
        End If
    End Sub
    Private Sub ChkExigeRequeridos_Click(sender As Object, e As RoutedEventArgs) Handles chkExigeRequeridos.Click
        MobjObjetoWin.ObjExigeRequeridosBln.ObjValorPro = chkExigeRequeridos.IsChecked
    End Sub
    Private Sub SEvnAvance(aobjSender As Object, e As ClsPanEventArgs) Handles MobjObjetoWin.EvnImportadoObj
        MdblCantObjetosProcesados += 1
        Dim lstrResultado = My.Resources.RegProcesados & Format(MdblCantObjetosProcesados, "#0") &
                My.Resources.De & Format(MdblCantObjetosAImportar, "#0")
        Dispatcher.Invoke(MdgtPgbActualiza,
             System.Windows.Threading.DispatcherPriority.Background,
             New Object() {ProgressBar.ValueProperty, MdblCantObjetosProcesados})
        Dispatcher.Invoke(MdgtLblActualiza,
                System.Windows.Threading.DispatcherPriority.Background,
                New Object() {Label.ContentProperty, lstrResultado})
    End Sub
    Private Sub EFinImportar() Handles MobjObjetoWin.EvnFinImportar
        MdblCantObjetosProcesados = 0
    End Sub
#End Region
End Class