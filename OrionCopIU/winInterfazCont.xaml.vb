Public Class WinInterfazCont
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidaEntradaDef As Integer
        enuFechaDesde
        enuFechaHasta
        enuNroCompInicio
    End Enum
#End Region
    ' Variables
    Private MobjInterfazCont As ClsCBInterfazContableOri = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomIntCon
    Private ReadOnly MblnFinMesAuto As Boolean = False
#End Region
#Region "Constructor"
    Public Sub New(ablnFinMesAuto As Boolean)
        MblnFinMesAuto = ablnFinMesAuto
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuInterfazCont
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SCargueForma(EnuElementosAdicionalesDef.None, 3,
                Nothing, Nothing, True)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        If Not GobjParametros.FblnParametrosInterfazOk Then
            Dim lstrMens = "Por favor revise la Parametrización de los Documentos Contables!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        Else
            SLeaInterfaceIni()
            If MblnFinMesAuto Then
                dtpFechaDesde.IsEnabled = False
                dtpFechaHasta.IsEnabled = False
                HbttAceptar.IsEnabled = False
                HbttCancelar.IsEnabled = False
                SCree()
                SGuarde()
                MyBase.SCerrarClic()
            Else
                SCrearClic()
            End If
        End If
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
        ObjObjetoWin = GobjParametros
        Select Case GobjParametros.ObjIdAppContableByt.ObjValorPro
            Case EnuAppConta.EnuApoloBD, EnuAppConta.EnuApoloAP
                MobjInterfazCont = New ClsInterfazApolo(GCOBJREGISTRO)
            Case EnuAppConta.EnuContaPyme
                MobjInterfazCont = New ClsInterfazContaPyme(GCOBJREGISTRO)
            Case EnuAppConta.EnuSIIGO
                MobjInterfazCont = New ClsInterfazContaSIIGO(GCOBJREGISTRO)
            Case EnuAppConta.EnuSIIGON
                Dim lobjIntCon As New ClsInterfazContaSIIGO(GCOBJREGISTRO)
                MobjInterfazCont = lobjIntCon
                lobjIntCon.BlnNube = True
            Case EnuAppConta.EnuColon
                MobjInterfazCont = New ClsInterfazContaColon(GCOBJREGISTRO)
            Case EnuAppConta.EnuMekano
                MobjInterfazCont = New ClsInterfazMekano(GCOBJREGISTRO)
            Case EnuAppConta.EnuPodium
                MobjInterfazCont = New ClsInterfazPodium(GCOBJREGISTRO)
        End Select
        EnuTipoPermisoObjWin = EnuPermisosDef.enuCrear
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidaEntradaDef.enuFechaDesde) = lblFechaDesde
        StcValidaControl(EnuValidaEntradaDef.enuFechaHasta) = lblFechaHasta
        StcValidaControl(EnuValidaEntradaDef.enuNroCompInicio) = lblIdComInicial
        '
        txtAppContable.Content = GobjParametros.ObjIdAppContableByt.StrNombreApp
        '
        If GobjParametros.ObjTipoInterfazByt.ObjValorPro =
                EnuTipoInterfazDef.EnuPorDocumento Then
            lblIdComInicial.Visibility = Visibility.Collapsed
            txtNroComprobanteInicial.Visibility = Visibility.Collapsed
            Height = 420
        End If
        If GobjParametros.ObjIdAppContableByt.ObjValorPro > EnuAppConta.EnuApoloBD Then
            chkIdTerceroStr.Visibility = Visibility.Collapsed
        End If
        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub
    Protected Overrides Sub SMuestreDatos()
        With MobjInterfazCont
            dtpFechaDesde.SelectedDate = .DtmFechaDesde
            dtpFechaHasta.SelectedDate = .DtmFechaHasta
            txtNroComprobanteInicial.Text = .EntIdComprobanteInicial
            If GobjParametros.ObjTipoInterfazByt.ObjValorPro = EnuTipoInterfazDef.EnuPorDocumento Then
                txtTipoInterface.Content = "Interfaz por Documento"
            Else
                txtTipoInterface.Content = "Interfaz por Comprobante"
            End If
        End With
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        With MobjInterfazCont
            StcValidValido(EnuValidaEntradaDef.enuFechaDesde) = .BlnEsValidaFechaDesde
            StcValidValido(EnuValidaEntradaDef.enuFechaHasta) = .BlnEsValidaFechaHasta
            StcValidValido(EnuValidaEntradaDef.enuNroCompInicio) = .BlnEsValidaIdComprobanteInicial
        End With
        '
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjInterfazCont
            .DtmFechaDesde = dtpFechaDesde.SelectedDate
            .DtmFechaHasta = dtpFechaHasta.SelectedDate
            .BlnIdTerceroStr = chkIdTerceroStr.IsChecked
            If txtNroComprobanteInicial.Visibility = Visibility.Visible Then
                If IsNumeric(txtNroComprobanteInicial.Text) Then
                    .EntIdComprobanteInicial = CType(txtNroComprobanteInicial.Text, Integer)
                Else
                    .EntIdComprobanteInicial = 0
                End If
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
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            With MobjInterfazCont
                If GobjParametros.ObjTipoInterfazByt.ObjValorPro =
                        EnuTipoInterfazDef.EnuPorComprobante Then
                    .EntIdComprobanteInicial = .FentIdUltimoComprobanteInterfazAnterior + 1
                End If
                Dim lobjPeriodoAnt As ClsPeriodo = ClsOrionCop.FobjPeriodoAnterior
                .DtmFechaDesde = lobjPeriodoAnt.DtmFechaInicioPeriodo
                .DtmFechaHasta = lobjPeriodoAnt.DtmFechaFinPeriodo
            End With
            lblnNoHayError = True
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando
                SMuestreDatos()
                SRegistre()
                SFrmAdicione()
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Protected Overrides Sub SGuarde()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Mouse.OverrideCursor = Cursors.Wait
            SRegistre()
            SValide()
            If FblnEstanTodosBien() Then
                SLevanteEveNoti("Generando interfaz contable!", String.Empty, 0,
                        EnuSeveridadNot.EnuInformacion)
                SLevanteEveNoti("Generando interfaz contable!", String.Empty, 0,
                        EnuSeveridadNot.EnuInformacion)
                If MobjInterfazCont.FblnHayDatos Then
                    MobjInterfazCont.SGenerereInterfazContable(MblnFinMesAuto, lstrMens)
                Else
                    lstrMens = "No hay Movimiento entre las Fecha seleccionadas!"
                End If
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PathTooLongException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As DirectoryNotFoundException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As NotSupportedException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                If String.IsNullOrEmpty(lstrMens) Then
                    lstrMens = "La Interfaz Contable se generó exitosamente!"
                End If
                SFinaliceOperacion()
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SLeaInterfaceIni()
        Dim lsrArchivoIni As StreamReader
        Dim lstrLinea As String
        Dim lstrArg() As String
        Dim lstrArchivoIni As String = GstrTrayDatPrg & "Interface.ini"
        lsrArchivoIni = ClsPanorama.FsrStreamReader(lstrArchivoIni)
        If IsNothing(lsrArchivoIni) Then Exit Sub
        lstrLinea = lsrArchivoIni.ReadLine
        Do While Not IsNothing(lstrLinea)
            lstrArg = lstrLinea.Split("=")
            Select Case lstrArg(0)
                Case "IdTerAF"
                    If lstrArg(1) = "0" Then
                        chkIdTerceroStr.IsChecked = False
                    Else
                        chkIdTerceroStr.IsChecked = True
                    End If
            End Select
            lstrLinea = lsrArchivoIni.ReadLine
        Loop
        lsrArchivoIni.Close()
    End Sub
    Private Sub SEscribaInterfaceIni()
        Dim lstrIdTer = "IdTerAF="
        If chkIdTerceroStr.IsChecked Then
            lstrIdTer &= "1"
        Else
            lstrIdTer &= "0"
        End If
        Dim lstrArchivo = GstrTrayDatPrg & "Interface.ini"
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            If My.Computer.FileSystem.FileExists(lstrArchivo) Then
                My.Computer.FileSystem.DeleteFile(lstrArchivo)
            End If
            Using lswArchivoErr = File.AppendText(lstrArchivo)
                lswArchivoErr.WriteLine(lstrIdTer)
                lswArchivoErr.Flush()
            End Using
            lblnNoHayError = True
        Catch ex As ArgumentException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PathTooLongException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As DirectoryNotFoundException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As NotSupportedException
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
            If TypeOf lelmElemento Is DatePicker OrElse TypeOf lelmElemento Is TextBox Then
                If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                    With MobjInterfazCont
                        Select Case lelmElemento.Name
                            Case "dtpFechaDesde"
                                .DtmFechaDesde = dtpFechaDesde.SelectedDate
                            Case "dtpFechaHasta"
                                .DtmFechaHasta = dtpFechaHasta.SelectedDate
                            Case "txtNroComprobanteInicial"
                                If IsNumeric(txtNroComprobanteInicial.Text) Then
                                    .EntIdComprobanteInicial = CType(txtNroComprobanteInicial.Text, Integer)
                                Else
                                    .EntIdComprobanteInicial = 0
                                End If
                        End Select
                    End With
                    SMuestreDatos()
                End If
            End If
        End If
    End Sub
    Protected Overrides Sub EwinClosed(sender As Object, e As EventArgs) Handles Me.Closed
        SEscribaInterfaceIni()
        MyBase.EwinClosed(sender, e)
    End Sub
#End Region
End Class
