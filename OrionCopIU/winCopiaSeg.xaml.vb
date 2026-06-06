Imports Microsoft.Win32
Public Class WinCopiaSeg
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuArchivoCopia
    End Enum
#End Region
    ' Variables
    Private MstrTrayectoriaCopia As String = String.Empty
    Private MstrArchivoCopia As String = String.Empty
    Friend Property BlnReiniciar As Boolean = False
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomCopSeg
    '
    Private WithEvents MbgwWorker As New BackgroundWorker
    Private WithEvents MtmrControl As Forms.Timer = Nothing
    Private MblnExportando As Boolean = False
    Private MblnImportando As Boolean = False
    Private MblnCancele As Boolean = False
    '
    Public Property BlnAutom As Boolean = False
#End Region
#Region "Constructor"
    Friend Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuCopiaSeg
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            SCargueForma(EnuElementosAdicionalesDef.None, 1,
                Nothing, Nothing, True)
            SPuebleBarraEstado(HcolLabelsBarraEstado)
            SCree()
            If BlnAutom Then
                HbttAceptar.IsEnabled = False
                HbttCancelar.IsEnabled = False
                SGuarde()
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
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
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
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
        EnuTipoPermisoObjWin = EnuPermisosDef.enuCrear
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuArchivoCopia) = lblRestaurar
        '
        HbttAceptar.TabIndex = 20
        HbttCancelar.TabIndex = 21
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If Not HblnCargandoForma Then
            txtCopiaSeguridad.Text = MstrArchivoCopia
            txtCopiaSeguridad.ToolTip = MstrArchivoCopia
        End If
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        If MblnExportando Then
            StcValidValido(EnuValidEntradaDef.enuArchivoCopia) = True
        Else
            StcValidValido(EnuValidEntradaDef.enuArchivoCopia) = My.Computer.FileSystem.FileExists(MstrArchivoCopia)
        End If
        '
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
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
    Protected Overrides Sub SHabiliteMenues()
        MyBase.SHabiliteMenues()
    End Sub
    ''' <summary>
    ''' Sub que prepara a la ventana y a su objeto para crear un nuevo objeto. Invalida el Sub
    ''' "SCree" de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SCree()
        EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuCreando
        SFrmAdicione()
        If GstrIdUsuario = GCSTRUSUARIOU Then
            rdbRestaurarBk.Visibility = Visibility.Visible
            bttExaminar.Visibility = Visibility.Visible
        Else
            rdbRestaurarBk.Visibility = Visibility.Collapsed
            bttExaminar.Visibility = Visibility.Collapsed
        End If
        SSeleccioneAccion()
        MtmrControl = New Forms.Timer With {
            .Interval = 100
            }
        rdbGenerabK.Focus()
    End Sub
    ''' <summary>
    ''' . Invalida el procedimiento "SGuarde" de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SGuarde()
        HbttAceptar.IsEnabled = False
        rdbGenerabK.IsEnabled = False
        rdbRestaurarBk.IsEnabled = False
        If rdbGenerabK.IsChecked Then
            SGenereCopiaSeg()
        Else
            If MsgBox("Esta seguro de restaurar la Copia de Seguridad seleccionada?", MsgBoxStyle.Question +
                      MsgBoxStyle.YesNo, "Restaurar Copia de Seguridad") = vbYes Then
                SRestaureCopiaSeg()
            Else
                rdbGenerabK.IsEnabled = True
                HbttAceptar.IsEnabled = True
            End If
        End If
    End Sub
    Protected Overrides Sub SCancele()
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
            If MtmrControl.Enabled Then
                MtmrControl.Stop()
                GobjPanDat.SCanceleAccionCopia()
                Threading.Thread.Sleep(200)
                SFinaliceOperacion()
                MblnCancele = True
            Else
                SCerrarClic()
            End If
        Else
                If ObjObjetoWin.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                ObjObjetoWin.SNormaliceEstado(False)
            End If
            SCerrarClic()
        End If
    End Sub
    Protected Overrides Sub SFinaliceOperacion()
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
            EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando
            HbttCancelar.Content = "_Cierre"
            HbttAceptar.IsEnabled = True
            HbttCancelar.Focus()
        End If
        If BlnAutom Then
            Me.Close()
        End If
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SNombreArchivoCopia()
        MstrTrayectoriaCopia = ClsAdministrador.FobjAppActual.StrTrayCopiaSeguridad
        MstrArchivoCopia = ClsOrionCop.FstrNombreArchivoCopia(BlnAutom)
    End Sub
    Private Sub SSeleccioneAccion()
        If rdbGenerabK.IsChecked Then
            MblnImportando = False
            MblnExportando = True
            SNombreArchivoCopia()
            lblBackUp.Visibility = Visibility.Visible
            lblRestaurar.Visibility = Visibility.Collapsed
            bttExaminar.Visibility = Visibility.Collapsed
        Else
            MblnImportando = True
            MblnExportando = False
            MstrArchivoCopia = String.Empty
            lblBackUp.Visibility = Visibility.Collapsed
            lblRestaurar.Visibility = Visibility.Visible
            bttExaminar.Visibility = Visibility.Visible
        End If
        SMuestreDatos()
    End Sub
    Private Sub SGenereCopiaSeg()
        txtProceso.Content = "Generando Copia de Seguridad"
        pgbAvance.Minimum = 0
        pgbAvance.Maximum = 100
        pgbAvance.Value = 0.0
        MtmrControl.Start()
        MbgwWorker.RunWorkerAsync()
    End Sub
    Private Sub SRestaureCopiaSeg()
        txtProceso.Content = "Restaurando Copia de Seguridad"
        pgbAvance.Minimum = 0
        pgbAvance.Maximum = 100
        pgbAvance.Value = 0.0
        MtmrControl.Start()
        MbgwWorker.RunWorkerAsync()
        BlnReiniciar = True
        Mouse.OverrideCursor = Input.Cursors.Arrow
    End Sub
#End Region
#Region "Eventos del proceso"
    Private Sub EvnTimer(sender As Object, e As EventArgs) Handles MtmrControl.Tick
        If Not MblnCancele Then
            Dim lentAvance = GobjPanDat.EntProcientoAvance
            txtAvance.Content = lentAvance.ToString & " %"
            pgbAvance.Value = lentAvance
        End If
    End Sub
    Private Sub Bgw_DoWork(sender As System.Object, e As DoWorkEventArgs) Handles MbgwWorker.DoWork
        If MblnImportando Then
            GobjPanDat.SRestaureBkPan(MstrArchivoCopia)
        ElseIf MblnExportando Then
            GobjPanDat.SGenereBkPan(MstrArchivoCopia)
        End If
    End Sub
    Private Sub Bgw_RunWorkerCompleted(sender As System.Object,
            e As RunWorkerCompletedEventArgs) Handles MbgwWorker.RunWorkerCompleted
        Dim lstrCopia = MstrArchivoCopia.Substring(MstrArchivoCopia.IndexOf("2", 1))
        Dim lstrMens As String
        MtmrControl.Stop()
        If MblnCancele Then
            lstrMens = "Proceso cancelado por el Usuario"
            If MblnExportando Then
                lstrMens &= "!"
                If My.Computer.FileSystem.FileExists(MstrArchivoCopia) Then
                    My.Computer.FileSystem.DeleteFile(MstrArchivoCopia)
                End If
            Else
                lstrMens &= ". Debe reiniciar Orión Plus!"
            End If
        Else
            txtAvance.Content = "100 %"
            pgbAvance.Value = 100
            If MblnImportando Then
                lstrCopia = "Restaura " & lstrCopia
                GobjPanorama.SRegistreAccionLogApp("Restaurar Copia Seguridad", lstrCopia)
                GobjPanorama.ObjAppActual.SRegistreLogOffOrigen(My.Computer.Name,
                        EnuOrigenInstanciamientoDef.enuEstacionTrabajo)
                lstrMens = "Copia restaurada exitosamente. Debe reiniciar Orión Plus!"
            Else
                lstrCopia = "Genera " & lstrCopia
                GobjPanorama.SRegistreAccionLogApp("Generar Copia Seguridad", lstrCopia)
                lstrMens = "Copia generada exitosamente!"
            End If
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            txtProceso.Content = lstrMens
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
        SFinaliceOperacion()
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
    Private Sub Rdb_Click(sender As Object, e As RoutedEventArgs) Handles rdbGenerabK.Click, rdbRestaurarBk.Click
        If TypeOf sender Is RadioButton Then
            SSeleccioneAccion()
        End If
    End Sub
    Private Sub BttExaminar_Click(sender As Object, e As RoutedEventArgs) Handles bttExaminar.Click
        Dim lofdOrigenDatos As New OpenFileDialog With {
            .DefaultExt = ".sql",
            .Filter = "MySqlBk(.sql)|*.sql",
            .InitialDirectory = MstrTrayectoriaCopia
        }
        Dim lblnOk As Boolean = lofdOrigenDatos.ShowDialog
        If lblnOk Then
            txtCopiaSeguridad.Text = lofdOrigenDatos.FileName
            MstrArchivoCopia = txtCopiaSeguridad.Text
        End If
        SValide()
    End Sub
#End Region
End Class
