Public Class WinParametrosRep
#Region "Definiciones"
    Private MdtmFechaDesde As Date = Date.Today.AddDays(-Today.Day + 1)
    Private MdtmFechaHasta As Date = Date.Today.Date, mdtmFechaCartera As Date = Date.Today
    Private MblnFechasOk As Boolean = False
    Private MblnLimitesOk As Boolean = False
    Private MblnCerrando As Boolean = False
    Private MstrMensaje As String = String.Empty
    Private MshrIdAno As Short = 0
    Private MentIdServicio As Integer = 0
    Private MblnPoblandoCbos As Boolean = False
    Private MblnServicioOk As Boolean = False
    Private MenuTipoRepEdadCart As EnuTipoRepEdadCartera = EnuTipoRepEdadCartera.None
    ' Limites edad cartera
    Private MentLimite1 As Integer = GentLimite1
    Private MentLimite2 As Integer = GentLimite2
    Private MentLimite3 As Integer = GentLimite3
    Private MentLimite4 As Integer = GentLimite4
    Friend Property ObjRepOrionCop As ClsRepOrionCop = Nothing
    Friend Property EnuReporte As EnuReporteDef = EnuReporteDef.None
#End Region

#Region "Genera Reportes"
    Private Sub SGenereCajaBancos()
        Mouse.OverrideCursor = Cursors.Wait
        If MblnFechasOk Then
            Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaDesde)
            Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaHasta)
            With ObjRepOrionCop
                .StrFechaDesde = lstrFechaDesde
                .StrFechaHasta = lstrFechaHasta
                .EnuReporte = EnuReporteDef.enuCajaBancos
                MstrMensaje = .SGenereCajaBancos()
            End With
        End If
        SMuestreMensaje()
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Sub SGenereRCFechas()
        Mouse.OverrideCursor = Cursors.Wait
        If MblnFechasOk Then
            Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaDesde)
            Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaHasta.AddDays(1))
            With ObjRepOrionCop
                .StrFechaDesde = lstrFechaDesde
                .StrFechaHasta = lstrFechaHasta
                .EnuReporte = EnuReporteDef.enuRCFechas
                MstrMensaje = .SGenereRCFechas()
            End With
        End If
        SMuestreMensaje()
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Sub SGenereRCReversados()
        Mouse.OverrideCursor = Cursors.Wait
        If MblnFechasOk Then
            Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaDesde)
            Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaHasta)
            With ObjRepOrionCop
                .StrFechaDesde = lstrFechaDesde
                .StrFechaHasta = lstrFechaHasta
                .EnuReporte = EnuReporteDef.enuRecCajaReversados
                MstrMensaje = .SGenereRCReversados()
            End With
        End If
        SMuestreMensaje()
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Sub SGenereRelDocs()
        Mouse.OverrideCursor = Cursors.Wait
        If MblnFechasOk Then
            Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaDesde)
            Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaHasta)
            With ObjRepOrionCop
                .StrFechaDesde = lstrFechaDesde
                .StrFechaHasta = lstrFechaHasta
                .EnuReporte = EnuReporteDef.enuRelDocs
                .SGenereRelDocs()
            End With
        End If
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Sub SGenereInfDiario()
        Mouse.OverrideCursor = Cursors.Wait
        If MblnFechasOk Then
            Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaDesde)
            Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaHasta)
            With ObjRepOrionCop
                .StrFechaDesde = lstrFechaDesde
                .StrFechaHasta = lstrFechaHasta
                .EnuReporte = EnuReporteDef.enuInformeDiario
                MstrMensaje = .SGenereInformeDiario()
            End With
        End If
        SMuestreMensaje()
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Sub SGenereCarteraCliente()
        Mouse.OverrideCursor = Cursors.Wait
        If MblnFechasOk Then
            Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(mdtmFechaCartera)
            Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(mdtmFechaCartera)
            With ObjRepOrionCop
                .StrFechaDesde = lstrFechaDesde
                .StrFechaHasta = lstrFechaHasta
                .EnuReporte = EnuReporteDef.enuCarteraPorCliente
                MstrMensaje = .SGenereCarteraPorCliente()
            End With
        End If
        SMuestreMensaje()
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Sub SGenereCxCDetPorSer()
        Mouse.OverrideCursor = Cursors.Wait
        If MblnFechasOk Then
            Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(mdtmFechaCartera)
            Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(mdtmFechaCartera)
            With ObjRepOrionCop
                .StrFechaDesde = lstrFechaDesde
                .StrFechaHasta = lstrFechaHasta
                .EnuReporte = EnuReporteDef.enuCxCDetPorSer
                MstrMensaje = .SGenereCxCDetPorSer()
            End With
        End If
        SMuestreMensaje()
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Sub SGenereEdadCartera()
        Mouse.OverrideCursor = Cursors.Wait
        If MblnLimitesOk Then
            MstrMensaje = "Generando y exportando Reporte Edad de la Cartera!"
            SMuestreMensaje()
            MstrMensaje = ObjRepOrionCop.SGenereEdadCartera(GentLimite1, GentLimite2, GentLimite3,
                    GentLimite4, MenuTipoRepEdadCart, Date.Now, False, False)
            If String.IsNullOrEmpty(MstrMensaje) Then
                MstrMensaje = "Reporte generado y exportado exitosamente"
            End If
        End If
        SMuestreMensaje()
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Sub SGenereItemsProgramaFact()
        Mouse.OverrideCursor = Cursors.Wait
        If MblnServicioOk Then
            With ObjRepOrionCop
                .ShrIdAno = MshrIdAno
                .EntIdServicio = MentIdServicio
                .EnuReporte = EnuReporteDef.enuItemsProgramaFact
                MstrMensaje = .SGenereItemsProgramaFact()
            End With
        End If
        SMuestreMensaje()
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Sub SGenereResumenMovCont()
        Mouse.OverrideCursor = Cursors.Wait
        If MblnFechasOk Then
            Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaDesde)
            Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaHasta)
            With ObjRepOrionCop
                .StrFechaDesde = lstrFechaDesde
                .StrFechaHasta = lstrFechaHasta
                .EnuReporte = EnuReporteDef.enuResumenMovCont
                MstrMensaje = .SGenereResumenMovCont()
            End With
        End If
        SMuestreMensaje()
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Sub SExporteRecibosCaja()
        If MblnFechasOk Then
            Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaDesde)
            Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaHasta)
            With ObjRepOrionCop
                .StrFechaDesde = lstrFechaDesde
                .StrFechaHasta = lstrFechaHasta
                .SExporteRecibosCaja()
            End With
        End If
    End Sub
    Private Sub SExporteFacturas()
        If MblnFechasOk Then
            Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaDesde)
            Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaHasta)
            With ObjRepOrionCop
                .StrFechaDesde = lstrFechaDesde
                .StrFechaHasta = lstrFechaHasta
                .SExporteFacturas()
            End With
        End If
    End Sub
    Private Sub SGenereValoresFacturados()
        Mouse.OverrideCursor = Cursors.Wait
        If MblnFechasOk Then
            Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaDesde)
            Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaHasta)
            With ObjRepOrionCop
                .StrFechaDesde = lstrFechaDesde
                .StrFechaHasta = lstrFechaHasta
                .EnuReporte = EnuReporteDef.enuValoresFactTodos
                MstrMensaje = .SGenereRepVlrsFacturadosTodos()
            End With
        End If
        SMuestreMensaje()
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
#End Region

#Region "Habilita Controles"
    Private Sub SInvisibiliceCnvs()
        cnvCajaBancos.Visibility = System.Windows.Visibility.Collapsed
        cnvCarteraCliente.Visibility = System.Windows.Visibility.Collapsed
        cnvEdadCartera.Visibility = System.Windows.Visibility.Collapsed
        cnvInfDiario.Visibility = System.Windows.Visibility.Collapsed
        cnvItemsProgramaFact.Visibility = System.Windows.Visibility.Collapsed
    End Sub
    Private Sub SHabiliteControles()
        Height = 200
        Select Case EnuReporte
            Case EnuReporteDef.enuCajaBancos
                SHabiliteCajaBancos()
            Case EnuReporteDef.enuInformeDiario
                SHabiliteInfDiario()
            Case EnuReporteDef.enuCarteraPorCliente, EnuReporteDef.enuCxCDetPorSer
                SHabiliteCarteraClientes()
            Case EnuReporteDef.enuEdadCartera
                SHabiliteEdadCartera()
            Case EnuReporteDef.enuItemsProgramaFact
                SHabiliteItemsProgramaFact()
            Case EnuReporteDef.enuResumenMovCont
                SHabiliteResumenMovCont()
            Case EnuReporteDef.enuExpRecsCaja, EnuReporteDef.enuRCFechas, EnuReporteDef.enuRelDocs,
                    EnuReporteDef.enuExpFacsFechas, EnuReporteDef.enuValoresFactTodos,
                    EnuReporteDef.enuRecCajaReversados
                SHabiliteFechas()
        End Select
    End Sub
    Private Sub SHabiliteCajaBancos()
        Me.Title = "Reporte de Ingresos a Caja y Bancos"
        cnvCajaBancos.Visibility = System.Windows.Visibility.Visible
        lblFechaDesde.Content = "Mostrar los Ingresos entre el:"
    End Sub
    Private Sub SHabiliteFechas()
        Select Case EnuReporte
            Case EnuReporteDef.enuExpRecsCaja
                Me.Title = "Exportar Recibos de Caja"
                lblInforme.Content = "Exportar Recibos de Caja entre el:"
                dtpFechaDesdeIng.SelectedDate = MdtmFechaDesde
                dtpFechaHastaIng.SelectedDate = MdtmFechaHasta.AddDays(-1)
                cnvInfDiario.Visibility = System.Windows.Visibility.Visible
            Case EnuReporteDef.enuExpFacsFechas
                Me.Title = "Exportar Facturas"
                lblInforme.Content = "Exportar Facturas entre el:"
                dtpFechaDesdeIng.SelectedDate = MdtmFechaDesde
                dtpFechaHastaIng.SelectedDate = MdtmFechaHasta.AddDays(-1)
                cnvInfDiario.Visibility = System.Windows.Visibility.Visible
            Case EnuReporteDef.enuRCFechas
                Me.Title = "Reporte Recibos de Caja entre Fechas"
                lblInforme.Content = "Informe entre el:"
                dtpFechaDesdeIng.SelectedDate = MdtmFechaDesde
                dtpFechaHastaIng.SelectedDate = MdtmFechaHasta
                cnvInfDiario.Visibility = System.Windows.Visibility.Visible
            Case EnuReporteDef.enuRelDocs
                Me.Title = "Reporte Relación de Documentos"
                lblInforme.Content = "Informe entre el:"
                dtpFechaDesdeIng.SelectedDate = MdtmFechaDesde
                dtpFechaHastaIng.SelectedDate = MdtmFechaHasta.AddDays(-1)
                cnvInfDiario.Visibility = System.Windows.Visibility.Visible
            Case EnuReporteDef.enuValoresFactTodos
                Me.Title = "Valores facturados por Servicio, Cliente y Predio Agrupador"
                lblInforme.Content = "Informe entre el:"
                dtpFechaDesdeIng.SelectedDate = MdtmFechaDesde
                dtpFechaHastaIng.SelectedDate = MdtmFechaHasta.AddDays(-1)
                cnvInfDiario.Visibility = System.Windows.Visibility.Visible
            Case EnuReporteDef.enuRecCajaReversados
                Me.Title = "Valores de Medios de Pago reversados"
                lblInforme.Content = "Informe entre el:"
                dtpFechaDesdeIng.SelectedDate = MdtmFechaDesde
                dtpFechaHastaIng.SelectedDate = MdtmFechaHasta.AddDays(-1)
                cnvCajaBancos.Visibility = System.Windows.Visibility.Visible
        End Select
    End Sub
    Private Sub SHabiliteInfDiario()
        Me.Title = "Informe Diario"
        lblInforme.Content = "Generar el informe entre el:"
        dtpFechaDesdeIng.SelectedDate = MdtmFechaDesde
        dtpFechaHastaIng.SelectedDate = MdtmFechaHasta.AddDays(-1)
        cnvInfDiario.Visibility = System.Windows.Visibility.Visible
    End Sub
    Private Sub SHabiliteCarteraClientes()
        Me.Title = "Reporte de la Cartera por Cliente"
        If EnuReporte = EnuReporteDef.enuCxCDetPorSer Then
            Me.Title = "Reporte detallado por Servicio"
            lblFechaCarHasta.Content = "Fecha de la Cartera Detallada:"
        End If
        cnvCarteraCliente.Visibility = System.Windows.Visibility.Visible
    End Sub
    Private Sub SHabiliteEdadCartera()
        Me.Title = "Reporte de la Cartera por Edad"
        Me.Height = 260
        cnvEdadCartera.Visibility = System.Windows.Visibility.Visible
        SLeaLimites()
    End Sub
    Private Sub SHabiliteItemsProgramaFact()
        Me.Title = "Reporte de los Items del Programa de Facturación"
        SPuebleCombos()
        cnvItemsProgramaFact.Visibility = System.Windows.Visibility.Visible
    End Sub
    Private Sub SHabiliteLimites()
        If MenuTipoRepEdadCart = EnuTipoRepEdadCartera.enuGrafico Then
            txtLimite1.IsEnabled = False
            txtLimite2.IsEnabled = False
            txtLimite3.IsEnabled = False
            txtLimite4.IsEnabled = False
        Else
            txtLimite1.IsEnabled = True
            txtLimite2.IsEnabled = True
            txtLimite3.IsEnabled = True
            txtLimite4.IsEnabled = True
        End If
    End Sub
    Private Sub SHabiliteResumenMovCont()
        Me.Title = "Reporte Resumen Movimiento Contable"
        MdtmFechaDesde = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo.AddMonths(-1)
        MdtmFechaHasta = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo.AddMonths(-1)
        Dim lentUltimoDiaMes = Date.DaysInMonth(MdtmFechaHasta.Year, MdtmFechaDesde.Month)
        MdtmFechaHasta = DateSerial(MdtmFechaDesde.Year, MdtmFechaHasta.Month, lentUltimoDiaMes)
        dtpFechaDesde.SelectedDate = MdtmFechaDesde
        dtpFechaHasta.SelectedDate = MdtmFechaHasta
        cnvCajaBancos.Visibility = System.Windows.Visibility.Visible
        lblFechaDesde.Content = "Generar el reporte entre el:"
    End Sub
#End Region

#Region "Procedimientos"
    Private Sub SValideFechas()
        MblnFechasOk = False
        Select Case EnuReporte
            Case EnuReporteDef.enuCajaBancos, EnuReporteDef.enuResumenMovCont
                MdtmFechaDesde = dtpFechaDesde.SelectedDate
                MdtmFechaHasta = dtpFechaHasta.SelectedDate
                If MdtmFechaHasta <= Date.Today AndAlso MdtmFechaDesde <= MdtmFechaHasta Then
                    MblnFechasOk = True
                End If
            Case EnuReporteDef.enuInformeDiario
                MdtmFechaDesde = dtpFechaDesdeIng.SelectedDate
                MdtmFechaHasta = dtpFechaHastaIng.SelectedDate
                If MdtmFechaHasta <= Date.Today AndAlso MdtmFechaDesde <= MdtmFechaHasta Then
                    MblnFechasOk = True
                End If
            Case EnuReporteDef.enuCarteraPorCliente, EnuReporteDef.enuCxCDetPorSer
                mdtmFechaCartera = dtpFechaCartera.SelectedDate
                If mdtmFechaCartera <= Date.Today Then
                    MblnFechasOk = True
                End If
            Case EnuReporteDef.enuExpRecsCaja, EnuReporteDef.enuRCFechas,
                    EnuReporteDef.enuRelDocs, EnuReporteDef.enuExpFacsFechas,
                    EnuReporteDef.enuValoresFactTodos
                MdtmFechaDesde = dtpFechaDesdeIng.SelectedDate
                MdtmFechaHasta = dtpFechaHastaIng.SelectedDate
                If MdtmFechaHasta <= Date.Today.AddDays(1) AndAlso MdtmFechaDesde <=
                        MdtmFechaHasta Then
                    MblnFechasOk = True
                End If
            Case EnuReporteDef.enuRecCajaReversados
                MdtmFechaDesde = dtpFechaDesde.SelectedDate
                MdtmFechaHasta = dtpFechaHasta.SelectedDate
                If MdtmFechaHasta <= Date.Today.AddDays(1) AndAlso MdtmFechaDesde <=
                        MdtmFechaHasta Then
                    MblnFechasOk = True
                End If
        End Select
        If Not MblnFechasOk Then
            MstrMensaje = "INFORMACION: Las Fechas no son validas."
        End If
    End Sub
    Private Sub SValideLimites()
        Dim lobjLimite As Object = txtLimite1.Text
        MblnLimitesOk = ClsPanorama.FblnEsValidoNumero(lobjLimite, 1, 360, True, EnuTipoValor.enuInteger)
        If MblnLimitesOk Then
            GentLimite1 = lobjLimite
            MentLimite1 = GentLimite1
            lblEntre2.Content = "Entre " & (GentLimite1 + 1).ToString
            lobjLimite = txtLimite2.Text
            MblnLimitesOk = ClsPanorama.FblnEsValidoNumero(lobjLimite, 1, 360, True, EnuTipoValor.enuInteger)
        End If
        If MblnLimitesOk Then
            GentLimite2 = lobjLimite
            MentLimite2 = GentLimite2
            lblEntre3.Content = "Entre " & (GentLimite2 + 1).ToString
            lobjLimite = txtLimite3.Text
            MblnLimitesOk = ClsPanorama.FblnEsValidoNumero(lobjLimite, 1, 360, True, EnuTipoValor.enuInteger)
        End If
        If MblnLimitesOk Then
            GentLimite3 = lobjLimite
            MentLimite3 = GentLimite3
            lblEntre4.Content = "Entre " & (GentLimite3 + 1).ToString
            lobjLimite = txtLimite4.Text
            MblnLimitesOk = ClsPanorama.FblnEsValidoNumero(lobjLimite, 1, 360, True, EnuTipoValor.enuInteger)
        End If
        If MblnLimitesOk Then
            GentLimite4 = lobjLimite
            MentLimite4 = GentLimite4
            txtLimite5.Text = GentLimite4.ToString
            MblnLimitesOk = (GentLimite4 > GentLimite3 AndAlso GentLimite3 > GentLimite2 AndAlso GentLimite2 > GentLimite1)
        End If
        If Not MblnLimitesOk Then
            MstrMensaje = "Todos o alguno de los limites no son validos!"
        End If
        SHabiliteLimites()
    End Sub
    Private Sub SValideServicio()
        MblnServicioOk = (MentIdServicio > 0)
        If Not MblnServicioOk Then
            MstrMensaje = "INFORMACION: El Servicio seleccionado no son validas."
        End If
    End Sub
    Private Sub SLeaLimites()
        If MenuTipoRepEdadCart = EnuTipoRepEdadCartera.enuGrafico Then
            txtLimite1.Text = 30
            txtLimite2.Text = 60
            txtLimite3.Text = 120
            txtLimite4.Text = 180
        Else
            txtLimite1.Text = MentLimite1
            txtLimite2.Text = MentLimite2
            txtLimite3.Text = MentLimite3
            txtLimite4.Text = MentLimite4
        End If
        SValideLimites()
    End Sub
    Private Sub SMuestreMensaje()
        lblMensajes.Background = System.Windows.Media.Brushes.SlateGray
        If MstrMensaje.Length > 0 Then
            lblMensajes.Background = System.Windows.Media.Brushes.DarkBlue
        End If
        lblMensajes.Content = MstrMensaje
    End Sub
    Private Sub SPuebleCombos()
        MblnPoblandoCbos = True
        With cboAnos
            .Items.Clear()
            .Items.Add(0)
            For Each lobjAno As ClsAno In GobjParametros.ColAnos
                .Items.Add(lobjAno.ObjIdAnoShr.ObjValorPro)
            Next
        End With
        With cboServicios
            .Items.Clear()
            .Items.Add((My.Resources.Ninguno))
        End With
        MblnPoblandoCbos = False
        cboAnos.SelectedIndex = 0
        SValideServicio()
    End Sub
    Private Sub SPuebleComboServicios()
        MblnPoblandoCbos = True
        Dim lcolServicios As New Collection
        If MshrIdAno = 0 Then
            lcolServicios = GobjParametros.ColServiciosPer
        Else
            Dim lobjAno As ClsAno = GobjParametros.ColAnos(MshrIdAno.ToString)
            lcolServicios = lobjAno.ColServiciosAno
        End If
        cboServicios.Items.Clear()
        cboServicios.Items.Add(My.Resources.Ninguno)
        If Not IsNothing(lcolServicios) Then
            Dim lstrServicio = String.Empty
            For Each lobjServicio As ClsServicio In lcolServicios
                lstrServicio = lobjServicio.ObjIdServicioShr.ObjValorPro & "-" &
                        lobjServicio.ObjNombreServicioStr.ObjValorPro
                cboServicios.Items.Add(lstrServicio)
            Next
        End If
        MblnPoblandoCbos = False
        cboServicios.SelectedIndex = 0
    End Sub
    Private Sub SEstablezcaServicio()
        If Not IsNothing(cboServicios.SelectedItem) Then
            Dim lstrServicio As String = cboServicios.SelectedItem
            If lstrServicio <> My.Resources.Ninguno Then
                MentIdServicio = CType(lstrServicio.Split("-")(0), Integer)
            Else
                MentIdServicio = 0
            End If
        Else
            MentIdServicio = 0
        End If
    End Sub
#End Region

#Region "Eventos de la Ventana"
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        SInvisibiliceCnvs()
        SHabiliteControles()
        dtpFechaDesde.SelectedDate = MdtmFechaDesde
        dtpFechaHasta.SelectedDate = MdtmFechaHasta
        dtpFechaCartera.SelectedDate = mdtmFechaCartera
        rdbDetallado.IsChecked = True
        MenuTipoRepEdadCart = EnuTipoRepEdadCartera.enuDetallado
        dtpFechaDesde.Focus()
    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            Mouse.OverrideCursor = Cursors.Wait
            Select Case lelmElemento.Name
                Case "bttAceptar"
                    Select Case EnuReporte
                        Case EnuReporteDef.enuCajaBancos
                            SValideFechas()
                            SGenereCajaBancos()
                        Case EnuReporteDef.enuRecCajaReversados
                            SValideFechas()
                            SGenereRCReversados()
                        Case EnuReporteDef.enuInformeDiario
                            SValideFechas()
                            SGenereInfDiario()
                        Case EnuReporteDef.enuCarteraPorCliente
                            SValideFechas()
                            SGenereCarteraCliente()
                        Case EnuReporteDef.enuEdadCartera
                            SValideLimites()
                            SGenereEdadCartera()
                        Case EnuReporteDef.enuItemsProgramaFact
                            SValideServicio()
                            SGenereItemsProgramaFact()
                        Case EnuReporteDef.enuResumenMovCont
                            SValideFechas()
                            SGenereResumenMovCont()
                        Case EnuReporteDef.enuExpRecsCaja
                            SValideFechas()
                            SExporteRecibosCaja()
                        Case EnuReporteDef.enuExpFacsFechas
                            SValideFechas()
                            SExporteFacturas()
                        Case EnuReporteDef.enuRCFechas
                            SValideFechas()
                            SGenereRCFechas()
                        Case EnuReporteDef.enuRelDocs
                            SValideFechas()
                            SGenereRelDocs()
                        Case EnuReporteDef.enuCxCDetPorSer
                            SValideFechas()
                            SGenereCxCDetPorSer()
                        Case EnuReporteDef.enuValoresFactTodos
                            SValideFechas()
                            SGenereValoresFacturados()
                    End Select
                Case "bttCancelar"
                    MblnCerrando = True
                    MstrMensaje = String.Empty
                    If EnuReporte = EnuReporteDef.enuExpRecsCaja OrElse
                            EnuReporte = EnuReporteDef.enuExpFacsFechas Then
                        GblnOK = False
                    End If
            End Select
            If String.IsNullOrEmpty(MstrMensaje) Then
                Close()
            End If
            Mouse.OverrideCursor = Cursors.Arrow
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
        If Not MblnCerrando Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is DatePicker OrElse TypeOf lelmElemento Is TextBox Then
                Select Case lelmElemento.Name
                    Case "dtpFechaDesde", "dtpFechaHasta", "dtpFechaCartera"
                        SValideFechas()
                    Case "txtLimite1", "txtLimite2", "txtLimite3", "txtLimite4"
                        SValideLimites()
                End Select
            End If
        End If
    End Sub

    Private Sub OnRdbClick(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is RadioButton Then
            Select Case lelmElemento.Name
                Case "rdbDetallado"
                    MenuTipoRepEdadCart = EnuTipoRepEdadCartera.enuDetallado
                Case "rdbResumido"
                    MenuTipoRepEdadCart = EnuTipoRepEdadCartera.enuResumido
                Case "rdbGrafico"
                    MenuTipoRepEdadCart = EnuTipoRepEdadCartera.enuGrafico
            End Select
            SLeaLimites()
        End If
    End Sub

    Private Sub Cbo_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cboAnos.SelectionChanged,
            cboServicios.SelectionChanged
        If Not MblnPoblandoCbos Then
            Dim lcboCombo As ComboBox = sender
            Select Case lcboCombo.Name
                Case "cboAnos"
                    MshrIdAno = cboAnos.SelectedItem
                    SPuebleComboServicios()
                Case "cboServicios"
                    SEstablezcaServicio()
            End Select
        End If
        SValideServicio()
    End Sub
#End Region
End Class
