Imports System.Text
Friend Class ClsRepOrionCop
    Implements IDisposable
#Region "Definiciones"
#Region "Enumeradores"
    Private Enum EnuTipoArchivoDef As Integer
        None = 0
        EnuPdf
        EnuExcelx
        EnuCrystalRepo
    End Enum
#End Region
    Private ReadOnly MstrCampoBdCarpeta As String = PanL.ClsIdCarpetaShr.SstrNombreCampoBd
    Private ReadOnly MstrCampoBdCentroUtil As String = PanL.ClsIdCentroUtilShr.SstrNombreCampoBd
    Private ReadOnly MstbExpresionSql As New StringBuilder
    Private MbytLogoCenUtilidad As Byte() = Nothing
    Private MbytCodigoQR As Byte() = Nothing
    Private MdsAuxiliarCon As DataSet = Nothing
    Private MdtbAuxiliarCon As DataTable = Nothing
    Private MdtbResMovCont As DataTable = Nothing
    Private MblnDisposed As Boolean = False
    '
    Public Event EvnInicioExportacion As EventHandler(Of ClsPanEventArgs)
    Public Event EvnInicio As EventHandler(Of ClsPanEventArgs)
    Public Event EvnAvance As EventHandler(Of ClsPanEventArgs)
    Private MobjArgumentoEventoPan As ClsPanEventArgs = Nothing
#End Region
#Region "Constructores"
    Public Sub New(aobjRegistro As Object)
        If aobjRegistro Is Nothing OrElse Not (aobjRegistro.GetType.Name = "String") Then
            Throw New ModuloNoRegistradoPanException()
        ElseIf Not (aobjRegistro = GCOBJREGISTRO) Then
            Throw New ModuloNoRegistradoPanException()
        End If
        MstbExpresionSql.Capacity = 500
    End Sub
#End Region
#Region "Propiedades"
    Friend Property EnuReporte As EnuReporteDef = EnuReporteDef.None
    Friend Property ObjParRepDocs As ClsParametrosReportesDocs = Nothing
    ' Propiedades Reporte Auxilir Contable
    Friend Property StrIdCuentaContIni As String = String.Empty
    Friend Property StrIdCuentaContFin As String = String.Empty
    Friend Property BlnCalculaSaldo As Boolean = False
    Friend Property DblIdCliente As Double = 0.0
    Friend Property StrIdPredioAgru As String = String.Empty
    Friend Property StrFechaDesde As String = ClsPanoramaDat.FstrFechaNormalizada(GCDTMFECHANULA)
    Friend Property StrFechaHasta As String = ClsPanoramaDat.FstrFechaNormalizada(GCDTMFECHANULA)
    Friend Property EntIdServicio As Integer = 0
    Friend Property ShrIdAno As Short = 0
    Friend Property DecIntPortCausar As Decimal = 0
    Friend Property BlnEstadoDeCuenta As Boolean = False
    ' Propiedas Reporte Movimiento Cuenta
    Friend Property DsMovimiento As DataSet = Nothing
    Friend ReadOnly Property ObjArgumentoEventoPan As ClsPanEventArgs
        Get
            If IsNothing(MobjArgumentoEventoPan) Then
                MobjArgumentoEventoPan = New ClsPanEventArgs With {
                    .BlnCancele = False,
                    .BlnVaciandoObjeto = False
                }
            End If
            Return MobjArgumentoEventoPan
        End Get
    End Property
    Friend Shared Function FfrmRep() As FrmReportes
        Dim lfrmRep As New FrmReportes
        Return lfrmRep
    End Function
    Friend Sub SVacie()
        StrIdCuentaContIni = String.Empty
        StrIdCuentaContFin = String.Empty
        BlnCalculaSaldo = False
        DblIdCliente = 0
        StrIdPredioAgru = String.Empty
        StrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(GCDTMFECHANULA)
        StrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(GCDTMFECHANULA)
        EntIdServicio = 0
        ShrIdAno = 0
        MblnDisposed = False
        Dispose()
    End Sub
#End Region
#Region "Genera Reportes"
    Friend Sub SGenereReporteDialog()
        Try
            Dim lwinParRep As WinParametrosRep = Nothing
            SRefresqueLogo()
            SRefresqueCodigoQR()
            Select Case EnuReporte
                Case EnuReporteDef.enuFacturaEFac, EnuReporteDef.enuFactura,
                        EnuReporteDef.enuFacturaDscto
                    SGenereReporteFacturas(True, ObjParRepDocs.BlnExcluirFacEnvEmail)
                Case EnuReporteDef.enuCtaCobroDet
                    SGenereReporteCtaCobro(True, True)
                Case EnuReporteDef.enuNotasCon
                    SGenereReporteNotasCon(True)
                Case EnuReporteDef.enuNotasDb
                    SGenereReporteNotasDb(True)
            End Select
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        End Try
    End Sub
    Friend Sub SGenereReporte()
        Try
            Dim lwinParRep As WinParametrosRep = Nothing
            SRefresqueLogo()
            If Not FblnGeneroReporte() Then
                Select Case EnuReporte
                    Case EnuReporteDef.enuResumenMovCont
                        lwinParRep = New WinParametrosRep With {
                            .EnuReporte = EnuReporteDef.enuResumenMovCont,
                            .ObjRepOrionCop = Me
                        }
                        lwinParRep.Show()
                    Case EnuReporteDef.enuCajaBancos
                        lwinParRep = New WinParametrosRep With {
                            .EnuReporte = EnuReporteDef.enuCajaBancos,
                            .ObjRepOrionCop = Me
                        }
                        lwinParRep.Show()
                    Case EnuReporteDef.enuRCFechas
                        lwinParRep = New WinParametrosRep With {
                            .EnuReporte = EnuReporteDef.enuRCFechas,
                            .ObjRepOrionCop = Me
                        }
                        lwinParRep.Show()
                    Case EnuReporteDef.enuRecCajaReversados
                        lwinParRep = New WinParametrosRep With {
                            .EnuReporte = EnuReporteDef.enuRecCajaReversados,
                            .ObjRepOrionCop = Me
                        }
                        lwinParRep.Show()
                    Case EnuReporteDef.enuValoresFactTodos
                        lwinParRep = New WinParametrosRep With {
                            .EnuReporte = EnuReporteDef.enuValoresFactTodos,
                            .ObjRepOrionCop = Me
                        }
                        lwinParRep.Show()
                    Case EnuReporteDef.enuExpRecsCaja
                        lwinParRep = New WinParametrosRep With {
                            .ObjRepOrionCop = Me,
                            .EnuReporte = EnuReporteDef.enuExpRecsCaja
                        }
                        lwinParRep.Show()
                    Case EnuReporteDef.enuExpFacsFechas
                        lwinParRep = New WinParametrosRep With {
                            .ObjRepOrionCop = Me,
                            .EnuReporte = EnuReporteDef.enuExpFacsFechas
                        }
                        lwinParRep.ShowDialog()
                    Case EnuReporteDef.enuInformeDiario
                        lwinParRep = New WinParametrosRep With {
                            .EnuReporte = EnuReporteDef.enuInformeDiario,
                            .ObjRepOrionCop = Me
                        }
                        lwinParRep.Show()
                    Case EnuReporteDef.enuCarteraPorCliente
                        lwinParRep = New WinParametrosRep With {
                            .EnuReporte = EnuReporteDef.enuCarteraPorCliente,
                            .ObjRepOrionCop = Me
                        }
                        lwinParRep.Show()
                    Case EnuReporteDef.enuEdadCartera
                        lwinParRep = New WinParametrosRep With {
                            .EnuReporte = EnuReporteDef.enuEdadCartera,
                            .ObjRepOrionCop = Me
                        }
                        lwinParRep.Show()
                    Case EnuReporteDef.enuCarteraPorPredioAgr
                        SGenereCarteraPorPredioAgr()
                    Case EnuReporteDef.enuCarteraPorPredio
                        SGenereCarteraPorPredio()
                    Case EnuReporteDef.enuCxCDetPorSer
                        lwinParRep = New WinParametrosRep With {
                            .EnuReporte = EnuReporteDef.enuCxCDetPorSer,
                            .ObjRepOrionCop = Me
                        }
                        lwinParRep.Show()
                    Case EnuReporteDef.enuCarteraPorServicio
                        SGenereCarteraPorServicio()
                    Case EnuReporteDef.enuEstadoCuentas
                        SGenereEstadoCuentas()
                    Case EnuReporteDef.enuEstadoCtaCli
                        SGenereEstadoCtaCli()
                    Case EnuReporteDef.enuPrediosSector
                        SGenerePrediosSector()
                    Case EnuReporteDef.enuPrediosPropietario
                        SGenerePrediosPropietario()
                    Case EnuReporteDef.enuCuotasAdminPropi
                        SGenereCuotasAdminPropietario()
                    Case EnuReporteDef.enuItemsProgramaFact
                        lwinParRep = New WinParametrosRep With {
                            .EnuReporte = EnuReporteDef.enuItemsProgramaFact,
                            .ObjRepOrionCop = Me
                        }
                        lwinParRep.Show()
                    Case EnuReporteDef.enuRelDocs
                        lwinParRep = New WinParametrosRep With {
                            .EnuReporte = EnuReporteDef.enuRelDocs,
                            .ObjRepOrionCop = Me
                        }
                        lwinParRep.Show()
                    Case Else
                        Throw New ErrorInesperadoPanLException("Reporte no determinado")
                End Select
            End If
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        End Try
    End Sub
    Private Function FblnGeneroReporte() As Boolean
        Dim lblnGenero = FblnGeneroReporte2()
        If Not lblnGenero Then
            lblnGenero = True
            Select Case EnuReporte
                Case EnuReporteDef.enuNotaDevAnt
                    SGenereReporteNotasDevAnt()
                Case EnuReporteDef.enuNotaReverCr
                    SGenereReporteNotaRevCr()
                Case EnuReporteDef.enuCtaCobro
                    SGenereReporteCtaCobro(False, False)
                Case EnuReporteDef.enuCtaCobroDet
                    SGenereReporteCtaCobro(True, False)
                Case EnuReporteDef.enuAuxiliar
                    SGenereAuxiliar()
                Case EnuReporteDef.enuAuxTer
                    SGenereAuxTer()
                Case EnuReporteDef.enuMovimCuenta
                    SGenereMovCuenta()
                Case EnuReporteDef.enuMovimAntici
                    SGenereMovAnt()
                Case EnuReporteDef.enuDirTf
                    SGenereDirTf()
                Case EnuReporteDef.enuDirClientes
                    SGenereDirClientes()
                Case EnuReporteDef.enuAnticiposPorAplicar
                    SGenereAntPorAplicar()
                Case EnuReporteDef.enuDocsNoRegEFac
                    SGenereDocsNoRegEFac()
                Case EnuReporteDef.enuValoresFacturados
                    SGenereRepVlrsFacturados()
                Case Else
                    lblnGenero = False
            End Select
        End If
        Return lblnGenero
    End Function
    Private Function FblnGeneroReporte2() As Boolean
        Dim lblnGenero = True
        Select Case EnuReporte
            Case EnuReporteDef.enuFactura, EnuReporteDef.enuFacturaEFac,
                    EnuReporteDef.enuFacturaDscto, EnuReporteDef.enuFactImportada
                SGenereReporteFacturas(False, False)
            Case EnuReporteDef.enuRecCaja
                SGenereReporteRecCaja()
            Case EnuReporteDef.enuNotasAjuste
                SGenereReporteNotasAjuste()
            Case EnuReporteDef.enuNotasCon
                SGenereReporteNotasCon(False)
            Case EnuReporteDef.enuNotasDb
                SGenereReporteNotasDb(False)
            Case EnuReporteDef.enuNotaCr
                SGenereReporteNotasCr()
            Case EnuReporteDef.enuFacVivas
                SGenereReporteFacVIvas()
            Case EnuReporteDef.enuFactAutoMes
                SGenereRepFacsAutoMes()
            Case EnuReporteDef.enuPropietariosXCP_Res
                SGenerePropietariosXCP(True)
            Case EnuReporteDef.enuPropietariosXCP
                SGenerePropietariosXCP(False)
            Case Else
                lblnGenero = False
        End Select
        Return lblnGenero
    End Function
    Private Sub SGenereReporteFacturas(ablnDialog As Boolean, ablnExcluirEnvMail As Boolean)
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        SRefresqueLogo()
        SRefresqueCodigoQR()
        With lfrmRep
            If EnuReporte = EnuReporteDef.enuFacturaEFac Then
                .RcReporte = New repFacturaEFac
            Else
                If EnuReporte = EnuReporteDef.enuFactura Then
                    .RcReporte = New repFacturaLogo
                ElseIf EnuReporte = EnuReporteDef.enuFactImportada Then
                    .RcReporte = New repFactImportadaLogo
                Else
                    .RcReporte = New repFacturaDsctoPPLogo
                End If
            End If
            Using ldsFacturas As New DataSet
                SGenereDataSetFacturas(ldsFacturas, ablnExcluirEnvMail)
                .RcReporte.SetDataSource(ldsFacturas)
                If ablnDialog Then
                    .ShowDialog()
                Else
                    .Show()
                End If
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereRepFacsAutoMes()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        SRefresqueLogo()
        SRefresqueCodigoQR()
        With lfrmRep
            If GobjParametros.BlnEFacAutorizado Then
                .RcReporte = New repFacturaEFac
            Else
                If GobjParametros.ObjAnoActual.ObjAplicaDsctoPPBln.ObjValorPro Then
                    .RcReporte = New repFacturaDsctoPPLogo
                Else
                    .RcReporte = New repFacturaLogo
                End If
            End If
            Using ldsFacturas As New DataSet
                SGenereDataSetFacsAutoMes(ldsFacturas)
                .RcReporte.SetDataSource(ldsFacturas)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereReporteRecCaja()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repRecCajaConLogo
            Using ldsRecCaja As New DataSet
                SGenereDataSetRecCaja(ldsRecCaja, False)
                .RcReporte.SetDataSource(ldsRecCaja)
                .Show()
                .Activate()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereReporteNotasAjuste()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repNotasAjuste
            Using ldsNotasAjuste As New DataSet
                SGenereDataSetNotasAjuste(ldsNotasAjuste)
                .RcReporte.SetDataSource(ldsNotasAjuste)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereReporteNotasCon(ablnDialog As Boolean)
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repNotasConLogo
            Using ldsNotasCon As New DataSet
                SGenereDataSetNotasCon(ldsNotasCon)
                .RcReporte.SetDataSource(ldsNotasCon)
                If ablnDialog Then
                    .ShowDialog()
                Else
                    .Show()
                End If
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereReporteNotasDb(ablnDialog As Boolean)
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repNotasDbLogo
            Using ldsNotasDb As New DataSet
                SGenereDataSetNotasDb(ldsNotasDb)
                .RcReporte.SetDataSource(ldsNotasDb)
                If ablnDialog Then
                    .ShowDialog()
                Else
                    .Show()
                End If
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereReporteNotasCr()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repNotasCrLogo
            Using ldsNotasCr As New DataSet
                SGenereDataSetNotasCr(ldsNotasCr)
                .RcReporte.SetDataSource(ldsNotasCr)
                .Show()
                .Activate()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereReporteNotasDevAnt()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repNotasDevAntLogo
            Using ldsNotasDevAnt As New DataSet
                SGenereDataSetNotasDevAnt(ldsNotasDevAnt)
                .RcReporte.SetDataSource(ldsNotasDevAnt)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereReporteNotaRevCr()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repNotasRevCrLogo
            Using ldsNotasDevAnt As New DataSet
                SGenereDataSetNotasRevCr(ldsNotasDevAnt)
                .RcReporte.SetDataSource(ldsNotasDevAnt)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereReporteCtaCobro(ablnDetallado As Boolean, ablnDialog As Boolean)
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            If ablnDetallado Then
                .RcReporte = New repCtaCobroDetLogo
            Else
                .RcReporte = New repCtaCobroLogo
            End If
            Using ldsCtasCobro As New DataSet
                SGenereDataSetCuentasCobro(ldsCtasCobro, ablnDetallado)
                .RcReporte.SetDataSource(ldsCtasCobro)
                If ablnDialog Then
                    .ShowDialog()
                Else
                    .Show()
                End If
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereAuxiliar()
        GobjPanDat.SControleProcesoObj(True)
        SGenereDataSetAux()
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repAuxContLogo
            .RcReporte.SetDataSource(MdsAuxiliarCon)
            .Show()
            MdsAuxiliarCon.Dispose()
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereAuxTer()
        GobjPanDat.SControleProcesoObj(True)
        SGenereDataSetAux()
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repAuxTerLogo
            .RcReporte.SetDataSource(MdsAuxiliarCon)
            .Show()
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Friend Sub SGenereRelDocs()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repRelDocs
            Using ldsRelDocs As New DataSet
                SGenereDataSetRelDocs(ldsRelDocs)
                .RcReporte.SetDataSource(ldsRelDocs)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Friend Function SGenereResumenMovCont() As String
        Dim lstrMens = String.Empty
        GobjPanDat.SControleProcesoObj(True)
        BlnCalculaSaldo = True
        StrIdCuentaContIni = GobjPanorama.ObjCarpetaActual.StrIdCtaPrimera
        StrIdCuentaContFin = GobjPanorama.ObjCarpetaActual.StrIdCtaUltima
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repResumenMovCont
            Using ldsResumenMovCon As New DataSet
                SGenereDataSetResumenMovCon(ldsResumenMovCon)
                Dim ldtbResMovCont = ldsResumenMovCon.Tables("OriResumenMovCon")
                If ldtbResMovCont.Rows.Count = 0 Then
                    lstrMens = "No hay Movimiento Contable en el presente Intervalo de Tiempo!"
                Else
                    .RcReporte.SetDataSource(ldsResumenMovCon)
                    .Show()
                End If
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
        Return lstrMens
    End Function
    Private Sub SGenerePrediosSector()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repPrediosSectorLogo
            Using ldsPrediosSector As New DataSet
                SGenereDataSetPrediosSector(ldsPrediosSector)
                .RcReporte.SetDataSource(ldsPrediosSector)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)

    End Sub
    Private Sub SGenerePrediosPropietario()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repPrediosPropietarioLogo
            Using ldsPrediosPropietario As New DataSet
                SGenereDataSetPrediosPropietario(ldsPrediosPropietario)
                .RcReporte.SetDataSource(ldsPrediosPropietario)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenerePropietariosXCP(ablnResumido As Boolean)
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            If ablnResumido Then
                .RcReporte = New repPropietariosXCP_res
            Else
                .RcReporte = New repPropietariosXCP
            End If
            Using ldsPropietarioXCP As New DataSet
                SGenereDataSetPropietariosXCP(ldsPropietarioXCP, ablnResumido)
                .RcReporte.SetDataSource(ldsPropietarioXCP)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereCuotasAdminPropietario()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repCuotaAdminPropietarioLogo
            Using ldsCuotasAdminPropietario As New DataSet
                SGenereDataSetCuotasPropietario(ldsCuotasAdminPropietario)
                .RcReporte.SetDataSource(ldsCuotasAdminPropietario)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)

    End Sub
    Friend Function SGenereCajaBancos() As String
        Dim lstrMens = String.Empty
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repCajaBancosLogo
            Using ldsCajaBancos As New DataSet
                SGenereDataSetMediosPago(ldsCajaBancos)
                Dim ldtbMediosPago = ldsCajaBancos.Tables("OriMediosPago")
                If ldtbMediosPago.Rows.Count = 0 Then
                    lstrMens = "No hay Recibos de Caja en el presente Intervalo de Tiempo!"
                Else
                    .RcReporte.SetDataSource(ldsCajaBancos)
                    .Show()
                End If
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
        Return lstrMens
    End Function
    Friend Function SGenereRCFechas() As String
        Dim lstrMens = String.Empty
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repRCFechas
            Using ldsRCFechas As New DataSet
                SGenereDataSetRCFechas(ldsRCFechas)
                Dim ldtbRCFechas = ldsRCFechas.Tables("OriRecibosCaja")
                If ldtbRCFechas.Rows.Count = 0 Then
                    lstrMens = "No hay Recibos de Caja en el presente Intervalo de Tiempo!"
                Else
                    .RcReporte.SetDataSource(ldsRCFechas)
                    .Show()
                End If
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
        Return lstrMens
    End Function
    Friend Function SGenereRCReversados() As String
        Dim lstrMens = String.Empty
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repRCReversados
            Using ldsRCReversados As New DataSet
                SGenereDataSetRecCajaRev(ldsRCReversados)
                Dim ldtbMediosPago = ldsRCReversados.Tables("OriMediosPago")
                If ldtbMediosPago.Rows.Count = 0 Then
                    lstrMens = "No hay Recibos de Caja Reversados en el presente Intervalo de Tiempo!"
                Else
                    .RcReporte.SetDataSource(ldsRCReversados)
                    .Show()
                End If
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
        Return lstrMens
    End Function
    Friend Function SGenereInformeDiario() As String
        Dim lstrMens = String.Empty
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repInfDiario
            Using ldsInfDiario As New DataSet
                SGenereDataSetInfDiario(ldsInfDiario)
                Dim ldtbFrasDia = ldsInfDiario.Tables("OriIdFacturasDias")
                If ldtbFrasDia.Rows.Count = 0 Then
                    lstrMens = "No se han generado Facturas en el presente Intervalo de Tiempo!"
                Else
                    .RcReporte.SetDataSource(ldsInfDiario)
                    .Show()
                End If
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
        Return lstrMens
    End Function
    Friend Function SGenereCarteraPorCliente() As String
        Dim lstrMens = String.Empty
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repCarteraClientesLogo
            Using ldsCarteraPorCliente As New DataSet
                SGenereDataSetCarPorCli(ldsCarteraPorCliente)
                Dim ldtbCarCli = ldsCarteraPorCliente.Tables("OriCarteraPorCliente")
                If ldtbCarCli.Rows.Count = 0 Then
                    lstrMens = "No hay información en la fecha señalada!"
                Else
                    .RcReporte.SetDataSource(ldsCarteraPorCliente)
                    .Show()
                End If
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
        Return lstrMens
    End Function
    Friend Function SGenereCxCDetPorSer() As String
        Dim lstrMens = String.Empty
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repCxCDetPorSer
            Using ldsCxCDetPorSer As New DataSet
                SGenereDataSetCxCDetPorServicio(ldsCxCDetPorSer)
                Dim ldtbCxCDet = ldsCxCDetPorSer.Tables("OriCarteraXServicios")
                If ldtbCxCDet.Rows.Count = 0 Then
                    lstrMens = "No hay información en la fecha señalada!"
                Else
                    .RcReporte.SetDataSource(ldsCxCDetPorSer)
                    .Show()
                End If
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
        Return lstrMens
    End Function
    Private Sub SGenereCarteraPorPredioAgr()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repCarteraPredioAgrLogo
            Using ldsCarteraPorPredioAgr As New DataSet
                SGenereDataSetCarPorPredioAgr(ldsCarteraPorPredioAgr)
                .RcReporte.SetDataSource(ldsCarteraPorPredioAgr)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereCarteraPorPredio()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repCarteraPrediosLogo
            Using ldsCarteraPorPredio As New DataSet
                SGenereDataSetCarPorPredio(ldsCarteraPorPredio)
                .RcReporte.SetDataSource(ldsCarteraPorPredio)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereCarteraPorServicio()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repCarteraServicios
            Using ldsCarteraPorServicio As New DataSet
                SGenereDataSetCarPorServicio(ldsCarteraPorServicio)
                .RcReporte.SetDataSource(ldsCarteraPorServicio)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Friend Function SGenereEdadCartera(aentLimite1 As Integer, aentLimite2 As Integer,
            aentLimite3 As Integer, aentLimite4 As Integer,
            aenuTipoRepEdadCart As EnuTipoRepEdadCartera, adtmFechaDatos As Date,
            ablnCierreMes As Boolean, ablnDialog As Boolean) As String
        Dim lstrMens = String.Empty
        Dim lblnDetallado = (aenuTipoRepEdadCart = EnuTipoRepEdadCartera.enuDetallado)
        GobjPanDat.SControleProcesoObj(True)
        Dim lstrNombresColumnas As String() = Array.Empty(Of String)()
        ObjArgumentoEventoPan.BlnCancele = False
        ObjArgumentoEventoPan.DblCantAProcesar = 1.0
        ObjArgumentoEventoPan.DblCantProcesada = 0.0
        ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.EnuRepEdadCar
        RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            Select Case aenuTipoRepEdadCart
                Case EnuTipoRepEdadCartera.None
                    Throw New ErrorInesperadoPanLException("Tipo de Reporte no esperado!")
                Case EnuTipoRepEdadCartera.enuDetallado
                    .RcReporte = New repEdadCarteraDetLogo
                Case EnuTipoRepEdadCartera.enuResumido
                    .RcReporte = New repEdadCarteraLogo
                Case EnuTipoRepEdadCartera.enuGrafico
                    .RcReporte = New repEdadCarGrafico
            End Select
            Using ldsEdadCartera As New DataSet
                RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                If aenuTipoRepEdadCart = EnuTipoRepEdadCartera.enuGrafico Then
                    SGenereDataSetEdadCartera(ldsEdadCartera)
                    lstrNombresColumnas = FstrNombresColumnas()
                Else
                    SGenereDataSetEdadCartera(ldsEdadCartera, aentLimite1, aentLimite2,
                                    aentLimite3, aentLimite4, lblnDetallado, adtmFechaDatos)
                    lstrNombresColumnas = FstrNombresColumnas(lblnDetallado, aentLimite1,
                                    aentLimite2, aentLimite3, aentLimite4)
                End If
                SComplementeDsEdadCartera(ldsEdadCartera, adtmFechaDatos)
                Dim ldtbEdadCar = ldsEdadCartera.Tables("OriEdadCartera")
                If ldtbEdadCar.Rows.Count = 0 Then
                    lstrMens = "No hay información a la fecha de hoy!"
                Else
                    .RcReporte.SetDataSource(ldsEdadCartera)
                    SExporteEdadCartera(.RcReporte, ablnCierreMes, ldtbEdadCar,
                            lstrNombresColumnas, adtmFechaDatos, lblnDetallado)
                    If ablnDialog Then
                        .ShowDialog()
                    Else
                        .Show()
                    End If
                End If
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
        Return lstrMens
    End Function
    Private Sub SGenereEstadoCuentas()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repEstadoDeudasLogo
            Using ldsEstadoCuentas As New DataSet
                SGenereDataSetEstadoCuenta(ldsEstadoCuentas)
                .RcReporte.SetDataSource(ldsEstadoCuentas)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereEstadoCtaCli()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repEstadoCtaCli
            Using ldsEstadoCtaCli As New DataSet
                SGenereDataSetEstadoCtaCli(ldsEstadoCtaCli)
                .RcReporte.SetDataSource(ldsEstadoCtaCli)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereMovCuenta()
        Dim lfrmRep = FfrmRep()
        GobjPanDat.SControleProcesoObj(True)
        With lfrmRep
            .RcReporte = New repMoviCtaLogo
            SGenereDataSetMovimiento(DsMovimiento)
            .RcReporte.SetDataSource(DsMovimiento)
            .Show()
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereMovAnt()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repMoviAntLogo
            SGenereDataSetMovimiento(DsMovimiento)
            .RcReporte.SetDataSource(DsMovimiento)
            .Show()
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Friend Function SGenereItemsProgramaFact() As String
        Dim lstrMens = String.Empty
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repItemsProgFactLogo
            Using ldsItemsProgFact As New DataSet
                SGenereDataSetItemsProgFact(ldsItemsProgFact)
                Dim ldtbItemsProgFact = ldsItemsProgFact.Tables("OriItemsProgFact")
                If ldtbItemsProgFact.Rows.Count = 0 Then
                    lstrMens = "No hay información en la fecha señalada!"
                Else
                    .RcReporte.SetDataSource(ldsItemsProgFact)
                    .Show()
                End If
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
        Return lstrMens
    End Function
    Private Sub SGenereDirTf()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repDirTfLogo
            Using ldsDirTf As New DataSet
                SGenereDataSetDirTf(ldsDirTf)
                .RcReporte.SetDataSource(ldsDirTf)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereDirClientes()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repDirClientes
            Using ldsDirClientes As New DataSet
                SGenereDataSetDirClientes(ldsDirClientes)
                .RcReporte.SetDataSource(ldsDirClientes)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereAntPorAplicar()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repAntPorAplicar
            Using ldsAntPorAplicar As New DataSet
                SGenereAntPorAplicar(ldsAntPorAplicar, Now)
                .RcReporte.SetDataSource(ldsAntPorAplicar)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Friend Sub SGenereAntPorAplicar(adtmFechaDatos As Date, ablnDialog As Boolean)
        GobjPanDat.SControleProcesoObj(True)
        ObjArgumentoEventoPan.BlnCancele = False
        ObjArgumentoEventoPan.DblCantAProcesar = 1.0
        ObjArgumentoEventoPan.DblCantProcesada = 0.0
        ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.EnuRepAntXApl
        RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
        Dim lfrmRep = FfrmRep()
        Dim lstrNombresColumnas As String() = {"Nro.", "Fecha", "PredioAgru.", "Id. Cliente",
                "NombreCliente", "Servicios", "Valor Anticipo", "Por Aplicar"}
        With lfrmRep
            .RcReporte = New repAntPorAplicar
            Using ldsAntPorAplicar As New DataSet
                SGenereAntPorAplicar(ldsAntPorAplicar, adtmFechaDatos)
                .RcReporte.SetDataSource(ldsAntPorAplicar)
                Dim ldtbAntPorApli = ldsAntPorAplicar.Tables("OriAntPorAplicar")
                If ldtbAntPorApli.Rows.Count > 0 Then
                    .RcReporte.SetDataSource(ldsAntPorAplicar)
                    SExporteAntPorAplicar(.RcReporte, ldtbAntPorApli, lstrNombresColumnas,
                                adtmFechaDatos)
                    If ablnDialog Then
                        .ShowDialog()
                    Else
                        .Show()
                    End If
                End If
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereDocsNoRegEFac()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New RepDocsNoRegEFac
            Using ldsDocsNoRegEFac As New DataSet
                SGenereDataSetEFacNoReg(ldsDocsNoRegEFac)
                .RcReporte.SetDataSource(ldsDocsNoRegEFac)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereReporteFacVIvas()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repFactConMora
            Using ldsFactConMora As New DataSet
                SGenereFacturasVivas(ldsFactConMora)
                .RcReporte.SetDataSource(ldsFactConMora)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SGenereRepVlrsFacturados()
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repVlrsFacturados
            Using ldsVlrsFacturados As New DataSet
                SGenereDataSetValFact(ldsVlrsFacturados)
                .RcReporte.SetDataSource(ldsVlrsFacturados)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Friend Function SGenereRepVlrsFacturadosTodos() As String
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep(), lstrMens = String.Empty
        With lfrmRep
            .RcReporte = New repVlrsFacturadosTodos
            Using ldsVlrsFacturados As New DataSet
                SGenereDataSetValFactTodos(ldsVlrsFacturados)
                Dim ldtbRes = ldsVlrsFacturados.Tables("OriVlrsFact")
                If ldtbRes.Rows.Count = 0 Then
                    lstrMens = "No hay Facturas en el presente Intervalo de Tiempo!"
                Else
                    .RcReporte.SetDataSource(ldsVlrsFacturados)
                End If
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
        Return lstrMens
    End Function
    Friend Sub SGenerePazYSalvo(aobjPredio As ClsPredio)
        GobjPanDat.SControleProcesoObj(True)
        Dim lfrmRep = FfrmRep()
        With lfrmRep
            .RcReporte = New repPazYSalvo
            Using ldsPazYSaldo As New DataSet
                SGenereDataSetPazYSalvo(ldsPazYSaldo, aobjPredio)
                .RcReporte.SetDataSource(ldsPazYSaldo)
                .Show()
            End Using
        End With
        GobjPanDat.SControleProcesoObj(False)
    End Sub
#End Region
#Region "Exportar reportes"
    Friend Sub SExporteFacsMes(adtbFactAExpor As DataTable)
        Dim lstrPref = String.Empty, lentIdFac As Integer, i = 0
        Dim lobjValorLlave As Object()
        Dim lobjFac As New ClsFactura()
        SRefresqueLogo()
        SRefresqueCodigoQR()
        Dim ldtbFrasMes = ClsOrionCop.FdtbFacsMesExportar()
        Dim ldblCanAProcesar = ldtbFrasMes.Rows.Count
        Dim lrcFactura As ReportClass = Nothing
        ObjParRepDocs = New ClsParametrosReportesDocs(String.Empty, 0, 0)
        If ldblCanAProcesar > 0 Then
            If GobjParametros.BlnEFacAutorizado Then
                lrcFactura = New repFacturaEFac
            Else
                If GobjParametros.ObjAnoActual.ObjAplicaDsctoPPBln.ObjValorPro Then
                    lrcFactura = New repFacturaDsctoPPLogo
                Else
                    lrcFactura = New repFacturaLogo
                End If
            End If
            ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.EnuExpFras
            ObjArgumentoEventoPan.DblCantAProcesar = ldblCanAProcesar
            RaiseEvent EvnInicioExportacion(Me, ObjArgumentoEventoPan)
        Else
            Exit Sub
        End If
        For Each ldrwFac As DataRow In ldtbFrasMes.Rows
            i += 1
            ObjArgumentoEventoPan.BlnCancele = False
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            If ObjArgumentoEventoPan.BlnCancele Then
                Exit For
            End If
            lstrPref = ClsPanorama.FobjValorCampo(ldrwFac(ClsPrefijo_FactStr.SstrNombreCampoBd),
                    EnuTipoValor.EnuString)
            lentIdFac = ClsPanorama.FobjValorCampo(ldrwFac(ClsIdFacturaEnt.SstrNombreCampoBd),
                    EnuTipoValor.EnuInteger)
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFac}
            lobjFac.SAbra(lobjValorLlave)
            With ObjParRepDocs
                .StrPrefijoDocsRep = lstrPref
                .EntIdDocInicial = lentIdFac
                .EntIdDocFinal = lentIdFac
            End With
            SExporteFactura(lstrPref, lentIdFac, False, lrcFactura)
        Next
        If lrcFactura IsNot Nothing Then
            lrcFactura.Close()
            lrcFactura.Dispose()
        End If
    End Sub
    ''' <summary>
    ''' Exporta a un archivo PDF una sola factura independiente de la ubicacion actual
    ''' </summary>
    ''' <param name="aobjFactura">Factura a ser exportada</param>
    ''' ''' <param name="ablnEnvioEmail">Indica si es para ser enviada por email</param>
    ''' <remarks></remarks>
    Friend Sub SExporteFactura(aobjFactura As ClsFactura, aobjTerCentroUtil As ClsTercero,
                                ablnEnvioEmail As Boolean)
        Try
            SRefresqueLogo(aobjTerCentroUtil)
            SRefresqueCodigoQR()
            If Not aobjFactura.BlnExiste Then
                Throw New ErrorInesperadoPanLException("Factura a exportar no existe")
            End If
            Dim lstrPrefFac As String = aobjFactura.ObjPrefijo_FactStr.ObjValorPro
            Dim lentIdFac As Integer = aobjFactura.ObjIdFacturaEnt.ObjValorPro
            ObjParRepDocs = New ClsParametrosReportesDocs(lstrPrefFac, 0, 0)
            Dim lrcFactura As ReportClass = Nothing
            If GobjParametros.BlnEFacAutorizado Then
                lrcFactura = New repFacturaEFac
            Else
                If aobjFactura.ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuManual Then
                    lrcFactura = New repFactImportadaLogo
                ElseIf GobjParametros.ObjAnoActual.ObjAplicaDsctoPPBln.ObjValorPro Then
                    lrcFactura = New repFacturaDsctoPPLogo
                Else
                    lrcFactura = New repFacturaLogo
                End If
            End If
            If Not ablnEnvioEmail Then
                ObjArgumentoEventoPan.BlnCancele = False
                ObjArgumentoEventoPan.DblCantAProcesar = 1
                RaiseEvent EvnInicioExportacion(Me, ObjArgumentoEventoPan)
                ObjArgumentoEventoPan.BlnCancele = False
                ObjArgumentoEventoPan.DblCantProcesada = 1
                RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                If ObjArgumentoEventoPan.BlnCancele Then
                    Exit Sub
                End If
            End If
            With ObjParRepDocs
                .EntIdDocInicial = aobjFactura.ObjIdFacturaEnt.ObjValorPro
                .EntIdDocFinal = aobjFactura.ObjIdFacturaEnt.ObjValorPro
            End With
            If ablnEnvioEmail Then
                If aobjFactura.BlnEnviarPorCorreo Then
                    SExporteFactura(lstrPrefFac, lentIdFac, ablnEnvioEmail, lrcFactura)
                End If
            Else
                SExporteFactura(lstrPrefFac, lentIdFac, ablnEnvioEmail, lrcFactura)
            End If
            If lrcFactura IsNot Nothing Then
                lrcFactura.Close()
                lrcFactura.Dispose()
            End If
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
    ''' <summary>
    ''' Exporta la factura pasada en el argumento a la carpeta de email para ser enviada 
    ''' al FTP de eFac
    ''' </summary>
    Friend Sub SExporteFactura(aobjFactura As ClsFactura, aobjTerCentroUtil As ClsTercero)
        Try
            SRefresqueLogo(aobjTerCentroUtil)
            SRefresqueCodigoQR()
            If Not aobjFactura.BlnExiste Then
                Throw New ErrorInesperadoPanLException("Factura a exportar no existe")
            End If
            Dim lstrPrefFac As String = aobjFactura.ObjPrefijo_FactStr.ObjValorPro
            Dim lentIdFac As Integer = aobjFactura.ObjIdFacturaEnt.ObjValorPro
            ObjParRepDocs = New ClsParametrosReportesDocs(lstrPrefFac, 0, 0)
            Dim lrcFactura As ReportClass = Nothing
            If GobjParametros.BlnEFacAutorizado Then
                lrcFactura = New repFacturaEFac
            Else
                If aobjFactura.ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuManual OrElse
                aobjFactura.ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuImportada Then
                    lrcFactura = New repFactImportadaLogo
                Else
                    lrcFactura = New repFacturaLogo
                End If
            End If
            With ObjParRepDocs
                .EntIdDocInicial = aobjFactura.ObjIdFacturaEnt.ObjValorPro
                .EntIdDocFinal = aobjFactura.ObjIdFacturaEnt.ObjValorPro
            End With
            SExporteFactura(lstrPrefFac, lentIdFac, True, lrcFactura)
            If lrcFactura IsNot Nothing Then
                lrcFactura.Close()
                lrcFactura.Dispose()
            End If
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Private Sub SExporteFactura(astrPrefFac As String, aentIdFact As Integer,
            ablnEnvioEmail As Boolean, arcFactura As ReportClass)
        Using ldsFacturas As New DataSet
            SGenereDataSetFacturas(ldsFacturas, False)
            arcFactura.SetDataSource(ldsFacturas)
        End Using
        Dim leopOpcionesExp As New ExportOptions
        Dim lddoOpcionesDisco As DiskFileDestinationOptions =
                    ExportOptions.CreateDiskFileDestinationOptions()
        If ablnEnvioEmail Then
            lddoOpcionesDisco.DiskFileName = ClsOrionCop.FstrArchivoPdfDcto(astrPrefFac,
                    aentIdFact, EnuTipoDocOri.EnuFactura)
        Else
            lddoOpcionesDisco.DiskFileName = ClsOrionCop.FstrNombreArchFactura(astrPrefFac,
                    aentIdFact)
        End If
        leopOpcionesExp.ExportFormatType = ExportFormatType.PortableDocFormat
        leopOpcionesExp.ExportDestinationType = ExportDestinationType.DiskFile
        leopOpcionesExp.ExportDestinationOptions = lddoOpcionesDisco
        arcFactura.Export(leopOpcionesExp)
    End Sub
    ''' <summary>
    ''' Exporta un recibo de caja como .Pdf a la carpetas de emails para ser enviado por internet
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub SExporteUnReciboCaja(aobjRecCaja As ClsReciboCaja,
            aobjTerCentroUtil As ClsTercero, ablnFirma As Boolean)
        Try
            SRefresqueLogo(aobjTerCentroUtil)
            Dim lstrPrefRC As String = aobjRecCaja.ObjPrefijo_RecStr.ObjValorPro
            Dim lentIdRC As Integer = aobjRecCaja.ObjIdRecCajaEnt.ObjValorPro
            Dim lrcRecCaja As New repRecCajaConLogo
            SExporteRecCaja(lstrPrefRC, lentIdRC, lrcRecCaja, ablnFirma)
            lrcRecCaja.Close()
            lrcRecCaja.Dispose()
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Private Sub SExporteRecCaja(astrPrefRec As String, aentIdRec As Integer,
                arcRecCaja As ReportClass, ablnFirma As Boolean)
        Dim lblnNoHayError = False, lstrNombreArch As String
        ObjParRepDocs = New ClsParametrosReportesDocs(astrPrefRec, aentIdRec, aentIdRec)
        lstrNombreArch = ClsOrionCop.FstrNombreArchRecibo(astrPrefRec, aentIdRec)
        Try
            Using ldsRecCaja As New DataSet
                SGenereDataSetRecCaja(ldsRecCaja, ablnFirma)
                arcRecCaja.SetDataSource(ldsRecCaja)
            End Using
            lblnNoHayError = True
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                Dim leopOpcionesExp As New ExportOptions
                Dim lddoOpcionesDisco As DiskFileDestinationOptions =
                        ExportOptions.CreateDiskFileDestinationOptions()
                lddoOpcionesDisco.DiskFileName = lstrNombreArch
                leopOpcionesExp.ExportFormatType = ExportFormatType.PortableDocFormat
                leopOpcionesExp.ExportDestinationType = ExportDestinationType.DiskFile
                leopOpcionesExp.ExportDestinationOptions = lddoOpcionesDisco
                arcRecCaja.Export(leopOpcionesExp)
            End If
        End Try
    End Sub
    Friend Sub SExporteRecibosCaja()
        Try
            SRefresqueLogo()
            Dim lrcRecCaja As ReportClass
            lrcRecCaja = New repRecCajaConLogo
            SExporteRecibosCaja(lrcRecCaja)
            lrcRecCaja.Close()
            lrcRecCaja.Dispose()
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Private Sub SExporteRecibosCaja(arcRecCaja As ReportClass)
        Dim lblnNoHayError = False, lstrNombreArch As String
        Dim lstrFechaDesde = StrFechaDesde.Replace("-", "")
        Dim lstrFechaHasta = StrFechaHasta.Replace("-", "")
        lstrNombreArch = "RecibosCaja_Entre_" & lstrFechaDesde & "_y_" & lstrFechaHasta
        lstrNombreArch = GstrTrayRecibosCajaPdf & "\" & lstrNombreArch & ".Pdf"
        Try
            Using ldsRecCaja As New DataSet
                SGenereDataSetRecibosCaja(ldsRecCaja, False)
                arcRecCaja.SetDataSource(ldsRecCaja)
            End Using
            lblnNoHayError = True
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                Dim leopOpcionesExp As New ExportOptions
                Dim lddoOpcionesDisco As DiskFileDestinationOptions =
                        ExportOptions.CreateDiskFileDestinationOptions()
                lddoOpcionesDisco.DiskFileName = lstrNombreArch
                leopOpcionesExp.ExportFormatType = ExportFormatType.PortableDocFormat
                leopOpcionesExp.ExportDestinationType = ExportDestinationType.DiskFile
                leopOpcionesExp.ExportDestinationOptions = lddoOpcionesDisco
                arcRecCaja.Export(leopOpcionesExp)
            End If
        End Try
    End Sub
    Friend Sub SExporteFacturas()
        SRefresqueLogo()
        SRefresqueCodigoQR()
        Dim lrcFac As ReportClass
        If GobjParametros.BlnEFacAutorizado Then
            lrcFac = New repFacturaEFac
        Else
            If GobjParametros.ObjAnoActual.ObjTipoDsctoPPByt.ObjValorPro <>
                        EnuTipoDsctoPP.None Then
                lrcFac = New repFacturaDsctoPPLogo
            Else
                lrcFac = New repFacturaLogo
            End If
        End If
        SExporteFactsEntreFechas(lrcFac)
        lrcFac.Close()
        lrcFac.Dispose()
    End Sub
    Private Sub SExporteFactsEntreFechas(arcFacs As ReportClass)
        Dim lblnNoHayError = False, lstrNombreArch As String
        Dim lstrFechaDesde = StrFechaDesde.Replace("-", "")
        Dim lstrFechaHasta = StrFechaHasta.Replace("-", "")
        lstrNombreArch = "Facturas_Entre_" & lstrFechaDesde & "_y_" &
                lstrFechaHasta
        lstrNombreArch = GstrTrayFacturasPdf & "\" & lstrNombreArch & ".Pdf"
        Try
            Using ldsFacs As New DataSet
                SGenereDataSetFactsFechas(ldsFacs, False)
                arcFacs.SetDataSource(ldsFacs)
            End Using
            lblnNoHayError = True
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                Dim leopOpcionesExp As New ExportOptions
                Dim lddoOpcionesDisco As DiskFileDestinationOptions =
                            ExportOptions.CreateDiskFileDestinationOptions()
                lddoOpcionesDisco.DiskFileName = lstrNombreArch
                leopOpcionesExp.ExportFormatType = ExportFormatType.PortableDocFormat
                leopOpcionesExp.ExportDestinationType = ExportDestinationType.DiskFile
                leopOpcionesExp.ExportDestinationOptions = lddoOpcionesDisco
                arcFacs.Export(leopOpcionesExp)
            End If
        End Try
    End Sub
    Friend Sub SExporteNotaCr(aobjNotaCr As ClsNotaCr, aobjTerCenUtil As ClsTercero)
        Dim lrcNotaCr As ReportClass = Nothing
        Try
            SRefresqueLogo(aobjTerCenUtil)
            Dim lstrPrefNcr As String = aobjNotaCr.ObjPrefijo_NotaCrStr.ObjValorPro
            Dim lentIdNotaCr As Integer = aobjNotaCr.ObjIdNotaCrEnt.ObjValorPro
            ObjParRepDocs = New ClsParametrosReportesDocs(lstrPrefNcr, lentIdNotaCr, lentIdNotaCr)
            lrcNotaCr = New repNotasCrLogo
            Using ldsNotaCr As New DataSet
                SGenereDataSetNotasCr(ldsNotaCr)
                lrcNotaCr.SetDataSource(ldsNotaCr)
            End Using
            Dim leopOpcionesExp As New ExportOptions
            Dim lddoOpcionesDisco As DiskFileDestinationOptions =
                        ExportOptions.CreateDiskFileDestinationOptions()
            lddoOpcionesDisco.DiskFileName = ClsOrionCop.FstrArchivoPdfDcto(lstrPrefNcr,
                        lentIdNotaCr, EnuTipoDocOri.EnuNotaCr)
            leopOpcionesExp.ExportFormatType = ExportFormatType.PortableDocFormat
            leopOpcionesExp.ExportDestinationType = ExportDestinationType.DiskFile
            leopOpcionesExp.ExportDestinationOptions = lddoOpcionesDisco
            lrcNotaCr.Export(leopOpcionesExp)
            lrcNotaCr.Close()
            lrcNotaCr.Dispose()
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lrcNotaCr IsNot Nothing Then
                lrcNotaCr.Dispose()
            End If
        End Try
    End Sub
    Friend Sub SExporteNotaDb(aobjNotaDb As ClsNotaDb, aobjTerCenUtil As ClsTercero)
        Dim lstrPrefNotaDb As String = aobjNotaDb.ObjPrefijo_NotaDbStr.ObjValorPro
        Dim lentIdNotaDb As Integer = aobjNotaDb.ObjIdNotaDbEnt.ObjValorPro
        SRefresqueLogo(aobjTerCenUtil)
        ObjParRepDocs = New ClsParametrosReportesDocs(lstrPrefNotaDb, lentIdNotaDb, lentIdNotaDb)
        Dim lrcNotaDb As New repNotasDbLogo
        Try
            Using ldsNotaDb As New DataSet
                SGenereDataSetNotasDb(ldsNotaDb)
                lrcNotaDb.SetDataSource(ldsNotaDb)
            End Using
            Dim leopOpcionesExp As New ExportOptions
            Dim lddoOpcionesDisco As DiskFileDestinationOptions =
                        ExportOptions.CreateDiskFileDestinationOptions()
            lddoOpcionesDisco.DiskFileName = ClsOrionCop.FstrArchivoPdfDcto(lstrPrefNotaDb,
                        lentIdNotaDb, EnuTipoDocOri.EnuNotaDb)
            leopOpcionesExp.ExportFormatType = ExportFormatType.PortableDocFormat
            leopOpcionesExp.ExportDestinationType = ExportDestinationType.DiskFile
            leopOpcionesExp.ExportDestinationOptions = lddoOpcionesDisco
            lrcNotaDb.Export(leopOpcionesExp)
            lrcNotaDb.Close()
            lrcNotaDb.Dispose()
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lrcNotaDb IsNot Nothing Then
                lrcNotaDb.Dispose()
            End If
        End Try
    End Sub
    Friend Sub SExporteNotaCon(aobjNotaCon As ClsNotaCon, aobjTerCenutil As ClsTercero)
        Dim lrcNotaCon As ReportClass = Nothing
        Try
            Dim lstrPrefNotaCon As String = aobjNotaCon.ObjPrefijo_NotaConStr.ObjValorPro
            Dim lentIdNotaCon As Integer = aobjNotaCon.ObjIdNotaConEnt.ObjValorPro
            SRefresqueLogo(aobjTerCenutil)
            ObjParRepDocs = New ClsParametrosReportesDocs(lstrPrefNotaCon, lentIdNotaCon, lentIdNotaCon)
            lrcNotaCon = New repNotasConLogo
            Using ldsNotaCon As New DataSet
                SGenereDataSetNotasCon(ldsNotaCon)
                lrcNotaCon.SetDataSource(ldsNotaCon)
            End Using
            Dim leopOpcionesExp As New ExportOptions
            Dim lddoOpcionesDisco As DiskFileDestinationOptions =
                        ExportOptions.CreateDiskFileDestinationOptions()
            lddoOpcionesDisco.DiskFileName = ClsOrionCop.FstrArchivoPdfDcto(lstrPrefNotaCon,
                        lentIdNotaCon, EnuTipoDocOri.EnuNotaCon)
            leopOpcionesExp.ExportFormatType = ExportFormatType.PortableDocFormat
            leopOpcionesExp.ExportDestinationType = ExportDestinationType.DiskFile
            leopOpcionesExp.ExportDestinationOptions = lddoOpcionesDisco
            lrcNotaCon.Export(leopOpcionesExp)
            lrcNotaCon.Close()
            lrcNotaCon.Dispose()
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lrcNotaCon IsNot Nothing Then
                lrcNotaCon.Dispose()
            End If
        End Try
    End Sub
    Friend Sub SExporteNotaRCr(aobjNotaRCr As ClsNotaReversionCr, aobjTerCenUtil As ClsTercero)
        Dim lstrPrefNotaRCr As String = aobjNotaRCr.ObjPrefijo_NotaReversaCrStr.ObjValorPro
        Dim lentIdNotaRCr As Integer = aobjNotaRCr.ObjIdNotaReversaCrEnt.ObjValorPro
        SRefresqueLogo(aobjTerCenUtil)
        Dim lrcNotaRCr As ReportClass
        ObjParRepDocs = New ClsParametrosReportesDocs(lstrPrefNotaRCr, lentIdNotaRCr, lentIdNotaRCr)
        lrcNotaRCr = New repNotasRevCrLogo
        Try
            Using ldsNotaRCr As New DataSet
                SGenereDataSetNotasRevCr(ldsNotaRCr)
                lrcNotaRCr.SetDataSource(ldsNotaRCr)
            End Using
            Dim leopOpcionesExp As New ExportOptions
            Dim lddoOpcionesDisco As DiskFileDestinationOptions =
                        ExportOptions.CreateDiskFileDestinationOptions()
            lddoOpcionesDisco.DiskFileName = ClsOrionCop.FstrArchivoPdfDcto(lstrPrefNotaRCr,
                        lentIdNotaRCr, EnuTipoDocOri.EnuNotaRevCr)
            leopOpcionesExp.ExportFormatType = ExportFormatType.PortableDocFormat
            leopOpcionesExp.ExportDestinationType = ExportDestinationType.DiskFile
            leopOpcionesExp.ExportDestinationOptions = lddoOpcionesDisco
            lrcNotaRCr.Export(leopOpcionesExp)
            lrcNotaRCr.Close()
            lrcNotaRCr.Dispose()
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lrcNotaRCr IsNot Nothing Then
                lrcNotaRCr.Dispose()
            End If
        End Try
    End Sub
    Private Sub SExporteEdadCartera(arcEdadCartera As ReportClass, ablnCierreMes As Boolean,
            adtbEdadCartera As DataTable, astrNombresColumnas As String(),
            adtmFechaDatos As Date, ablnDetallado As Boolean)
        Dim leopOpcionesExp As New ExportOptions
        Dim lddoOpcionesDisco As DiskFileDestinationOptions =
                ExportOptions.CreateDiskFileDestinationOptions()
        lddoOpcionesDisco.DiskFileName = FstrNombreArchEdadCartera(EnuTipoArchivoDef.EnuPdf,
                adtmFechaDatos, ablnCierreMes)
        leopOpcionesExp.ExportFormatType = ExportFormatType.PortableDocFormat
        leopOpcionesExp.ExportDestinationType = ExportDestinationType.DiskFile
        leopOpcionesExp.ExportDestinationOptions = lddoOpcionesDisco
        arcEdadCartera.Export(leopOpcionesExp)
        '
        lddoOpcionesDisco.DiskFileName = FstrNombreArchEdadCartera(
                EnuTipoArchivoDef.EnuCrystalRepo, adtmFechaDatos, ablnCierreMes)
        leopOpcionesExp.ExportFormatType = ExportFormatType.CrystalReport
        leopOpcionesExp.ExportDestinationOptions = lddoOpcionesDisco
        arcEdadCartera.Export(leopOpcionesExp)
        Dim lobjExp As New ClsExportToExcell
        Dim lstrNomArchivo = FstrNombreArchEdadCartera(EnuTipoArchivoDef.EnuExcelx,
                adtmFechaDatos, ablnCierreMes)
        lobjExp.SExporteToExcel(lstrNomArchivo, adtbEdadCartera, astrNombresColumnas,
                True, ablnDetallado)
    End Sub
    Private Function FstrNombreArchEdadCartera(aenuTipoArchivo As EnuTipoArchivoDef,
                adtmFechaDatos As Date, ablnCierreMes As Boolean) As String
        Dim lstrExtension = String.Empty
        Select Case aenuTipoArchivo
            Case EnuTipoArchivoDef.EnuPdf
                lstrExtension = ".pdf"
            Case EnuTipoArchivoDef.EnuCrystalRepo
                lstrExtension = ".rpt"
            Case EnuTipoArchivoDef.EnuExcelx
                lstrExtension = ".xlsx"
        End Select
        Dim lstrTrayFM As String
        If ablnCierreMes Then
            lstrTrayFM = GstrTrayReportes & "\FinMes"
        Else
            lstrTrayFM = GstrTrayReportes
        End If
        If Not My.Computer.FileSystem.DirectoryExists(lstrTrayFM) Then
            My.Computer.FileSystem.CreateDirectory(lstrTrayFM)
        End If
        With MstbExpresionSql
            .Clear.Append(lstrTrayFM).Append("\EdadCartera_")
            If ablnCierreMes Then
                .Append("FM_")
            End If
            .Append(adtmFechaDatos.Year.ToString).Append(Format(adtmFechaDatos.Month, "0#"))
            .Append(Format(adtmFechaDatos.Day, "0#")).Append("_")
        End With
        Dim lstrNomArchivoBase = MstbExpresionSql.ToString
        Dim lstrNomArchivo = String.Empty
        Dim i = 1
        Do While True
            lstrNomArchivo = lstrNomArchivoBase & Format(i, "0#") & lstrExtension
            If My.Computer.FileSystem.FileExists(lstrNomArchivo) Then
                i += 1
                If i > 50 Then
                    Throw New ErrorInesperadoPanLException("Error conformando nombre de archivo")
                End If
            Else
                Exit Do
            End If
        Loop
        Return lstrNomArchivo
    End Function
    Private Sub SExporteAntPorAplicar(arcAntPorAplicar As ReportClass,
            adtbAntPorAplicar As DataTable, astrNombresColumnas As String(),
            adtmFechaDatos As Date)
        Dim leopOpcionesExp As New ExportOptions
        Dim lddoOpcionesDisco As DiskFileDestinationOptions =
                ExportOptions.CreateDiskFileDestinationOptions()
        lddoOpcionesDisco.DiskFileName = FstrNombreArchAntPorAplicar(EnuTipoArchivoDef.EnuPdf,
                adtmFechaDatos)
        leopOpcionesExp.ExportFormatType = ExportFormatType.PortableDocFormat
        leopOpcionesExp.ExportDestinationType = ExportDestinationType.DiskFile
        leopOpcionesExp.ExportDestinationOptions = lddoOpcionesDisco
        arcAntPorAplicar.Export(leopOpcionesExp)
        '
        lddoOpcionesDisco.DiskFileName = FstrNombreArchAntPorAplicar(
                EnuTipoArchivoDef.EnuCrystalRepo, adtmFechaDatos)
        leopOpcionesExp.ExportFormatType = ExportFormatType.CrystalReport
        leopOpcionesExp.ExportDestinationOptions = lddoOpcionesDisco
        arcAntPorAplicar.Export(leopOpcionesExp)
        Dim lobjExp As New ClsExportToExcell
        Dim lstrNomArchivo = FstrNombreArchAntPorAplicar(EnuTipoArchivoDef.EnuExcelx,
                adtmFechaDatos)
        lobjExp.SExporteToExcel(lstrNomArchivo, adtbAntPorAplicar, astrNombresColumnas)
    End Sub
    Private Function FstrNombreArchAntPorAplicar(aenuTipoArchivo As EnuTipoArchivoDef,
                adtmFechaDatos As Date) As String
        Dim lstrExtension = String.Empty
        Select Case aenuTipoArchivo
            Case EnuTipoArchivoDef.EnuPdf
                lstrExtension = ".pdf"
            Case EnuTipoArchivoDef.EnuCrystalRepo
                lstrExtension = ".rpt"
            Case EnuTipoArchivoDef.EnuExcelx
                lstrExtension = ".xlsx"
        End Select
        Dim lstrTrayFM = GstrTrayReportes & "\FinMes"
        If Not My.Computer.FileSystem.DirectoryExists(lstrTrayFM) Then
            My.Computer.FileSystem.CreateDirectory(lstrTrayFM)
        End If
        With MstbExpresionSql
            .Clear.Append(lstrTrayFM).Append("\AntXAplicar_FM_")
            .Append(adtmFechaDatos.Year.ToString).Append(Format(adtmFechaDatos.Month, "0#"))
            .Append(Format(adtmFechaDatos.Day, "0#")).Append("_")
        End With
        Dim lstrNomArchivoBase = MstbExpresionSql.ToString
        Dim lstrNomArchivo = String.Empty
        Dim i = 1
        Do While True
            lstrNomArchivo = lstrNomArchivoBase & Format(i, "0#") & lstrExtension
            If My.Computer.FileSystem.FileExists(lstrNomArchivo) Then
                i += 1
                If i > 50 Then
                    Throw New ErrorInesperadoPanLException("Error conformando nombre de archivo")
                End If
            Else
                Exit Do
            End If
        Loop
        Return lstrNomArchivo
    End Function
    Friend Sub SExporteEstadoCta(ByRef astrNombreArchivo As String)
        Dim lrcEstadoCta As New repEstadoCtaCli
        Using ldsEstadoCta As New DataSet
            SGenereDataSetEstadoCtaCli(ldsEstadoCta)
            lrcEstadoCta.SetDataSource(ldsEstadoCta)
        End Using
        astrNombreArchivo = ClsOrionCop.FstrNomArchEstadoCtaPreAgr(StrIdPredioAgru)
        Dim leopOpcionesExp As New ExportOptions
        Dim lddoOpcionesDisco As DiskFileDestinationOptions =
                        ExportOptions.CreateDiskFileDestinationOptions()
        lddoOpcionesDisco.DiskFileName = astrNombreArchivo
        leopOpcionesExp.ExportFormatType = ExportFormatType.PortableDocFormat
        leopOpcionesExp.ExportDestinationType = ExportDestinationType.DiskFile
        leopOpcionesExp.ExportDestinationOptions = lddoOpcionesDisco
        lrcEstadoCta.Export(leopOpcionesExp)
        lrcEstadoCta.Close()
        lrcEstadoCta.Dispose()
    End Sub
#End Region
#Region "Abrir Reportes ya Generados"
    Friend Function SAbraReporte(astrNombreComplArchivo As String) As String
        Dim lstrMens = String.Empty
        Dim lfrmRep = FfrmRep()
        If My.Computer.FileSystem.FileExists(astrNombreComplArchivo) Then
            Dim lrcReporteEnDisco As New ReportClass With {
                .FileName = astrNombreComplArchivo
            }
            lrcReporteEnDisco.Load(astrNombreComplArchivo, OpenReportMethod.OpenReportByDefault)
            lfrmRep.RcReporte = lrcReporteEnDisco
            lfrmRep.Show()
        Else
            lstrMens = "El archivo solicitado no existe en este Equipo!"
        End If
        Return lstrMens
    End Function
#End Region
#Region "DataSets"
#Region "Documentos (FAC, REC, NDB, Etc.)"
#Region "Factura"
    Private Sub SGenereDataSetFacturas(adsFactura As DataSet, ablnExcluirEnvMail As Boolean)
        If IsNothing(ObjParRepDocs) Then
            Throw New ErrorInesperadoPanLException("Sin Parametros para el Reporte de Facturas")
        End If
        MstbExpresionSql.Clear()
        Dim lstrPrefFacturas = ObjParRepDocs.StrPrefijoDocsRep
        Dim lstrIdFacIni = ObjParRepDocs.EntIdDocInicial.ToString
        Dim lstrIdFacFin = ObjParRepDocs.EntIdDocFinal.ToString
        Dim lstrSqlFactura = FstrExpSqlFactura(lstrPrefFacturas, lstrIdFacIni, lstrIdFacFin,
                ablnExcluirEnvMail)
        Dim lstrSqlDetalleFac = FstrExpSqlDetalleFac(lstrPrefFacturas, lstrIdFacIni, lstrIdFacFin)
        Dim lstrSqlEstadoCuenta = FstrExpSqlEstadoCuenta(lstrPrefFacturas, lstrIdFacIni, lstrIdFacFin)
        Dim lstrSqlDetEstadoCta = FstrExpDetEstadoCuenta(lstrPrefFacturas, lstrIdFacIni, lstrIdFacFin)
        Dim lstrSqlAnticiposApl = FstrExpSqlAntApli(lstrPrefFacturas, lstrIdFacIni, lstrIdFacFin)
        Dim lstrSqlServicios = FstrExpSqlServicios()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlFactura)
        lcolExpresionesSql.Add(lstrSqlDetalleFac)
        lcolExpresionesSql.Add(lstrSqlEstadoCuenta)
        lcolExpresionesSql.Add(lstrSqlDetEstadoCta)
        lcolExpresionesSql.Add(lstrSqlAnticiposApl)
        lcolExpresionesSql.Add(lstrSqlServicios)
        lcolNombresTablas.Add("OriFactura")
        lcolNombresTablas.Add("OriItemsFactura")
        lcolNombresTablas.Add("OriEstadosCuenta")
        lcolNombresTablas.Add("OriDetEstadosCta")
        lcolNombresTablas.Add("OriAnticipos")
        lcolNombresTablas.Add("OriServicios")
        GobjPanDat.SdsDataSet(adsFactura, lcolExpresionesSql, lcolNombresTablas)
        SComplementeDtbFacturas(adsFactura)
        adsFactura.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "Facturas" & ".XML"
        'adsFactura.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlFactura(astrPrefFacturas As String, astrIdFacIni As String,
            astrIdFacFin As String, ablnExcluirEnvMail As Boolean) As String
        With MstbExpresionSql
            .Clear.Append("SELECT F.").Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFacturaEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(" AS FechaPlazo").Append(", ")
            .Append(ClsFechaGraciaDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaDctoProntoPagoDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd).Append(", ")
            .Append("'' AS NomPreAgr").Append(", ")
            .Append(ClsReferenciaPago_FacStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPieFacturaUno_FactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPieFacturaDos_FactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsAnuladoBln.SstrNombreCampoBd).Append(", ")
            .Append(ClsDctoProntoPago_FacDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsValor_FactDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsCUFEStr.SstrNombreCampoBd).Append(", 0.00 AS ValorIva")
            .Append(", 0.00 AS ValorBaseIva").Append(", 0.00 AS AntiApli")
            .Append(", 0.00 AS TotalAPagar").Append(", 0.00 AS TasaMora")
            .Append(" FROM ").Append(ClsFactura.SstrNombreTabla)
            .Append(" AS F INNER JOIN ").Append(ClsCliente.SstrNombreTabla).Append(" AS C ON F.")
            .Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(" = ").Append("C.")
            .Append(ClsIdClienteDbl.SstrNombreCampoBd).Append(" AND F.").Append(MstrCampoBdCarpeta)
            .Append(" = C.").Append(MstrCampoBdCarpeta).Append(" AND F.")
            .Append(MstrCampoBdCentroUtil).Append(" = C.").Append(MstrCampoBdCentroUtil)
            .Append(" WHERE F.").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND F.").Append(MstrCampoBdCentroUtil).Append(" = ")
            .Append(GshrIdCentroUtil.ToString).Append(" AND F.")
            .Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(" = '")
            .Append(astrPrefFacturas).Append("' AND F.")
            .Append(ClsIdFacturaEnt.SstrNombreCampoBd).Append(" >= ").Append(astrIdFacIni).Append(" AND F.")
            .Append(ClsIdFacturaEnt.SstrNombreCampoBd).Append(" <= ").Append(astrIdFacFin)
            If ablnExcluirEnvMail Then
                .Append(" AND F.").Append(ClsEnviadaMailBln.SstrNombreCampoBd)
                .Append(" = FALSE")
            End If
            .Append(FstrFiltroFrasElec)
            .Append(" ORDER BY ").Append(ClsIdFacturaEnt.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpSqlFactsFechas(ablnExcluirEnvMail As Boolean) As String
        With MstbExpresionSql
            .Clear.Append("SELECT F.").Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFacturaEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(" AS FechaPlazo").Append(", ")
            .Append(ClsFechaGraciaDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaDctoProntoPagoDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd).Append(", ")
            .Append("'' AS NomPreAgr").Append(", ")
            .Append(ClsReferenciaPago_FacStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPieFacturaUno_FactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPieFacturaDos_FactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsAnuladoBln.SstrNombreCampoBd).Append(", ")
            .Append(ClsDctoProntoPago_FacDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsValor_FactDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsCUFEStr.SstrNombreCampoBd).Append(", 0.00 AS ValorIva")
            .Append(", 0.00 AS ValorBaseIva").Append(", 0.00 AS AntiApli")
            .Append(", 0.00 AS TotalAPagar").Append(", 0.00 AS TasaMora")
            .Append(" FROM ").Append(ClsFactura.SstrNombreTabla)
            .Append(" AS F INNER JOIN ").Append(ClsCliente.SstrNombreTabla).Append(" AS C ON F.")
            .Append(MstrCampoBdCarpeta).Append(" = ").Append("C.").Append(MstrCampoBdCarpeta)
            .Append(" AND F.").Append(MstrCampoBdCentroUtil).Append(" = ").Append("C.")
            .Append(MstrCampoBdCentroUtil).Append(" AND F.")
            .Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(" = ").Append("C.")
            .Append(ClsIdClienteDbl.SstrNombreCampoBd)
            .Append(" WHERE F.").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND F.").Append(MstrCampoBdCentroUtil).Append(" = ")
            .Append(GshrIdCentroUtil.ToString).Append(" AND F.")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(" BETWEEN '")
            .Append(StrFechaDesde).Append("' AND '").Append(StrFechaHasta).Append("'")
            If ablnExcluirEnvMail Then
                .Append(" AND F.").Append(ClsEnviadaMailBln.SstrNombreCampoBd)
                .Append(" = FALSE")
            End If
            .Append(FstrFiltroFrasElec)
            .Append(" ORDER BY ").Append(ClsIdFacturaEnt.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpSqlDetalleFacsFechas() As String
        Dim lstrTablbaPri = ClsItemFactura.SstrNombreTabla
        Dim lstrTablaSec = ClsFactura.SstrNombreTabla
        Dim lstrCampSelPri As String() = {ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
                ClsIdFactura_ItemFactEnt.SstrNombreCampoBd, ClsIdItemFacturaShr.SstrNombreCampoBd,
                ClsDetalle_ItemFactStr.SstrNombreCampoBd,
                ClsTarifaIva_ItemFactDbl.SstrNombreCampoBd & " * 100 AS TarifaIva",
                ClsValor_ItemFactDec.SstrNombreCampoBd}
        Dim lstrCampSelSec As String() = {}
        Dim lstrCampRelPri As String() = {MstrCampoBdCarpeta, MstrCampoBdCentroUtil,
                ClsPrefijo_ItemFactStr.SstrNombreCampoBd, ClsIdFactura_ItemFactEnt.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {MstrCampoBdCarpeta, MstrCampoBdCentroUtil,
                ClsPrefijo_FactStr.SstrNombreCampoBd, ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{ClsPrefijo_ItemFactStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFactura_ItemFactEnt.SstrNombreCampoBd, "ASC"},
                {ClsIdItemFacturaShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = "P." & MstrCampoBdCarpeta & " = " & GshrIdCarpeta & " AND P." &
                MstrCampoBdCentroUtil & " = " & GshrIdCentroUtil & " AND S." &
                ClsFechaFacturaDtm.SstrNombreCampoBd & " BETWEEN '" & StrFechaDesde & "' AND '" &
                StrFechaHasta & "'"
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablbaPri, lstrCampSelPri,
                lstrTablaSec, lstrCampSelSec, lstrCampRelPri, lstrCampRelSec, lstrOrden,
                lstrFiltro, {})
        Return lstrExpSql
    End Function
    Private Function FstrExpSqlEstadoCtaFechas() As String
        Dim ldtmFechaDesde As Date = DateSerial(CInt(StrFechaDesde.Substring(0, 4)),
                CInt(StrFechaDesde.Substring(5, 2)), CInt(StrFechaDesde.Substring(8, 2)))
        ldtmFechaDesde = ldtmFechaDesde.AddDays(-1)
        Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaDesde)
        With MstbExpresionSql
            .Clear.Append("SELECT ").Append(ClsIdEstadoCuentaEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijoFac_EstadoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsDeudaCapitalDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsDeudaIntMoraDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsAntPorAplDec.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsEstadoCuenta.SstrNombreTabla)
            .Append(" WHERE ").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND ").Append(MstrCampoBdCentroUtil).Append(" = ")
            .Append(GshrIdCentroUtil.ToString).Append(" AND ")
            .Append(ClsFechaEstadoDtm.SstrNombreCampoBd).Append(" BETWEEN '").Append(lstrFechaDesde)
            .Append("' AND '").Append(StrFechaHasta).Append("' ORDER BY ")
            .Append(ClsIdEstadoCuentaEnt.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpDetEstadoCtaFechas() As String
        Dim lshrIdAnoAct As Short = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
        Dim ldtmFechaDesde As Date = DateSerial(CInt(StrFechaDesde.Substring(0, 4)),
                CInt(StrFechaDesde.Substring(5, 2)), CInt(StrFechaDesde.Substring(8, 2)))
        ldtmFechaDesde = ldtmFechaDesde.AddDays(-1)
        Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaDesde)
        With MstbExpresionSql
            .Clear()
            .Append("SELECT ").Append(ClsPrefijoFac_EstadoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd).Append(", IF(")
            .Append(ClsIdAno_ItemFactEstadoShr.SstrNombreCampoBd).Append(">0,").Append(lshrIdAnoAct)
            .Append(", 0) AS IdAnos, ").Append(ClsIdServicioItemFac_EstadoShr.SstrNombreCampoBd)
            .Append(", SUM(F.").Append(ClsDeudaCap_ItFacEstDec.SstrNombreCampoBd)
            .Append(") AS DeudaCap, ").Append("SUM(F.").Append(ClsDeudaIntMora_ItFacEstDec.
            SstrNombreCampoBd).Append(" - ").Append(ClsDeudaIntMes_ItFacEstDec.SstrNombreCampoBd)
            .Append(") AS ValorInt, SUM(").Append(ClsDeudaIntMes_ItFacEstDec.SstrNombreCampoBd)
            .Append(") AS IntMes ").Append(" FROM ").Append(ClsFacturaEstado.SstrNombreTabla)
            .Append(" AS F INNER JOIN ").Append(ClsEstadoCuenta.SstrNombreTabla).Append(" AS E ON F.")
            .Append(MstrCampoBdCarpeta).Append(" = E.").Append(MstrCampoBdCarpeta).Append(" AND F.")
            .Append(MstrCampoBdCentroUtil).Append(" = E.").Append(MstrCampoBdCentroUtil)
            .Append(" AND F.").Append(ClsIdEstado_FactEstadoEnt.SstrNombreCampoBd).Append(" = E.")
            .Append(ClsIdEstadoCuentaEnt.SstrNombreCampoBd).Append(" WHERE F.")
            .Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND F.").Append(MstrCampoBdCentroUtil).Append(" = ")
            .Append(GshrIdCentroUtil.ToString).Append(" AND ")
            .Append(ClsFechaEstadoDtm.SstrNombreCampoBd).Append(" BETWEEN '").Append(lstrFechaDesde)
            .Append("' AND '").Append(StrFechaHasta).Append("' GROUP BY ")
            .Append(ClsPrefijoFac_EstadoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd).Append(", ").Append("IdAnos")
            .Append(", ").Append(ClsIdServicioItemFac_EstadoShr.SstrNombreCampoBd)
            .Append(" ORDER BY ").Append(ClsPrefijoFac_EstadoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd).Append(", IdAnos DESC ,")
            .Append(ClsIdServicioItemFac_EstadoShr.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString()
    End Function
    Private Function FstrExpSqlAntApliFechas() As String
        Dim lstrTablaPri = ClsItemNotaCon.SstrNombreTabla
        Dim lstrTablaSec = ClsNotaCon.SstrNombreTabla
        Dim lstrCampSelPri As String() = {ClsPrefijoFact_ItemNotaConStr.SstrNombreCampoBd,
                ClsIdFactura_ItemNotaConEnt.SstrNombreCampoBd,
                "SUM(" & ClsValor_ItemNotaConDec.SstrNombreCampoBd & ") AS Valor "}
        Dim lstrCampSelSec As String() = {}
        Dim lstrCamRelPri As String() = {MstrCampoBdCarpeta, MstrCampoBdCentroUtil,
                ClsPrefijo_NotaConStr.SstrNombreCampoBd,
                ClsIdNotaCon_ItemNotaConEnt.SstrNombreCampoBd}
        Dim LstrCampRelSec As String() = {MstrCampoBdCarpeta, MstrCampoBdCentroUtil,
                ClsPrefijo_NotaConStr.SstrNombreCampoBd, ClsIdNotaConEnt.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{ClsIdFactura_ItemNotaConEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrfiltro As String = "P." & MstrCampoBdCarpeta & " = " & GshrIdCarpeta & " AND P." &
                MstrCampoBdCentroUtil & " = " & GshrIdCentroUtil & " AND " &
                ClsFecha_NotaConDtm.SstrNombreCampoBd & " BETWEEN '" & StrFechaDesde & "' AND '" &
                StrFechaHasta & "'"
        Dim lstrCampGrupo As String() = {ClsPrefijoFact_ItemNotaConStr.SstrNombreCampoBd,
                ClsIdFactura_ItemNotaConEnt.SstrNombreCampoBd}
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri, lstrCampSelPri,
                lstrTablaSec, lstrCampSelSec, lstrCamRelPri, LstrCampRelSec, lstrOrden, lstrfiltro,
                lstrCampGrupo)
        lstrExpSql = lstrExpSql.Replace("AS Valor)", "AS Valor")
        Return lstrExpSql
    End Function
    Private Function FstrFiltroFrasElec() As String
        Dim lstrFiltro = String.Empty
        If GobjParametros.BlnEFacAutorizado Then
            lstrFiltro = " AND IF(" & ClsIdEstadoEDocEnt.SstrNombreCampoBd & " <> " &
                    EnuEstadoEDoc.EnuNoEDoc & ", " &
                    ClsIdEstadoEDocEnt.SstrNombreCampoBd & " >= 4, " &
                    ClsIdEstadoEDocEnt.SstrNombreCampoBd & " >= 0)"
        End If
        Return lstrFiltro
    End Function
    Private Function FstrExpSqlDetalleFac(astrPrefFacturas As String, astrIdFacIni As String,
            astrIdFacFin As String) As String
        With MstbExpresionSql
            .Clear.Append("SELECT ").Append(ClsPrefijo_ItemFactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_ItemFactEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdItemFacturaShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsDetalle_ItemFactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsTarifaIva_ItemFactDbl.SstrNombreCampoBd).Append(" * 100 AS TarifaIva, ")
            .Append(ClsValor_ItemFactDec.SstrNombreCampoBd).Append(" FROM ")
            .Append(ClsItemFactura.SstrNombreTabla)
            .Append(" WHERE ").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND ").Append(MstrCampoBdCentroUtil).Append(" = ")
            .Append(GshrIdCentroUtil.ToString).Append(" AND ")
            .Append(ClsPrefijo_ItemFactStr.SstrNombreCampoBd).Append(" = '").Append(astrPrefFacturas)
            .Append("' AND ").Append(ClsIdFactura_ItemFactEnt.SstrNombreCampoBd)
            .Append(" >= ").Append(astrIdFacIni).Append(" AND ")
            .Append(ClsIdFactura_ItemFactEnt.SstrNombreCampoBd).Append(" <= ").Append(astrIdFacFin)
            .Append(" ORDER BY ").Append(ClsPrefijo_ItemFactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_ItemFactEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdItemFacturaShr.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpSqlServicios() As String
        Dim lshrIdAnoActual As Short = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
        Dim lstrCamSel As String() = {ClsIdAno_ServicioShr.SstrNombreCampoBd,
                ClsIdServicioShr.SstrNombreCampoBd,
                ClsConceptoServicioStr.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{ClsIdAno_ServicioShr.SstrNombreCampoBd, "DESC"},
                {ClsIdServicioShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion
        Dim lstrExpSqlAgrSer = ClsPanoramaDat.FstrConstruyaExpSqlSelect(
                ClsServicio.SstrNombreTabla, lstrCamSel, lstrOrden, lstrFiltro,
                Array.Empty(Of String))
        Return lstrExpSqlAgrSer
    End Function
    Private Function FstrExpSqlEstadoCuenta(astrPrefFacturas As String, astrIdFacIni As String,
            astrIdFacFin As String) As String
        With MstbExpresionSql
            .Clear.Append("SELECT ").Append(ClsIdEstadoCuentaEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijoFac_EstadoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsDeudaCapitalDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsDeudaIntMoraDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsAntPorAplDec.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsEstadoCuenta.SstrNombreTabla)
            .Append(" WHERE ").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND ").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil.ToString)
            .Append(" AND ").Append(ClsPrefijoFac_EstadoStr.SstrNombreCampoBd).Append(" = '")
            .Append(astrPrefFacturas).Append("' AND ").Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd)
            .Append(" >= ").Append(astrIdFacIni).Append(" AND ")
            .Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd).Append(" <= ").Append(astrIdFacFin)
            .Append(" ORDER BY ").Append(ClsIdEstadoCuentaEnt.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpDetEstadoCuenta(astrPrefFacturas As String, astrIdFacIni As String,
            astrIdFacFin As String) As String
        Dim lshrIdAnoAct As Short = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
        With MstbExpresionSql
            .Clear()
            .Append("SELECT ").Append(ClsPrefijoFac_EstadoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd).Append(", IF(")
            .Append(ClsIdAno_ItemFactEstadoShr.SstrNombreCampoBd).Append(">0,").Append(lshrIdAnoAct)
            .Append(", 0) AS IdAnos, ").Append(ClsIdServicioItemFac_EstadoShr.SstrNombreCampoBd)
            .Append(", SUM(F.").Append(ClsDeudaCap_ItFacEstDec.SstrNombreCampoBd)
            .Append(") AS DeudaCap, ").Append("SUM(F.").Append(ClsDeudaIntMora_ItFacEstDec.
            SstrNombreCampoBd).Append(" - ").Append(ClsDeudaIntMes_ItFacEstDec.SstrNombreCampoBd)
            .Append(") AS ValorInt, SUM(").Append(ClsDeudaIntMes_ItFacEstDec.SstrNombreCampoBd)
            .Append(") AS IntMes ").Append(" FROM ").Append(ClsFacturaEstado.SstrNombreTabla)
            .Append(" AS F INNER JOIN ").Append(ClsEstadoCuenta.SstrNombreTabla).Append(" AS E ON F.")
            .Append(MstrCampoBdCarpeta).Append(" = E.").Append(MstrCampoBdCarpeta).Append(" AND F.")
            .Append(MstrCampoBdCentroUtil).Append(" = E.").Append(MstrCampoBdCentroUtil)
            .Append(" AND F.").Append(ClsIdEstado_FactEstadoEnt.SstrNombreCampoBd).Append(" = E.")
            .Append(ClsIdEstadoCuentaEnt.SstrNombreCampoBd).Append(" WHERE F.")
            .Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND F.").Append(MstrCampoBdCentroUtil).Append(" = ")
            .Append(GshrIdCentroUtil.ToString).Append(" AND ")
            .Append(ClsPrefijoFac_EstadoStr.SstrNombreCampoBd).Append(" = '")
            .Append(astrPrefFacturas).Append("' AND ")
            .Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd).Append(" BETWEEN ")
            .Append(astrIdFacIni).Append(" AND ").Append(astrIdFacFin).Append(" GROUP BY ")
            .Append(ClsPrefijoFac_EstadoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd).Append(", ").Append("IdAnos")
            .Append(", ").Append(ClsIdServicioItemFac_EstadoShr.SstrNombreCampoBd)
            .Append(" ORDER BY ").Append(ClsPrefijoFac_EstadoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd).Append(", IdAnos DESC ,")
            .Append(ClsIdServicioItemFac_EstadoShr.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString()
    End Function
    Private Function FstrExpSqlAntApli(astrPrefFacturas As String, astrIdFacIni As String,
            astrIdFacFin As String) As String
        With MstbExpresionSql
            .Clear().Append("SELECT ").Append(ClsPrefijoFact_ItemNotaConStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_ItemNotaConEnt.SstrNombreCampoBd).Append(", ")
            .Append("SUM(").Append(ClsValor_ItemNotaConDec.SstrNombreCampoBd)
            .Append(") AS ").Append(ClsValor_ItemNotaConDec.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsItemNotaCon.SstrNombreTabla)
            .Append(" WHERE ").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND ").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil.ToString)
            .Append(" AND ").Append(ClsPrefijoFact_ItemNotaConStr.SstrNombreCampoBd).Append(" = '")
            .Append(astrPrefFacturas).Append("' AND ").Append(ClsIdFactura_ItemNotaConEnt.SstrNombreCampoBd)
            .Append(" >= ").Append(astrIdFacIni).Append(" AND ")
            .Append(ClsIdFactura_ItemNotaConEnt.SstrNombreCampoBd).Append(" <= ").Append(astrIdFacFin)
            .Append(" GROUP BY ").Append(ClsPrefijoFact_ItemNotaConStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_ItemNotaConEnt.SstrNombreCampoBd)
            .Append(" ORDER BY ").Append(ClsIdFactura_ItemNotaConEnt.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Sub SComplementeDtbFacturas(adsDataSetFacturas As DataSet)
        Dim ldtmFechaFac = GCDTMFECHANULA, ldblTasaMora = 0.0
        If Not IsNothing(adsDataSetFacturas) Then
            Dim ldtbAnticiposApl As DataTable = adsDataSetFacturas.Tables("OriAnticipos")
            Dim ldtbFacturas As DataTable = adsDataSetFacturas.Tables("OriFactura")
            Dim ldtbNombresPreAgr As DataTable = adsDataSetFacturas.Tables("OriPredioAgr")
            Dim ldclQR As New DataColumn("QR", System.Type.GetType("System.Byte[]"))
            ldtbFacturas.Columns.Add(ldclQR)
            Dim lobjValorLlave As Object()
            Dim lobjFra As New ClsFactura()
            Dim lColDrwFacExluir As New Collection
            Dim ldecValorIva As Decimal, ldecValorBaseIva As Decimal
            For Each ldrwFra As DataRow In ldtbFacturas.Rows
                Dim lstrPrefFra As String = ClsPanorama.FobjValorCampo(ldrwFra(ClsPrefijo_FactStr.SstrNombreCampoBd),
                            EnuTipoValor.enuString)
                Dim lentIdFact As Integer = ClsPanorama.FobjValorCampo(ldrwFra(ClsIdFacturaEnt.SstrNombreCampoBd),
                            EnuTipoValor.enuInteger)
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFra, lentIdFact}
                lobjFra.SAbra(lobjValorLlave)
                ldrwFra("NomPreAgr") = lobjFra.StrNombrePredioAgr
                ldecValorIva = lobjFra.FdecIvaFactura
                ldecValorBaseIva = lobjFra.FdecBaseIvaCapital
                ldrwFra("ValorIva") = ldecValorIva
                ldrwFra("ValorBaseIva") = ldecValorBaseIva
                Dim ldecTotalAPagar As Decimal = lobjFra.DecTotalAPagar
                ldrwFra("TotalAPagar") = ldecTotalAPagar
                ldrwFra("FechaPlazo") = lobjFra.DtmFechaPlazo
                ldrwFra("QR") = lobjFra.FbytQRFact
                ' Tasa mora
                If ldtmFechaFac <> lobjFra.ObjFechaFacturaDtm.ObjValorPro Then
                    ldtmFechaFac = lobjFra.ObjFechaFacturaDtm.ObjValorPro
                    ldblTasaMora = GobjParametros.FdblTasaMoraFecha(ldtmFechaFac)
                End If
                ldrwFra("TasaMora") = ldblTasaMora
                'Anticipos
                Dim lstrFiltro = ClsPrefijoFact_ItemNotaConStr.SstrNombreCampoBd & " = '" & lstrPrefFra &
                            "' AND " & ClsIdFactura_ItemNotaConEnt.SstrNombreCampoBd & " = " & lentIdFact
                Dim ldrwAnticipos As DataRow() = ldtbAnticiposApl.Select(lstrFiltro)
                Dim ldrwAnticipo As DataRow
                Dim ldecVlrAnticipo = 0D
                If ldrwAnticipos.Length > 0 Then
                    For i = 0 To ldrwAnticipos.Length - 1
                        ldrwAnticipo = ldrwAnticipos(i)
                        ldecVlrAnticipo += ClsPanorama.FobjValorCampo(ldrwAnticipo(
                                ClsValor_ItemNotaConDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                    Next
                End If
                ldrwFra("AntiApli") = ldecVlrAnticipo
            Next
            adsDataSetFacturas.Tables.Remove("OriAnticipos")
        End If
    End Sub
    Private Sub SGenereDataSetFactsFechas(adsFactura As DataSet, ablnExcluirEnvMail As Boolean)
        MstbExpresionSql.Clear()
        Dim lstrSqlFactura = FstrExpSqlFactsFechas(ablnExcluirEnvMail)
        Dim lstrSqlDetalleFac = FstrExpSqlDetalleFacsFechas()
        Dim lstrSqlEstadoCuenta = FstrExpSqlEstadoCtaFechas()
        Dim lstrSqlDetEstadoCta = FstrExpDetEstadoCtaFechas()
        Dim lstrSqlAnticiposApl = FstrExpSqlAntApliFechas()
        Dim lstrSqlServicios = FstrExpSqlServicios()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlFactura)
        lcolExpresionesSql.Add(lstrSqlDetalleFac)
        lcolExpresionesSql.Add(lstrSqlEstadoCuenta)
        lcolExpresionesSql.Add(lstrSqlDetEstadoCta)
        lcolExpresionesSql.Add(lstrSqlAnticiposApl)
        lcolExpresionesSql.Add(lstrSqlServicios)
        lcolNombresTablas.Add("OriFactura")
        lcolNombresTablas.Add("OriItemsFactura")
        lcolNombresTablas.Add("OriEstadosCuenta")
        lcolNombresTablas.Add("OriDetEstadosCta")
        lcolNombresTablas.Add("OriAnticipos")
        lcolNombresTablas.Add("OriServicios")
        GobjPanDat.SdsDataSet(adsFactura, lcolExpresionesSql, lcolNombresTablas)
        SComplementeDtbFacturas(adsFactura)
        adsFactura.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "Facturas" & ".XML"
        'adsFactura.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
#End Region
#Region "Facuras automáticas del Mes"
    Private Sub SGenereDataSetFacsAutoMes(adsFacsAutoMes As DataSet)
        If IsNothing(ObjParRepDocs) Then
            Throw New ErrorInesperadoPanLException("Sin Parametros para el Reporte de Facturas")
        End If
        MstbExpresionSql.Clear()
        Dim lstrSqlFactura = FstrExpSqlFacsAutoMes()
        Dim lstrSqlDetalleFac = FstrExpSqlDetFacAutoMes()
        Dim lstrSqlEstadoCuenta = FstrExpSqlEstCtaFacAutoMes()
        Dim lstrSqlDetEstadoCta = FstrExpDetEstCtaFacsAutoMes()
        Dim lstrSqlAnticiposApl = FstrExpSqlAntApliFacsAutoMes()
        Dim lstrSqlServicios = FstrExpSqlServicios()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlFactura)
        lcolExpresionesSql.Add(lstrSqlDetalleFac)
        lcolExpresionesSql.Add(lstrSqlEstadoCuenta)
        lcolExpresionesSql.Add(lstrSqlDetEstadoCta)
        lcolExpresionesSql.Add(lstrSqlAnticiposApl)
        lcolExpresionesSql.Add(lstrSqlServicios)
        lcolNombresTablas.Add("OriFactura")
        lcolNombresTablas.Add("OriItemsFactura")
        lcolNombresTablas.Add("OriEstadosCuenta")
        lcolNombresTablas.Add("OriDetEstadosCta")
        lcolNombresTablas.Add("OriAnticipos")
        lcolNombresTablas.Add("OriServicios")
        GobjPanDat.SdsDataSet(adsFacsAutoMes, lcolExpresionesSql, lcolNombresTablas)
        SComplementeDtbFacturas(adsFacsAutoMes)
        adsFacsAutoMes.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "Facturas" & ".XML"
        'adsFacsAutoMes.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlFacsAutoMes() As String
        Dim ldtmFechaIni = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim lstrFechaIni = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaIni) & "'"
        Dim lblnExcluirEnvMail = ObjParRepDocs.BlnExcluirFacEnvEmail
        With MstbExpresionSql
            .Clear.Append("SELECT F.").Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFacturaEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(" AS FechaPlazo").Append(", ")
            .Append(ClsFechaGraciaDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaDctoProntoPagoDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd).Append(", ")
            .Append("'' AS NomPreAgr").Append(", ")
            .Append(ClsReferenciaPago_FacStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPieFacturaUno_FactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPieFacturaDos_FactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsAnuladoBln.SstrNombreCampoBd).Append(", ")
            .Append(ClsDctoProntoPago_FacDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsValor_FactDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsCUFEStr.SstrNombreCampoBd).Append(", 0.00 AS ValorIva")
            .Append(", 0.00 AS ValorBaseIva").Append(", 0.00 AS AntiApli")
            .Append(", 0.00 AS TotalAPagar").Append(", 0.00 AS TasaMora")
            .Append(" FROM ").Append(ClsFactura.SstrNombreTabla)
            .Append(" AS F INNER JOIN ").Append(ClsCliente.SstrNombreTabla).Append(" AS C ON F.")
            .Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(" = ").Append("C.")
            .Append(ClsIdClienteDbl.SstrNombreCampoBd).Append(" AND F.").Append(MstrCampoBdCarpeta)
            .Append(" = C.").Append(MstrCampoBdCarpeta).Append(" AND F.")
            .Append(MstrCampoBdCentroUtil).Append(" = C.").Append(MstrCampoBdCentroUtil)
            .Append(" WHERE F.").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND F.").Append(MstrCampoBdCentroUtil).Append(" = ")
            .Append(GshrIdCentroUtil.ToString).Append(" AND F.")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaIni).Append(" AND F.")
            .Append(ClsIdModoFacturacionByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuModoFacturacionDef.EnuSistema)
            If lblnExcluirEnvMail Then
                .Append(" AND F.").Append(ClsEnviadaMailBln.SstrNombreCampoBd)
                .Append(" = FALSE")
            End If
            .Append(FstrFiltroFrasElec)
            .Append(" ORDER BY ").Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFacturaEnt.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpSqlDetFacAutoMes() As String
        Dim ldtmFechaIni = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim lstrFechaIni = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaIni) & "'"
        With MstbExpresionSql
            .Clear.Append("SELECT IT.").Append(ClsPrefijo_ItemFactStr.SstrNombreCampoBd)
            .Append(", ").Append("IT." & ClsIdFactura_ItemFactEnt.SstrNombreCampoBd)
            .Append(", ")
            .Append(ClsIdItemFacturaShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsDetalle_ItemFactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsTarifaIva_ItemFactDbl.SstrNombreCampoBd).Append(" * 100 AS TarifaIva, ")
            .Append(ClsValor_ItemFactDec.SstrNombreCampoBd).Append(" FROM ")
            .Append(ClsItemFactura.SstrNombreTabla).Append(" AS IT INNER JOIN ")
            .Append(ClsFactura.SstrNombreTabla).Append(" AS F ON IT.")
            .Append(MstrCampoBdCarpeta).Append(" = F.").Append(MstrCampoBdCarpeta).Append(" AND IT.")
            .Append(MstrCampoBdCentroUtil).Append(" = F.").Append(MstrCampoBdCentroUtil)
            .Append(" AND IT.").Append(ClsPrefijo_ItemFactStr.SstrNombreCampoBd).Append(" = F.")
            .Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(" AND IT.")
            .Append(ClsIdFactura_ItemFactEnt.SstrNombreCampoBd).Append(" = F.")
            .Append(ClsIdFacturaEnt.SstrNombreCampoBd)
            .Append(" WHERE IT.").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND IT.").Append(MstrCampoBdCentroUtil).Append(" = ")
            .Append(GshrIdCentroUtil.ToString).Append(" AND ")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(" >= ").Append(lstrFechaIni)
            .Append(" ORDER BY ").Append(ClsPrefijo_ItemFactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_ItemFactEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdItemFacturaShr.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpSqlEstCtaFacAutoMes() As String
        Dim ldtmFechaIni = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        ldtmFechaIni = ldtmFechaIni.AddDays(-1)
        Dim lstrFechaIni = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaIni) & "'"
        With MstbExpresionSql
            .Clear.Append("SELECT ").Append(ClsIdEstadoCuentaEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijoFac_EstadoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsDeudaCapitalDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsDeudaIntMoraDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsAntPorAplDec.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsEstadoCuenta.SstrNombreTabla)
            .Append(" WHERE ").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND ").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil.ToString)
            .Append(" AND ").Append(ClsFechaEstadoDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaIni).Append(" ORDER BY ").Append(ClsIdEstadoCuentaEnt.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpDetEstCtaFacsAutoMes() As String
        Dim lshrIdAnoAct As Short = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
        Dim ldtmFechaIni = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        ldtmFechaIni = ldtmFechaIni.AddDays(-1)
        Dim lstrFechaIni = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaIni) & "'"
        With MstbExpresionSql
            .Clear()
            .Append("SELECT ").Append(ClsPrefijoFac_EstadoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd).Append(", ").Append("IF(")
            .Append(ClsIdAno_ItemFactEstadoShr.SstrNombreCampoBd).Append(">0,").Append(lshrIdAnoAct)
            .Append(", 0) AS IdAnos").Append(", ")
            .Append(ClsIdServicioItemFac_EstadoShr.SstrNombreCampoBd).Append(", ")
            .Append("SUM(F.").Append(ClsDeudaCap_ItFacEstDec.SstrNombreCampoBd).Append(") AS DeudaCap,")
            .Append("SUM(F.").Append(ClsDeudaIntMora_ItFacEstDec.SstrNombreCampoBd).Append(" - ")
            .Append(ClsDeudaIntMes_ItFacEstDec.SstrNombreCampoBd).Append(") AS ValorInt, SUM(")
            .Append(ClsDeudaIntMes_ItFacEstDec.SstrNombreCampoBd)
            .Append(") AS IntMes ").Append(" FROM ").Append(ClsFacturaEstado.SstrNombreTabla)
            .Append(" AS F INNER JOIN ").Append(ClsEstadoCuenta.SstrNombreTabla).Append(" AS E ON F.")
            .Append(MstrCampoBdCarpeta).Append(" = E.").Append(MstrCampoBdCarpeta).Append(" AND F.")
            .Append(MstrCampoBdCentroUtil).Append(" = E.").Append(MstrCampoBdCentroUtil)
            .Append(" AND F.").Append(ClsIdEstado_FactEstadoEnt.SstrNombreCampoBd).Append(" = E.")
            .Append(ClsIdEstadoCuentaEnt.SstrNombreCampoBd).Append(" WHERE F.")
            .Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString).Append(" AND F.")
            .Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil.ToString)
            .Append(" AND ").Append(ClsFechaEstadoDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaIni).Append(" GROUP BY ").Append(ClsPrefijoFac_EstadoStr.
             SstrNombreCampoBd).Append(", ").Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd)
            .Append(", ").Append("IdAnos").Append(", ").Append(ClsIdServicioItemFac_EstadoShr.
             SstrNombreCampoBd).Append(" ORDER BY ")
            .Append(ClsPrefijoFac_EstadoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString()
    End Function
    Private Function FstrExpSqlAntApliFacsAutoMes() As String
        Dim ldtmFechaIni = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim lstrFechaIni = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaIni) & "'"
        With MstbExpresionSql
            .Clear().Append("SELECT ").Append(ClsPrefijoFact_ItemNotaConStr.SstrNombreCampoBd)
            .Append(", ").Append(ClsIdFactura_ItemNotaConEnt.SstrNombreCampoBd).Append(", ")
            .Append("SUM(I.").Append(ClsValor_ItemNotaConDec.SstrNombreCampoBd)
            .Append(") AS ").Append(ClsValor_ItemNotaConDec.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsItemNotaCon.SstrNombreTabla).Append(" AS I INNER JOIN ")
            .Append(ClsNotaCon.SstrNombreTabla).Append(" AS NC ON I.")
            .Append(ClsPrefijo_NotaConStr.SstrNombreCampoBd).Append(" = NC.")
            .Append(ClsPrefijo_NotaConStr.SstrNombreCampoBd).Append(" AND I.")
            .Append(ClsIdNotaCon_ItemNotaConEnt.SstrNombreCampoBd).Append(" = NC.")
            .Append(ClsIdNotaConEnt.SstrNombreCampoBd)
            .Append(" WHERE I.").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND I.").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil.ToString)
            .Append(" AND ").Append(ClsFecha_NotaConDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaIni)
            .Append(" GROUP BY ").Append(ClsPrefijoFact_ItemNotaConStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_ItemNotaConEnt.SstrNombreCampoBd)
            .Append(" ORDER BY ").Append(ClsPrefijoFact_ItemNotaConStr.SstrNombreCampoBd)
            .Append(", ").Append(ClsIdFactura_ItemNotaConEnt.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
#End Region
#Region "Recibo de Caja"
    Private Sub SGenereDataSetRecCaja(adsRecCaja As DataSet, ablnFirma As Boolean)
        If IsNothing(ObjParRepDocs) Then
            Throw New ErrorInesperadoPanLException("Sin Parametros para el Reporte de Recibos de Caja")
        End If
        Dim lstrSqlRecCaja = FstrExpSqlRC()
        Dim lstrSqlItemsDsctosRecCaja = FstrExpSqlItemDsctosRC()
        Dim lstrSqlMediosPago = FstrExpSqlMedPago()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlRecCaja)
        lcolExpresionesSql.Add(lstrSqlItemsDsctosRecCaja)
        lcolExpresionesSql.Add(lstrSqlMediosPago)
        lcolNombresTablas.Add("OriRecCaja")
        lcolNombresTablas.Add("OriItemsDsctosRecCaja")
        lcolNombresTablas.Add("OriMediosPago")
        GobjPanDat.SdsDataSet(adsRecCaja, lcolExpresionesSql, lcolNombresTablas)
        adsRecCaja.Tables.Add(FdtbInfItemsRC)
        If ablnFirma Then
            adsRecCaja.Tables.Add(FdtbCentroUtilidad)
        Else
            adsRecCaja.Tables.Add(FdtbCentroUtilidadSinFirma)
        End If
        SComplementeDtbRecCaja(adsRecCaja)
        SComplementeDtblItemsDsctosRecCaja(adsRecCaja)
        SComplementeDsMedPago(adsRecCaja)
        'Dim lstrNomArch = GstrTrayDatPrg & "ReciboCaja" & ".XML"
        'adsRecCaja.WriteXml(lstrNomArch, XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlRC() As String
        Dim lstrPrefRecCaja = ObjParRepDocs.StrPrefijoDocsRep
        Dim lstrIdRecCajaIni = ObjParRepDocs.EntIdDocInicial.ToString
        Dim lstrIdRecCajaFin = ObjParRepDocs.EntIdDocFinal.ToString
        With MstbExpresionSql
            .Clear()
            .Append("SELECT RC.").Append(ClsIdCliente_RecDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdRecCajaEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaRecDtm.SstrNombreCampoBd).Append(", ")
            .Append("RC.").Append(ClsFechaCreacionDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd).Append(", ")
            .Append("'' as NomPreAgr").Append(", ")
            .Append(ClsSaldo_RecDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsValorAnticipoDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsValor_RecDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsAnuladoBln.SstrNombreCampoBd).Append(", ")
            .Append("RC.").Append(ClsComentario_RecStr.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsReciboCaja.SstrNombreTabla)
            .Append(" AS RC INNER JOIN ").Append(ClsCliente.SstrNombreTabla).Append(" AS C ON RC.")
            .Append(MstrCampoBdCarpeta).Append(" = C.").Append(MstrCampoBdCarpeta).Append(" AND RC.")
            .Append(MstrCampoBdCentroUtil).Append(" = C.").Append(MstrCampoBdCentroUtil)
            .Append(" AND RC.").Append(ClsIdCliente_RecDbl.SstrNombreCampoBd).Append(" = C.")
            .Append(ClsIdClienteDbl.SstrNombreCampoBd)
            .Append(" WHERE RC.").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND RC.").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil)
            .Append(" AND ").Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(" = '").Append(lstrPrefRecCaja)
            .Append("' AND RC.").Append(ClsIdRecCajaEnt.SstrNombreCampoBd).Append(" >= ").Append(lstrIdRecCajaIni)
            .Append(" AND RC.").Append(ClsIdRecCajaEnt.SstrNombreCampoBd).Append(" <= ").Append(lstrIdRecCajaFin)
            .Append(" ORDER BY ").Append(ClsIdRecCajaEnt.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpSqlItemDsctosRC() As String
        Dim lstrPrefRecCaja = ObjParRepDocs.StrPrefijoDocsRep
        Dim lstrIdRecCajaIni = ObjParRepDocs.EntIdDocInicial.ToString
        Dim lstrIdRecCajaFin = ObjParRepDocs.EntIdDocFinal.ToString
        With MstbExpresionSql
            .Clear()
            .Append("SELECT ").Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdTipoItemRecByt.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijoFact_ItemRecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_ItemRecEnt.SstrNombreCampoBd)
            .Append(", SUM(").Append(ClsValor_ItemRecDec.SstrNombreCampoBd).Append("), '' as Detalle")
            .Append(" FROM ").Append(ClsItemRecCaja.SstrNombreTabla)
            .Append(" WHERE ").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND ").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil.ToString)
            .Append(" AND ").Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(" = '")
            .Append(lstrPrefRecCaja).Append("' AND ").Append(ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd)
            .Append(" >= ").Append(lstrIdRecCajaIni).Append(" AND ")
            .Append(ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd).Append(" <= ").Append(lstrIdRecCajaFin)
            .Append(" AND ").Append(ClsIdTipoItemRecByt.SstrNombreCampoBd).Append(" >= ")
            .Append(EnuTipoItemRecCajaDef.EnuDsctoIntMora)
            .Append(" GROUP BY ").Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdTipoItemRecByt.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijoFact_ItemRecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_ItemRecEnt.SstrNombreCampoBd).Append(", Detalle ")
            .Append(" ORDER BY ").Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd)
        End With
        Dim lstrSqlItemsDsctosRecCaja = MstbExpresionSql.ToString
        Return lstrSqlItemsDsctosRecCaja
    End Function
    Private Function FstrExpSqlMedPago() As String
        Dim lstrPrefRecCaja = ObjParRepDocs.StrPrefijoDocsRep
        Dim lstrIdRecCajaIni = ObjParRepDocs.EntIdDocInicial.ToString
        Dim lstrIdRecCajaFin = ObjParRepDocs.EntIdDocFinal.ToString
        With MstbExpresionSql
            .Clear()
            .Append("SELECT *, '' as Detalle").Append(" FROM ").Append(ClsMedioPago.SstrNombreTabla)
            .Append(" WHERE ").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND ").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil.ToString)
            .Append(" AND ").Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(" = '")
            .Append(lstrPrefRecCaja).Append("' AND ")
            .Append(ClsIdRecCaja_MedPagoEnt.SstrNombreCampoBd).Append(" >= ").Append(lstrIdRecCajaIni)
            .Append(" AND ").Append(ClsIdRecCaja_MedPagoEnt.SstrNombreCampoBd).Append(" <= ")
            .Append(lstrIdRecCajaFin).Append(" ORDER BY ")
            .Append(ClsIdRecCaja_MedPagoEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsOrdinal_MedPagoShr.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString()
    End Function
    Private Shared Sub SComplementeDtbRecCaja(adsRecCaja As DataSet)
        If Not IsNothing(adsRecCaja) Then
            Dim lstrPredioAgr As String
            Dim lstrPrediosAgr As String()
            Dim lstrIdPredio As String, lstrNomPre As String
            Dim ldtbRec As DataTable = adsRecCaja.Tables("OriRecCaja")
            Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
            For Each ldrwRC As DataRow In ldtbRec.Rows
                lstrPredioAgr = String.Empty
                lstrPrediosAgr = ClsPanorama.FobjValorCampo(ldrwRC(ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd),
                            EnuTipoValor.enuString).ToString.Split(",")
                For i As Integer = 0 To lstrPrediosAgr.Length - 1
                    lstrIdPredio = lstrPrediosAgr(i)
                    If String.IsNullOrEmpty(lstrIdPredio) Then lstrIdPredio = GCSTRSINPA
                    lstrPredioAgr += lstrIdPredio & " / "
                Next i
                If lstrPredioAgr.EndsWith(" / ") Then
                    lstrPredioAgr = lstrPredioAgr.Substring(0, lstrPredioAgr.Length - 3)
                End If
                If lstrPrediosAgr.Length > 1 Then
                    lstrNomPre = lstrPredioAgr
                Else
                    lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPredioAgr})
                    If lobjPredio.BlnExiste Then
                        lstrNomPre = lobjPredio.ObjNombrePredioStr.ToString()
                    Else
                        lstrNomPre = lstrPredioAgr
                    End If
                End If
                ldrwRC("NomPreAgr") = lstrNomPre
                ldrwRC(ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd) = lstrPredioAgr
            Next
        End If
    End Sub
    Private Function FdtbInfItemsRC()
        Dim lstrPrefRecCaja = ObjParRepDocs.StrPrefijoDocsRep
        Dim lstrIdRecCajaIni = ObjParRepDocs.EntIdDocInicial.ToString
        Dim lstrIdRecCajaFin = ObjParRepDocs.EntIdDocFinal.ToString
        Dim lstrPreRec As String, lentIdRec As Integer, lstrNroRec As String, lstrNroRecAct = String.Empty
        Dim lstrFiltro As String
        Dim ldtbInfItemsRC = ClsReciboCaja.FdtbInfItems()
        Dim ldtbItemsRC = ClsReciboCaja.FdtbInfItemsRecCaja(lstrPrefRecCaja, lstrIdRecCajaIni,
                lstrIdRecCajaFin)
        SPuebleInfPagoCap(ldtbInfItemsRC, ldtbItemsRC)
        For Each ldrwItemRec As DataRow In ldtbItemsRC.Rows
            lstrPreRec = ClsPanorama.FobjValorCampo(ldrwItemRec(ClsPrefijo_RecStr.SstrNombreCampoBd),
                            EnuTipoValor.enuString)
            lentIdRec = ClsPanorama.FobjValorCampo(ldrwItemRec(ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd),
                            EnuTipoValor.enuInteger)
            lstrNroRec = ClsPanorama.FstrNumeroDcto(lstrPrefRecCaja, lentIdRec)
            If lstrNroRecAct <> lstrNroRec Then
                lstrFiltro = ClsPrefijo_RecStr.SstrNombreCampoBd & " = '" & lstrPreRec & "' AND " &
                    ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd & " = " & lentIdRec
                SAdicioneItemsRC(ldtbItemsRC.Select(lstrFiltro), ldtbInfItemsRC)
                lstrNroRecAct = lstrNroRec
            End If
        Next
        Return ldtbInfItemsRC
    End Function
    Private Sub SAdicioneItemsRC(adrwsItemsRec As DataRow(), adtbInfItemsRec As DataTable)
        Const lstrFactMora = "Todas"
        Dim lstrPreFac = String.Empty, lentIdFac = 0, lshrIdItemFac = 0S, lstrFac = String.Empty
        Dim lstrPreRec = String.Empty, lentIdRec = 0, lstrNroRec = String.Empty, lstrNroRecAct = String.Empty
        Dim lenuTipoItemRec As EnuTipoItemRecCajaDef = EnuTipoItemRecCajaDef.None
        Dim ldrwNew As DataRow = Nothing
        Dim ldecValor(8) As Decimal, lstrDetalle(8) As String, lentOrdTipoItemRec(8) As Integer
        Dim lentOrdinal = 0
        For Each ldrwItemRec As DataRow In adrwsItemsRec
            lenuTipoItemRec = ClsPanorama.FobjValorCampo(ldrwItemRec(
                    ClsIdTipoItemRecByt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
            lstrPreRec = ClsPanorama.FobjValorCampo(ldrwItemRec(ClsPrefijo_RecStr.SstrNombreCampoBd),
                            EnuTipoValor.enuString)
            lentIdRec = ClsPanorama.FobjValorCampo(ldrwItemRec(ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd),
                            EnuTipoValor.enuInteger)
            lstrPreFac = ClsPanorama.FobjValorCampo(ldrwItemRec(ClsPrefijoFact_ItemRecStr.SstrNombreCampoBd),
                            EnuTipoValor.enuString)
            lentIdFac = ClsPanorama.FobjValorCampo(ldrwItemRec(ClsIdFactura_ItemRecEnt.SstrNombreCampoBd),
                            EnuTipoValor.enuInteger)
            lshrIdItemFac = ClsPanorama.FobjValorCampo(ldrwItemRec(ClsIdItemFac_ItemRecShr.SstrNombreCampoBd),
                            EnuTipoValor.enuShort)
            lstrFac = ClsPanorama.FstrNumeroDcto(lstrPreFac, lentIdFac)
            If lentOrdinal = 0 Then
                lentOrdinal = FentOrdinalTbl(adtbInfItemsRec, lstrPreRec, lentIdRec)
            End If
            Select Case lenuTipoItemRec
                Case EnuTipoItemRecCajaDef.EnuAbonoCapital
                    '
                Case EnuTipoItemRecCajaDef.EnuAbonoIntMora
                    lstrFac = "Todas"
                    lstrDetalle(0) = My.Resources.RCAbonoMor
                    ldecValor(0) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                            EnuTipoValor.enuDecimal)
                    If lentOrdTipoItemRec(0) = 0 Then
                        lentOrdinal += 1
                        lentOrdTipoItemRec(0) = lentOrdinal
                    End If
                Case EnuTipoItemRecCajaDef.EnuAnticipo
                    lstrDetalle(1) = My.Resources.RCAnticipo &
                                FstrComplementoSerAnticipo(lstrPreRec, lentIdRec)
                    ldecValor(1) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                            EnuTipoValor.enuDecimal)
                    If lentOrdTipoItemRec(1) = 0 Then
                        lentOrdinal += 1
                        lentOrdTipoItemRec(1) = lentOrdinal
                    End If
                Case EnuTipoItemRecCajaDef.EnuDsctoCapital
                    lstrDetalle(2) = My.Resources.RCDsctoCap
                    ldecValor(2) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                            EnuTipoValor.enuDecimal)
                    If lentOrdTipoItemRec(2) = 0 Then
                        lentOrdinal += 1
                        lentOrdTipoItemRec(2) = lentOrdinal
                    End If
                Case EnuTipoItemRecCajaDef.EnuDsctoPP
                    lstrDetalle(3) = My.Resources.RCDsctoPP
                    ldecValor(3) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                            EnuTipoValor.enuDecimal)
                    If lentOrdTipoItemRec(3) = 0 Then
                        lentOrdinal += 1
                        lentOrdTipoItemRec(3) = lentOrdinal
                    End If
                Case EnuTipoItemRecCajaDef.EnuDsctoIntMora
                    lstrDetalle(4) = My.Resources.RCDsctoMor
                    ldecValor(4) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                            EnuTipoValor.enuDecimal)
                    If lentOrdTipoItemRec(4) = 0 Then
                        lentOrdinal += 1
                        lentOrdTipoItemRec(4) = lentOrdinal
                    End If
                Case EnuTipoItemRecCajaDef.EnuReteFuente
                    lstrDetalle(5) = My.Resources.RCReteFte
                    ldecValor(5) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                            EnuTipoValor.enuDecimal)
                    If lentOrdTipoItemRec(5) = 0 Then
                        lentOrdinal += 1
                        lentOrdTipoItemRec(5) = lentOrdinal
                    End If
                Case EnuTipoItemRecCajaDef.EnuReteIca
                    lstrDetalle(6) = My.Resources.RCRetIca
                    ldecValor(6) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                            EnuTipoValor.enuDecimal)
                    If lentOrdTipoItemRec(6) = 0 Then
                        lentOrdinal += 1
                        lentOrdTipoItemRec(6) = lentOrdinal
                    End If
                Case EnuTipoItemRecCajaDef.EnuReteIva
                    lstrDetalle(7) = My.Resources.RCRetIva
                    ldecValor(7) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                            EnuTipoValor.enuDecimal)
                    If lentOrdTipoItemRec(7) = 0 Then
                        lentOrdinal += 1
                        lentOrdTipoItemRec(7) = lentOrdinal
                    End If
            End Select
        Next
        For i = 0 To 7
            If ldecValor(i) > 0 Then
                ldrwNew = adtbInfItemsRec.NewRow
                ldrwNew("Ordinal") = lentOrdTipoItemRec(i)
                ldrwNew("Prefijo") = lstrPreRec
                ldrwNew("IdReciboCaja") = lentIdRec
                If i = 0 Then
                    ldrwNew("NroFact") = lstrFactMora
                ElseIf i <> 1 Then
                    ldrwNew("NroFact") = lstrFac
                End If
                ldrwNew("Detalle") = lstrDetalle(i)
                ldrwNew("Valor") = ldecValor(i)
                adtbInfItemsRec.Rows.Add(ldrwNew)
            End If
        Next
    End Sub
    Private Function FentOrdinalTbl(adtbInfItemsRec As DataTable, astrPrefRec As String,
            aentIdRec As Integer) As Integer
        Dim lentOrdReg As Integer, lentOrdinal = 0
        Dim lstrfiltro = "Prefijo = '" & astrPrefRec & "' AND IdReciboCaja = " & aentIdRec
        Dim ldrwRegs As DataRow() = adtbInfItemsRec.Select(lstrfiltro)
        For Each ldrwReg As DataRow In ldrwRegs
            lentOrdReg = ClsPanorama.FobjValorCampo(ldrwReg("Ordinal"), EnuTipoValor.enuInteger)
            If lentOrdReg > lentOrdinal Then
                lentOrdinal = lentOrdReg
            End If
        Next
        Return lentOrdinal
    End Function
    Private Sub SPuebleInfPagoCap(adtbInfItemsRC As DataTable, adtbItemsRC As DataTable)
        Dim lshrIdItem = 0S, i = 0
        Dim lstrPreFac = String.Empty, lentIdFac = 0, lshrIdItemFac = 0S
        Dim lstrNroFact = String.Empty, lstrPreRec = String.Empty, lentIdRec = 0
        Dim lstrNroRec = String.Empty, lstrNroRecAct = String.Empty
        Dim lstrDetalle = "", lstrDetItem = String.Empty, ldecValor = 0D
        Dim lobjFact = New ClsFactura()
        Dim lobjItemFac As ClsItemFactura = Nothing
        Dim lobjValorLlave As Object() = Nothing
        Dim lenuTipoItemRec As EnuTipoItemRecCajaDef = EnuTipoItemRecCajaDef.None
        For Each ldrwItemRec As DataRow In adtbItemsRC.Rows
            lenuTipoItemRec = ClsPanorama.FobjValorCampo(ldrwItemRec(
                    "IdTipoItemRec"), EnuTipoValor.enuInteger)
            If lenuTipoItemRec = EnuTipoItemRecCajaDef.EnuAbonoCapital Then
                lstrPreRec = ClsPanorama.FobjValorCampo(ldrwItemRec(
                        ClsPrefijo_RecStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                lentIdRec = ClsPanorama.FobjValorCampo(ldrwItemRec(
                        ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd),
                            EnuTipoValor.enuInteger)
                lstrNroRec = ClsPanorama.FstrNumeroDcto(lstrPreRec, lentIdRec)
                If lenuTipoItemRec = EnuTipoItemRecCajaDef.EnuAbonoCapital Then
                    lstrPreFac = ClsPanorama.FobjValorCampo(ldrwItemRec(
                            ClsPrefijoFact_ItemRecStr.SstrNombreCampoBd),
                            EnuTipoValor.enuString)
                    lentIdFac = ClsPanorama.FobjValorCampo(ldrwItemRec(
                            ClsIdFactura_ItemRecEnt.SstrNombreCampoBd),
                            EnuTipoValor.enuInteger)
                    lshrIdItemFac = ClsPanorama.FobjValorCampo(ldrwItemRec(
                            ClsIdItemFac_ItemRecShr.SstrNombreCampoBd),
                            EnuTipoValor.enuShort)
                    lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPreFac,
                                lentIdFac}
                    lobjFact.SAbra(lobjValorLlave)
                End If
                If lshrIdItemFac > 0 Then
                    lobjItemFac = lobjFact.ColItemsFactura(lshrIdItemFac.ToString)
                    lstrDetItem = lobjItemFac.ObjDetalle_ItemFactStr.ObjValorNuevo
                Else
                    lstrDetItem = ""
                End If
                If lstrNroRecAct <> lstrNroRec Then
                    i = 1
                    lstrNroRecAct = lstrNroRec
                Else
                    i += 1
                End If
                If lenuTipoItemRec = EnuTipoItemRecCajaDef.EnuAbonoCapital Then
                    lstrDetalle = My.Resources.RCAbonoCap & " " & lstrDetItem
                Else
                    If lenuTipoItemRec = EnuTipoItemRecCajaDef.EnuAbonoIntMora Then
                        lstrDetalle = My.Resources.RCAbonoMor
                    Else
                        lstrDetalle = "Anticipo recibido"
                    End If
                End If
                ldecValor = ClsPanorama.FobjValorCampo(ldrwItemRec(
                    ClsValor_ItemRecDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
                Dim ldrwNew As DataRow = adtbInfItemsRC.NewRow
                Dim lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPreFac, lentIdFac)
                ldrwNew("Ordinal") = i
                ldrwNew("Prefijo") = lstrPreRec
                ldrwNew("IdReciboCaja") = lentIdRec
                ldrwNew("NroFact") = lstrNroFac
                ldrwNew("Detalle") = lstrDetalle
                ldrwNew("Valor") = ldecValor
                adtbInfItemsRC.Rows.Add(ldrwNew)
            End If
        Next
    End Sub
    Friend Function FstrComplementoSerAnticipo(astrPrefRC As String, aentIdRC As Integer)
        Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, astrPrefRC, aentIdRC}
        Dim lobjRC = New ClsReciboCaja()
        lobjRC.SAbra(lobjValorLlave)
        Dim lstrServicios As String = lobjRC.ObjAnticipo.ObjServicios_AntStr.ObjValorPro
        Dim lstrComp = " - Para aplicar a "
        If lstrServicios.Contains("A") Then
            lstrComp &= "todos los Servicios"
        ElseIf lstrServicios = "0" Then
            lstrComp &= "las Cuotas de Administración"
        ElseIf lstrServicios.Contains(",") Then
            lstrComp &= "todos los servicios relacionados en el Recibo de Caja!"
        Else
            Dim lstrKeySer = "0," & lstrServicios
            If GobjParametros.ColServiciosPer.Contains(lstrKeySer) Then
                Dim lobjSer As ClsServicio = GobjParametros.ColServiciosPer(lstrKeySer)
                lstrComp &= lobjSer.ObjNombreServicioStr.ToString()
            Else
                lstrComp = String.Empty
            End If
        End If
        Return lstrComp
    End Function
    Private Shared Sub SComplementeDtblItemsDsctosRecCaja(adsRecCaja As DataSet)
        If Not IsNothing(adsRecCaja) Then
            Dim ldtbItemsRec As DataTable = adsRecCaja.Tables("OriItemsDsctosRecCaja")
            Dim ldrwItems As DataRow() = ldtbItemsRec.Select()
            Dim lstrDetalle = String.Empty, lstrPreFac = String.Empty,
                    lentIdFac = 0, lstrNroFact = String.Empty
            For Each ldrwItem As DataRow In ldrwItems
                Dim lenuTipoItemRec As EnuTipoItemRecCajaDef = ClsPanorama.FobjValorCampo(ldrwItem("IdTipoItemRec"),
                            EnuTipoValor.enuInteger)
                lstrNroFact = String.Empty
                Select Case lenuTipoItemRec
                    Case EnuTipoItemRecCajaDef.EnuDsctoCapital
                        lstrDetalle = My.Resources.RCDsctoCap
                    Case EnuTipoItemRecCajaDef.EnuDsctoPP
                        lstrDetalle = My.Resources.RCDsctoPP
                    Case EnuTipoItemRecCajaDef.EnuDsctoIntMora
                        lstrDetalle = My.Resources.RCDsctoMor
                    Case EnuTipoItemRecCajaDef.EnuReteCree
                        lstrDetalle = My.Resources.RCRetCre
                    Case EnuTipoItemRecCajaDef.EnuReteFuente
                        lstrDetalle = My.Resources.RCReteFte
                    Case EnuTipoItemRecCajaDef.EnuReteIca
                        lstrDetalle = My.Resources.RCRetIca
                    Case EnuTipoItemRecCajaDef.EnuReteIva
                        lstrDetalle = My.Resources.RCRetIva
                End Select
                lstrPreFac = ClsPanorama.FobjValorCampo(ldrwItem(ClsPrefijoFact_ItemRecStr.SstrNombreCampoBd),
                                EnuTipoValor.enuString)
                lentIdFac = ClsPanorama.FobjValorCampo(ldrwItem(ClsIdFactura_ItemRecEnt.SstrNombreCampoBd),
                            EnuTipoValor.enuInteger)
                lstrNroFact = ClsPanorama.FstrNumeroDcto(lstrPreFac, lentIdFac)
                lstrDetalle &= " Factura " & lstrNroFact
                ldrwItem("Detalle") = lstrDetalle
            Next
        End If
    End Sub
    Private Shared Sub SComplementeDsMedPago(adsRecCaja As DataSet)
        Dim ldtbMedPago = adsRecCaja.Tables("OriMediosPago")
        Dim ldrwMedPagos() As DataRow = ldtbMedPago.Select
        If ldrwMedPagos.Count > 0 Then
            Dim lbytIdTipoMedPago As Byte, lstrNombreMedPago As String
            Dim lstrCuentaIngreso As String, lstrIdCtaCont As String
            For Each ldrwMedPago As DataRow In ldrwMedPagos
                lbytIdTipoMedPago = ClsPanorama.FobjValorCampo(ldrwMedPago(
                         ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd),
                         EnuTipoValor.enuByte)
                lstrIdCtaCont = ClsPanorama.FobjValorCampo(ldrwMedPago(
                        ClsIdCtaContabIngresoStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                lstrNombreMedPago = ClsOrionCop.FstrNombreDatoConstanteOri(EnuGrupoConstantesOriDef.EnuMediosPago,
                        lbytIdTipoMedPago)
                If IsNumeric(lstrNombreMedPago.Substring(0, 1)) Then
                    lstrNombreMedPago = lstrNombreMedPago.Substring(lstrNombreMedPago.IndexOf("-") + 1)
                End If
                lstrCuentaIngreso = ClsOrionCop.FstrCuentaBanco(lstrIdCtaCont)
                ldrwMedPago("Detalle") = lstrNombreMedPago & " a " & lstrCuentaIngreso
            Next
        End If
    End Sub
#End Region
#Region "Exportar Recibos Caja"
    Private Sub SGenereDataSetRecibosCaja(adsRecCaja As DataSet, ablnFirma As Boolean)
        Dim lstrSqlRecCaja = FstrExpSqlRecsC()
        Dim lstrSqlItemsDsctosRecCaja = FstrExpSqlItemDsctosRecsC()
        Dim lstrSqlMediosPago = FstrExpSqlMedPagoRecsC()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlRecCaja)
        lcolExpresionesSql.Add(lstrSqlItemsDsctosRecCaja)
        lcolExpresionesSql.Add(lstrSqlMediosPago)
        lcolNombresTablas.Add("OriRecCaja")
        lcolNombresTablas.Add("OriItemsDsctosRecCaja")
        lcolNombresTablas.Add("OriMediosPago")
        GobjPanDat.SdsDataSet(adsRecCaja, lcolExpresionesSql, lcolNombresTablas)
        adsRecCaja.Tables.Add(FdtbInfItemsRecsC)
        If ablnFirma Then
            adsRecCaja.Tables.Add(FdtbCentroUtilidad)
        Else
            adsRecCaja.Tables.Add(FdtbCentroUtilidadSinFirma)
        End If
        SComplementeDtbRecCaja(adsRecCaja)
        SComplementeDtblItemsDsctosRecCaja(adsRecCaja)
        SComplementeDsMedPago(adsRecCaja)
        'Dim lstrNomArch = GstrTrayDatPrg & "ReciboCaja" & ".XML"
        'adsRecCaja.WriteXml(lstrNomArch, XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlRecsC() As String
        With MstbExpresionSql
            .Clear()
            .Append("SELECT RC.").Append(ClsIdCliente_RecDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdRecCajaEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaRecDtm.SstrNombreCampoBd).Append(", ")
            .Append("RC.").Append(ClsFechaCreacionDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd).Append(", ")
            .Append("'' as NomPreAgr").Append(", ")
            .Append(ClsSaldo_RecDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsValorAnticipoDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsValor_RecDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsAnuladoBln.SstrNombreCampoBd).Append(", ")
            .Append("RC.").Append(ClsComentario_RecStr.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsReciboCaja.SstrNombreTabla)
            .Append(" AS RC INNER JOIN ").Append(ClsCliente.SstrNombreTabla).Append(" AS C ON RC.")
            .Append(MstrCampoBdCarpeta).Append(" = C.").Append(MstrCampoBdCarpeta).Append(" AND RC.")
            .Append(MstrCampoBdCentroUtil).Append(" = C.").Append(MstrCampoBdCentroUtil)
            .Append(" AND RC.").Append(ClsIdCliente_RecDbl.SstrNombreCampoBd).Append(" = C.")
            .Append(ClsIdClienteDbl.SstrNombreCampoBd)
            .Append(" WHERE RC.").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND RC.").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil)
            .Append(" AND ").Append(ClsFechaRecDtm.SstrNombreCampoBd).Append(" BETWEEN '")
            .Append(StrFechaDesde).Append("' AND '").Append(StrFechaHasta).Append("' ORDER BY ")
            .Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdRecCajaEnt.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpSqlItemDsctosRecsC() As String
        With MstbExpresionSql
            .Clear()
            .Append("SELECT I.").Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(", I.")
            .Append(ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdTipoItemRecByt.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijoFact_ItemRecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_ItemRecEnt.SstrNombreCampoBd)
            .Append(", SUM(I.").Append(ClsValor_ItemRecDec.SstrNombreCampoBd).Append("), '' as Detalle")
            .Append(" FROM ").Append(ClsItemRecCaja.SstrNombreTabla).Append(" AS I INNER JOIN ")
            .Append(ClsReciboCaja.SstrNombreTabla).Append(" AS R ON I.").Append(MstrCampoBdCarpeta)
            .Append(" = R.").Append(MstrCampoBdCarpeta).Append(" AND I.").Append(MstrCampoBdCentroUtil)
            .Append(" = R.").Append(MstrCampoBdCentroUtil).Append(" AND I.")
            .Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(" = R.")
            .Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(" AND I.")
            .Append(ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd).Append(" = R.")
            .Append(ClsIdRecCajaEnt.SstrNombreCampoBd)
            .Append(" WHERE I.").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND I.").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil.ToString)
            .Append(" AND ").Append(ClsFechaRecDtm.SstrNombreCampoBd).Append(" BETWEEN '")
            .Append(StrFechaDesde).Append("' AND '").Append(StrFechaHasta)
            .Append("' AND ").Append(ClsIdTipoItemRecByt.SstrNombreCampoBd).Append(" >= ")
            .Append(EnuTipoItemRecCajaDef.EnuDsctoIntMora)
            .Append(" GROUP BY ").Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdTipoItemRecByt.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijoFact_ItemRecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_ItemRecEnt.SstrNombreCampoBd).Append(", Detalle ")
            .Append(" ORDER BY ").Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd)
        End With
        Dim lstrSqlItemsDsctosRecCaja = MstbExpresionSql.ToString
        Return lstrSqlItemsDsctosRecCaja
    End Function
    Private Function FstrExpSqlMedPagoRecsC() As String
        With MstbExpresionSql
            .Clear()
            .Append("SELECT *, '' as Detalle").Append(" FROM ").Append(ClsMedioPago.SstrNombreTabla)
            .Append(" AS I INNER JOIN ").Append(ClsReciboCaja.SstrNombreTabla).Append(" AS R ON I.")
            .Append(MstrCampoBdCarpeta).Append(" = R.").Append(MstrCampoBdCarpeta).Append(" AND I.")
            .Append(MstrCampoBdCentroUtil).Append(" = R.").Append(MstrCampoBdCentroUtil)
            .Append(" AND I.").Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(" = R.")
            .Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(" AND I.")
            .Append(ClsIdRecCaja_MedPagoEnt.SstrNombreCampoBd).Append(" = R.")
            .Append(ClsIdRecCajaEnt.SstrNombreCampoBd)
            .Append(" WHERE I.").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND I.").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil.ToString)
            .Append(" AND ").Append(ClsFechaRecDtm.SstrNombreCampoBd).Append(" BETWEEN '")
            .Append(StrFechaDesde).Append("' AND '").Append(StrFechaHasta)
            .Append("' ORDER BY I.").Append(ClsIdRecCaja_MedPagoEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsOrdinal_MedPagoShr.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString()
    End Function
    Private Function FdtbInfItemsRecsC()
        Dim lstrPreRec As String, lentIdRec As Integer, lstrNroRec As String,
                lstrNroRecAct = String.Empty
        Dim lstrFiltro As String
        Dim ldtbInfItemsRC = ClsReciboCaja.FdtbInfItems()
        Dim ldtbItemsRC = FdtbInfItemsRecCaja()
        SPuebleInfPagoCap(ldtbInfItemsRC, ldtbItemsRC)
        For Each ldrwItemRec As DataRow In ldtbItemsRC.Rows
            lstrPreRec = ClsPanorama.FobjValorCampo(ldrwItemRec(ClsPrefijo_RecStr.SstrNombreCampoBd),
                            EnuTipoValor.enuString)
            lentIdRec = ClsPanorama.FobjValorCampo(ldrwItemRec(ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd),
                            EnuTipoValor.enuInteger)
            lstrNroRec = ClsPanorama.FstrNumeroDcto(lstrPreRec, lentIdRec)
            If lstrNroRecAct <> lstrNroRec Then
                lstrFiltro = ClsPrefijo_RecStr.SstrNombreCampoBd & " = '" & lstrPreRec & "' AND " &
                    ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd & " = " & lentIdRec
                SAdicioneItemsRC(ldtbItemsRC.Select(lstrFiltro), ldtbInfItemsRC)
                lstrNroRecAct = lstrNroRec
            End If
        Next
        Return ldtbInfItemsRC
    End Function
    Private Function FdtbInfItemsRecCaja() As DataTable
        With MstbExpresionSql
            .Clear().Append("SELECT I.").Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(", I.")
            .Append(ClsIdRecCajaEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijoFact_ItemRecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_ItemRecEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdItemFac_ItemRecShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdTipoItemRecByt.SstrNombreCampoBd).Append(", ").Append("SUM(I.")
            .Append(ClsValor_ItemRecDec.SstrNombreCampoBd).Append(") AS Valor FROM ")
            .Append(ClsItemRecCaja.SstrNombreTabla).Append(" AS I INNER JOIN ")
            .Append(ClsReciboCaja.SstrNombreTabla).Append(" AS R ON I.")
            .Append(MstrCampoBdCarpeta).Append(" = R.").Append(MstrCampoBdCarpeta).Append(" AND I.")
            .Append(MstrCampoBdCentroUtil).Append(" = R.").Append(MstrCampoBdCentroUtil)
            .Append(" AND I.").Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(" = R.")
            .Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(" AND I.")
            .Append(ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd).Append(" = R.")
            .Append(ClsIdRecCajaEnt.SstrNombreCampoBd).Append(" WHERE I.").Append(MstrCampoBdCarpeta)
            .Append(" = ").Append(GshrIdCarpeta.ToString).Append(" AND I.")
            .Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil.ToString)
            .Append(" AND ").Append(ClsFechaRecDtm.SstrNombreCampoBd).Append(" BETWEEN '")
            .Append(StrFechaDesde).Append("' AND '").Append(StrFechaHasta).Append("' GROUP BY ")
            .Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdRecCajaEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijoFact_ItemRecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_ItemRecEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdItemFac_ItemRecShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdTipoItemRecByt.SstrNombreCampoBd).Append(" ORDER BY ")
            .Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdRecCajaEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijoFact_ItemRecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_ItemRecEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdItemFac_ItemRecShr.SstrNombreCampoBd)
        End With
        Dim lstrExpSql = MstbExpresionSql.ToString()
        Dim ldtbInfItemsRC = ClsPanorama.FdtbDataTable(lstrExpSql)
        Return ldtbInfItemsRC
    End Function
#End Region
#Region "Notas de Ajuste"
    Private Sub SGenereDataSetNotasAjuste(adsNotaAjuste As DataSet)
        If IsNothing(ObjParRepDocs) Then
            Throw New ErrorInesperadoPanLException("Sin Parametros para el Reporte de Recibos de Caja")
        End If
        Dim lstrSqlNotaAjuste = FstrExpSqlNotaAjuste()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlNotaAjuste)
        lcolNombresTablas.Add("OriNotasAjusteCuota")
        GobjPanDat.SdsDataSet(adsNotaAjuste, lcolExpresionesSql, lcolNombresTablas)
        adsNotaAjuste.Tables.Add(FdtbCentroUtilidad)
        SComplementeDtbNAjuste(adsNotaAjuste)
        'Dim lstrNomArch = GstrTrayDatPrg & "NotasAjuste" & ".XML"
        'adsNotaAjuste.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlNotaAjuste() As String
        Dim lstrPrefNotaAjuste = ObjParRepDocs.StrPrefijoDocsRep
        Dim lstrIdNotaAjusteIni = ObjParRepDocs.EntIdDocInicial.ToString
        Dim lstrIdNotaAjusteFin = ObjParRepDocs.EntIdDocFinal.ToString
        With MstbExpresionSql
            .Append("Select NA.").Append(ClsIdCliente_NotaAjusteDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijo_NotaAjusteStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdNotaAjusteEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsFecha_NotaAjusteDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredio_NotaAjusteStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_NotaAjusteStr.SstrNombreCampoBd).Append(", ")
            .Append("'' AS NomPreAgr").Append(", ")
            .Append(ClsIdAnticipo_NotaAjusteEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsValor_NotaAjusteDec.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsNotaAjusteCuotaAdmin.SstrNombreTabla)
            .Append(" As NA INNER JOIN ").Append(ClsCliente.SstrNombreTabla).Append(" As C On NA.")
            .Append(MstrCampoBdCarpeta).Append(" = C.").Append(MstrCampoBdCarpeta)
            .Append(" AND NA.").Append(MstrCampoBdCentroUtil).Append(" = C.")
            .Append(MstrCampoBdCentroUtil).Append(" AND NA.").Append(ClsIdCliente_NotaAjusteDbl.SstrNombreCampoBd)
            .Append(" = C.").Append(ClsIdClienteDbl.SstrNombreCampoBd)
            .Append(" WHERE NA.").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND NA.").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil)
            .Append(" AND ").Append(ClsPrefijo_NotaAjusteStr.SstrNombreCampoBd).Append(" = '")
            .Append(lstrPrefNotaAjuste).Append("' AND NA.").Append(ClsIdNotaAjusteEnt.SstrNombreCampoBd)
            .Append(" >= ").Append(lstrIdNotaAjusteIni).Append(" AND NA.")
            .Append(ClsIdNotaAjusteEnt.SstrNombreCampoBd).Append(" <= ").Append(lstrIdNotaAjusteFin)
            .Append(" ORDER BY ").Append(ClsPrefijo_NotaAjusteStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdNotaAjusteEnt.SstrNombreCampoBd)
        End With
        Dim lstrSqlNotaAjuste = MstbExpresionSql.ToString
        Return lstrSqlNotaAjuste
    End Function
    Private Shared Sub SComplementeDtbNAjuste(adsNotaAjuste As DataSet)
        If Not IsNothing(adsNotaAjuste) Then
            Dim lstrIdPredioAgr As String, lstrNomPre As String
            Dim ldtbNA As DataTable = adsNotaAjuste.Tables("OriNotasAjusteCuota")
            Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
            For Each ldrwNA As DataRow In ldtbNA.Rows
                lstrIdPredioAgr = ClsPanorama.FobjValorCampo(ldrwNA(
                        ClsIdPredioAgrupador_NotaAjusteStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
                lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrIdPredioAgr})
                If lobjPredio.BlnExiste Then
                    lstrNomPre = lobjPredio.ObjNombrePredioStr.ToString()
                Else
                    lstrNomPre = lstrIdPredioAgr
                End If
                ldrwNA("NomPreAgr") = lstrNomPre
            Next
        End If
    End Sub
#End Region
#Region "Notas de Intereses"
    Private Sub SGenereDataSetNotasDb(adsNotaInt As DataSet)
        If IsNothing(ObjParRepDocs) Then
            Throw New ErrorInesperadoPanLException("Sin Parametros para el Reporte de Notas Db")
        End If
        Dim lstrPrefNotaDb = ObjParRepDocs.StrPrefijoDocsRep
        Dim lstrIdNotaDbIni = ObjParRepDocs.EntIdDocInicial.ToString
        Dim lstrIdNotaDbFin = ObjParRepDocs.EntIdDocFinal.ToString
        MstbExpresionSql.Clear()
        With MstbExpresionSql
            .Append("SELECT ND.").Append(ClsIdCliente_NotaDbDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijo_NotaDbStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdNotaDbEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsFecha_NotaDbDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_NotaDbStr.SstrNombreCampoBd).Append(", ")
            .Append("'' AS NomPreAgr").Append(", ")
            .Append(ClsAnuladoBln.SstrNombreCampoBd).Append(", ")
            .Append(ClsValor_NotaDbDec.SstrNombreCampoBd)
            .Append(", 0.00 AS ValorIva").Append(", 0.00 AS ValorBaseIva").Append(", ")
            .Append(ClsCUDE_NDbStr.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsNotaDb.SstrNombreTabla)
            .Append(" AS ND INNER JOIN ").Append(ClsCliente.SstrNombreTabla).Append(" AS C ON ND.")
            .Append(MstrCampoBdCarpeta).Append(" = C.").Append(MstrCampoBdCarpeta)
            .Append(" AND ND.").Append(MstrCampoBdCentroUtil).Append(" = C.")
            .Append(MstrCampoBdCentroUtil).Append(" AND ND.").Append(ClsIdCliente_NotaDbDbl.SstrNombreCampoBd)
            .Append(" = C.").Append(ClsIdClienteDbl.SstrNombreCampoBd)
            .Append(" WHERE ND.").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND ND.").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil)
            .Append(" AND ").Append(ClsPrefijo_NotaDbStr.SstrNombreCampoBd).Append(" = '").Append(lstrPrefNotaDb)
            .Append("' AND ND.").Append(ClsIdNotaDbEnt.SstrNombreCampoBd).Append(" >= ").Append(lstrIdNotaDbIni)
            .Append(" AND ND.").Append(ClsIdNotaDbEnt.SstrNombreCampoBd).Append(" <= ").Append(lstrIdNotaDbFin)
            .Append(FstrFiltroNotasElec)
            .Append(" ORDER BY ").Append(ClsIdNotaDbEnt.SstrNombreCampoBd)
        End With
        Dim lstrSqlNotaDb = MstbExpresionSql.ToString
        MstbExpresionSql.Clear()
        Dim lstrSqlItemsNotaDb = FstrExpSqlItemsNotaDb()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlNotaDb)
        lcolExpresionesSql.Add(lstrSqlItemsNotaDb)
        lcolNombresTablas.Add("OriNotasDb")
        lcolNombresTablas.Add("OriItemsNotaDb")
        GobjPanDat.SdsDataSet(adsNotaInt, lcolExpresionesSql, lcolNombresTablas)
        adsNotaInt.Tables.Add(FdtbCentroUtilidad)
        SComplementeDsNotasDb(adsNotaInt)
        SComplementeTablaItems(adsNotaInt)
        'Dim lstrNomArch = GstrTrayDatPrg & "NotasDb" & ".XML"
        'adsNotaInt.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrFiltroNotasElec() As String
        Dim lstrFiltro = String.Empty
        If GobjParametros.BlnEFacAutorizado Then
            lstrFiltro = " AND IF(" & ClsIdEstadoEDocEnt.SstrNombreCampoBd & " < " &
                    EnuEstadoEDoc.EnuNoEDoc & ", " &
                    ClsIdEstadoEDocEnt.SstrNombreCampoBd & " >= 4, " &
                    ClsIdEstadoEDocEnt.SstrNombreCampoBd & " >= 0)"
        End If
        Return lstrFiltro
    End Function
    Private Shared Sub SComplementeDsNotasDb(adsNotasDb As DataSet)
        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
        Dim lobjNDb As New ClsNotaDb()
        Dim ldecValorIva As Decimal, ldecValorBaseIva As Decimal
        If Not IsNothing(adsNotasDb) Then
            Dim ldtbNotaDb As DataTable = adsNotasDb.Tables("OriNotasDb")
            Dim lstrNomPredio As String, lstrPredAgr As String
            Dim ldclQR As New DataColumn("QR", System.Type.GetType("System.Byte[]"))
            ldtbNotaDb.Columns.Add(ldclQR)
            For Each ldrwNDb As DataRow In ldtbNotaDb.Rows
                lstrPredAgr = ClsPanorama.FobjValorCampo(ldrwNDb(
                        ClsIdPredioAgrupador_NotaDbStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPredAgr})
                If lobjPredio.BlnExiste Then
                    lstrNomPredio = lobjPredio.ObjNombrePredioStr.ObjValorPro
                Else
                    lstrNomPredio = lstrPredAgr
                End If
                ldrwNDb("NomPreAgr") = lstrNomPredio
                ' IVA
                Dim lstrPrefNdb As String = ClsPanorama.FobjValorCampo(ldrwNDb(
                        ClsPrefijo_NotaDbStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                Dim lentIdNDb As Integer = ClsPanorama.FobjValorCampo(ldrwNDb(
                        ClsIdNotaDbEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                lobjNDb.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefNdb, lentIdNDb})
                ldecValorIva = lobjNDb.DecValorIvaNota
                ldecValorBaseIva = lobjNDb.DecValorBaseIvaNota
                ldrwNDb("ValorIva") = ldecValorIva
                ldrwNDb("ValorBaseIva") = ldecValorBaseIva
                ldrwNDb("QR") = lobjNDb.FbytQRNdb
            Next
        End If
    End Sub
    Private Function FstrExpSqlItemsNotaDb()
        Dim lstrPrefNotaDb = ObjParRepDocs.StrPrefijoDocsRep
        Dim lstrIdNotaDbIni = ObjParRepDocs.EntIdDocInicial.ToString
        Dim lstrIdNotaDbFin = ObjParRepDocs.EntIdDocFinal.ToString
        Dim lstrCamposSelect() = {"*", "'' as NroFact"}
        Dim lstrIndice = {{ClsPrefijo_NotaDbStr.SstrNombreCampoBd, "ASC"},
                          {ClsIdNotaDb_ItemNotaDbEnt.SstrNombreCampoBd, "ASC"},
                          {ClsIdItemNotaDbShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsPrefijo_NotaDbStr.SstrNombreCampoBd &
                    " = '" & lstrPrefNotaDb & "' AND " &
                    ClsIdNotaDb_ItemNotaDbEnt.SstrNombreCampoBd & " BETWEEN " & lstrIdNotaDbIni & " AND " &
                    lstrIdNotaDbFin
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(ClsItemNotaDb.SstrNombreTabla,
                lstrCamposSelect, lstrIndice, lstrFiltro, Array.Empty(Of String))
        Return lstrExpSql
    End Function
    Private Sub SComplementeTablaItems(adsNotaInt As DataSet)
        If adsNotaInt IsNot Nothing Then
            Dim ldtbItemsNotaDb As DataTable = adsNotaInt.Tables("OriItemsNotaDb")
            Dim lstrPrefFact As String, lentIdFact As Integer, lstrNroFact As String
            For Each ldrwItNotaDb As DataRow In ldtbItemsNotaDb.Rows
                lstrPrefFact = ClsPanorama.FobjValorCampo(ldrwItNotaDb(
                        ClsPrefijoFact_ItemNotaDbStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                lentIdFact = ClsPanorama.FobjValorCampo(ldrwItNotaDb(
                        ClsIdFactura_ItemNotaDbEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                lstrNroFact = ClsPanorama.FstrNumeroDcto(lstrPrefFact, lentIdFact)
                ldrwItNotaDb("NroFact") = lstrNroFact
            Next
        End If
    End Sub
#End Region
#Region "Notas Crédito"
    Private Sub SGenereDataSetNotasCr(adsNotasCr As DataSet)
        If IsNothing(ObjParRepDocs) Then
            Throw New ErrorInesperadoPanLException("Sin Parametros para el Reporte de Recibos de Caja")
        End If
        Dim lstrPrefNotaCr = ObjParRepDocs.StrPrefijoDocsRep
        Dim lstrIdNotaCrIni = ObjParRepDocs.EntIdDocInicial.ToString
        Dim lstrIdNotaCrFin = ObjParRepDocs.EntIdDocFinal.ToString
        Const CLSTRCOMA = ", "
        Dim lstbExpresion As New StringBuilder
        lstbExpresion.EnsureCapacity(500)
        ' Nota Credito
        With lstbExpresion
            .Append("SELECT NC.").Append(ClsIdCliente_NotaCrDbl.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsPrefijo_NotaCrStr.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsIdNotaCrEnt.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsValor_NotaCrDec.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsFecha_NotaCrDtm.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsIdPredioAgrupador_NotaCrStr.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append("'' AS NomPreAgr").Append(", ")
            .Append(ClsAnuladoBln.SstrNombreCampoBd).Append(", ")
            .Append(ClsCUDEStr.SstrNombreCampoBd).Append(", ")
            .Append("NC.").Append(ClsComentario_NotaCrStr.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsNotaCr.SstrNombreTabla).Append(" AS NC INNER JOIN ")
            .Append(ClsCliente.SstrNombreTabla).Append(" AS C ON NC.").Append(MstrCampoBdCarpeta)
            .Append(" = C.").Append(MstrCampoBdCarpeta).Append(" AND NC.").Append(MstrCampoBdCentroUtil)
            .Append(" = C.").Append(MstrCampoBdCentroUtil).Append(" AND NC.")
            .Append(ClsIdCliente_NotaCrDbl.SstrNombreCampoBd).Append(" = C.")
            .Append(ClsIdClienteDbl.SstrNombreCampoBd)
            .Append(" WHERE NC.").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND NC.").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil)
            .Append(" AND NC.").Append(ClsPrefijo_NotaCrStr.SstrNombreCampoBd).Append(" = '")
            .Append(lstrPrefNotaCr).Append("' AND NC.").Append(ClsIdNotaCrEnt.SstrNombreCampoBd)
            .Append(" >= ").Append(lstrIdNotaCrIni).Append(" AND NC.")
            .Append(ClsIdNotaCrEnt.SstrNombreCampoBd).Append(" <= ").Append(lstrIdNotaCrFin)
            .Append(" ORDER BY ").Append(ClsIdNotaCrEnt.SstrNombreCampoBd)
        End With
        Dim lstrSqlNotaCr = lstbExpresion.ToString
        lstbExpresion.Clear()
        With lstbExpresion
            .Append("SELECT ").Append(ClsPrefijo_NotaCrStr.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsIdNotaCr_ItemNotaCrEnt.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsPrefijoFact_ItemNotaCrStr.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsIdFactura_ItemNotaCrEnt.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsIdTipoDscto_ItemNotaCrByt.SstrNombreCampoBd).Append(", '' AS Detalle, ")
            .Append(ClsValor_ItemNotaCrDec.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsItemNotaCr.SstrNombreTabla)
            .Append(" WHERE ").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND ").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil.ToString)
            .Append(" AND ").Append(ClsPrefijo_NotaCrStr.SstrNombreCampoBd).Append(" = '")
            .Append(lstrPrefNotaCr).Append("' AND ").Append(ClsIdNotaCr_ItemNotaCrEnt.SstrNombreCampoBd)
            .Append(" >= ").Append(lstrIdNotaCrIni).Append(" AND ")
            .Append(ClsIdNotaCr_ItemNotaCrEnt.SstrNombreCampoBd).Append(" <= ").Append(lstrIdNotaCrFin)
            .Append(" ORDER BY ").Append(ClsPrefijoFact_ItemNotaCrStr.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsIdFactura_ItemNotaCrEnt.SstrNombreCampoBd)
        End With
        Dim lstrSqlItemsNotaCr = lstbExpresion.ToString
        lstbExpresion.Clear()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlNotaCr)
        lcolExpresionesSql.Add(lstrSqlItemsNotaCr)
        lcolNombresTablas.Add("OriNotasCr")
        lcolNombresTablas.Add("OriItemsNotaCr")
        GobjPanDat.SdsDataSet(adsNotasCr, lcolExpresionesSql, lcolNombresTablas)
        SComplementeDsNotasCr(adsNotasCr)
        adsNotasCr.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "NotasCredito" & ".XML"
        'adsNotasCr.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Shared Sub SComplementeDsNotasCr(adsNotasCr As DataSet)
        If Not IsNothing(adsNotasCr) Then
            Dim ldtbNCr As DataTable = adsNotasCr.Tables("OriNotasCr")
            Dim lstrNomPredio As String, lstrPredAgr As String
            Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
            Dim ldclQR As New DataColumn("QR", System.Type.GetType("System.Byte[]"))
            ldtbNCr.Columns.Add(ldclQR)
            Dim lobjNcr As New ClsNotaCr(), lobjValorLlave As Object()
            For Each ldrwNCr As DataRow In ldtbNCr.Rows
                Dim lstrPrefNCr As String = ClsPanorama.FobjValorCampo(ldrwNCr(
                        ClsPrefijo_NotaCrStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                Dim lentIdNcr As Integer = ClsPanorama.FobjValorCampo(ldrwNCr(
                        ClsIdNotaCrEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefNCr, lentIdNcr}
                lobjNcr.SAbra(lobjValorLlave)
                lstrPredAgr = ClsPanorama.FobjValorCampo(ldrwNCr(
                        ClsIdPredioAgrupador_NotaCrStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPredAgr})
                If lobjPredio.BlnExiste Then
                    lstrNomPredio = lobjPredio.ObjNombrePredioStr.ObjValorPro
                Else
                    If String.IsNullOrEmpty(lstrPredAgr) Then
                        lstrNomPredio = "Sin Predio Agrupador"
                    Else
                        lstrNomPredio = lstrPredAgr
                    End If
                End If
                ldrwNCr("NomPreAgr") = lstrNomPredio
                ldrwNCr("QR") = lobjNcr.FbytQRNcr
            Next
            Dim ldtbItemsNotaCr As DataTable = adsNotasCr.Tables("OriItemsNotaCr")
            For Each ldrwItem As DataRow In ldtbItemsNotaCr.Rows
                Dim lenuTipoDescuento As EnuTipoDescuentoDef = ClsPanorama.FobjValorCampo(ldrwItem(
                        ClsIdTipoDscto_ItemNotaCrByt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                Dim lstrPrefFac As String = ClsPanorama.FobjValorCampo(ldrwItem(
                        ClsPrefijoFact_ItemNotaConStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                Dim lentIdFac As Integer = ClsPanorama.FobjValorCampo(ldrwItem(
                        ClsIdFactura_ItemNotaConEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                Dim lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPrefFac, lentIdFac)
                Dim lstrDetalle = String.Empty
                Select Case lenuTipoDescuento
                    Case EnuTipoDescuentoDef.EnuDsctoCapital
                        lstrDetalle = "Descuento a Capital "
                    Case EnuTipoDescuentoDef.EnuDsctoIntMora
                        lstrDetalle = "Descuento a Intereses de Mora"
                    Case EnuTipoDescuentoDef.EnuReteCree
                        lstrDetalle = "Retención del CREE"
                    Case EnuTipoDescuentoDef.EnuReteFuente
                        lstrDetalle = "Retención en la Fuente"
                    Case EnuTipoDescuentoDef.EnuReteIca
                        lstrDetalle = "Retención de Industria y Comercio"
                    Case EnuTipoDescuentoDef.EnuReteIva
                        lstrDetalle = "Retención de IVA"
                    Case EnuTipoDescuentoDef.EnuDsctoPP
                        lstrDetalle = "Decuento por Pronto Pago"
                    Case EnuTipoDescuentoDef.EnuCancelaIva
                        lstrDetalle = "IVA llevado al Gasto"
                End Select
                lstrDetalle &= " Factura " & lstrNroFac
                ldrwItem("Detalle") = lstrDetalle
            Next
        End If
    End Sub
#End Region
#Region "Notas Devolución Anticipo"
    Private Sub SGenereDataSetNotasDevAnt(adsNotasDevAnt As DataSet)
        If IsNothing(ObjParRepDocs) Then
            Throw New ErrorInesperadoPanLException("Sin Parametros para el Reporte de Notas de Dev Anticipos")
        End If
        Dim lstrIdNotaDevAntFin = ObjParRepDocs.EntIdDocFinal.ToString
        Dim lstrPrefNotaDevAnt = ObjParRepDocs.StrPrefijoDocsRep
        Const CLSTRCOMA = ", "
        Dim lstbExpresion As New StringBuilder
        lstbExpresion.EnsureCapacity(500)
        ' Nota Devolucion Anticipo
        With lstbExpresion
            .Clear()
            .Append("SELECT NDA.").Append(ClsIdCliente_NotaDevAntDbl.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsPrefijo_NotaDevAntStr.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsIdNotaDevAntEnt.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsIdAnticipo_NotaDevAntEnt.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsValor_NotaDevAntDec.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsFecha_NotaDevAntDtm.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsIdPredioAgrupador_NotaDevAntStr.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsAnuladoBln.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append("NDA.").Append(ClsComentario_NotaDevAntStr.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsNotaDevAnt.SstrNombreTabla).Append(" AS NDA INNER JOIN ")
            .Append(ClsCliente.SstrNombreTabla).Append(" AS C ON NDA.").Append(MstrCampoBdCarpeta)
            .Append(" = C.").Append(MstrCampoBdCarpeta).Append(" AND NDA.").Append(MstrCampoBdCentroUtil)
            .Append(" = C.").Append(MstrCampoBdCentroUtil).Append(" AND NDA.")
            .Append(ClsIdCliente_NotaDevAntDbl.SstrNombreCampoBd).Append(" = C.")
            .Append(ClsIdClienteDbl.SstrNombreCampoBd)
            .Append(" WHERE NDA.").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND NDA.").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil)
            .Append(" AND NDA.").Append(ClsPrefijo_NotaDevAntStr.SstrNombreCampoBd)
            .Append(" = '").Append(lstrPrefNotaDevAnt)
            .Append("' AND NDA.").Append(ClsIdNotaDevAntEnt.SstrNombreCampoBd)
            .Append(" >= ").Append(lstrIdNotaDevAntFin).Append(" AND NDA.")
            .Append(ClsIdNotaDevAntEnt.SstrNombreCampoBd).Append(" <= ").Append(lstrIdNotaDevAntFin)
            .Append(" ORDER BY ").Append(ClsIdNotaDevAntEnt.SstrNombreCampoBd)
        End With
        Dim lstrSqlNotaDevAnt = lstbExpresion.ToString
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlNotaDevAnt)
        lcolNombresTablas.Add("OriNotasDevAnt")
        GobjPanDat.SdsDataSet(adsNotasDevAnt, lcolExpresionesSql, lcolNombresTablas)
        adsNotasDevAnt.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "NotasDevolucionAnt" & ".XML"
        'adsNotasDevAnt.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
#End Region
#Region "Notas Reversión Créditos"
    Private Sub SGenereDataSetNotasRevCr(adsNotasRevCr As DataSet)
        If IsNothing(ObjParRepDocs) Then
            Throw New ErrorInesperadoPanLException("Sin Parametros para el Reporte de NotasReversión Cr")
        End If
        Dim lstrSqlNotaRevCr = FstrExpNotaRevCr()
        Dim lstrSqlNovedadesNotaRevRc = FstrExpNovNotaReverCr()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlNotaRevCr)
        lcolNombresTablas.Add("OriNotasRevCr")
        lcolExpresionesSql.Add(lstrSqlNovedadesNotaRevRc)
        lcolNombresTablas.Add("OriNovedadesNotaRevCr")
        GobjPanDat.SdsDataSet(adsNotasRevCr, lcolExpresionesSql, lcolNombresTablas)
        adsNotasRevCr.Tables.Add(FdtbCentroUtilidad)
        SComplementeTablas(adsNotasRevCr)
        'Dim lstrNomArch = GstrTrayDatPrg & "NotasReversionCr" & ".XML"
        'adsNotasRevCr.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpNotaRevCr() As String
        Dim lstrPrefNotaRevCr = ObjParRepDocs.StrPrefijoDocsRep
        Dim lstrIdNotaRevCrFin = ObjParRepDocs.EntIdDocFinal.ToString
        Const CLSTRCOMA = ", "
        Dim lstbExpresion As New StringBuilder
        lstbExpresion.EnsureCapacity(500)
        ' Nota Reversión Créditos
        With lstbExpresion
            .Clear()
            .Append("SELECT NRCR.").Append(ClsIdCliente_NotaReversaCrDbl.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsIdNotaReversaCrEnt.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsPrefijoDoc_NotaReversaCrStr.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsIdDoc_NotaReversaCrEnt.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsValor_NotaReversaCrDec.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsFecha_NotaReversaCrDtm.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsIdPredioAgr_NotaReversaCrStr.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append("'' AS NomPreAgr").Append(", ")
            .Append(ClsDetalle_NotaReversaCrStr.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsTipoDocReversadoByt.SstrNombreCampoBd).Append(CLSTRCOMA)
            .Append(ClsCUDEStr.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsNotaReversionCr.SstrNombreTabla).Append(" AS NRCR INNER JOIN ")
            .Append(ClsCliente.SstrNombreTabla).Append(" AS C ON NRCR.").Append(MstrCampoBdCarpeta)
            .Append(" = C.").Append(MstrCampoBdCarpeta).Append(" AND NRCR.").Append(MstrCampoBdCentroUtil)
            .Append(" = C.").Append(MstrCampoBdCentroUtil).Append(" AND NRCR.")
            .Append(ClsIdCliente_NotaReversaCrDbl.SstrNombreCampoBd).Append(" = C.")
            .Append(ClsIdClienteDbl.SstrNombreCampoBd)
            .Append(" WHERE NRCR.").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND NRCR.").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil)
            .Append(" AND NRCR.").Append(ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd)
            .Append(" = '").Append(lstrPrefNotaRevCr)
            .Append("' AND NRCR.").Append(ClsIdNotaReversaCrEnt.SstrNombreCampoBd)
            .Append(" = ").Append(lstrIdNotaRevCrFin)
        End With
        Return lstbExpresion.ToString
    End Function
    Private Function FstrExpNovNotaReverCr() As String
        Dim lstrExpSqlNovedades = FstrExpSqlNovedades()
        Dim lstrExpSqlNovedadesAnt = FstrExpSqlNovedadesAnt()
        Dim lstrExpSql = (lstrExpSqlNovedades) & " UNION ALL " & (lstrExpSqlNovedadesAnt)
        Return lstrExpSql
    End Function
    Private Function FstrExpSqlNovedades() As String
        Dim lstrPrefNotaRevCr = ObjParRepDocs.StrPrefijoDocsRep
        Dim lstrIdNotaRevRC = ObjParRepDocs.EntIdDocFinal.ToString
        Dim lstrNombreTabla = ClsNovedad.SstrNombreTabla
        Dim lstrCamposSelect As String() = {ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd,
                ClsIdDocOrigenEnt.SstrNombreCampoBd, ClsFechaCreacionDtm.SstrNombreCampoBd,
                ClsIdCuentaDb_NovStr.SstrNombreCampoBd, ClsIdCuentaCr_NovStr.SstrNombreCampoBd,
                ClsPrefijoFact_NovStr.SstrNombreCampoBd, ClsIdFactura_NovEnt.SstrNombreCampoBd,
                ClsValor_NovDec.SstrNombreCampoBd, ClsIdTipoNovedadByt.SstrNombreCampoBd,
                "'' AS NroFac", "'' AS Detalle"}
        Dim lstrIndice = {{"", ""}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " & EnuTipoDocOri.EnuNotaRevCr &
                " AND " & ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" &
                lstrPrefNotaRevCr & "' AND " & ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " &
                lstrIdNotaRevRC
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, Array.Empty(Of String)())
        Return lstrExpSql
    End Function
    Private Function FstrExpSqlNovedadesAnt() As String
        Dim lstrPrefNotaRevCr = ObjParRepDocs.StrPrefijoDocsRep
        Dim lstrIdNotaRevRC = ObjParRepDocs.EntIdDocFinal.ToString
        Dim lstrNombreTabla = ClsNovedadAnticipo.SstrNombreTabla
        Dim lstrCamposSelect As String() = {ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd,
                                            ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd,
                                            ClsFechaCreacionDtm.SstrNombreCampoBd,
                                            ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd,
                                            ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd,
                                            "'' AS " & ClsPrefijoFact_NovStr.SstrNombreCampoBd,
                                            "0 AS " & ClsIdFactura_NovEnt.SstrNombreCampoBd,
                                            ClsValor_NovAntDec.SstrNombreCampoBd,
                                            ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd,
                                            "'' AS NroFac", "'' AS Detalle"}
        Dim lstrIndice = {{"", ""}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd &
                       " = " & EnuTipoDocOri.EnuNotaRevCr & " AND " &
                       ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd & " = '" & lstrPrefNotaRevCr & "' AND " &
                       ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd & " = " & lstrIdNotaRevRC
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelect, lstrIndice,
                lstrFiltro, Array.Empty(Of String)())
        Return lstrExpSql
    End Function
    Private Shared Sub SComplementeTablas(adsNotasRevCr As DataSet)
        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
        Dim ldtbNRCr As DataTable = adsNotasRevCr.Tables("OriNotasRevCr")
        Dim lstrNomPredio As String, lstrPredAgr As String
        Dim ldclQR As New DataColumn("QR", System.Type.GetType("System.Byte[]"))
        ldtbNRCr.Columns.Add(ldclQR)
        Dim lobjNrcr As New ClsNotaReversionCr(), lobjValorLlave As Object()
        For Each ldrwNRCr As DataRow In ldtbNRCr.Rows
            Dim lstrPrefNrcr As String = ClsPanorama.FobjValorCampo(ldrwNRCr(
                    ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd), EnuTipoValor.enuString)
            Dim lentIdNrcr As Integer = ClsPanorama.FobjValorCampo(ldrwNRCr(
                    ClsIdNotaReversaCrEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
            lstrPredAgr = ClsPanorama.FobjValorCampo(ldrwNRCr(
                    ClsIdPredioAgr_NotaReversaCrStr.SstrNombreCampoBd), EnuTipoValor.enuString)
            lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPredAgr})
            If lobjPredio.BlnExiste Then
                lstrNomPredio = lobjPredio.ObjNombrePredioStr.ObjValorPro
            Else
                If String.IsNullOrEmpty(lstrPredAgr) Then
                    lstrNomPredio = "Sin Predio Agrupador"
                Else
                    lstrNomPredio = lstrPredAgr
                End If
            End If
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefNrcr, lentIdNrcr}
            lobjNrcr.SAbra(lobjValorLlave)
            ldrwNRCr("NomPreAgr") = lstrNomPredio
            ldrwNRCr("QR") = lobjNrcr.FbytQRNcr
        Next
        Dim ldtbNovedades = adsNotasRevCr.Tables("OriNovedadesNotaRevCr")
        Dim lstrConceptoNovedad As String
        For Each ldrwNovedad As DataRow In ldtbNovedades.Rows
            Dim lstrPrefFac As String = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsPrefijoFact_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            Dim lentIdFac As Integer = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdFactura_NovEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            Dim lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPrefFac, lentIdFac)
            Dim lenuTipoNovedad As EnuTipoNov =
                    ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdTipoNovedadByt.SstrNombreCampoBd),
                    EnuTipoValor.enuByte)
            Select Case lenuTipoNovedad
                Case EnuTipoNov.EnuRCrAnApCap
                    lstrConceptoNovedad = "Reversión Anticipo aplicado a Capital"
                Case EnuTipoNov.EnuRCrAnApInt
                    lstrConceptoNovedad = "Reversión Anticipo aplicado a Int. de Mora"
                Case EnuTipoNov.EnuRCrAntRec
                    lstrConceptoNovedad = "Reversión Anticipo recibido"
                Case EnuTipoNov.EnuRCrDctoCap
                    lstrConceptoNovedad = "Reversión Descuento a Capital"
                Case EnuTipoNov.EnuRCrDctoInt
                    lstrConceptoNovedad = "Reversión Descuento a Int. de Mora"
                Case EnuTipoNov.EnuRCrPagoCap
                    lstrConceptoNovedad = "Reversión Abono a Capital"
                Case EnuTipoNov.EnuRCrPagoInt
                    lstrConceptoNovedad = "Reversión Abono a Int. de Mora"
                Case EnuTipoNov.EnuRCrRetCre
                    lstrConceptoNovedad = "Reversión ReteCree aplicado"
                Case EnuTipoNov.EnuRCrRetFte
                    lstrConceptoNovedad = "Reversión Retefuente aplicado"
                Case EnuTipoNov.EnuRCrRetIca
                    lstrConceptoNovedad = "Reversión ReteIca aplicado"
                Case EnuTipoNov.EnuRCrRetIva
                    lstrConceptoNovedad = "Reversión ReteIva aplicado"
                Case EnuTipoNov.EnuRDbCap
                    lstrConceptoNovedad = "Reversión Débito a Deuda de Capital"
                Case EnuTipoNov.EnuRDbInt
                    lstrConceptoNovedad = "Reversión Débito a Int. de Mora"
                Case EnuTipoNov.EnuRDbIva
                    lstrConceptoNovedad = "Reversión a Débito Iva Generado."
                Case EnuTipoNov.EnuDbIvaInt
                    lstrConceptoNovedad = "Reversión a Débito Iva Intereses Generado."
                Case EnuTipoNov.EnuRDbAntApl
                    lstrConceptoNovedad = "Reversión Anticipo aplicado."
                Case EnuTipoNov.EnuRCrIvaGas
                    lstrConceptoNovedad = "Reversión IVA al Gasto."
                Case Else
                    Throw New ErrorInesperadoPanLException("Tipo de Novedad inconsistente!")
            End Select
            ldrwNovedad("NroFac") = lstrNroFac
            ldrwNovedad("Detalle") = lstrConceptoNovedad
        Next
    End Sub
#End Region
#Region "Notas Contables"
    ' DataSet Notas Contables
    Private Sub SGenereDataSetNotasCon(adsNotasCon As DataSet)
        If IsNothing(ObjParRepDocs) Then
            Throw New ErrorInesperadoPanLException("Sin Parametros para el Reporte de Notas Contables")
        End If
        Dim lstrPrefNotasCon = ObjParRepDocs.StrPrefijoDocsRep
        Dim lstrIdNotaConIni = ObjParRepDocs.EntIdDocInicial.ToString
        Dim lstrIdNotaConFin = ObjParRepDocs.EntIdDocFinal.ToString
        With MstbExpresionSql
            .Clear()
            .Append("SELECT NC.").Append(ClsIdCliente_NotaConDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijo_NotaConStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdNotaConEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsFecha_NotaConDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_NotaConStr.SstrNombreCampoBd).Append(", ")
            .Append("'' AS NomPreAgr").Append(", ")
            .Append(ClsIdAnticipo_NotaConEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsAnuladoBln.SstrNombreCampoBd).Append(", ")
            .Append(ClsCUDEStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsValor_NotaConDec.SstrNombreCampoBd).Append(" FROM ")
            .Append(ClsNotaCon.SstrNombreTabla).Append(" AS NC INNER JOIN ")
            .Append(ClsCliente.SstrNombreTabla).Append(" AS C ON NC.")
            .Append(ClsIdCliente_NotaConDbl.SstrNombreCampoBd).Append(" = ").Append("C.")
            .Append(ClsIdClienteDbl.SstrNombreCampoBd).Append(" AND NC.").Append(MstrCampoBdCarpeta)
            .Append(" = C.").Append(MstrCampoBdCarpeta).Append(" AND NC.")
            .Append(MstrCampoBdCentroUtil).Append(" = C.").Append(MstrCampoBdCentroUtil)
            .Append(" WHERE NC.").Append(MstrCampoBdCarpeta).Append(" = ")
            .Append(GshrIdCarpeta.ToString).Append(" AND NC.").Append(MstrCampoBdCentroUtil)
            .Append(" = ").Append(GshrIdCentroUtil.ToString).Append(" AND NC.")
            .Append(ClsPrefijo_NotaConStr.SstrNombreCampoBd).Append(" = '").Append(lstrPrefNotasCon)
            .Append("' AND NC.").Append(ClsIdNotaConEnt.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrIdNotaConIni).Append(" AND NC.")
            .Append(ClsIdNotaConEnt.SstrNombreCampoBd).Append(" <= ").Append(lstrIdNotaConFin)
            .Append(" ORDER BY ").Append(ClsIdNotaConEnt.SstrNombreCampoBd)
        End With
        Dim lstrSqlNotasCon = MstbExpresionSql.ToString
        With MstbExpresionSql
            .Clear()
            .Append("SELECT *, '' as Detalle").Append(" FROM ").Append(ClsItemNotaCon.SstrNombreTabla)
            .Append(" WHERE ").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta.ToString)
            .Append(" AND ").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil.ToString)
            .Append(" AND ").Append(ClsPrefijo_NotaConStr.SstrNombreCampoBd).Append(" = '")
            .Append(lstrPrefNotasCon).Append("' AND ").Append(ClsIdNotaCon_ItemNotaConEnt.SstrNombreCampoBd)
            .Append(" >= ").Append(lstrIdNotaConIni).Append(" AND ")
            .Append(ClsIdNotaCon_ItemNotaConEnt.SstrNombreCampoBd).Append(" <= ").Append(lstrIdNotaConFin)
            .Append(" ORDER BY ").Append(ClsIdNotaCon_ItemNotaConEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdItemNotaConShr.SstrNombreCampoBd)
        End With
        Dim lstrSqlDetalleNotaCon = MstbExpresionSql.ToString
        MstbExpresionSql.Clear()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlNotasCon)
        lcolExpresionesSql.Add(lstrSqlDetalleNotaCon)
        lcolNombresTablas.Add("OriNotasCon")
        lcolNombresTablas.Add("OriItemsNotaCon")
        GobjPanDat.SdsDataSet(adsNotasCon, lcolExpresionesSql, lcolNombresTablas)
        adsNotasCon.Tables.Add(FdtbCentroUtilidad)
        SComplementeDsNotasCon(adsNotasCon)
        'Dim lstrNomArch = GstrTrayDatPrg & "NotaContable" & ".XML"
        'adsNotasCon.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Shared Sub SComplementeDsNotasCon(adsNotasCon As DataSet)
        If Not IsNothing(adsNotasCon) Then
            Dim ldtbNCon As DataTable = adsNotasCon.Tables("OriNotasCon")
            Dim ldtbItemsNotaCon As DataTable = adsNotasCon.Tables("OriItemsNotaCon")
            Dim ldrwItems As DataRow() = ldtbItemsNotaCon.Select()
            Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
            Dim lstrPredAgr As String, lstrNomPredio As String
            Dim ldclQR As New DataColumn("QR", System.Type.GetType("System.Byte[]"))
            ldtbNCon.Columns.Add(ldclQR)
            Dim lobjNcon As New ClsNotaCon(), lobjValorLlave As Object()
            For Each ldrwNCon As DataRow In ldtbNCon.Rows
                Dim lstrPrefNCon As String = ClsPanorama.FobjValorCampo(ldrwNCon(
                        ClsPrefijo_NotaConStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                Dim lentIdNcon As Integer = ClsPanorama.FobjValorCampo(ldrwNCon(
                        ClsIdNotaConEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefNCon, lentIdNcon}
                lobjNcon.SAbra(lobjValorLlave)
                lstrPredAgr = ClsPanorama.FobjValorCampo(ldrwNCon(
                        ClsIdPredioAgrupador_NotaConStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
                lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPredAgr})
                If lobjPredio.BlnExiste Then
                    lstrNomPredio = lobjPredio.ObjNombrePredioStr.ObjValorPro
                Else
                    If String.IsNullOrEmpty(lstrPredAgr) Then
                        lstrNomPredio = "Sin Predio Agrupador"
                    Else
                        lstrNomPredio = lstrPredAgr
                    End If
                End If
                ldrwNCon("NomPreAgr") = lstrNomPredio
                ldrwNCon("QR") = lobjNcon.FbytQRNcon
            Next
            For Each ldrwItem As DataRow In ldrwItems
                Dim lstrPrefFac As String = ClsPanorama.FobjValorCampo(ldrwItem(
                        ClsPrefijoFact_ItemNotaConStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                Dim lentIdFac As Integer = ClsPanorama.FobjValorCampo(ldrwItem(
                        ClsIdFactura_ItemNotaConEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                Dim lshrIdItemFac As Short = ClsPanorama.FobjValorCampo(ldrwItem(
                        ClsIdItemFac_ItemNotaConShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
                Dim lenuTipoItemNcon As Byte = ClsPanorama.FobjValorCampo(ldrwItem(
                        ClsIdTipoItemNotaConByt.SstrNombreCampoBd), EnuTipoValor.enuByte)
                Dim lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPrefFac, lentIdFac)
                ldrwItem("Detalle") = FstrDetalleNCon(lenuTipoItemNcon, lshrIdItemFac, lstrNroFac)
            Next
        End If
    End Sub
    Private Shared Function FstrDetalleNCon(aenuTipoItemNCon As EnuTipoItemNotaConDef, ashrIdItemFac As Short,
                                     astrNroFac As String)
        Dim lstrDetalle As String
        lstrDetalle = ClsNotaCon.FstrDetalle(aenuTipoItemNCon)
        lstrDetalle &= " a Fac. Nro. " & astrNroFac & " Item " & ashrIdItemFac.ToString
        Return lstrDetalle
    End Function
#End Region
#Region "Cuenta de Cobro"
    Private Sub SGenereDataSetCuentasCobro(adsCtasCobro As DataSet, ablnDetallada As Boolean)
        If IsNothing(ObjParRepDocs) Then
            Throw New ErrorInesperadoPanLException("Sin Parametros para el Reporte de Cuentas de Cobro")
        End If
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        Dim lstrSqlCuentaCobro = FstrExpSqlCtaCobro()
        lcolExpresionesSql.Add(lstrSqlCuentaCobro)
        lcolNombresTablas.Add("OriCtasCobro")
        If ablnDetallada Then
            Dim lstrSqlFacturasEstado = FstrExpSqlFrasEstado()
            lcolExpresionesSql.Add(lstrSqlFacturasEstado)
            lcolNombresTablas.Add("OriFacturasEstado")
        End If
        GobjPanDat.SdsDataSet(adsCtasCobro, lcolExpresionesSql, lcolNombresTablas)
        adsCtasCobro.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "CuentasCobro" & ".XML"
        'Dim lstrNomArch = GstrTrayDatPrg & "CuentasCobroDetalle" & ".XML"
        'adsCtasCobro.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlCtaCobro() As String
        Dim lstrIdEstadoCtaIni = ObjParRepDocs.EntIdDocInicial.ToString
        Dim lstrIdEstadoCtaFin = ObjParRepDocs.EntIdDocFinal.ToString
        With MstbExpresionSql
            .Clear()
            .Append("SELECT CC.").Append(ClsIdCliente_EstadoDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdEstadoCuentaEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgr_EstadoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsDeudaCapitalDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsAntPorAplDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsDeudaIntMoraDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaEstadoDtm.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsEstadoCuenta.SstrNombreTabla).Append(" AS CC INNER JOIN ")
            .Append(ClsCliente.SstrNombreTabla).Append(" AS C ON CC.").Append(MstrCampoBdCarpeta)
            .Append(" = C.").Append(MstrCampoBdCarpeta).Append(" AND CC.").Append(MstrCampoBdCentroUtil)
            .Append(" = C.").Append(MstrCampoBdCentroUtil).Append(" AND CC.")
            .Append(ClsIdCliente_EstadoDbl.SstrNombreCampoBd).Append(" = C.")
            .Append(ClsIdClienteDbl.SstrNombreCampoBd)
            .Append(" WHERE CC.").Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND CC.").Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil)
            .Append(" AND CC.").Append(ClsIdEstadoCuentaEnt.SstrNombreCampoBd)
            .Append(" >= ").Append(lstrIdEstadoCtaIni).Append(" AND CC.")
            .Append(ClsIdEstadoCuentaEnt.SstrNombreCampoBd).Append(" <= ").Append(lstrIdEstadoCtaFin)
            If Not BlnEstadoDeCuenta Then
                .Append(" AND ").Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd).Append(" = 0")
            End If
            .Append(" ORDER BY ").Append(ClsIdEstadoCuentaEnt.SstrNombreCampoBd)
        End With
        Dim lstrSqlCtaCobro = MstbExpresionSql.ToString
        Return lstrSqlCtaCobro
    End Function
    Private Function FstrExpSqlFrasEstado()
        Dim lstrIdEstadoCtaIni = ObjParRepDocs.EntIdDocInicial.ToString
        Dim lstrIdEstadoCtaFin = ObjParRepDocs.EntIdDocFinal.ToString
        With MstbExpresionSql
            .Clear().Append("SELECT FE.").Append(ClsIdEstado_FactEstadoEnt.SstrNombreCampoBd)
            .Append(", ").Append(ClsDetalleItemFac_EstadoStr.SstrNombreCampoBd).Append(", ")
            .Append("FE.").Append(ClsPrefijoFacturaVivaStr.SstrNombreCampoBd).Append(", ")
            .Append("FE.").Append(ClsIdFacturaVivaEnt.SstrNombreCampoBd).Append(", ")
            .Append("FE.").Append(ClsDebitos_ItFacEstadoDec.SstrNombreCampoBd).Append(", ")
            .Append("FE.").Append(ClsCreditos_ItFacEstadoDec.SstrNombreCampoBd).Append(", ")
            .Append("FE.").Append(ClsVlrItemFac_EstadoDec.SstrNombreCampoBd).Append(", ").Append("F.")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(", ").Append("F.")
            .Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" FROM ").Append(ClsFacturaEstado.SstrNombreTabla)
            .Append(" AS FE INNER JOIN ").Append(ClsFactura.SstrNombreTabla).Append(" AS F ON FE.")
            .Append(MstrCampoBdCarpeta).Append(" = F.").Append(MstrCampoBdCarpeta).Append(" AND FE.")
            .Append(MstrCampoBdCentroUtil).Append(" = F.").Append(MstrCampoBdCentroUtil)
            .Append(" AND FE.").Append(ClsPrefijoFacturaVivaStr.SstrNombreCampoBd).Append(" = F.")
            .Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(" AND FE.")
            .Append(ClsIdFacturaVivaEnt.SstrNombreCampoBd).Append(" = F.")
            .Append(ClsIdFacturaEnt.SstrNombreCampoBd)
            .Append(" WHERE FE.").Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCarpeta).Append(" AND FE.")
            .Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd)
            .Append(" = ").Append(GshrIdCentroUtil).Append(" AND FE.")
            .Append(ClsIdEstado_FactEstadoEnt.SstrNombreCampoBd)
            .Append(" >= ").Append(lstrIdEstadoCtaIni).Append(" AND FE.")
            .Append(ClsIdEstado_FactEstadoEnt.SstrNombreCampoBd).Append(" <= ").Append(lstrIdEstadoCtaFin)
            .Append(" ORDER BY ").Append(ClsIdEstado_FactEstadoEnt.SstrNombreCampoBd)
        End With
        Dim lstrExpSql = MstbExpresionSql.ToString
        Return lstrExpSql
    End Function
#End Region
#End Region
#Region "Predios"
    ' Predios por Sector
    Private Sub SGenereDataSetPrediosSector(adsPrediosSector As DataSet)
        Dim lstrExpSelectPredios = FstrExpSqlPrediosSector()
        Dim lstrExpSelectSectores = FstrExpSqlSectores()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrExpSelectPredios)
        lcolExpresionesSql.Add(lstrExpSelectSectores)
        lcolNombresTablas.Add("OriPredios")
        lcolNombresTablas.Add("OriSectores")
        GobjPanDat.SdsDataSet(adsPrediosSector, lcolExpresionesSql, lcolNombresTablas)
        adsPrediosSector.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "PrediosPorSector" & ".XML"
        'adsPrediosSector.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
        'adsPrediosSector.Dispose()
    End Sub
    Private Function FstrExpSqlPrediosSector() As String
        With MstbExpresionSql
            .Clear().Append("SELECT ").Append(ClsIdSector_PredioShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsReferenciaPagoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsAreaPredioDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsFactorPonderaCPDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsCoeficientePropiedadDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFichaCatastralStr.SstrNombreCampoBd)
            .Append(", ").Append(ClsIdMatriculaInmobiliariaStr.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsPredio.SstrNombreTabla).Append(" WHERE ").Append(ClsOrionCop.StrFiltroUbicacion)
            .Append(" ORDER BY ").Append(ClsIdSector_PredioShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioStr.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpSqlSectores() As String
        With MstbExpresionSql
            .Clear().Append("SELECT * FROM ").Append(ClsSector.SstrNombreTabla).Append(" WHERE ")
            .Append(ClsOrionCop.StrFiltroUbicacion).Append(" ORDER BY ")
            .Append(ClsIdSectorShr.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    'Predios Por Propietario
    Private Sub SGenereDataSetPrediosPropietario(adsPrediosPropietario As DataSet)
        Dim lstrExpSelectPredios = FstrExpSqlPrediosProp()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrExpSelectPredios)
        lcolNombresTablas.Add("OriPredios")
        GobjPanDat.SdsDataSet(adsPrediosPropietario, lcolExpresionesSql, lcolNombresTablas)
        adsPrediosPropietario.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "PrediosPorPropietario" & ".XML"
        'adsPrediosPropietario.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlPrediosProp() As String
        Dim lstrTablaPri = ClsPropietario.SstrNombreTabla
        Dim lstrTablaSec = ClsPredio.SstrNombreTabla
        Dim lstrCampSelPri As String() = {ClsIdCliente_PropDbl.SstrNombreCampoBd,
                ClsNombreCompleto_PropStr.SstrNombreCampoBd,
                ClsPorcentajePartiDbl.SstrNombreCampoBd}
        Dim lstrCampSelSec As String() = {ClsIdPredioAgrupadorStr.SstrNombreCampoBd,
                ClsIdPredioStr.SstrNombreCampoBd,
                ClsAreaPredioDec.SstrNombreCampoBd,
                ClsFactorPonderaCPDbl.SstrNombreCampoBd,
                ClsCoeficientePropiedadDec.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                ClsIdPredio_PropStr.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{ClsIdPredioAgrupadorStr.SstrNombreCampoBd, "ASC"},
                {"S." & ClsIdPredioStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsAreaPredioDec.SstrNombreCampoBd & " > 0"
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri,
                lstrCampSelPri, lstrTablaSec, lstrCampSelSec, lstrCampRelPri,
                lstrCampRelSec, lstrOrden, lstrFiltro, {})
        Return lstrExpSql
    End Function
    ' Propietarios por coeficiente de propiedad
    Private Sub SGenereDataSetPropietariosXCP(adsPropietariosXCP As DataSet, ablnResumido As Boolean)
        Dim lstrExpSelectPropXCP As String
        Dim lcolNombresTablas As New Collection
        If ablnResumido Then
            lstrExpSelectPropXCP = FstrExpSqlPropXCP_Res()
            lcolNombresTablas.Add("OriPropXCP_Res")
        Else
            lstrExpSelectPropXCP = FstrExpSqlPropXCP()
            lcolNombresTablas.Add("OriPropXCP")
        End If
        Dim lcolExpresionesSql As New Collection From {
                lstrExpSelectPropXCP}
        GobjPanDat.SdsDataSet(adsPropietariosXCP, lcolExpresionesSql, lcolNombresTablas)
        adsPropietariosXCP.Tables.Add(FdtbCentroUtilidad)
        Dim lstrNomArch As String
        If ablnResumido Then
            lstrNomArch = GstrTrayDatPrg & "ProietariosXCP_Res" & ".XML"
        Else
            lstrNomArch = GstrTrayDatPrg & "ProietariosXCP" & ".XML"
        End If
        adsPropietariosXCP.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlPropXCP() As String
        Dim lstrExpSql = "SELECT PR." & ClsIdCliente_PropDbl.SstrNombreCampoBd & ", " &
                ClsNombreCompleto_PropStr.SstrNombreCampoBd & ", P." &
                ClsIdPredioStr.SstrNombreCampoBd & ", " &
                ClsCoeficientePropiedadDec.SstrNombreCampoBd & " * " &
                ClsPorcentajePartiDbl.SstrNombreCampoBd & " * 100 AS CP FROM " &
                ClsPredio.SstrNombreTabla & " AS P INNER JOIN " & ClsPropietario.SstrNombreTabla &
                " AS PR ON P." & OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd & " = PR." &
                OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd & " AND P." &
                OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd & " = PR." &
                OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd & " AND P." &
                ClsIdPredioStr.SstrNombreCampoBd & " = PR." & ClsIdPredio_PropStr.SstrNombreCampoBd &
                " WHERE P." & OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd & " = " & GshrIdCarpeta &
                " AND P." & OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd & " = " & GshrIdCentroUtil
        Return lstrExpSql
    End Function
    Private Function FstrExpSqlPropXCP_Res() As String
        Dim lstrExpSql = "SELECT PR." & ClsIdCliente_PropDbl.SstrNombreCampoBd & ", " &
                ClsNombreCompleto_PropStr.SstrNombreCampoBd & ", SUM(" &
                ClsCoeficientePropiedadDec.SstrNombreCampoBd & " * " &
                ClsPorcentajePartiDbl.SstrNombreCampoBd & " * 100) AS CP FROM " &
                ClsPredio.SstrNombreTabla & " AS P INNER JOIN " & ClsPropietario.SstrNombreTabla &
                " AS PR ON P." & OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd & " = PR." &
                OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd & " AND P." &
                OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd & " = PR." &
                OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd & " AND P." &
                ClsIdPredioStr.SstrNombreCampoBd & " = PR." & ClsIdPredio_PropStr.SstrNombreCampoBd &
                " WHERE P." & OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd & " = " & GshrIdCarpeta &
                " AND P." & OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd & " = " &
                GshrIdCentroUtil & " GROUP BY " & ClsIdCliente_PropDbl.SstrNombreCampoBd &
                " ORDER BY CP DESC"
        Return lstrExpSql
    End Function
    ' Cuotas de Administración por propietario
    Private Sub SGenereDataSetCuotasPropietario(adsCuotasAdminPropietario As DataSet)
        Dim lstrExpSelectPrediosProp = FstrExpSqlPrediosPropResum()
        Dim lstrExpSelectCuotasPredios = FstrExpSqlCuotasAdminPredios()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrExpSelectPrediosProp)
        lcolExpresionesSql.Add(lstrExpSelectCuotasPredios)
        lcolNombresTablas.Add("OriPredios")
        lcolNombresTablas.Add("OriCuotasAdmin")
        GobjPanDat.SdsDataSet(adsCuotasAdminPropietario, lcolExpresionesSql, lcolNombresTablas)
        adsCuotasAdminPropietario.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "CuotasAdminPorPropietario" & ".XML"
        'adsCuotasAdminPropietario.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlCuotasAdminPredios() As String
        Dim lentIdAno = ObjParRepDocs.EntIdDocInicial
        Dim lstrCamposPri =
                {"DISTINCT " & ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd,
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd,
                ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd,
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd}
        Dim lstrCamposSec As String() = Array.Empty(Of String)()
        Dim lstrTablaPri = ClsItemProgramaFact.SstrNombreTabla
        Dim lstrTablaSec = ClsServicio.SstrNombreTabla
        Dim lstrCampoePriRel = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd}
        Dim lstrOrden = {{ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = "P." & PanL.ClsIdCarpetaShr.SstrNombreCampoBd & " = " & GshrIdCarpeta.ToString &
                " AND P." & PanL.ClsIdCentroUtilShr.SstrNombreCampoBd & " = " & GshrIdCentroUtil.ToString &
                " AND P." & ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & lentIdAno.ToString &
                " AND " & ClsEsAjusteBln.SstrNombreCampoBd & " = " & False.ToString
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri, lstrCamposPri,
                lstrTablaSec, lstrCamposSec, lstrCampoePriRel, lstrCamposRelSec, lstrOrden,
                lstrFiltro, Array.Empty(Of String)())
        Return lstrExpSql
    End Function
    Private Function FstrExpSqlPrediosPropResum() As String
        Dim lstrTablaPri = ClsPropietario.SstrNombreTabla
        Dim lstrTablaSec = ClsPredio.SstrNombreTabla
        Dim lstrCampSelPri As String() = {ClsIdCliente_PropDbl.SstrNombreCampoBd,
                ClsNombreCompleto_PropStr.SstrNombreCampoBd,
                ClsPorcentajePartiDbl.SstrNombreCampoBd}
        Dim lstrCampSelSec As String() = {ClsIdPredioStr.SstrNombreCampoBd,
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                ClsIdPredio_PropStr.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri
        Dim lstrorden As String(,) = {{ClsIdPredioAgrupadorStr.SstrNombreCampoBd, "ASC"},
                {ClsIdPredioStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri,
                lstrCampSelPri, lstrTablaSec, lstrCampSelSec, lstrCampRelPri,
                lstrCampRelSec, lstrorden, lstrFiltro, {})
        Return lstrExpSql
    End Function
    ' Items Programa Facturación
    Private Sub SGenereDataSetItemsProgFact(adsItemsProgFact As DataSet)
        Dim lstrSqlItemsProgFact = FstrExpItemsProgFact()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlItemsProgFact)
        lcolNombresTablas.Add("OriItemsProgFact")
        GobjPanDat.SdsDataSet(adsItemsProgFact, lcolExpresionesSql, lcolNombresTablas)
        adsItemsProgFact.Tables.Add(FdtbCentroUtilidad)
        SComplementeDtbItemsProgFact(adsItemsProgFact)
        'Dim lstrNomArch = GstrTrayDatPrg & "ItemsProgramaFact" & ".XML"
        'adsItemsProgFact.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
        'adsItemsProgFact.Dispose()
    End Sub
    Private Function FstrExpItemsProgFact() As String
        With MstbExpresionSql
            .Clear()
            .Append("SELECT ").Append(ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd).Append(", ")
            .Append(" '' AS Servicio").Append(", ")
            .Append(ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd).Append(", ")
            .Append("'' AS NombreCliente").Append(", ")
            .Append(ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPeriodoIni_ItemProgStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsCantidadPeriodosShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd).Append(", ")
            .Append(ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd).Append(" From ")
            .Append(ClsItemProgramaFact.SstrNombreTabla).Append(" WHERE ")
            .Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND ").Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCentroUtil).Append(" AND ").Append(ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd)
            .Append(" = ").Append(ShrIdAno).Append(" AND ")
            .Append(ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd).Append(" = ").Append(EntIdServicio)
            .Append(" ORDER BY ").Append(ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Sub SComplementeDtbItemsProgFact(adsItesmProgFact As DataSet)
        Dim ldtbItemsProgFact = adsItesmProgFact.Tables("OriItemsProgFact")
        Dim ldrwItemsProFact() = ldtbItemsProgFact.Select()
        Dim lcolServicios As Collection
        If ShrIdAno = 0 Then
            lcolServicios = GobjParametros.ColServiciosPer
        Else
            Dim lobjAno As ClsAno = GobjParametros.ColAnos(ShrIdAno.ToString)
            lcolServicios = lobjAno.ColServiciosAno
        End If
        Dim lstrKey = ShrIdAno.ToString & "," & EntIdServicio.ToString
        Dim lobjServicio As ClsServicio = lcolServicios(lstrKey)
        Dim lstrNombreSer = lobjServicio.ObjNombreServicioStr.ToString
        Dim ldblIdCliente As Double, lstrIdPredio As String, lstrNombreCliente = String.Empty
        Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
        Dim lobjPropietario As ClsPropietario
        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
        For Each ldrwItemProFac As DataRow In ldrwItemsProFact
            lstrIdPredio = ClsPanorama.FobjValorCampo(ldrwItemProFac(
                    ClsIdPredio_ItemFactStr.SstrNombreCampoBd), EnuTipoValor.enuString)
            If String.IsNullOrEmpty(lstrIdPredio) Then
                ldblIdCliente = ClsPanorama.FobjValorCampo(ldrwItemProFac(
                        ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd),
                        EnuTipoValor.enuDouble)
                lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente})
                If lobjCliente.BlnExiste Then
                    lstrNombreCliente = lobjCliente.ObjNombreCompletoStr.ToString
                End If
            Else
                lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrIdPredio})
                lobjPropietario = lobjPredio.ColPropietarios(1)
                lstrNombreCliente = lobjPropietario.ObjNombreCompleto_PropStr.ToString()
                If lobjPredio.ColPropietarios.Count > 1 Then
                    lstrNombreCliente &= " y otro(s)"
                End If
            End If
            ldrwItemProFac("Servicio") = lstrNombreSer
            ldrwItemProFac("NombreCliente") = lstrNombreCliente
        Next
    End Sub
    ' Paz y Salvo
    Private Sub SGenereDataSetPazYSalvo(adsPazYSalvo As DataSet, aobjPredio As ClsPredio)
        SRefresqueLogo()
        Dim lstrExpSqlPredio = FstrExpSqlPred(aobjPredio)
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrExpSqlPredio)
        lcolNombresTablas.Add("OriPredioProp")
        GobjPanDat.SdsDataSet(adsPazYSalvo, lcolExpresionesSql, lcolNombresTablas)
        adsPazYSalvo.Tables.Add(FdtbCentroUtilidad)
        Dim lstrIdPrediosPyS As String = FstrIdPrediosPyS(aobjPredio)
        SComplementeDSPazYsalvo(adsPazYSalvo, lstrIdPrediosPyS)
        'Dim lstrNomArch = GstrTrayDatPrg & "DsetPazYSalvo" & ".XML"
        'adsPazYSalvo.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlPred(aobjPredio As ClsPredio) As String
        Dim lstrTablaPri = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec = ClsPropietario.SstrNombreTabla
        Dim lstrCampPri = {"DISTINCT '' AS Predios"}
        Dim lstrCampSec = {ClsIdCliente_PropDbl.SstrNombreCampoBd,
                ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamRelPri = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCamRelSec = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                ClsIdPredio_PropStr.SstrNombreCampoBd}
        Dim lstrOrden = {{"", ""}}
        Dim lstrFiltroPred = " AND " &
                aobjPredio.ObjIdPredioAgrupadorStr.StrNombreCampoBD & " = '" &
                aobjPredio.ObjIdPredioAgrupadorStr.ToString() & "'"
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & lstrFiltroPred
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri,
                lstrCampPri, lstrTablaSec, lstrCampSec, lstrCamRelPri, lstrCamRelSec,
                lstrOrden, lstrFiltro, {})
        Return lstrExpSql
    End Function
    Private Sub SComplementeDSPazYsalvo(adsPazYSalvo As DataSet, astrPredios As String)
        Dim ldtbPredCli = adsPazYSalvo.Tables("OriPredioProp")
        Dim lTipoString As System.Type = System.Type.GetType("System.String")
        Dim ldclTipoId As New DataColumn("TipoId", lTipoString)
        Dim lstrPredio As String
        If astrPredios.Contains(" y ") Then
            lstrPredio = "de los Predios identificados con los nombres " & astrPredios &
                    ", los cuales forman "
        Else
            lstrPredio = "del Predio identificado con el nombre " & astrPredios &
                    ", el cual forma "
        End If
        ldtbPredCli.Columns.Add(ldclTipoId)
        Dim ldblIdCliente As Double = ClsPanorama.FobjValorCampo(ldtbPredCli(0)(
                ClsIdCliente_PropDbl.SstrNombreCampoBd), EnuTipoValor.enuDouble)
        Dim lstrNomTipoDocId = FstrNomTipoDocId(ldblIdCliente)
        For Each ldrwRes As DataRow In ldtbPredCli.Rows
            ldrwRes("TipoId") = lstrNomTipoDocId
            ldrwRes("Predios") = lstrPredio
        Next
        Dim ldtbCenUti = adsPazYSalvo.Tables("PanCentrosUtilidad")
        Dim ldtmFechaFinPer As Date = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
        ldtmFechaFinPer = DateSerial(ldtmFechaFinPer.Year, ldtmFechaFinPer.Month, ldtmFechaFinPer.Day)
        Dim lstrFehaHasta As String = ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaFinPer)
        ldtbCenUti(0)("FechaHasta") = lstrFehaHasta
    End Sub
    Private Function FstrNomTipoDocId(adblIdTercero As Double) As String
        Dim lstrTabla = ClsTercero.SstrNombreTabla
        Dim lstrCamSel = {ClsTipoDocIdentidadByt.SstrNombreCampoBd}
        Dim lstrOrden = {{"", ""}}
        Dim lstrFiltro = ClsIdTerceroDbl.SstrNombreCampoBd & " = " & adblIdTercero
        Dim ldtbTer = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSel, lstrOrden, lstrFiltro)
        Dim lbytTipoDocId As Byte = ClsPanorama.FobjValorCampo(
                ldtbTer(0)(ClsTipoDocIdentidadByt.SstrNombreCampoBd), EnuTipoValor.enuByte)
        Dim lstrNomTipoDocId = FstrNombreTipoDocId(lbytTipoDocId)
        Return lstrNomTipoDocId
    End Function
    Private Function FstrIdPrediosPyS(aobjPredio As ClsPredio) As String
        Dim lstrIdPredio As String, lstrIdPrediosPyS = String.Empty
        Dim lblnUltimo As Boolean, i = 0
        If aobjPredio.ObjIdPredioStr.ObjValorPro = aobjPredio.ObjIdPredioAgrupadorStr.ObjValorPro Then
            For Each lobjPredio As ClsPredio In aobjPredio.ColPrediosAgrupados
                i += 1
                lblnUltimo = i = aobjPredio.ColPrediosAgrupados.Count - 1
                lstrIdPredio = lobjPredio.ObjIdPredioStr.ToString()
                If aobjPredio.FdecSaldoDeudaPredio(lstrIdPredio) = 0 Then
                    If aobjPredio.ColPrediosAgrupados.Count > 1 Then
                        If lblnUltimo Then
                            lstrIdPrediosPyS &= lstrIdPredio & " y "
                        Else
                            lstrIdPrediosPyS &= lstrIdPredio & ", "
                        End If
                    Else
                        lstrIdPrediosPyS &= lstrIdPredio
                    End If
                End If
            Next
            If lstrIdPrediosPyS.EndsWith(", ") Then
                lstrIdPrediosPyS = lstrIdPrediosPyS.Substring(0, lstrIdPrediosPyS.Length - 2)
            End If
        Else
            lstrIdPrediosPyS = aobjPredio.ObjIdPredioStr.ToString()
        End If
        Return lstrIdPrediosPyS
    End Function
#End Region
#Region "Caja y Bancos"
    ' DataSet Caja y Bancos
    Private Sub SGenereDataSetMediosPago(adsCajaBancos As DataSet)
        Dim lstrSqlMediosPago = FstrExpresionMediosPago()
        Dim lstrSqlTotalesMO = FstrExpresionTotalesMedPago()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlMediosPago)
        lcolExpresionesSql.Add(lstrSqlTotalesMO)
        lcolNombresTablas.Add("OriMediosPago")
        lcolNombresTablas.Add("OriTotalesMediosPago")
        GobjPanDat.SdsDataSet(adsCajaBancos, lcolExpresionesSql, lcolNombresTablas)
        adsCajaBancos.Tables.Add(FdtbCentroUtilidad)
        SComplementeDtbMediosPago(adsCajaBancos)
        SComplementeDtbTotalesMediosPago(adsCajaBancos)
        Dim ldtbReintegroAnt = FdtbReintegroAnt()
        SIntegereAnticiposReint(adsCajaBancos, ldtbReintegroAnt)
        'Dim lstrNomArch = GstrTrayDatPrg & "CajaBancos" & ".XML"
        'adsCajaBancos.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Shared Sub SComplementeDtbMediosPago(adsCajaBancos As DataSet)
        Dim ldtbMediosPago = adsCajaBancos.Tables("OriMediosPago")
        Dim ldrwMediosPago = ldtbMediosPago.Select()
        If ldrwMediosPago.Length > 0 Then
            Dim lenuTipoMedioPago As EnuTipoMedioPagoDef
            Dim lstrNombreTipoMedPago As String
            For Each ldrwMedioPago As DataRow In ldrwMediosPago
                lenuTipoMedioPago = ClsPanorama.FobjValorCampo(ldrwMedioPago(ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd),
                        EnuTipoValor.enuByte)
                lstrNombreTipoMedPago = ClsOrionCop.FstrNombreDatoConstanteOri(EnuGrupoConstantesOriDef.EnuMediosPago,
                        lenuTipoMedioPago)
                ldrwMedioPago("TipoMedioPago") = lstrNombreTipoMedPago
            Next
        End If
    End Sub
    Private Shared Sub SComplementeDtbTotalesMediosPago(adsCajaBancos As DataSet)
        Dim ldtbTotalesMediosPago = adsCajaBancos.Tables("OriTotalesMediosPago")
        Dim ldrwMediosPago = ldtbTotalesMediosPago.Select()
        If ldrwMediosPago.Length > 0 Then
            Dim lenuTipoMedioPago As EnuTipoMedioPagoDef
            Dim lstrNombreTipoMedPago As String
            For Each ldrwMedioPago As DataRow In ldrwMediosPago
                lenuTipoMedioPago = ClsPanorama.FobjValorCampo(ldrwMedioPago(
                        ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd),
                        EnuTipoValor.enuByte)
                lstrNombreTipoMedPago = ClsOrionCop.FstrNombreDatoConstanteOri(
                        EnuGrupoConstantesOriDef.EnuMediosPago,
                        lenuTipoMedioPago)
                ldrwMedioPago("TipoMedioPago") = lstrNombreTipoMedPago
            Next
        End If
    End Sub
    Private Function FstrExpresionMediosPago() As String
        With MstbExpresionSql
            .Clear()
            .Append("SELECT 'Rec. Caja' AS Dcto, NroRecCaja, R.")
            .Append(ClsFechaRecDtm.SstrNombreCampoBd)
            .Append(", R.").Append(ClsFechaCreacionDtm.SstrNombreCampoBd)
            .Append(", R.").Append(ClsIdCliente_RecDbl.SstrNombreCampoBd)
            .Append(", ").Append(ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd)
            .Append(", ").Append(ClsNombreCompletoStr.SstrNombreCampoBd)
            .Append(", ").Append(ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd)
            .Append(", ").Append(ClsIdCtaContabIngresoStr.SstrNombreCampoBd)
            .Append(", ").Append("'' AS TipoMedioPago").Append(", SUM(")
            .Append(ClsValor_MedPagoDec.SstrNombreCampoBd).Append(") As Valor")
            .Append(" FROM (SELECT RC.").Append(MstrCampoBdCarpeta).Append(", ")
            .Append("RC.").Append(MstrCampoBdCentroUtil).Append(", ")
            .Append("CONCAT(RC.").Append(ClsPrefijo_RecStr.SstrNombreCampoBd)
            .Append(",'-',RC.").Append(ClsIdRecCajaEnt.SstrNombreCampoBd)
            .Append(") AS NroRecCaja,").Append("RC.")
            .Append(ClsFechaRecDtm.SstrNombreCampoBd).Append(", ").Append("RC.")
            .Append(ClsFechaCreacionDtm.SstrNombreCampoBd).Append(", ")
            .Append("RC.").Append(ClsIdCliente_RecDbl.SstrNombreCampoBd).Append(", ")
            .Append("RC.").Append(ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd)
            .Append(", C.").Append(ClsNombreCompletoStr.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsReciboCaja.SstrNombreTabla)
            .Append(" AS RC INNER JOIN ").Append(ClsCliente.SstrNombreTabla)
            .Append(" AS C ON RC.").Append(MstrCampoBdCarpeta).Append(" = C.")
            .Append(MstrCampoBdCarpeta).Append(" AND RC.").Append(MstrCampoBdCentroUtil)
            .Append(" = C.").Append(MstrCampoBdCentroUtil).Append(" AND RC.")
            .Append(ClsIdCliente_RecDbl.SstrNombreCampoBd).Append(" = C.")
            .Append(ClsIdClienteDbl.SstrNombreCampoBd).Append(" WHERE RC.")
            .Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND RC.").Append(MstrCampoBdCentroUtil).Append(" = ")
            .Append(GshrIdCentroUtil).Append(" AND ").Append(ClsAnuladoBln.SstrNombreCampoBd)
            .Append(" = FALSE AND ")
            .Append(ClsFechaRecDtm.SstrNombreCampoBd).Append(" BETWEEN '")
            .Append(StrFechaDesde).Append("' AND '").Append(StrFechaHasta).Append("'")
            .Append(") AS R INNER JOIN ")
            .Append(ClsMedioPago.SstrNombreTabla).Append(" AS M ON R.")
            .Append(MstrCampoBdCarpeta).Append(" = M.").Append(MstrCampoBdCarpeta)
            .Append(" AND R.").Append(MstrCampoBdCentroUtil)
            .Append(" = M.").Append(MstrCampoBdCentroUtil)
            .Append(" AND R.NroRecCaja = CONCAT(M.")
            .Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(",'-',M.")
            .Append(ClsIdRecCaja_MedPagoEnt.SstrNombreCampoBd).Append(") WHERE R.")
            .Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND R.").Append(MstrCampoBdCentroUtil).Append(" = ")
            .Append(GshrIdCentroUtil).Append(" GROUP BY NroRecCaja, ")
            .Append(ClsFechaRecDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaCreacionDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdCliente_RecDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdCtaContabIngresoStr.SstrNombreCampoBd).Append(", ")
            .Append("TipoMedioPago").Append(" ORDER BY ")
            .Append(ClsFechaRecDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd).Append(", ")
            .Append("NroRecCaja")
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FdtbReintegroAnt() As DataTable
        Dim lstrCtaCaja = GobjParametros.ObjIdCtaCajaStr.ToString()
        Dim lstrTabla = ClsNotaDevAnt.SstrNombreTabla
        Dim lstrCampSel As String() = {"'Rein. Ant.' AS Dcto",
                ClsPrefijo_NotaDevAntStr.SstrNombreCampoBd & " + '-' + " &
                ClsIdNotaDevAntEnt.SstrNombreCampoBd & " AS NroRecCaja ",
                ClsFecha_NotaDevAntDtm.SstrNombreCampoBd,
                ClsFechaCreacionDtm.SstrNombreCampoBd,
                ClsIdCliente_NotaDevAntDbl.SstrNombreCampoBd,
                ClsIdPredioAgrupador_NotaDevAntStr.SstrNombreCampoBd,
                "'Anticipo reintegrado' AS " & ClsNombreCompletoStr.SstrNombreCampoBd,
                EnuTipoMedioPagoDef.EnuEfectivo & " AS " &
                ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd, "'" & lstrCtaCaja &
                "' AS " & ClsIdCtaContabIngresoStr.SstrNombreCampoBd,
                "'Efectivo' AS TipoMedioPago", ClsValor_NotaDevAntDec.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{ClsFecha_NotaDevAntDtm.SstrNombreCampoBd, "ASC"},
                {"NroRecCaja", "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsFecha_NotaDevAntDtm.SstrNombreCampoBd & " BETWEEN '" & StrFechaDesde &
                "' AND '" & StrFechaHasta & "'"
        Dim ldtbReintAnt = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSel, lstrOrden,
                lstrFiltro)
        Return ldtbReintAnt
    End Function
    Private Function FstrExpresionTotalesMedPago() As String
        ' Filtro Ubicación
        With MstbExpresionSql
            .Clear().Append("M.").Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCarpeta.ToString).Append(" AND M.")
            .Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCentroUtil.ToString)
        End With
        Dim lstrFiltroUbicacion = MstbExpresionSql.ToString
        With MstbExpresionSql
            .Clear()
            .Append("SELECT ").Append("SUM(M.").Append(ClsValor_MedPagoDec.SstrNombreCampoBd)
            .Append(") AS Valor, M.").Append(ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd)
            .Append(", '' as TipoMedioPago")
            .Append(" FROM ").Append(ClsReciboCaja.SstrNombreTabla).Append(" AS R INNER JOIN ")
            .Append(ClsMedioPago.SstrNombreTabla).Append(" AS M ON R.")
            .Append(MstrCampoBdCarpeta).Append(" = M.").Append(MstrCampoBdCarpeta).Append(" AND R.")
            .Append(MstrCampoBdCentroUtil).Append(" = M.").Append(MstrCampoBdCentroUtil)
            .Append(" AND R.").Append(ClsPrefijo_RecStr.SstrNombreCampoBd)
            .Append(" = M.").Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(" AND R.")
            .Append(ClsIdRecCajaEnt.SstrNombreCampoBd).Append(" = M.").Append(ClsIdRecCaja_MedPagoEnt.SstrNombreCampoBd)
            .Append(" WHERE ").Append(lstrFiltroUbicacion).Append(" AND ")
            .Append(ClsAnuladoBln.SstrNombreCampoBd).Append(" = FALSE AND ")
            .Append(ClsFechaRecDtm.SstrNombreCampoBd).Append(" >= '").Append(StrFechaDesde).Append("' AND ")
            .Append(ClsFechaRecDtm.SstrNombreCampoBd).Append(" <= '").Append(StrFechaHasta).Append("'")
            .Append(" GROUP BY M.").Append(ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd)
            .Append(" ORDER BY M.").Append(ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Sub SIntegereAnticiposReint(adsCajaBancoas As DataSet,
            adtbReintegroAnt As DataTable)
        Dim ldtbMediodPago = adsCajaBancoas.Tables("OriMediosPago")
        Dim lstrDocto As String, lstrNro As String, ldtmFecha As Date
        Dim ldtmFecCrea As Date, ldblIdCliente As Double, lstrIdpredioAgr As String
        Dim lstrDeta As String, lbytIdTipoMP As Byte, lstrIdCtaCaja As String
        Dim lstrTipoMP As String, ldecValor As Decimal
        For Each ldrwReiAnt As DataRow In adtbReintegroAnt.Rows
            lstrDocto = ldrwReiAnt("Dcto")
            lstrNro = ldrwReiAnt("NroRecCaja")
            ldtmFecha = ldrwReiAnt(ClsFecha_NotaDevAntDtm.SstrNombreCampoBd)
            ldtmFecCrea = ldrwReiAnt(ClsFechaCreacionDtm.SstrNombreCampoBd)
            ldblIdCliente = ldrwReiAnt(ClsIdCliente_NotaDevAntDbl.SstrNombreCampoBd)
            lstrIdpredioAgr = ldrwReiAnt(
                    ClsIdPredioAgrupador_NotaDevAntStr.SstrNombreCampoBd)
            lstrDeta = ldrwReiAnt(ClsNombreCompletoStr.SstrNombreCampoBd)
            lbytIdTipoMP = ldrwReiAnt(ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd)
            lstrIdCtaCaja = ldrwReiAnt(ClsIdCtaContabIngresoStr.SstrNombreCampoBd)
            lstrTipoMP = ldrwReiAnt("TipoMedioPago")
            ldecValor = ldrwReiAnt(ClsValor_NotaDevAntDec.SstrNombreCampoBd) * -1
            Dim ldrNewMedPag = ldtbMediodPago.NewRow
            ldrNewMedPag("Dcto") = lstrDocto
            ldrNewMedPag("NroRecCaja") = lstrNro
            ldrNewMedPag(ClsFechaRecDtm.SstrNombreCampoBd) = ldtmFecha
            ldrNewMedPag(ClsFechaCreacionDtm.SstrNombreCampoBd) = ldtmFecCrea
            ldrNewMedPag(ClsIdCliente_RecDbl.SstrNombreCampoBd) = ldblIdCliente
            ldrNewMedPag(ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd) =
                    lstrIdpredioAgr
            ldrNewMedPag(ClsNombreCompletoStr.SstrNombreCampoBd) = lstrDeta
            ldrNewMedPag(ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd) = lbytIdTipoMP
            ldrNewMedPag(ClsIdCtaContabIngresoStr.SstrNombreCampoBd) = lstrIdCtaCaja
            ldrNewMedPag("TipoMedioPago") = lstrTipoMP
            ldrNewMedPag(ClsValor_MedPagoDec.SstrNombreCampoBd) = ldecValor
            ldtbMediodPago.Rows.Add(ldrNewMedPag)
        Next
    End Sub
#End Region
#Region "Recibos de Caja entre Fechas"
    Private Sub SGenereDataSetRCFechas(adsRCFechas As DataSet)
        Dim lstrSqlRCFechas = FstrExpRCFechas()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlRCFechas)
        lcolNombresTablas.Add("OriRecibosCaja")
        GobjPanDat.SdsDataSet(adsRCFechas, lcolExpresionesSql, lcolNombresTablas)
        adsRCFechas.Tables.Add(FdtbCentroUtilidad)
        SComplementeDtbRCFechas(adsRCFechas)
        'Dim lstrNomArch = GstrTrayDatPrg & "RCFechas" & ".XML"
        'adsRCFechas.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpRCFechas() As String
        Dim lstrTabla = ClsReciboCaja.SstrNombreTabla
        Dim lstrCamposSelect = {ClsFechaCreacionDtm.SstrNombreCampoBd,
                                ClsFechaRecDtm.SstrNombreCampoBd,
                                ClsPrefijo_RecStr.SstrNombreCampoBd,
                                ClsIdRecCajaEnt.SstrNombreCampoBd, "'' as NroRC",
                                ClsIdCliente_RecDbl.SstrNombreCampoBd,
                                ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd,
                                PanL.ClsIdUsuarioStr.SstrNombreCampoBd,
                                ClsAnuladoBln.SstrNombreCampoBd,
                                ClsIdUsuarioAnuloStr.SstrNombreCampoBd,
                                ClsValor_RecDec.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsFechaCreacionDtm.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsFechaCreacionDtm.SstrNombreCampoBd &
                " > '" & StrFechaDesde & "' AND " & ClsFechaCreacionDtm.SstrNombreCampoBd &
                " < '" & StrFechaHasta & "'"
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, Array.Empty(Of String)())
        Return lstrSql
    End Function
    Private Shared Sub SComplementeDtbRCFechas(adsRCFechas As DataSet)
        Dim ldtbRCFechas As DataTable = adsRCFechas.Tables("OriRecibosCaja")
        Dim lstrPref As String, lentIdRC As Integer, lstrNroRC As String
        For Each ldrwRC As DataRow In ldtbRCFechas.Rows
            lstrPref = ClsPanorama.FobjValorCampo(ldrwRC(ClsPrefijo_RecStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
            lentIdRC = ClsPanorama.FobjValorCampo(ldrwRC(ClsIdRecCajaEnt.SstrNombreCampoBd),
                        EnuTipoValor.enuInteger)
            lstrNroRC = ClsPanorama.FstrNumeroDcto(lstrPref, lentIdRC)
            ldrwRC("NroRC") = lstrNroRC
        Next
    End Sub
#End Region
#Region "Recibos de Caja reversados"
    Private Sub SGenereDataSetRecCajaRev(adsRecCajaRev As DataSet)
        Dim lstrSqlMedPagoRev = FstrExpMedPagoReversados()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlMedPagoRev)
        lcolNombresTablas.Add("OriMediosPago")
        GobjPanDat.SdsDataSet(adsRecCajaRev, lcolExpresionesSql, lcolNombresTablas)
        adsRecCajaRev.Tables.Add(FdtbCentroUtilidad)
        SComplementeDtbMediosPago(adsRecCajaRev)
        Dim lstrNomArch = GstrTrayDatPrg & "RecCajaReversados" & ".XML"
        adsRecCajaRev.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpMedPagoReversados() As String
        With MstbExpresionSql
            .Clear()
            .Append("SELECT NroNotaRCr, NroRecCaja, ")
            .Append(ClsFecha_NotaReversaCrDtm.SstrNombreCampoBd).Append(",  R.")
            .Append(ClsFechaCreacionDtm.SstrNombreCampoBd).Append(", R.")
            .Append(ClsIdCliente_NotaReversaCrDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgr_NotaReversaCrStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd).Append(", ")
            .Append("'' AS TipoMedioPago, ").Append(ClsIdCtaContabIngresoStr.SstrNombreCampoBd)
            .Append(", SUM(").Append(ClsValor_MedPagoDec.SstrNombreCampoBd)
            .Append(") As Valor").Append(" FROM (SELECT NR.").Append(MstrCampoBdCarpeta).Append(", ")
            .Append("NR.").Append(MstrCampoBdCentroUtil).Append(", ")
            .Append("CONCAT(NR.").Append(ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd)
            .Append(",'-',NR.").Append(ClsIdNotaReversaCrEnt.SstrNombreCampoBd)
            .Append(") AS NroNotaRCr, ").Append("CONCAT(NR.")
            .Append(ClsPrefijoDoc_NotaReversaCrStr.SstrNombreCampoBd).Append(", '-', NR.")
            .Append(ClsIdDoc_NotaReversaCrEnt.SstrNombreCampoBd).Append(") As NroRecCaja, NR.")
            .Append(ClsFecha_NotaReversaCrDtm.SstrNombreCampoBd).Append(", NR.")
            .Append(ClsFechaCreacionDtm.SstrNombreCampoBd).Append(", NR.")
            .Append(ClsIdCliente_NotaReversaCrDbl.SstrNombreCampoBd).Append(", NR.")
            .Append(ClsIdPredioAgr_NotaReversaCrStr.SstrNombreCampoBd).Append(", C.")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(" FROM ")
            .Append(ClsNotaReversionCr.SstrNombreTabla).Append(" AS NR INNER JOIN ")
            .Append(ClsCliente.SstrNombreTabla).Append(" AS C ON NR.")
            .Append(MstrCampoBdCarpeta).Append(" = C.").Append(MstrCampoBdCarpeta).Append(" AND NR.")
            .Append(MstrCampoBdCentroUtil).Append(" = C.").Append(MstrCampoBdCentroUtil)
            .Append(" AND NR.").Append(ClsIdCliente_NotaReversaCrDbl.SstrNombreCampoBd)
            .Append(" = C.").Append(ClsIdClienteDbl.SstrNombreCampoBd).Append(" WHERE NR.")
            .Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta).Append(" AND NR.")
            .Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil).Append(" AND ")
            .Append(ClsTipoDocReversadoByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuDocReversado.EnuReciboC).Append(" AND ")
            .Append(ClsFecha_NotaReversaCrDtm.SstrNombreCampoBd).Append(" BETWEEN '")
            .Append(StrFechaDesde).Append("' AND '").Append(StrFechaHasta)
            .Append("') AS R INNER JOIN ").Append(ClsMedioPago.SstrNombreTabla).Append(" AS M ON R.")
            .Append(MstrCampoBdCarpeta).Append(" = M.").Append(MstrCampoBdCarpeta).Append(" AND R.")
            .Append(MstrCampoBdCentroUtil).Append(" = M.").Append(MstrCampoBdCentroUtil)
            .Append(" AND R.NroRecCaja = ").Append("CONCAT(M.")
            .Append(ClsPrefijo_RecStr.SstrNombreCampoBd).Append(",'-',M.")
            .Append(ClsIdRecCaja_MedPagoEnt.SstrNombreCampoBd).Append(") WHERE R.")
            .Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta).Append(" AND R.")
            .Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil)
            .Append(" GROUP BY NroNotaRCr, NroRecCaja, ")
            .Append(ClsFecha_NotaReversaCrDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaCreacionDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdCliente_NotaReversaCrDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgr_NotaReversaCrStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdCtaContabIngresoStr.SstrNombreCampoBd).Append(", ")
            .Append("TipoMedioPago ORDER BY ").Append(ClsFecha_NotaReversaCrDtm.SstrNombreCampoBd)
            .Append(", ").Append(ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd).Append(", ")
            .Append("NroNotaRCr")
        End With
        Return MstbExpresionSql.ToString()
    End Function
#End Region
#Region "Informe Diario"
    ' DataSet Informe Diario
    Private Sub SGenereDataSetInfDiario(adsInfDiario As DataSet)
        Dim lstrSqlIdFacturasDia = FstrExpresionIdFacturasDia()
        Dim lstrSqlVentasSegunIva = FstrExpVentasSegunIva()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlIdFacturasDia)
        lcolExpresionesSql.Add(lstrSqlVentasSegunIva)
        lcolNombresTablas.Add("OriIdFacturasDias")
        lcolNombresTablas.Add("OriVentasSegunIva")
        GobjPanDat.SdsDataSet(adsInfDiario, lcolExpresionesSql, lcolNombresTablas)
        adsInfDiario.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "InformeDiario" & ".XML"
        'adsInfDiario.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpresionIdFacturasDia() As String
        With MstbExpresionSql
            .Clear()
            .Append("SELECT ").Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(", ")
            .Append("MIN(").Append(ClsIdFacturaEnt.SstrNombreCampoBd).Append(") AS FacIni").Append(", ")
            .Append("MAX(").Append(ClsIdFacturaEnt.SstrNombreCampoBd).Append(") AS FacFin")
            .Append(" FROM ").Append(ClsFactura.SstrNombreTabla)
            .Append(" WHERE ").Append(ClsOrionCop.StrFiltroUbicacion).Append(" AND ")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(" >= '").Append(StrFechaDesde).Append("' AND ")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(" <= '").Append(StrFechaHasta).Append("'")
            .Append(" GROUP BY ").Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(" ORDER BY ")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(", ").Append(ClsPrefijo_FactStr.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpVentasSegunIva() As String
        With MstbExpresionSql
            .Clear().Append("F.").Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCarpeta.ToString).Append(" AND F.")
            .Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCentroUtil.ToString)
        End With
        Dim lstrFiltroUbica = MstbExpresionSql.ToString
        With MstbExpresionSql
            .Clear()
            .Append("SELECT F.").Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(", ")
            .Append("FI.").Append(ClsEsExcluidoIva_ItemFactBln.SstrNombreCampoBd).Append(", ")
            .Append("FI.").Append(ClsTarifaIva_ItemFactDbl.SstrNombreCampoBd).Append(", ")
            .Append("SUM(FI.").Append(ClsValor_ItemFactDec.SstrNombreCampoBd).Append(") AS Valor")
            .Append(" FROM ").Append(ClsFactura.SstrNombreTabla).Append(" AS F INNER JOIN ")
            .Append(ClsItemFactura.SstrNombreTabla).Append(" AS FI ON F.")
            .Append(MstrCampoBdCarpeta).Append(" = FI.").Append(MstrCampoBdCarpeta).Append(" AND F.")
            .Append(MstrCampoBdCentroUtil).Append(" = FI.").Append(MstrCampoBdCentroUtil)
            .Append(" AND F.").Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(" = FI.")
            .Append(ClsPrefijo_ItemFactStr.SstrNombreCampoBd).Append(" AND F.")
            .Append(ClsIdFacturaEnt.SstrNombreCampoBd).Append(" = FI.")
            .Append(ClsIdFactura_ItemFactEnt.SstrNombreCampoBd)
            .Append(" WHERE ").Append(lstrFiltroUbica).Append(" AND ")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(" >= '").Append(StrFechaDesde).Append("' AND ")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(" <= '").Append(StrFechaHasta).Append("'")
            .Append(" GROUP BY ").Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsEsExcluidoIva_ItemFactBln.SstrNombreCampoBd).Append(", ")
            .Append(ClsTarifaIva_ItemFactDbl.SstrNombreCampoBd).Append(" ORDER BY ")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsEsExcluidoIva_ItemFactBln.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
#End Region
#Region "Clientes"
    'Cartera Por Cliente
    Private Sub SGenereDataSetCarPorCli(adsCarteraPorCliente As DataSet)
        Dim lstrSqlCarteraPorCliente = FstrExpCarteraPorCliente()
        Dim lstrSqlClientes = FstrExpClientes()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlCarteraPorCliente)
        lcolExpresionesSql.Add(lstrSqlClientes)
        lcolNombresTablas.Add("OriCarteraPorCliente")
        lcolNombresTablas.Add("OriClientes")
        GobjPanDat.SdsDataSet(adsCarteraPorCliente, lcolExpresionesSql, lcolNombresTablas)
        adsCarteraPorCliente.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "CarteraPorCliente" & ".XML"
        'adsCarteraPorCliente.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpCarteraPorCliente() As String 'Ok TipoNov
        ' Expresión Filtro Debitos
        Dim lstrCampoTipoNov = ClsIdTipoNovedadByt.SstrNombreCampoBd
        With MstbExpresionSql
            .Clear().Append("(")
            .Append(lstrCampoTipoNov).Append(" BETWEEN ").Append(EnuTipoNov.EnuDbCap)
            .Append(" AND ").Append(EnuTipoNov.EnuDbInt).Append(" OR ")
            .Append(lstrCampoTipoNov).Append(" BETWEEN ").Append(EnuTipoNov.EnuRCrPagoCap)
            .Append(" AND ").Append(EnuTipoNov.EnuRCrRetCre).Append(" OR ")
            .Append(lstrCampoTipoNov).Append(" = ").Append(EnuTipoNov.EnuDbIvaInt).Append(")")
        End With
        Dim lstrFiltroDb = MstbExpresionSql.ToString
        ' Expresión Filtro Creditos
        With MstbExpresionSql
            .Clear().Append("(")
            .Append(lstrCampoTipoNov).Append(" BETWEEN ").Append(EnuTipoNov.EnuCrPagoCap)
            .Append(" AND ").Append(EnuTipoNov.EnuCrRetCre).Append(" OR ")
            .Append(lstrCampoTipoNov).Append(" BETWEEN ").Append(EnuTipoNov.EnuRDbCap)
            .Append(" AND ").Append(EnuTipoNov.EnuRDbInt).Append(" OR ")
            .Append(lstrCampoTipoNov).Append(" = ").Append(EnuTipoNov.EnuCrIvaGas).Append(" OR ")
            .Append(lstrCampoTipoNov).Append(" = ").Append(EnuTipoNov.EnuRDbIvaInt).Append(")")
        End With
        Dim lstrFiltroCr = MstbExpresionSql.ToString
        ' Filtro Ubicación
        With MstbExpresionSql
            .Clear().Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCarpeta.ToString).Append(" AND ")
            .Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCentroUtil.ToString)
        End With
        Dim lstrFiltroUbicacion = MstbExpresionSql.ToString
        With MstbExpresionSql
            .Clear()
            .Append("SELECT ").Append(ClsIdTercero_NovDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd).Append(", ")
            .Append("'").Append(StrFechaHasta).Append("' as Fecha").Append(", ")
            .Append("SUM(IF(").Append(lstrFiltroDb).Append(", ").Append(ClsValor_NovDec.SstrNombreCampoBd)
            .Append(", 0) - IF(").Append(lstrFiltroCr).Append(", ").Append(ClsValor_NovDec.SstrNombreCampoBd)
            .Append(", 0)) AS Valor FROM ").Append(ClsNovedad.SstrNombreTabla).Append(" WHERE ")
            .Append(lstrFiltroUbicacion).Append(" AND ")
            .Append(ClsFechaNovedadDtm.SstrNombreCampoBd).Append(" <= '").Append(StrFechaHasta).Append("'")
            .Append(" GROUP BY ").Append(ClsIdTercero_NovDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd).Append(", ").Append("Fecha")
            .Append(" ORDER BY ").Append(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpClientes() As String
        Dim lstrTablaPri = ClsNovedad.SstrNombreTabla
        Dim lstrTablaSec = ClsCliente.SstrNombreTabla
        Dim lstrCamposPri = {"DISTINCT " & ClsIdTercero_NovDbl.SstrNombreCampoBd}
        Dim lstrCamposSec = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                                ClsIdTercero_NovDbl.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsIdClienteDbl.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = "P." & PanL.ClsIdCarpetaShr.SstrNombreCampoBd & " = " & GshrIdCarpeta & " AND P." &
                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd & " = " & GshrIdCentroUtil & " AND " &
                ClsFechaNovedadDtm.SstrNombreCampoBd & " <= '" & StrFechaHasta & "'"
        Dim lstrCamposGrupo As String() = Array.Empty(Of String)()
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri, lstrCamposPri, lstrTablaSec,
                lstrCamposSec, lstrCamposRelPri, lstrCamposRelSec, lstrIndice,
                lstrFiltro, lstrCamposGrupo)
        Return lstrSql
    End Function
    ' Cartera por predio Agrupador
    Private Sub SGenereDataSetCarPorPredioAgr(adsCarteraPorPredioAgr As DataSet)
        Dim lstrSqlCarteraPorPredioAgr = FstrExpCarteraPorPredioAgr()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlCarteraPorPredioAgr)
        lcolNombresTablas.Add("OriCarteraPorPredioAgr")
        GobjPanDat.SdsDataSet(adsCarteraPorPredioAgr, lcolExpresionesSql, lcolNombresTablas)
        adsCarteraPorPredioAgr.Tables.Add(FdtbCentroUtilidad)
        SComplementePredioAgr(adsCarteraPorPredioAgr)
        'Dim lstrNomArch = GstrTrayDatPrg & "CarteraPorPredioAgr" & ".XML"
        'adsCarteraPorPredioAgr.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpCarteraPorPredioAgr() As String
        With MstbExpresionSql
            .Clear().Append("SELECT ")
            .Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd).Append(", ")
            .Append("D." & ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", Saldo FROM (")
            .Append("SELECT ").Append(OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(", ")
            .Append(OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd)
            .Append(", ").Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(", ")
            .Append("SUM(").Append(ClsDebitos_FactDec.SstrNombreCampoBd).Append(" - ")
            .Append(ClsCreditos_FactDec.SstrNombreCampoBd).Append(") AS Saldo")
            .Append(" FROM ").Append(ClsFactura.SstrNombreTabla).Append(" WHERE ")
            .Append(ClsOrionCop.StrFiltroUbicacion).Append(" AND ")
            .Append(ClsDebitos_FactDec.SstrNombreCampoBd)
            .Append(" <> ").Append(ClsCreditos_FactDec.SstrNombreCampoBd)
            .Append(" AND ").Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd)
            .Append(" <> ''  GROUP BY ")
            .Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(") AS D INNER JOIN ")
            .Append(ClsCliente.SstrNombreTabla).Append(" AS C ON D.")
            .Append(OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = C.")
            .Append(OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" AND D.")
            .Append(OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = C.")
            .Append(OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" AND D.")
            .Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(" = C.")
            .Append(ClsIdClienteDbl.SstrNombreCampoBd).Append(" ORDER BY ")
            .Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdCliente_FactDbl.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Shared Sub SComplementePredioAgr(adsCarteraPorPredioAgr As DataSet)
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCamSel As String() = {"'Deudas de clientes sin predio' AS " &
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd, "0 AS " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd, "'Varios' AS " &
                ClsNombreCompletoStr.SstrNombreCampoBd, "SUM(" &
                ClsDebitos_FactDec.SstrNombreCampoBd & " - " &
                ClsCreditos_FactDec.SstrNombreCampoBd & ") AS Saldo "}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsDebitos_FactDec.SstrNombreCampoBd & " <> " &
                ClsCreditos_FactDec.SstrNombreCampoBd & " AND " &
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " = ''"
        Dim ldtbComp = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSel, {{"", ""}},
                lstrFiltro)
        Dim ldtbCartPorPre = adsCarteraPorPredioAgr.Tables("OriCarteraPorPredioAgr")
        If ldtbComp.Rows.Count > 0 Then
            Dim ldrwComp As DataRow = ldtbComp.Rows(0)
            Dim ldrwNewCar = ldtbCartPorPre.NewRow
            For i = 0 To 3
                ldrwNewCar(i) = ldrwComp(i)
            Next
            ldtbCartPorPre.Rows.Add(ldrwNewCar)
        End If
    End Sub
    'Cartera por Predio
    Private Sub SGenereDataSetCarPorPredio(adsCarteraPorPredio As DataSet)
        Dim lstrSqlCarteraPorPredio = FstrExpCarteraPorPredio()
        Dim lstrSqlClientes = FstrExpNombresClientes()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlCarteraPorPredio)
        lcolExpresionesSql.Add(lstrSqlClientes)
        lcolNombresTablas.Add("OriCarteraPorPredio")
        lcolNombresTablas.Add("OriClientes")
        GobjPanDat.SdsDataSet(adsCarteraPorPredio, lcolExpresionesSql, lcolNombresTablas)
        adsCarteraPorPredio.Tables.Add(FdtbCentroUtilidad)
        SComplementePredios(adsCarteraPorPredio)
        'Dim lstrNomArch = GstrTrayDatPrg & "CarteraPorPredio" & ".XML"
        'adsCarteraPorPredio.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpCarteraPorPredio() As String
        With MstbExpresionSql
            .Clear().Append("SELECT ").Append(ClsIdPredio_ItemFactStr.SstrNombreCampoBd)
            .Append(", ").Append(ClsIdCliente_FactDbl.SstrNombreCampoBd)
            .Append(", SUM(I.").Append(ClsDebitos_ItemFactDec.SstrNombreCampoBd).Append(" - I.")
            .Append(ClsCreditos_ItemFactDec.SstrNombreCampoBd).Append(") AS Saldo FROM ")
            .Append(ClsItemFactura.SstrNombreTabla).Append(" AS I INNER JOIN ")
            .Append(ClsFactura.SstrNombreTabla).Append("  AS F ON I.")
            .Append(OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = F.")
            .Append(OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" AND I.")
            .Append(OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = F.")
            .Append(OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" AND I.")
            .Append(ClsPrefijo_ItemFactStr.SstrNombreCampoBd).Append(" = F.")
            .Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append("  AND I.")
            .Append(ClsIdFactura_ItemFactEnt.SstrNombreCampoBd).Append(" = F.")
            .Append(ClsIdFacturaEnt.SstrNombreCampoBd).Append(" WHERE I.")
            .Append(OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCarpeta).Append(" AND I.")
            .Append(OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCentroUtil).Append(" AND I.")
            .Append(ClsDebitos_ItemFactDec.SstrNombreCampoBd).Append(" <> I.")
            .Append(ClsCreditos_ItemFactDec.SstrNombreCampoBd).Append(" GROUP BY ")
            .Append(ClsIdPredio_ItemFactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(" ORDER BY ")
            .Append(ClsIdPredio_ItemFactStr.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpNombresClientes() As String
        Dim lstrTabla = ClsCliente.SstrNombreTabla
        Dim lstrCampSel As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion
        Dim lstrOrden As String(,) = {{ClsIdClienteDbl.SstrNombreCampoBd, "ASC"}}
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCampSel, lstrOrden,
                lstrFiltro, Array.Empty(Of String))
        Return lstrExpSql
    End Function
    Private Shared Sub SComplementePredios(adsCarteraPorPredio As DataSet)
        Dim ldtbCartPorPre = adsCarteraPorPredio.Tables("OriCarteraPorPredio")
        Dim lstrIdPredio As String
        For Each ldrwFila As DataRow In ldtbCartPorPre.Rows
            lstrIdPredio = ClsPanorama.FobjValorCampo(ldrwFila(
                    ClsIdPredio_ItemFactStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            If lstrIdPredio = String.Empty Then
                ldrwFila(ClsIdPredio_ItemFactStr.SstrNombreCampoBd) = "Sin predio"
            Else
                Exit For
            End If
        Next
    End Sub
    ' Cartera por Servicio
    Private Sub SGenereDataSetCarPorServicio(adsCarteraPorServicio As DataSet)
        Dim lstrSqlCarteraPorServicio = FstrExpCarteraPorServicio()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlCarteraPorServicio)
        lcolNombresTablas.Add("OriCarteraPorServicio")
        GobjPanDat.SdsDataSet(adsCarteraPorServicio, lcolExpresionesSql, lcolNombresTablas)
        adsCarteraPorServicio.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "CarteraPorServicio" & ".XML"
        'adsCarteraPorServicio.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
        'adsCarteraPorServicio.Dispose()
    End Sub
    Private Function FstrExpCarteraPorServicio() As String
        ' Expresión Filtro 
        Dim lstrTablaPri = ClsItemFactura.SstrNombreTabla
        With MstbExpresionSql
            .Clear().Append("SUM(").Append(ClsDebitos_ItemFactDec.SstrNombreCampoBd).Append(")")
        End With
        Dim lstrDebitos = MstbExpresionSql.ToString
        With MstbExpresionSql
            .Clear().Append("SUM(").Append(ClsCreditos_ItemFactDec.SstrNombreCampoBd).Append(")")
        End With
        Dim lstrCreditos = MstbExpresionSql.ToString
        Dim lstrCamposPri As String() = {ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd,
                ClsIdServicio_ItemFactShr.SstrNombreCampoBd, lstrDebitos, lstrCreditos}
        Dim lstrTablaSec = ClsServicio.SstrNombreTabla
        Dim lstrCamposSec As String() = {ClsNombreServicioStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri As String() = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd, ClsIdServicio_ItemFactShr.SstrNombreCampoBd}
        Dim lsttrCamposSecRel As String() = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                 PanL.ClsIdCentroUtilShr.SstrNombreCampoBd, ClsIdAno_ServicioShr.SstrNombreCampoBd,
                ClsIdServicioShr.SstrNombreCampoBd}
        Dim lstrIndice As String(,) = {{ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd, "DESC"},
                                       {ClsIdServicioShr.SstrNombreCampoBd, "ASC"}}
        With MstbExpresionSql
            .Clear()
            .Append("P.").Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCarpeta)
            .Append(" AND P.").Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCentroUtil).Append(" AND ").Append(ClsDebitos_ItemFactDec.SstrNombreCampoBd)
            .Append(" <> ").Append(ClsCreditos_ItemFactDec.SstrNombreCampoBd)
        End With
        Dim lstrFiltro = MstbExpresionSql.ToString
        Dim lstrCamposGrupo As String() = {ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd,
                                          ClsIdServicio_ItemFactShr.SstrNombreCampoBd}
        Dim lstrExpSqlCarPorSer = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri, lstrCamposPri,
                lstrTablaSec, lstrCamposSec, lstrCamposRelPri, lsttrCamposSecRel, lstrIndice, lstrFiltro,
                lstrCamposGrupo)
        Return lstrExpSqlCarPorSer
    End Function
    'Edad Cartera
    Private Sub SGenereDataSetEdadCartera(adsEdadCartera As DataSet, aentLimite1 As Integer,
            aentLimite2 As Integer, aentLimite3 As Integer, aentLimite4 As Integer,
            ablnDetallado As Boolean, adtmFecha As Date)
        Dim lstrExpSqlEdadCartera As String
        ' El reporte de edad de cartera para el cierre siempre es el reporte resumido y
        ' requiere la fecha de cierre. Los demas reportes se hacen al día de hoy.
        ' El programa no permite hacer movimientos si no ha ha hecho el cierre de mes. 
        If ablnDetallado Then
            lstrExpSqlEdadCartera = FstrExpEdadCarteraDetallado(aentLimite1, aentLimite2,
                    aentLimite3, aentLimite4)
        Else
            lstrExpSqlEdadCartera = FstrExpEdadCartera(aentLimite1, aentLimite2,
                    aentLimite3, aentLimite4, adtmFecha)
        End If
        Dim lstrExpSqlClientes = FstrExpSqlClientesGen()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrExpSqlEdadCartera)
        lcolExpresionesSql.Add(lstrExpSqlClientes)
        lcolNombresTablas.Add("OriEdadCartera")
        lcolNombresTablas.Add("OriClientesFact")
        GobjPanDat.SdsDataSet(adsEdadCartera, lcolExpresionesSql, lcolNombresTablas)
        Dim ldtbNombreRangos = FdtbNombreRangos(aentLimite1, aentLimite2, aentLimite3, aentLimite4)
        adsEdadCartera.Tables.Add(ldtbNombreRangos)
        adsEdadCartera.Tables.Add(FdtbCentroUtilidad)
        'If ablnDetallado Then
        '    Dim lstrNomArch = GstrTrayDatPrg & "EdadCarteraDetallado" & ".XML"
        '    adsEdadCartera.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
        'Else
        '    Dim lstrNomArch = GstrTrayDatPrg & "EdadCartera" & ".XML"
        '    adsEdadCartera.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
        'End If
    End Sub
    Private Sub SGenereDataSetEdadCartera(adsEdadCartera As DataSet)
        Dim lstrExpSqlEdadCartera = FstrExpEdadCarteraResumen()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrExpSqlEdadCartera)
        lcolNombresTablas.Add("OriEdadCartera")
        GobjPanDat.SdsDataSet(adsEdadCartera, lcolExpresionesSql, lcolNombresTablas)
        adsEdadCartera.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "EdadCarteraResumen" & ".XML"
        'adsEdadCartera.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpEdadCarteraResumen() As String
        Dim lstrFechaHoy = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today) & "'"
        Dim lstrFechaLimite1 = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today.AddDays(-30)) & "'"
        Dim lstrFechaLimite2 = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today.AddDays(-60)) & "'"
        Dim lstrFechaLimite3 = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today.AddDays(-120)) & "'"
        Dim lstrFechaLimite4 = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today.AddDays(-180)) & "'"
        Dim lstrNomRango0 = "'Sin Vencer'"
        Dim lstrNomRango1 = "'De 1 a 30'"
        Dim lstrNomRango2 = "'De 31 a 60'"
        Dim lstrNomRango3 = "'De 61 a 120'"
        Dim lstrNomRango4 = "'De 121 a 180'"
        Dim lstrNomRango5 = "'Mas de 180'"
        ' Valor
        With MstbExpresionSql
            .Clear().Append(", (").Append(ClsDebitos_FactDec.SstrNombreCampoBd).Append(" - ")
            .Append(ClsCreditos_FactDec.SstrNombreCampoBd).Append("), 0")
        End With
        Dim lstrValor = MstbExpresionSql.ToString
        ' Rango 0: Sin vencer
        With MstbExpresionSql
            .Clear().Append("SUM(")
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" >= ").Append(lstrFechaHoy)
            .Append(lstrValor).Append("))").Append(" AS ").Append(lstrNomRango0)
        End With
        Dim lstrRango0 = MstbExpresionSql.ToString
        ' Rango 1 
        With MstbExpresionSql
            .Clear().Append("SUM(")
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" < ").Append(lstrFechaHoy)
            .Append(" AND ").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaLimite1).Append(lstrValor).Append("))").Append(" AS ").Append(lstrNomRango1)
        End With
        Dim lstrRango1 = MstbExpresionSql.ToString
        ' Rango 2
        With MstbExpresionSql
            .Clear().Append("SUM(")
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" < ").Append(lstrFechaLimite1)
            .Append(" AND ").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaLimite2).Append(lstrValor).Append("))").Append(" AS ").Append(lstrNomRango2)
        End With
        Dim lstrRango2 = MstbExpresionSql.ToString
        ' Rango 3
        With MstbExpresionSql
            .Clear().Append("SUM(")
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" < ").Append(lstrFechaLimite2)
            .Append(" AND ").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaLimite3).Append(lstrValor).Append("))").Append(" AS ").Append(lstrNomRango3)
        End With
        Dim lstrRango3 = MstbExpresionSql.ToString
        ' Rango 4
        With MstbExpresionSql
            .Clear().Append("SUM(")
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" < ").Append(lstrFechaLimite3)
            .Append(" AND ").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaLimite4).Append(lstrValor).Append("))").Append(" AS ").Append(lstrNomRango4)
        End With
        Dim lstrRango4 = MstbExpresionSql.ToString
        ' Rango 5
        With MstbExpresionSql
            .Clear().Append("SUM(")
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" < ").Append(lstrFechaLimite4)
            .Append(lstrValor).Append("))").Append(" AS ").Append(lstrNomRango5)
        End With
        Dim lstrRango5 = MstbExpresionSql.ToString
        ' Filtro Ubicación
        With MstbExpresionSql
            .Clear().Append(ClsFactura.SstrNombreTabla).Append(".").Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCarpeta.ToString).Append(" AND ").Append(ClsFactura.SstrNombreTabla).Append(".")
            .Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCentroUtil.ToString)
        End With
        Dim lstrFiltroUbicacion = MstbExpresionSql.ToString
        With MstbExpresionSql
            .Clear()
            .Append("SELECT ")
            .Append(lstrRango0).Append(", ").Append(lstrRango1).Append(", ")
            .Append(lstrRango2).Append(", ").Append(lstrRango3).Append(", ")
            .Append(lstrRango4).Append(", ").Append(lstrRango5)
            .Append(" FROM ").Append(ClsFactura.SstrNombreTabla)
            .Append(" WHERE ").Append(lstrFiltroUbicacion).Append(" AND ")
            .Append(ClsDebitos_FactDec.SstrNombreCampoBd).Append(" <> ")
            .Append(ClsCreditos_FactDec.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpEdadCartera(aentLimite1 As Integer, aentLimite2 As Integer,
            aentLimite3 As Integer, aentLimite4 As Integer, adtmFecha As Date) As String
        Dim ldtmFechaReporte As Date = Date.Today
        If adtmFecha <> GCDTMFECHANULA Then
            ldtmFechaReporte = adtmFecha
        End If
        Dim lstrFechaHoy = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today) & "'"
        Dim lstrFechaLimite1 = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaReporte.AddDays(-aentLimite1)) & "'"
        Dim lstrFechaLimite2 = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaReporte.AddDays(-aentLimite2)) & "'"
        Dim lstrFechaLimite3 = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaReporte.AddDays(-aentLimite3)) & "'"
        Dim lstrFechaLimite4 = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaReporte.AddDays(-aentLimite4)) & "'"
        Dim lstrNomRango0 = "'Rango0'"
        Dim lstrNomRango1 = "'Rango1'"
        Dim lstrNomRango2 = "'Rango2'"
        Dim lstrNomRango3 = "'Rango3'"
        Dim lstrNomRango4 = "'Rango4'"
        Dim lstrNomRango5 = "'Rango5'"
        ' Valor
        With MstbExpresionSql
            .Clear().Append(", (").Append(ClsDebitos_FactDec.SstrNombreCampoBd).Append(" - ")
            .Append(ClsCreditos_FactDec.SstrNombreCampoBd).Append("), 0")
        End With
        Dim lstrValor = MstbExpresionSql.ToString
        ' Rango 0: Sin vencer
        With MstbExpresionSql
            .Clear().Append("SUM(")
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaHoy).Append(lstrValor).Append("))").Append(" AS ").Append(lstrNomRango0)
        End With
        Dim lstrRango0 = MstbExpresionSql.ToString
        ' Rango 1 
        With MstbExpresionSql
            .Clear().Append("SUM(")
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" < ").Append(lstrFechaHoy)
            .Append(" AND ").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaLimite1).Append(lstrValor).Append("))").Append(" AS ").Append(lstrNomRango1)
        End With
        Dim lstrRango1 = MstbExpresionSql.ToString
        ' Rango 2
        With MstbExpresionSql
            .Clear().Append("SUM(")
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" < ").Append(lstrFechaLimite1)
            .Append(" AND ").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaLimite2).Append(lstrValor).Append("))").Append(" AS ").Append(lstrNomRango2)
        End With
        Dim lstrRango2 = MstbExpresionSql.ToString
        ' Rango 3
        With MstbExpresionSql
            .Clear().Append("SUM(")
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" < ").Append(lstrFechaLimite2)
            .Append(" AND ").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaLimite3).Append(lstrValor).Append("))").Append(" AS ").Append(lstrNomRango3)
        End With
        Dim lstrRango3 = MstbExpresionSql.ToString
        ' Rango 4
        With MstbExpresionSql
            .Clear().Append("SUM(")
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" < ").Append(lstrFechaLimite3)
            .Append(" AND ").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaLimite4).Append(lstrValor).Append("))").Append(" AS ").Append(lstrNomRango4)
        End With
        Dim lstrRango4 = MstbExpresionSql.ToString
        ' Rango 5
        With MstbExpresionSql
            .Clear().Append("SUM(")
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" < ").Append(lstrFechaLimite4)
            .Append(lstrValor).Append("))").Append(" AS ").Append(lstrNomRango5)
        End With
        Dim lstrRango5 = MstbExpresionSql.ToString
        ' Filtro Ubicación
        With MstbExpresionSql
            .Clear()
            .Append("F.").Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCarpeta.ToString).Append(" AND ").Append("F.")
            .Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCentroUtil.ToString)
        End With
        Dim lstrFiltroUbicacion = MstbExpresionSql.ToString
        With MstbExpresionSql
            .Clear()
            .Append("SELECT ").Append("F.")
            .Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd).Append(", ")
            .Append(lstrRango0).Append(", ").Append(lstrRango1).Append(", ")
            .Append(lstrRango2).Append(", ").Append(lstrRango3).Append(", ")
            .Append(lstrRango4).Append(", ").Append(lstrRango5).Append(" FROM ")
            .Append(ClsFactura.SstrNombreTabla).Append(" AS F INNER JOIN ")
            .Append(ClsCliente.SstrNombreTabla).Append(" AS C ON ").Append("F.")
            .Append(MstrCampoBdCarpeta).Append(" = ").Append("C.").Append(MstrCampoBdCarpeta)
            .Append(" AND ").Append("F.").Append(MstrCampoBdCentroUtil).Append(" = ")
            .Append("C.").Append(MstrCampoBdCentroUtil).Append(" AND ")
            .Append("F.").Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(" = ")
            .Append("C.").Append(ClsIdClienteDbl.SstrNombreCampoBd)
            .Append(" WHERE ").Append(lstrFiltroUbicacion).Append(" AND ")
            .Append(ClsDebitos_FactDec.SstrNombreCampoBd).Append(" <> ")
            .Append(ClsCreditos_FactDec.SstrNombreCampoBd)
            .Append(" GROUP BY ").Append("F.")
            .Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd)
            .Append(" ORDER BY ").Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpEdadCarteraDetallado(aentLimite1 As Integer, aentLimite2 As Integer,
                aentLimite3 As Integer, aentLimite4 As Integer) As String
        Dim lstrFechaHoy = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today) & "'"
        Dim lstrFechaLimite1 = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today.AddDays(-aentLimite1)) & "'"
        Dim lstrFechaLimite2 = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today.AddDays(-aentLimite2)) & "'"
        Dim lstrFechaLimite3 = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today.AddDays(-aentLimite3)) & "'"
        Dim lstrFechaLimite4 = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today.AddDays(-aentLimite4)) & "'"
        Dim lstrNomRango0 = "'Rango0'"
        Dim lstrNomRango1 = "'Rango1'"
        Dim lstrNomRango2 = "'Rango2'"
        Dim lstrNomRango3 = "'Rango3'"
        Dim lstrNomRango4 = "'Rango4'"
        Dim lstrNomRango5 = "'Rango5'"
        ' Valor
        With MstbExpresionSql
            .Clear().Append(", (").Append(ClsDebitos_FactDec.SstrNombreCampoBd).Append(" - ")
            .Append(ClsCreditos_FactDec.SstrNombreCampoBd).Append("), 0")
        End With
        Dim lstrValor = MstbExpresionSql.ToString
        ' Rango 0: Sin vencer
        With MstbExpresionSql
            .Clear()
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd)
            .Append(" >= ").Append(lstrFechaHoy)
            .Append(lstrValor).Append(")").Append(" AS ").Append(lstrNomRango0)
        End With
        Dim lstrRango0 = MstbExpresionSql.ToString
        ' Rango 1 
        With MstbExpresionSql
            .Clear()
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" < ").Append(lstrFechaHoy)
            .Append(" AND ").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaLimite1).Append(lstrValor).Append(")").Append(" AS ").Append(lstrNomRango1)
        End With
        Dim lstrRango1 = MstbExpresionSql.ToString
        ' Rango 2
        With MstbExpresionSql
            .Clear()
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" < ").Append(lstrFechaLimite1)
            .Append(" AND ").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaLimite2).Append(lstrValor).Append(")").Append(" AS ").Append(lstrNomRango2)
        End With
        Dim lstrRango2 = MstbExpresionSql.ToString
        ' Rango 3
        With MstbExpresionSql
            .Clear()
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" < ").Append(lstrFechaLimite2)
            .Append(" AND ").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaLimite3).Append(lstrValor).Append(")").Append(" AS ").Append(lstrNomRango3)
        End With
        Dim lstrRango3 = MstbExpresionSql.ToString
        ' Rango 4
        With MstbExpresionSql
            .Clear()
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" < ").Append(lstrFechaLimite3)
            .Append(" AND ").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" >= ")
            .Append(lstrFechaLimite4).Append(lstrValor).Append(")").Append(" AS ").Append(lstrNomRango4)
        End With
        Dim lstrRango4 = MstbExpresionSql.ToString
        ' Rango 5
        With MstbExpresionSql
            .Clear()
            .Append("IF(").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(" < ").Append(lstrFechaLimite4)
            .Append(lstrValor).Append(")").Append(" AS ").Append(lstrNomRango5)
        End With
        Dim lstrRango5 = MstbExpresionSql.ToString
        ' Filtro Ubicación
        With MstbExpresionSql
            .Clear().Append("F.").Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCarpeta.ToString).Append(" AND ").Append("F.")
            .Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCentroUtil.ToString)
        End With
        Dim lstrFiltroUbicacion = MstbExpresionSql.ToString
        With MstbExpresionSql
            .Clear()
            .Append("SELECT ").Append("F.")
            .Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFacturaEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(", ")
            .Append(lstrRango0).Append(", ").Append(lstrRango1).Append(", ")
            .Append(lstrRango2).Append(", ").Append(lstrRango3).Append(", ")
            .Append(lstrRango4).Append(", ").Append(lstrRango5).Append(", ").Append("DATEDIFF(")
            .Append(lstrFechaHoy).Append(", ").Append(ClsFechaVencimientoDtm.SstrNombreCampoBd)
            .Append(")").Append(" AS Edad FROM ").Append(ClsFactura.SstrNombreTabla)
            .Append(" AS F INNER JOIN ").Append(ClsCliente.SstrNombreTabla).Append(" AS C ON ")
            .Append("F.").Append(MstrCampoBdCarpeta).Append(" = ").Append("C.")
            .Append(MstrCampoBdCarpeta).Append(" AND ").Append("F.").Append(MstrCampoBdCentroUtil)
            .Append(" = ").Append("C.").Append(MstrCampoBdCentroUtil).Append(" AND ")
            .Append("F.").Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(" = ")
            .Append("C.").Append(ClsIdClienteDbl.SstrNombreCampoBd)
            .Append(" WHERE ").Append(lstrFiltroUbicacion).Append(" AND ")
            .Append(ClsDebitos_FactDec.SstrNombreCampoBd).Append(" <> ")
            .Append(ClsCreditos_FactDec.SstrNombreCampoBd)
            .Append(" ORDER BY ").Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFacturaEnt.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Shared Function FdtbNombreRangos(aentLimite1 As Integer, aentLimite2 As Integer,
            aentLimite3 As Integer, aentLimite4 As Integer) As DataTable
        Dim lTipoString As System.Type = System.Type.GetType("System.String")
        Dim ldtbNombreRangos As New DataTable("OriNombresRangos")
        Dim ldclRango0 As New DataColumn("Rango0", lTipoString)
        Dim ldclRango1 As New DataColumn("Rango1", lTipoString)
        Dim ldclRango2 As New DataColumn("Rango2", lTipoString)
        Dim ldclRango3 As New DataColumn("Rango3", lTipoString)
        Dim ldclRango4 As New DataColumn("Rango4", lTipoString)
        Dim ldclRango5 As New DataColumn("Rango5", lTipoString)
        Dim lstrNomRango0 = "Sin Vencer"
        Dim lstrNomRango1 = "De 1" & " a " & aentLimite1.ToString
        Dim lstrNomRango2 = "De " & (aentLimite1 + 1).ToString & " a " & aentLimite2.ToString
        Dim lstrNomRango3 = "De " & (aentLimite2 + 1).ToString & " a " & aentLimite3.ToString
        Dim lstrNomRango4 = "De " & (aentLimite3 + 1).ToString & " a " & aentLimite4.ToString
        Dim lstrNomRango5 = "Mas de " & aentLimite4.ToString
        ldtbNombreRangos.Columns.AddRange({ldclRango0, ldclRango1, ldclRango2, ldclRango3, ldclRango4, ldclRango5})
        Dim ldrwNewFila = ldtbNombreRangos.NewRow
        ldrwNewFila("Rango0") = lstrNomRango0
        ldrwNewFila("Rango1") = lstrNomRango1
        ldrwNewFila("Rango2") = lstrNomRango2
        ldrwNewFila("Rango3") = lstrNomRango3
        ldrwNewFila("Rango4") = lstrNomRango4
        ldrwNewFila("Rango5") = lstrNomRango5
        ldtbNombreRangos.Rows.Add(ldrwNewFila)
        Return ldtbNombreRangos
    End Function
    Private Shared Sub SComplementeDsEdadCartera(adsEdadCartera As DataSet, adtmFechaDatos As Date)
        Dim ldtbCentroUtil = adsEdadCartera.Tables(ClsCentroUtilidad.SstrNombreTabla)
        Dim ldrwCenUtil As DataRow = ldtbCentroUtil.Select()(0)
        ldrwCenUtil("FechaDatos") = adtmFechaDatos
    End Sub
    Private Shared Function FstrNombresColumnas(ablnDetallado As Boolean, aentLimite1 As Integer,
            aentLimite2 As Integer, aentLimite3 As Integer, aentLimite4 As Integer) As String()
        Dim lstrNomRango0 = "Sin Vencer"
        Dim lstrNomRango1 = "De 1" & " a " & aentLimite1.ToString
        Dim lstrNomRango2 = "De " & (aentLimite1 + 1).ToString & " a " & aentLimite2.ToString
        Dim lstrNomRango3 = "De " & (aentLimite2 + 1).ToString & " a " & aentLimite3.ToString
        Dim lstrNomRango4 = "De " & (aentLimite3 + 1).ToString & " a " & aentLimite4.ToString
        Dim lstrNomRango5 = "Mas de " & aentLimite4.ToString
        Dim lstrNombresColumnas As String()
        If ablnDetallado Then
            lstrNombresColumnas = {"Id. Cliente", "Nombre", "Predio Agrupador", "Pref.", "Nro. Fac.",
                    "Fec. Vcto.", lstrNomRango0, lstrNomRango1, lstrNomRango2, lstrNomRango3,
                    lstrNomRango4, lstrNomRango5, "Total", "Días"}
        Else
            lstrNombresColumnas = {"Id. Cliente", "Nombre", "Predio Agrupador", lstrNomRango0,
                    lstrNomRango1, lstrNomRango2, lstrNomRango3, lstrNomRango4, lstrNomRango5, "Total"}
        End If
        Return lstrNombresColumnas
    End Function
    Private Shared Function FstrNombresColumnas() As String()
        Dim lstrNomRango0 = "Sin Vencer"
        Dim lstrNomRango1 = "De 1 a 30"
        Dim lstrNomRango2 = "De 31 a 60"
        Dim lstrNomRango3 = "De 61 a 120"
        Dim lstrNomRango4 = "De 121 a 180"
        Dim lstrNomRango5 = "Mas de 180"
        Dim lstrNombresColumnas As String() = {lstrNomRango0, lstrNomRango1, lstrNomRango2, lstrNomRango3,
                    lstrNomRango4, lstrNomRango5}
        Return lstrNombresColumnas
    End Function
    ' Estado de la cuenta
    Private Sub SGenereDataSetEstadoCuenta(adsEstadosCuentas As DataSet)
        Dim lstrExpSqlEstadoSugerido = FstrExpSqlEstadoSugerido()
        Dim lstrExpSqlEstadoActual = FstrExpSqlEstadoActual()
        Dim lstrExpSqlNombreEstado = FstrExpSqlNombreEstado()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrExpSqlEstadoSugerido)
        lcolExpresionesSql.Add(lstrExpSqlEstadoActual)
        lcolExpresionesSql.Add(lstrExpSqlNombreEstado)
        lcolNombresTablas.Add("OriEstadoSugerido")
        lcolNombresTablas.Add("OriEstadoActual")
        lcolNombresTablas.Add("OriNombreEstado")
        GobjPanDat.SdsDataSet(adsEstadosCuentas, lcolExpresionesSql, lcolNombresTablas)
        adsEstadosCuentas.Tables.Add(FdtbCentroUtilidad)
        SComplementeEstadoCuenta(adsEstadosCuentas)
        'Dim lstrNomArch = GstrTrayDatPrg & "EstadoCuentas" & ".XML"
        'adsEstadosCuentas.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlEstadoSugerido() As String
        Dim lstrFechaHoy = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today) & "'"
        ' Filtro
        With MstbExpresionSql
            .Clear().Append(" WHERE ").Append("F.").Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCarpeta.ToString).Append(" AND F.")
            .Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCentroUtil.ToString).Append(" AND ").Append(ClsDebitos_FactDec.SstrNombreCampoBd)
            .Append(" <> ").Append(ClsCreditos_FactDec.SstrNombreCampoBd)
        End With
        Dim lstrFiltro = MstbExpresionSql.ToString
        With MstbExpresionSql
            .Clear().Append("SELECT ").Append(ClsNombreCompletoStr.SstrNombreCampoBd)
            .Append(", ").Append(" DATEDIFF(").Append(lstrFechaHoy).Append(", MIN(")
            .Append(ClsFechaVencimientoDtm.SstrNombreCampoBd).Append(")) AS Dias")
            .Append(", ").Append("F.").Append(ClsIdCliente_FactDbl.SstrNombreCampoBd)
            .Append(", ").Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd)
            .Append(", ").Append("'' AS Estado").Append(" FROM ")
            .Append(ClsFactura.SstrNombreTabla).Append(" AS F INNER JOIN ")
            .Append(ClsCliente.SstrNombreTabla).Append(" AS C ON ").Append("F.")
            .Append(MstrCampoBdCarpeta).Append(" = C.").Append(MstrCampoBdCarpeta)
            .Append(" AND F.").Append(MstrCampoBdCentroUtil).Append(" = C.")
            .Append(MstrCampoBdCentroUtil).Append(" AND F.")
            .Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(" = C.")
            .Append(ClsIdClienteDbl.SstrNombreCampoBd).Append(lstrFiltro)
            .Append(" GROUP BY ").Append(ClsNombreCompletoStr.SstrNombreCampoBd)
            .Append(", ").Append("F.").Append(ClsIdCliente_FactDbl.SstrNombreCampoBd)
            .Append(", ").Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd)
            .Append(" ORDER BY Dias DESC ,")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpSqlEstadoActual() As String
        With MstbExpresionSql
            .Clear().Append("SELECT ").Append(ClsIdClienteDbl.SstrNombreCampoBd)
            .Append(", ").Append("'' AS ")
            .Append(ClsIdPredioAgrupadorStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdEstadoDeudaByt.SstrNombreCampoBd).Append(" FROM ")
            .Append(ClsCliente.SstrNombreTabla).Append(" WHERE ")
            .Append(ClsOrionCop.StrFiltroUbicacion)
        End With
        Dim lstrExpEstadoSinPreAgr = MstbExpresionSql.ToString

        Dim lstrTablaPri = ClsPropietario.SstrNombreTabla
        Dim lstrCampSelPri As String() = {ClsIdCliente_PropDbl.SstrNombreCampoBd}
        Dim lstrTablaSec = ClsPredio.SstrNombreTabla
        Dim lstrCampSelSec As String() = {ClsIdPredioAgrupadorStr.SstrNombreCampoBd,
                ClsIdEstadoDeuda_PredioByt.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd,
                OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                ClsIdPredio_PropStr.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd,
                OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND S." &
                ClsIdPredioStr.SstrNombreCampoBd & " = " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        Dim lstrExpEstadoPreAgr = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri,
                lstrCampSelPri, lstrTablaSec, lstrCampSelSec, lstrCampRelPri,
                lstrCampRelSec, {{}}, lstrFiltro, {})
        ' Consulta Total
        With MstbExpresionSql
            .Clear.Append("(").Append(lstrExpEstadoSinPreAgr).Append(")").Append(" UNION ALL ")
            .Append("(").Append(lstrExpEstadoPreAgr).Append(")")
            .Append(" ORDER BY ").Append(ClsIdClienteDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupadorStr.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpSqlNombreEstado() As String
        With MstbExpresionSql
            .Clear().Append("SELECT IdConstante, Dato FROM OriTblConstantes WHERE IdGrupo = ")
            .Append(EnuGrupoConstantesOriDef.EnuEstadoDeuda).Append(" ORDER BY IdConstante")
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Shared Sub SComplementeEstadoCuenta(adsEstadoCuentas As DataSet)
        Dim lstrNomEstado = String.Empty, lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.None
        Dim ldtbEstadoSuger = adsEstadoCuentas.Tables("OriEstadoSugerido")
        Dim ldrwEstadosSugeridos = ldtbEstadoSuger.Select()
        For Each ldrwEstCue As DataRow In ldrwEstadosSugeridos
            Dim lentDiasVen As Integer = ClsPanorama.FobjValorCampo(ldrwEstCue("Dias"),
                    EnuTipoValor.enuInteger)
            With GobjParametros
                If lentDiasVen >= .ObjDiasParaPerdidaShr.ObjValorPro Then
                    lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuPerdida
                ElseIf lentDiasVen < .ObjDiasParaPerdidaShr.ObjValorPro AndAlso
                        lentDiasVen >= .ObjDiasParaJuridicoShr.ObjValorPro Then
                    lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuJuridico
                ElseIf lentDiasVen < .ObjDiasParaJuridicoShr.ObjValorPro AndAlso
                        lentDiasVen >= .ObjDiasParaPrejuridicoShr.ObjValorPro Then
                    lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuPrejuridico
                ElseIf lentDiasVen < .ObjDiasParaPrejuridicoShr.ObjValorPro AndAlso
                        lentDiasVen >= .ObjDiasParaPersuasivoShr.ObjValorPro Then
                    lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuPersuasivo
                Else
                    lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuNormal
                End If
            End With
            lstrNomEstado = ClsOrionCop.FstrNombreDatoConstanteOri(EnuGrupoConstantesOriDef.EnuEstadoDeuda,
                    lenuEstadoSugeridoCliente)
            ldrwEstCue("Estado") = lstrNomEstado
        Next
    End Sub
    ' Estado de cuenta de Cliente y Predio Agrupador
    Private Sub SGenereDataSetEstadoCtaCli(adsEstadoCtaCli As DataSet)
        Dim ldtbFacturasVivas = FdtbFacturasVivas(DblIdCliente, StrIdPredioAgru)
        Dim ldtbFact = ldtbFacturasVivas.Copy
        ldtbFact.TableName = "OriFacturas"
        Dim lstrSqlClientes = FstrSqlCliente()
        Dim ldtbClientes = ClsPanorama.FdtbDataTable(lstrSqlClientes)
        Dim ldtbCli = ldtbClientes.Copy
        ldtbCli.TableName = "OriClientes"
        SComplementeDtbCliente(ldtbCli)
        adsEstadoCtaCli.Tables.Add(ldtbCli)
        adsEstadoCtaCli.Tables.Add(ldtbFact)
        adsEstadoCtaCli.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "EstadoCtaCli" & ".XML"
        'adsEstadoCtaCli.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Shared Function FdtbFacturasVivas(adblIdCliente As Double,
            astrIdPredioAgru As String) As DataTable
        GobjPanDat.SControleProcesoObj(True)
        Dim lstrCamposSelect = {ClsPrefijo_FactStr.SstrNombreCampoBd,
                ClsIdFacturaEnt.SstrNombreCampoBd,
                ClsFechaFacturaDtm.SstrNombreCampoBd,
                ClsValor_FactDec.SstrNombreCampoBd,
                ClsFechaVencimientoDtm.SstrNombreCampoBd,
                ClsDebitos_FactDec.SstrNombreCampoBd,
                ClsCreditos_FactDec.SstrNombreCampoBd,
                ClsDebitos_FactDec.SstrNombreCampoBd & " - " &
                ClsCreditos_FactDec.SstrNombreCampoBd & " AS Saldo"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " &
                adblIdCliente & " AND " &
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " = '" &
                astrIdPredioAgru & "'"
        lstrFiltro &= " AND " & ClsDebitos_FactDec.SstrNombreCampoBd & " <> " &
                ClsCreditos_FactDec.SstrNombreCampoBd
        Dim lstrIndice = {{ClsIdCliente_FactDbl.SstrNombreCampoBd, "ASC"},
                          {ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd, "ASC"},
                          {ClsFechaFacturaDtm.SstrNombreCampoBd, "ASC"},
                          {ClsPrefijo_FactStr.SstrNombreCampoBd, "ASC"},
                          {ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}}
        Dim ldtbFacturas = ClsPanorama.FdtbDataTable(ClsFactura.SstrNombreTabla,
                lstrCamposSelect, lstrIndice, lstrFiltro)
        GobjPanDat.SControleProcesoObj(False)
        Return ldtbFacturas
    End Function
    Private Sub SComplementeDtbCliente(adtbCliente As DataTable)
        Dim lstrIdPredioAgr = StrIdPredioAgru
        If String.IsNullOrEmpty(lstrIdPredioAgr) Then lstrIdPredioAgr = "Sin Predio Agrupador"
        adtbCliente.Rows(0)("PredioAgrupador") = lstrIdPredioAgr
        adtbCliente.Rows(0)("IntPorCausar") = DecIntPortCausar
    End Sub
    ' Movimiento
    Private Sub SGenereDataSetMovimiento(adsMovimiento As DataSet)
        Dim lstrSqlClientes = FstrSqlCliente()
        Dim ldtbClientes = ClsPanorama.FdtbDataTable(lstrSqlClientes)
        Dim ldtbCli = ldtbClientes.Copy
        ldtbCli.TableName = "OriClientes"
        adsMovimiento.Tables.Add(ldtbCli)
        adsMovimiento.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "MovimientoCta" & ".XML"
        'adsMovimiento.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    'DirectorioTF
    Private Sub SGenereDataSetDirTf(adsDirTf As DataSet)
        Dim lstrSqlDirTf = FstrExpSqlDirTf()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlDirTf)
        lcolNombresTablas.Add("OriDirTf")
        GobjPanDat.SdsDataSet(adsDirTf, lcolExpresionesSql, lcolNombresTablas)
        adsDirTf.Tables.Add(FdtbCentroUtilidad)
        SComplementeDtbDirTf(adsDirTf)
        'Dim lstrNomArch = GstrTrayDatPrg & "DirectorioTf" & ".XML"
        'adsDirTf.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Shared Function FstrExpSqlDirTf() As String
        Dim lstrTablaPri = ClsPropietario.SstrNombreTabla
        Dim lstrTablaSec = ClsTercero.SstrNombreTabla
        Dim lstrCamposSelPri = {ClsIdPredio_PropStr.SstrNombreCampoBd,
                ClsNombreCompleto_PropStr.SstrNombreCampoBd}
        Dim lstrCamposSelSec = {ClsIdTerceroDbl.SstrNombreCampoBd, ClsTelefonoUnoStr.SstrNombreCampoBd,
                                ClsTelefonoDosStr.SstrNombreCampoBd, ClsCelularStr.SstrNombreCampoBd,
                                ClsCelular2Str.SstrNombreCampoBd, ClsEmailStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {ClsIdCliente_PropDbl.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {ClsIdTerceroDbl.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsIdPredio_PropStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFIltro = ClsOrionCop.StrFiltroUbicacion_Pri
        Dim lstrExpSqlDirTf = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri,
                lstrCamposSelPri, lstrTablaSec, lstrCamposSelSec, lstrCamposRelPri,
                lstrCamposRelSec, lstrIndice, lstrFIltro,
                Array.Empty(Of String)())
        Return lstrExpSqlDirTf
    End Function
    Private Shared Sub SComplementeDtbDirTf(adsDirTf As DataSet)
        Dim ldtbDirTf = adsDirTf.Tables("OriDirTf")
        Dim ldrwTfsClientes = ldtbDirTf.Select()
        For Each ldrwTfCli As DataRow In ldrwTfsClientes
            If IsDBNull(ldrwTfCli(ClsTelefonoUnoStr.SstrNombreCampoBd)) Then
                ldrwTfCli(ClsTelefonoUnoStr.SstrNombreCampoBd) = String.Empty
            End If
            If IsDBNull(ldrwTfCli(ClsTelefonoDosStr.SstrNombreCampoBd)) Then
                ldrwTfCli(ClsTelefonoDosStr.SstrNombreCampoBd) = String.Empty
            End If
            If IsDBNull(ldrwTfCli(ClsCelularStr.SstrNombreCampoBd)) Then
                ldrwTfCli(ClsCelularStr.SstrNombreCampoBd) = String.Empty
            End If
        Next
    End Sub
    'DirectorioClientes
    Private Sub SGenereDataSetDirClientes(adsDirClientes As DataSet)
        Dim lstrSqlDirClientes = FstrExpSqlDirClientes()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlDirClientes)
        lcolNombresTablas.Add("OriClientes")
        GobjPanDat.SdsDataSet(adsDirClientes, lcolExpresionesSql, lcolNombresTablas)
        adsDirClientes.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "DirectorioClientes" & ".XML"
        'adsDirClientes.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Shared Function FstrExpSqlDirClientes() As String
        Dim lstrTablaPri = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec = ClsTercero.SstrNombreTabla
        Dim lstrCamposSelPri = {ClsIdClienteDbl.SstrNombreCampoBd,
                                ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposSelSec = {ClsDireccionUnoStr.SstrNombreCampoBd,
                                ClsDireccionDosStr.SstrNombreCampoBd,
                                ClsTelefonoUnoStr.SstrNombreCampoBd,
                                ClsTelefonoDosStr.SstrNombreCampoBd,
                                ClsCelularStr.SstrNombreCampoBd,
                                ClsCelular2Str.SstrNombreCampoBd,
                                ClsEmailStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {ClsIdTerceroDbl.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsNombreCompletoStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFIltro = "P." & PanL.ClsIdCarpetaShr.SstrNombreCampoBd & " = " &
                GshrIdCarpeta & " AND P." & PanL.ClsIdCentroUtilShr.SstrNombreCampoBd &
                " = " & GshrIdCentroUtil
        Dim lstrExpSqlDirCliente = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri,
                lstrCamposSelPri, lstrTablaSec, lstrCamposSelSec, lstrCamposRelPri,
                lstrCamposRelSec, lstrIndice, lstrFIltro, Array.Empty(Of String)())
        Return lstrExpSqlDirCliente
    End Function
    ' Anticipos por aplicar
    Private Sub SGenereAntPorAplicar(adsAntPorAplicar As DataSet, adtmFechaRep As Date)
        Dim lstrSqlAntPorAplicar = FstrExpAntPorAplicar()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlAntPorAplicar)
        lcolNombresTablas.Add("OriAntPorAplicar")
        GobjPanDat.SdsDataSet(adsAntPorAplicar, lcolExpresionesSql, lcolNombresTablas)
        adsAntPorAplicar.Tables.Add(FdtbCentroUtilidad)
        SComplementeTablaAnt(adsAntPorAplicar, adtmFechaRep)
        'Dim lstrNomArch = GstrTrayDatPrg & "AntPorAplicar" & ".XML"
        'adsAntPorAplicar.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Shared Function FstrExpAntPorAplicar() As String
        Dim lstrTablaPri = ClsAnticipo.SstrNombreTabla
        Dim lstrTablaSec = ClsCliente.SstrNombreTabla
        Dim lstrCamposSelPri = {ClsIdAnticipoEnt.SstrNombreCampoBd, ClsFechaAnticipoDtm.SstrNombreCampoBd,
                                ClsIdCliente_AntDbl.SstrNombreCampoBd,
                                ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd,
                                ClsServicios_AntStr.SstrNombreCampoBd, "'' as Detalle",
                                ClsValor_AntDec.SstrNombreCampoBd, ClsValor_AntDec.SstrNombreCampoBd & " - " &
                                ClsDebitos_AntDec.SstrNombreCampoBd & " AS PorAplicar"}
        Dim lstrCamposSelSec = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                                ClsIdCliente_AntDbl.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsIdAnticipoEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFIltro = "P." & PanL.ClsIdCarpetaShr.SstrNombreCampoBd & " = " & GshrIdCarpeta &
                " AND P." & PanL.ClsIdCentroUtilShr.SstrNombreCampoBd & " = " & GshrIdCentroUtil &
                " AND " & ClsValor_AntDec.SstrNombreCampoBd & " - " &
                ClsDebitos_AntDec.SstrNombreCampoBd & " > 0"
        Dim lstrExpAntPorAplicar = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri, lstrCamposSelPri,
                lstrTablaSec, lstrCamposSelSec, lstrCamposRelPri, lstrCamposRelSec, lstrIndice, lstrFIltro,
                Array.Empty(Of String))
        Return lstrExpAntPorAplicar
    End Function
    Private Sub SComplementeTablaAnt(adstAnticipos As DataSet, adtmFechaRep As Date)
        Dim lstrServicio As String
        Dim lstrDetalle As String
        Dim ldtbAnti As DataTable = adstAnticipos.Tables("OriAntPorAplicar")
        For Each ldrwAnt As DataRow In ldtbAnti.Rows
            lstrServicio = ClsPanorama.FobjValorCampo(ldrwAnt(ClsServicios_AntStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            If lstrServicio.Contains(",") Then
                lstrDetalle = "Varios Servicios"
            ElseIf lstrServicio.Contains("A") Then
                lstrDetalle = "Todos los servicios"
            ElseIf lstrServicio.Contains("0") Then
                lstrDetalle = "Cuotas Administración"
            Else
                lstrDetalle = GobjParametros.FstrNombreServicio(lstrServicio)
            End If
            ldrwAnt("DEtalle") = lstrDetalle
        Next
        Dim ldtbCenUtil = adstAnticipos.Tables(ClsCentroUtilidad.SstrNombreTabla)
        ldtbCenUtil.Rows(0)("FechaDatos") = adtmFechaRep
    End Sub
    ' Listado facturas vivas de un cliente discrimina deuda mora por factura
    Private Sub SGenereFacturasVivas(adsFacCliMora As DataSet)
        Dim lstrSqlFacturas = FstrExpSqlFactCli()
        Dim lstrSqlMoraFac = FstrExpSqlMoraFac()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlFacturas)
        lcolNombresTablas.Add("OriFacturas")
        lcolExpresionesSql.Add(lstrSqlMoraFac)
        lcolNombresTablas.Add("OriMoraFac")
        GobjPanDat.SdsDataSet(adsFacCliMora, lcolExpresionesSql, lcolNombresTablas)
        adsFacCliMora.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "FactConMora" & ".XML"
        'adsFacCliMora.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlFactCli() As String
        Dim lstrTablaPri = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec = ClsFactura.SstrNombreTabla
        Dim lstrCamSelPri = {ClsIdClienteDbl.SstrNombreCampoBd,
                ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamSelSec = {ClsPrefijo_FactStr.SstrNombreCampoBd, ClsIdFacturaEnt.SstrNombreCampoBd,
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd,
                ClsFechaVencimientoDtm.SstrNombreCampoBd,
                "(" & ClsDebitos_FactDec.SstrNombreCampoBd & " - " &
                ClsCreditos_FactDec.SstrNombreCampoBd & ") AS Saldo"}
        Dim lstrCamRelPri = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd, ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamRelSec = {PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd, ClsIdCliente_FactDbl.SstrNombreCampoBd}
        Dim lstrFiltro = "P." & PanL.ClsIdCarpetaShr.SstrNombreCampoBd & " = " & GshrIdCarpeta & " AND P." &
                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd & " = " & GshrIdCentroUtil & " AND P." &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & DblIdCliente & " AND (" &
                ClsDebitos_FactDec.SstrNombreCampoBd & " - " &
                ClsCreditos_FactDec.SstrNombreCampoBd & ") > 0"
        Dim lstrOrden As String(,) = {{ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd, "ASC"},
                {ClsPrefijo_FactStr.SstrNombreCampoBd, "ASC"}, {ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri, lstrCamSelPri,
                lstrTablaSec, lstrCamSelSec, lstrCamRelPri, lstrCamRelSec, lstrOrden,
                lstrFiltro, Array.Empty(Of String))
        Return lstrExpSql
    End Function
    Private Function FstrExpSqlMoraFac() As String
        ' Saldo
        With MstbExpresionSql
            .Clear.Append("SUM(IF(")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 3 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 35 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 21 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 23 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 25").Append(", ")
            .Append("Valor").Append(", ").Append("0) - IF(")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 5 Or ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 36 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 7 Or ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 9 Or ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 19").Append(", ")
            .Append("Valor").Append(", ").Append("0)) AS Saldo")
        End With
        Dim lstrSaldo = MstbExpresionSql.ToString
        ' Sql Interno
        With MstbExpresionSql
            .Clear.Append("(SELECT ").Append(ClsPrefijoFact_NovStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_NovEnt.SstrNombreCampoBd).Append(", ").Append(lstrSaldo)
            .Append(" FROM ").Append(ClsNovedad.SstrNombreTabla).Append(" WHERE ")
            .Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND ").Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCentroUtil).Append(" AND ").Append(ClsIdTercero_NovDbl.SstrNombreCampoBd)
            .Append(" = ").Append(DblIdCliente).Append(" GROUP BY ")
            .Append(ClsPrefijoFact_NovStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_NovEnt.SstrNombreCampoBd).Append(")")
        End With
        Dim lstrSqlInt = MstbExpresionSql.ToString
        ' Sql Final
        With MstbExpresionSql
            .Clear().Append("SELECT ").Append(ClsPrefijoFact_NovStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_NovEnt.SstrNombreCampoBd).Append(", ").Append("Saldo FROM ")
            .Append(lstrSqlInt).Append(" AS F ").Append(" ORDER BY ")
            .Append(ClsPrefijoFact_NovStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_NovEnt.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    ' Expresion Sql General Clientes
    Private Shared Function FstrExpSqlClientesGen() As String
        Dim lstrTablaPri = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec = ClsTercero.SstrNombreTabla
        Dim lstrCamposSelPri = {ClsIdClienteDbl.SstrNombreCampoBd,
                                ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposSelSec = {ClsTelefonoUnoStr.SstrNombreCampoBd,
                                ClsTelefonoDosStr.SstrNombreCampoBd,
                                ClsCelularStr.SstrNombreCampoBd,
                                ClsCelular2Str.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {ClsIdTerceroDbl.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsIdClienteDbl.SstrNombreCampoBd, "ASC"}}
        Dim lstrFIltro = "P." & PanL.ClsIdCarpetaShr.SstrNombreCampoBd & " = " &
                GshrIdCarpeta & " AND P." & PanL.ClsIdCentroUtilShr.SstrNombreCampoBd &
                " = " & GshrIdCentroUtil
        Dim lstrExpSqlDirCliente = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri,
                lstrCamposSelPri, lstrTablaSec, lstrCamposSelSec, lstrCamposRelPri,
                lstrCamposRelSec, lstrIndice, lstrFIltro, Array.Empty(Of String)())
        Return lstrExpSqlDirCliente
    End Function
#End Region
#Region "Auxiliar Contable"
    Private Sub SGenereDataSetAux()
        If IsNothing(MdsAuxiliarCon) Then
            SRefresqueLogo()
            MdsAuxiliarCon = New DataSet
            Dim lstrExpresionSql = FstrEspresionSqlAux()
            Dim lstrOrden = FstrOrdenSqlAux()
            Dim lstrSqlAuxCont = lstrExpresionSql & lstrOrden
            Dim lstrSqlClientes = FstrSqlClientes()
            Dim lstrSqlCuentasCont = FstrExpresionCuentasCont()
            Dim lstrSqlAuxiliarCont = lstrSqlAuxCont
            Dim lcolExpresionesSql As New Collection
            Dim lcolNombresTablas As New Collection
            lcolExpresionesSql.Add(lstrSqlAuxiliarCont)
            lcolExpresionesSql.Add(lstrSqlCuentasCont)
            lcolExpresionesSql.Add(lstrSqlClientes)
            lcolNombresTablas.Add("OriNovedades")
            lcolNombresTablas.Add("PanCuentasCont")
            lcolNombresTablas.Add("OriClientes")
            GobjPanDat.SdsDataSet(MdsAuxiliarCon, lcolExpresionesSql, lcolNombresTablas)
            MdsAuxiliarCon.Tables.Add(FdtbCentroUtilidad)
            SComplementeDsAuxCont(MdsAuxiliarCon)
            'Dim lstrNomArch = GstrTrayDatPrg & "AuxiliarCont" & ".XML"
            'MdsAuxiliarCon.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
        End If
    End Sub
    Private Function FstrEspresionSqlAux() As String
        Dim lstrSqlDb = FstrExpresionSqlAux(True)
        Dim lstrSqlCr = FstrExpresionSqlAux(False)
        Dim lstrSqlDbAnt = FstrExpresionSqlAuxAnt(True)
        Dim lstrSqlCrAnt = FstrExpresionSqlAuxAnt(False)
        ' Consulta Total sin ordenar
        With MstbExpresionSql
            .Clear()
            .Append("(").Append(lstrSqlDb).Append(")").Append(" UNION ALL ")
            .Append("(").Append(lstrSqlDbAnt).Append(")").Append(" UNION ALL ")
            .Append("(").Append(lstrSqlCr).Append(")").Append(" UNION ALL ")
            .Append("(").Append(lstrSqlCrAnt).Append(")")
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpresionSqlAux(ablnDb As Boolean) As String
        Dim lstrNombreTabla = ClsNovedad.SstrNombreTabla
        Dim lstrCta = String.Empty, lstrTipo = String.Empty, lstrDetalle = "'' AS Detalle",
                lstrNroDocOri = "'' AS NroDocOrigen"
        Dim lstrDebito = String.Empty, lstrCredito = String.Empty, lstrSaldo = "0.00 AS Saldo"
        Dim lstrIndice(,) = {{"", ""}}
        MstbExpresionSql.Clear()
        If ablnDb Then
            MstbExpresionSql.Append(ClsIdCuentaDb_NovStr.SstrNombreCampoBd).Append(" AS IdCta")
            lstrCta = MstbExpresionSql.ToString
            lstrTipo = ("'D' as Tipo")
            MstbExpresionSql.Clear()
            With MstbExpresionSql
                .Append("SUM(").Append(ClsValor_NovDec.SstrNombreCampoBd).Append(") AS Debito")
            End With
            lstrDebito = MstbExpresionSql.ToString
            lstrCredito = "0.00 AS Credito"
            MstbExpresionSql.Clear()
        Else
            MstbExpresionSql.Append(ClsIdCuentaCr_NovStr.SstrNombreCampoBd).Append(" AS IdCta")
            lstrCta = MstbExpresionSql.ToString
            lstrTipo = ("'C' as Tipo")
            lstrDebito = "0.00 AS Debito"
            MstbExpresionSql.Clear()
            With MstbExpresionSql
                .Append("SUM(").Append(ClsValor_NovDec.SstrNombreCampoBd).Append(") AS Credito")
            End With
            lstrCredito = MstbExpresionSql.ToString
            MstbExpresionSql.Clear()
        End If
        Dim lstrCamposSelect = {ClsFechaNovedadDtm.SstrNombreCampoBd,
                        ClsIdTercero_NovDbl.SstrNombreCampoBd,
                        ClsAliasCont_NovStr.SstrNombreCampoBd,
                        ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd,
                        ClsIdTipoDocOrigenByt.SstrNombreCampoBd,
                        ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd,
                        ClsIdDocOrigenEnt.SstrNombreCampoBd,
                        lstrDebito, lstrCredito, lstrSaldo,
                        lstrTipo, lstrDetalle, lstrNroDocOri, lstrCta}
        Dim lstrCamposAgrup = {ClsFechaNovedadDtm.SstrNombreCampoBd,
                        ClsIdTercero_NovDbl.SstrNombreCampoBd,
                        ClsAliasCont_NovStr.SstrNombreCampoBd,
                        ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd,
                        ClsIdTipoDocOrigenByt.SstrNombreCampoBd,
                        ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd,
                        ClsIdDocOrigenEnt.SstrNombreCampoBd,
                        "Tipo", "Detalle", "NroDocOrigen", "IdCta"}
        Dim lstrFiltro = FstrExpresionFiltro(ablnDb)
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, lstrCamposAgrup)
        Return lstrSql
    End Function
    Private Function FstrExpresionSqlAuxAnt(ablnDb As Boolean) As String
        Dim lstrCta = String.Empty, lstrTipo = String.Empty, lstrDetalle = "'' AS Detalle",
                lstrTipoDoc = "'' AS NroDocOrigen"
        Dim lstrDebito = String.Empty, lstrCredito = String.Empty, lstrSaldo = "0 AS Saldo"
        MstbExpresionSql.Clear()
        If ablnDb Then
            MstbExpresionSql.Append(ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd).Append(" AS IdCta")
            lstrCta = MstbExpresionSql.ToString
            lstrTipo = ("'D' as Tipo")
            MstbExpresionSql.Clear()
            With MstbExpresionSql
                .Append("SUM(").Append(ClsValor_NovAntDec.SstrNombreCampoBd).Append(") AS Debito")
            End With
            lstrDebito = MstbExpresionSql.ToString
            lstrCredito = "0 AS Credito"
            MstbExpresionSql.Clear()
        Else
            MstbExpresionSql.Append(ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd).Append(" AS IdCta")
            lstrCta = MstbExpresionSql.ToString
            lstrTipo = ("'C' as Tipo")
            lstrDebito = "0 AS Debito"
            MstbExpresionSql.Clear()
            With MstbExpresionSql
                .Append("SUM(").Append(ClsValor_NovDec.SstrNombreCampoBd).Append(") AS Credito")
            End With
            lstrCredito = MstbExpresionSql.ToString
            MstbExpresionSql.Clear()
        End If
        ' Select
        With MstbExpresionSql
            .Clear().Append("SELECT ").Append(ClsFechaNovedadAntDtm.SstrNombreCampoBd).Append(", ")
            .Append("N.").Append(ClsIdTercero_NovAntDbl.SstrNombreCampoBd).Append(", ")
            .Append("N.").Append(ClsAliasCont_NovAntStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd).Append(", ")
            .Append("N.").Append(ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd).Append(", ")
            .Append("N.").Append(ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd).Append(", ")
            .Append("N.").Append(ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd).Append(", ")
            .Append(lstrDebito).Append(", ").Append(lstrCredito).Append(", ").Append(lstrSaldo)
            .Append(", ").Append(lstrTipo).Append(", ").Append(lstrDetalle).Append(", ")
            .Append(lstrTipoDoc).Append(", ").Append(lstrCta)
        End With
        Dim lstrSelect = MstbExpresionSql.ToString
        ' From
        With MstbExpresionSql
            .Clear().Append(" FROM ").Append(ClsNovedadAnticipo.SstrNombreTabla)
            .Append(" AS N INNER JOIN ").Append(ClsAnticipo.SstrNombreTabla).Append(" AS A ON N.")
            .Append(MstrCampoBdCarpeta).Append(" = A.").Append(MstrCampoBdCarpeta).Append(" AND N.")
            .Append(MstrCampoBdCentroUtil).Append(" = A.").Append(MstrCampoBdCentroUtil)
            .Append(" AND N.").Append(ClsIdAnticipo_NovEnt.SstrNombreCampoBd).Append(" = A.")
            .Append(ClsIdAnticipoEnt.SstrNombreCampoBd)
        End With
        Dim lstrFrom = MstbExpresionSql.ToString
        ' Where
        Dim lstrFiltro = " WHERE " & FstrExpresionFiltroAnt(ablnDb)
        ' Group
        With MstbExpresionSql
            .Clear().Append(" GROUP BY ").Append(ClsFechaNovedadAntDtm.SstrNombreCampoBd).Append(", ")
            .Append("N.").Append(ClsIdTercero_NovAntDbl.SstrNombreCampoBd).Append(", ")
            .Append("N.").Append(ClsAliasCont_NovAntStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd).Append(", ")
            .Append("N.").Append(ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd).Append(", ")
            .Append("N.").Append(ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd).Append(", ")
            .Append("N.").Append(ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd).Append(", ")
            .Append("Tipo").Append(", ").Append("Detalle").Append(", ")
            .Append("NroDocOrigen").Append(", ").Append("IdCta")
        End With
        Dim lstrGroup = MstbExpresionSql.ToString
        With MstbExpresionSql
            .Clear().Append(lstrSelect).Append(lstrFrom).Append(lstrFiltro).Append(lstrGroup)
        End With
        Dim lstrSql = MstbExpresionSql.ToString
        Return lstrSql
    End Function
    Private Function FstrExpresionFiltro(ablnDebito As Boolean)
        Dim lstrNombreCampoCta = String.Empty
        Dim lblnTodosClientes = (DblIdCliente = 0)
        MstbExpresionSql.Clear()
        With MstbExpresionSql
            .Append(ClsOrionCop.StrFiltroUbicacion).Append(" AND ").Append(ClsAnuladoBln.SstrNombreCampoBd)
            .Append(" <> TRUE AND ")
            .Append(ClsFechaNovedadDtm.SstrNombreCampoBd).Append(" >= '").Append(StrFechaDesde).Append("' AND ")
            .Append(ClsFechaNovedadDtm.SstrNombreCampoBd).Append(" <= '").Append(StrFechaHasta).Append("'")
            If Not lblnTodosClientes Then
                .Append(" AND ").Append(ClsIdTercero_NovDbl.SstrNombreCampoBd).Append(" = ")
                .Append(DblIdCliente)
                If StrIdPredioAgru <> My.Resources.Todos Then
                    .Append(" AND ").Append(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd)
                    .Append(" = '").Append(StrIdPredioAgru).Append("'")
                End If
            End If
            If ablnDebito Then
                lstrNombreCampoCta = ClsIdCuentaDb_NovStr.SstrNombreCampoBd
            Else
                lstrNombreCampoCta = ClsIdCuentaCr_NovStr.SstrNombreCampoBd
            End If
            .Append(" AND ").Append(lstrNombreCampoCta).Append(" >= '").Append(StrIdCuentaContIni)
            .Append("' AND ").Append(lstrNombreCampoCta).Append(" <= '").Append(StrIdCuentaContFin).Append("'")
        End With
        Dim lstrFiltro = MstbExpresionSql.ToString
        Return lstrFiltro
    End Function
    Private Function FstrOrdenSqlAux() As String
        With MstbExpresionSql
            .Clear.Append(" ORDER BY ")
            .Append("IdCta, ").Append(ClsFechaNovedadDtm.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdTipoDocOrigenByt.SstrNombreCampoBd).Append(", ")
            .Append(ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdDocOrigenEnt.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpresionFiltroAnt(ablnDebito As Boolean) 'Ok TipoNov
        Dim lstrNombreCampoCta = String.Empty
        Dim lblnTodosClientes = (DblIdCliente = 0)
        MstbExpresionSql.Clear()
        With MstbExpresionSql
            .Append("N.").Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND A.").Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND N.").Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCentroUtil).Append(" AND ")
            .Append(ClsFechaNovedadAntDtm.SstrNombreCampoBd).Append(" >= '").Append(StrFechaDesde).Append("' AND ")
            .Append(ClsFechaNovedadAntDtm.SstrNombreCampoBd).Append(" <= '").Append(StrFechaHasta).Append("'")
            .Append(" AND (").Append(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd).Append(" BETWEEN ")
            .Append(EnuTipoNov.EnuCrAntRec).Append(" AND ").Append(EnuTipoNov.EnuDbAntDev)
            .Append(" OR ").Append(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd).Append(" BETWEEN ")
            .Append(EnuTipoNov.EnuRCrAntRec).Append(" AND ").Append(EnuTipoNov.EnuRDbAntDev).Append(")")
            If Not lblnTodosClientes Then
                .Append(" AND N.").Append(ClsIdTercero_NovAntDbl.SstrNombreCampoBd)
                .Append(" = ").Append(DblIdCliente)
            End If
            If ablnDebito Then
                lstrNombreCampoCta = ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd
            Else
                lstrNombreCampoCta = ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd
            End If
            .Append(" AND ").Append(lstrNombreCampoCta).Append(" >= '").Append(StrIdCuentaContIni)
            .Append("' AND ").Append(lstrNombreCampoCta).Append(" <= '").Append(StrIdCuentaContFin).Append("'")
        End With
        Dim lstrFiltro = MstbExpresionSql.ToString
        Return lstrFiltro
    End Function
    Private Function FstrSqlCliente() As String
        MstbExpresionSql.Clear()
        ' Filtro
        With MstbExpresionSql
            .Append(ClsIdClienteDbl.SstrNombreCampoBd).Append(" = ").Append(DblIdCliente)
        End With
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                MstbExpresionSql.ToString
        Dim lstrCamposSelect = {ClsIdClienteDbl.SstrNombreCampoBd,
                              ClsNombreCompletoStr.SstrNombreCampoBd,
                                "'' AS PredioAgrupador", "0 as IntPorCausar"}
        Dim lstrNombreTabla = ClsCliente.SstrNombreTabla
        Dim lstrIndice = {{ClsIdClienteDbl.SstrNombreCampoBd, "ASC"}}
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, Array.Empty(Of String)())
        Return lstrSql
    End Function
    Private Function FstrSqlClientes() As String
        MstbExpresionSql.Clear()
        ' Filtro
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion
        Dim lstrCamposSelect = {ClsIdClienteDbl.SstrNombreCampoBd,
                              ClsNombreCompletoStr.SstrNombreCampoBd,
                                "'' AS PredioAgrupador", "0 as IntPorCausar", "'' AS AgrupadorServicios"}
        Dim lstrNombreTabla = ClsCliente.SstrNombreTabla
        Dim lstrIndice = {{ClsIdClienteDbl.SstrNombreCampoBd, "ASC"}}
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, Array.Empty(Of String)())
        Return lstrSql
    End Function
    Private Sub SComplementeDsAuxCont(adsAuxiliarCont As DataSet)
        MdtbAuxiliarCon = Nothing
        If DblIdCliente > 0 Then
            SComplementeTablaNovedadesCliente(adsAuxiliarCont)
        Else
            Dim ldtbCuentas = adsAuxiliarCont.Tables("PanCuentasCont")
            Dim ldrwCuentas = ldtbCuentas.Select()
            If ldrwCuentas.Length > 0 Then
                For Each ldrwCuenta As DataRow In ldrwCuentas
                    Dim lstrIdCta As String = ClsPanorama.FobjValorCampo(ldrwCuenta(
                            ClsIdCuentaContStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                    SComplementeTablaNovedades(adsAuxiliarCont, lstrIdCta)
                Next
            End If
        End If
        adsAuxiliarCont.Tables.Remove("OriNovedades")
        adsAuxiliarCont.Tables.Add(MdtbAuxiliarCon)
    End Sub
    ''' <summary>
    ''' Complementa la tabla de novedades cuando es un reporte Auxiliar sin Tercero
    ''' </summary>
    ''' <param name="adsAuxiliarCont">DataSet que contiene la DataTable de Novedades(Movimiento Contable)</param>
    ''' <param name="astrIdCta">El código de la cuenta de Contabilidad que va ha ser complementada.</param>
    ''' <remarks></remarks>
    Private Sub SComplementeTablaNovedades(adsAuxiliarCont As DataSet, astrIdCta As String) 'Ok TipoNov
        Dim ldtbNovedades = adsAuxiliarCont.Tables("OriNovedades")
        If IsNothing(MdtbAuxiliarCon) Then
            MdtbAuxiliarCon = ldtbNovedades.Clone
            MdtbAuxiliarCon.TableName = ldtbNovedades.TableName
        End If
        Dim ldecSaldo = 0D
        If BlnCalculaSaldo Then
            ldecSaldo = SInserteSaldoEnAux(astrIdCta)
        End If
        MstbExpresionSql.Clear()
        MstbExpresionSql.Append("IdCta = '").Append(astrIdCta).Append("'")
        Dim lstrFiltro = MstbExpresionSql.ToString
        Dim ldrwNovedades = ldtbNovedades.Select(lstrFiltro)
        If ldrwNovedades.Length > 0 Then
            Dim ldecValorDb As Decimal, ldecValorCr As Decimal
            Dim lstrPredioAgr As String
            Dim lstrPrefDocOri As String, lentIdDocOri As Integer, lstrNroDocOri As String,
                    lstrDocOrigen As String
            Dim lstrDetalle As String, lenuTipoDocOri As EnuTipoDocOri
            Dim lobjDocumento As ClsDocumento
            For Each ldrwNov As DataRow In ldrwNovedades
                lstrPrefDocOri = ClsPanorama.FobjValorCampo(ldrwNov(
                            ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                lentIdDocOri = ClsPanorama.FobjValorCampo(ldrwNov(
                        ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                lstrNroDocOri = ClsPanorama.FstrNumeroDcto(lstrPrefDocOri, lentIdDocOri)
                lenuTipoDocOri = ClsPanorama.FobjValorCampo(ldrwNov(
                        ClsIdTipoDocOrigenByt.SstrNombreCampoBd), EnuTipoValor.enuByte)
                lstrPredioAgr = ClsPanorama.FobjValorCampo(ldrwNov(
                        ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                lobjDocumento = GobjParametros.ColDocumentos(lenuTipoDocOri)
                lstrDetalle = lobjDocumento.ObjNombre_DocStr.ObjValorPro & " Nro. " & lstrNroDocOri
                ldecValorDb = ClsPanorama.FobjValorCampo(ldrwNov("Debito"), EnuTipoValor.enuDecimal)
                ldecValorCr = ClsPanorama.FobjValorCampo(ldrwNov("Credito"), EnuTipoValor.enuDecimal)
                lstrDocOrigen = ClsOrionCop.FstrDocOrigenNovedad(lenuTipoDocOri) & " " & lstrNroDocOri
                ldecSaldo += ldecValorDb - ldecValorCr
                ldrwNov("Detalle") = lstrDetalle
                ldrwNov("NroDocOrigen") = lstrDocOrigen
                ldrwNov("Saldo") = ldecSaldo
                If String.IsNullOrEmpty(lstrPredioAgr) Then
                    ldrwNov(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd) = GCSTRSINPA
                End If
                SCopieNovedadesAux(ldrwNov)
            Next
        End If
    End Sub
    Private Sub SComplementeTablaNovedadesCliente(adsAuxiliarCont As DataSet)
        If StrIdPredioAgru = My.Resources.Todos Then
            Dim lobjCliente = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
            lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, DblIdCliente})
            Dim lstrPrediosAgr = lobjCliente.FstrPrediosAgruClienteConFacturas(False)
            For Each lstrPredioAgr As String In lstrPrediosAgr
                SComplementeNovedadesPredio(adsAuxiliarCont, DblIdCliente,
                        lstrPredioAgr)
            Next
        Else
            SComplementeNovedadesPredio(adsAuxiliarCont, DblIdCliente,
                        StrIdPredioAgru)
        End If
    End Sub
    Private Sub SComplementeNovedadesPredio(adsAuxiliarCont As DataSet, adblIdCliente As Double,
            astrIdPredioAgr As String)
        Dim ldtbNovedades = adsAuxiliarCont.Tables("OriNovedades")
        Dim ldtbCuentasCon = adsAuxiliarCont.Tables("PanCuentasCont")
        Dim ldrwCuentas = ldtbCuentasCon.Select()
        Dim lstrFiltro = String.Empty, ldecSaldo = 0D
        Dim lstrIdCta = String.Empty, ldrwCuenta As DataRow = Nothing
        Dim ldrwNovedadesPredAgrYCta() As DataRow = Nothing
        If IsNothing(MdtbAuxiliarCon) Then
            MdtbAuxiliarCon = ldtbNovedades.Clone
            MdtbAuxiliarCon.TableName = ldtbNovedades.TableName
        End If
        For i = 0 To ldrwCuentas.Length - 1
            ldrwCuenta = ldrwCuentas(i)
            ' Filtro
            lstrIdCta = ldrwCuenta(ClsIdCuentaContStr.SstrNombreCampoBd)
            ' Filtro para cliente, predio y cuenta cont.
            With MstbExpresionSql
                .Clear().Append(ClsIdTercero_NovDbl.SstrNombreCampoBd).Append(" = ")
                .Append(adblIdCliente).Append(" AND ")
                .Append(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd).Append(" = '")
                .Append(astrIdPredioAgr).Append("' AND IdCta = '").Append(lstrIdCta).Append("'")
                lstrFiltro = .ToString
            End With
            ldecSaldo = 0
            If BlnCalculaSaldo Then
                ldecSaldo = SInserteSaldoEnAux(lstrIdCta, adblIdCliente, astrIdPredioAgr, ldecSaldo)
            End If
            ldrwNovedadesPredAgrYCta = ldtbNovedades.Select(lstrFiltro)
            If ldrwNovedadesPredAgrYCta.Length > 0 Then
                Dim ldecValorDb = 0D, ldecValorCr = 0D
                Dim lstrPrefDocOri = String.Empty, lentIdDocOri = 0, lstrNroDocOri = String.Empty,
                        lstrDocOrigen = String.Empty, lstrDetalle = String.Empty, lenuTipoDocOri = EnuTipoDocOri.None
                Dim lobjDocumento As ClsDocumento = Nothing
                For Each ldrwNov As DataRow In ldrwNovedadesPredAgrYCta
                    lstrPrefDocOri = ClsPanorama.FobjValorCampo(ldrwNov(
                            ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                    lentIdDocOri = ClsPanorama.FobjValorCampo(ldrwNov(
                            ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                    lstrNroDocOri = ClsPanorama.FstrNumeroDcto(lstrPrefDocOri, lentIdDocOri)
                    lenuTipoDocOri = ClsPanorama.FobjValorCampo(ldrwNov(
                            ClsIdTipoDocOrigenByt.SstrNombreCampoBd), EnuTipoValor.enuByte)
                    ldecValorDb = ClsPanorama.FobjValorCampo(ldrwNov("Debito"), EnuTipoValor.enuDecimal)
                    ldecValorCr = ClsPanorama.FobjValorCampo(ldrwNov("Credito"), EnuTipoValor.enuDecimal)
                    lobjDocumento = GobjParametros.ColDocumentos(lenuTipoDocOri)
                    lstrDetalle = lobjDocumento.ObjNombre_DocStr.ObjValorPro & " Nro. " & lstrNroDocOri
                    lstrDocOrigen = ClsOrionCop.FstrDocOrigenNovedad(lenuTipoDocOri) & " " & lstrNroDocOri
                    ldecSaldo += ldecValorDb - ldecValorCr
                    ldrwNov("Detalle") = lstrDetalle
                    ldrwNov("NroDocOrigen") = lstrDocOrigen
                    ldrwNov("Saldo") = ldecSaldo
                    If String.IsNullOrEmpty(astrIdPredioAgr) Then
                        ldrwNov(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd) = GCSTRSINPA
                    End If
                    SCopieNovedadesAux(ldrwNov)
                Next
            End If
        Next
    End Sub
    ''' <summary>
    ''' Inserta un nuevo datarow en la  datatable mdtbAuxiliarCon que contiene el saldo al fecha "strFechaDesde" 
    ''' menos un dia para la cuenta pasada en el argumento "astrIdCta" y devuelve su valor.
    ''' </summary>
    ''' <remarks>Debe ser el primer DataRow del DataTable</remarks>
    Private Function SInserteSaldoEnAux(astrIdCta As String) As Decimal
        Dim ldtmFechaSaldo = CType(StrFechaDesde, Date).AddDays(-1)
        Dim ldrwSaldo = MdtbAuxiliarCon.NewRow
        Dim ldecSaldo = FdecSaldoNov(astrIdCta)
        Dim ldecSaldoNovAnt = FdecSaldoNovAnt(astrIdCta)
        ldecSaldo += ldecSaldoNovAnt
        ldrwSaldo(ClsIdTipoDocOrigenByt.SstrNombreCampoBd) = 0
        ldrwSaldo(ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd) = String.Empty
        ldrwSaldo(ClsIdDocOrigenEnt.SstrNombreCampoBd) = 0
        ldrwSaldo(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd) = "Saldo"
        ldrwSaldo(ClsFechaNovedadDtm.SstrNombreCampoBd) = ldtmFechaSaldo
        ldrwSaldo(ClsIdTercero_NovDbl.SstrNombreCampoBd) = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjIdTerceroCentroUtilDbl.ObjValorPro
        ldrwSaldo("Detalle") = "Saldo a la fecha: " & Format(ldtmFechaSaldo, "dd/MM/yyyy")
        ldrwSaldo("NroDocOrigen") = String.Empty
        If ldecSaldo < 0 Then
            ldrwSaldo("Credito") = ldecSaldo * -1
            ldrwSaldo("Debito") = 0
        Else
            ldrwSaldo("Credito") = 0
            ldrwSaldo("Debito") = ldecSaldo
        End If
        ldrwSaldo("Saldo") = ldrwSaldo("Debito") - ldrwSaldo("Credito")
        ldrwSaldo("IdCta") = astrIdCta
        MdtbAuxiliarCon.Rows.Add(ldrwSaldo)
        Return ldecSaldo
    End Function
    ''' <summary>
    ''' Calcula el saldo a la fecha "strFechaDesde" menos un dia para la cuenta y el predio agrupador 
    ''' pasados en los parametreos "astrIdCta" y "astrIdPredioAgr" repectivamente e inserta un nuevo 
    ''' registro en la tabla mdtbAuxiliarCon donde el campo "Saldo" se calcula con base en el valor 
    ''' del saldo calculado y el valor del saldo acumulado pasado en el argumento "adecSaldoAcum"
    ''' </summary>
    ''' <param name="astrIdCta"></param>
    ''' <param name="astrIdPredioAgr"></param>
    ''' <param name="adecSaldoAcum"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SInserteSaldoEnAux(astrIdCta As String, adblIdCliente As Double,
                astrIdPredioAgr As String,
                adecSaldoAcum As Decimal) As Decimal
        Dim ldtmFechaSaldo = CType(StrFechaDesde, Date).AddDays(-1)
        Dim ldrwSaldo = MdtbAuxiliarCon.NewRow
        Dim ldecSaldo = FdecSaldoNov(astrIdCta, adblIdCliente, astrIdPredioAgr)
        Dim lstrIdPredioAgr = astrIdPredioAgr
        If String.IsNullOrEmpty(lstrIdPredioAgr) Then
            lstrIdPredioAgr = GCSTRSINPA
        End If
        ldrwSaldo(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd) = lstrIdPredioAgr
        ldrwSaldo(ClsFechaNovedadDtm.SstrNombreCampoBd) = ldtmFechaSaldo
        ldrwSaldo(ClsIdTercero_NovDbl.SstrNombreCampoBd) = adblIdCliente
        ldrwSaldo("Detalle") = "Saldo a la fecha: " & Format(ldtmFechaSaldo, "dd/MM/yyyy")
        ldrwSaldo("NroDocOrigen") = String.Empty
        If ldecSaldo < 0 Then
            ldrwSaldo("Credito") = ldecSaldo * -1
            ldrwSaldo("Debito") = 0
        Else
            ldrwSaldo("Credito") = 0
            ldrwSaldo("Debito") = ldecSaldo
        End If
        ldrwSaldo("Saldo") = adecSaldoAcum + ldecSaldo
        ldrwSaldo("IdCta") = astrIdCta
        MdtbAuxiliarCon.Rows.Add(ldrwSaldo)
        Return ldecSaldo
    End Function
    ''' <summary>
    ''' Copia el DataRow pasado en el argumento a la DataTable mdtbAuxiliarCon
    ''' </summary>
    ''' <param name="adrwNovedad">DataRow que sera copiado</param>
    Private Sub SCopieNovedadesAux(adrwNovedad As DataRow)
        Dim ldrwNuevoDataRow = MdtbAuxiliarCon.NewRow
        For Each ldclCampo As DataColumn In adrwNovedad.Table.Columns
            Dim lstrNomCam = ldclCampo.ColumnName
            ldrwNuevoDataRow(lstrNomCam) = adrwNovedad(lstrNomCam)
        Next
        MdtbAuxiliarCon.Rows.Add(ldrwNuevoDataRow)
    End Sub
    Private Function FstrExpresionCuentasCont() As String
        MstbExpresionSql.Clear()
        ' Filtro
        With MstbExpresionSql
            .Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ").Append(GshrIdCarpeta).Append(" AND ")
            .Append(ClsIdCuentaContStr.SstrNombreCampoBd).Append(" >= '").Append(StrIdCuentaContIni).Append("' AND ")
            .Append(ClsIdCuentaContStr.SstrNombreCampoBd).Append(" <= '").Append(StrIdCuentaContFin).Append("'")
        End With
        Dim lstrFiltro = MstbExpresionSql.ToString
        Dim lstrCamposSelect = {ClsIdCuentaContStr.SstrNombreCampoBd,
                                ClsNombreCuentaStr.SstrNombreCampoBd,
                                "CAST('1900-01-01' AS DATE) AS Fecha", "0 as Saldo"}
        Dim lstrNombreTabla = ClsCuentaContabilidad.SstrNombreTabla
        Dim lstrIndice = {{ClsIdCuentaContStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, Array.Empty(Of String)())
        Return lstrSql
    End Function
    ''' <summary>
    ''' Devuelve el saldo de la cuenta pasada en el argumento "astrIdCuentaCont" 
    ''' a la fecha dada por el campo "strFechaDesde" menos un dia.
    ''' </summary>
    ''' <param name="astrIdCuentaCont">Id de la cuenta a la cual se le calcula el saldo.</param>
    Private Function FdecSaldoNov(astrIdCuentaCont As String) As Decimal
        Dim ldtmFechaSaldo = CType(StrFechaDesde, Date).AddDays(-1)
        Dim ldtmFechaDesde = GCDTMFECHANULA
        If Not astrIdCuentaCont.StartsWith("13") Then
            If ldtmFechaSaldo.Month = 12 And ldtmFechaSaldo.Day = 31 Then
                Return 0
            Else
                ldtmFechaDesde = DateSerial(ldtmFechaSaldo.Year, 1, 1)
            End If
        End If
        Dim lstrNombreTabla = ClsNovedad.SstrNombreTabla
        Dim lstrFechaSaldo = ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaSaldo)
        Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaDesde)
        MstbExpresionSql.Clear()
        With MstbExpresionSql
            .Append("SUM(").Append(ClsValor_NovDec.SstrNombreCampoBd).Append(")")
        End With
        Dim lstrTotal = MstbExpresionSql.ToString
        Dim lstrTipoDb = "'D' as Tipo"
        Dim lstrCamposSelectDb = {lstrTotal, lstrTipoDb}
        MstbExpresionSql.Clear()
        ' Filtro
        With MstbExpresionSql
            .Clear()
            .Append(ClsOrionCop.StrFiltroUbicacion)
            .Append(" AND ").Append(ClsIdCuentaDb_NovStr.SstrNombreCampoBd)
            .Append(" = '").Append(astrIdCuentaCont).Append("' AND ").Append(ClsFechaNovedadDtm.SstrNombreCampoBd)
            .Append(" >= '").Append(lstrFechaDesde).Append("' AND ").Append(ClsFechaNovedadDtm.SstrNombreCampoBd)
            .Append(" <= '").Append(lstrFechaSaldo).Append("'")
        End With
        Dim lstrFiltro = MstbExpresionSql.ToString
        Dim lstrSqlDb = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelectDb,
                {{"", ""}}, lstrFiltro, Array.Empty(Of String)())
        MstbExpresionSql.Clear()
        Dim lstrTipoCr = "'C' as Tipo"
        Dim lstrCamposSelectCr = {lstrTotal, lstrTipoCr}
        MstbExpresionSql.Clear()
        With MstbExpresionSql
            .Append(ClsOrionCop.StrFiltroUbicacion)
            .Append(" AND ").Append(ClsIdCuentaCr_NovStr.SstrNombreCampoBd)
            .Append(" = '").Append(astrIdCuentaCont).Append("' AND ").Append(ClsFechaNovedadDtm.SstrNombreCampoBd)
            .Append(" >= '").Append(lstrFechaDesde).Append("' AND ").Append(ClsFechaNovedadDtm.SstrNombreCampoBd)
            .Append(" <= '").Append(lstrFechaSaldo).Append("'")
        End With
        lstrFiltro = MstbExpresionSql.ToString
        Dim lstrSqlCr = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelectCr,
                {{"", ""}}, lstrFiltro, Array.Empty(Of String)())
        MstbExpresionSql.Clear()
        With MstbExpresionSql
            .Append("(").Append(lstrSqlDb).Append(") UNION ALL (").Append(lstrSqlCr).Append(")")
        End With
        Dim lstrSql = MstbExpresionSql.ToString
        Dim ldtbSaldo = ClsPanorama.FdtbDataTable(lstrSql)
        Dim ldrwTotales As DataRow() = ldtbSaldo.Select()
        Dim ldecSaldo = 0D
        For Each ldrwTotal As DataRow In ldrwTotales
            If ldrwTotal("Tipo") = "D" Then
                ldecSaldo += ClsPanorama.FobjValorCampo(ldrwTotal("Valor"), EnuTipoValor.enuDecimal)
            Else
                ldecSaldo -= ClsPanorama.FobjValorCampo(ldrwTotal("Valor"), EnuTipoValor.enuDecimal)
            End If
        Next
        Return ldecSaldo
    End Function
    ''' <summary>
    ''' Devuelve el saldo de la cuenta pasada en el argumento "astrIdCuentaCont" correspondiente al
    ''' predio agrupador pasado en el argumento "astrIdPredioAgrupador"
    ''' a la fecha dada por el campo "strFechaDesde" menos un dia.
    ''' </summary>
    ''' <param name="astrIdCuentaCont">Id de la cuenta a la cual se le calcula el saldo.</param>
    Private Function FdecSaldoNov(astrIdCuentaCont As String, adblIdCiente As Double,
            astrIdPredioAgrupador As String) As Decimal
        Dim ldtmFechaSaldo = CType(StrFechaDesde, Date).AddDays(-1)
        Dim ldtmFechaDesde = GCDTMFECHANULA
        If Not astrIdCuentaCont.StartsWith("13") Then
            If ldtmFechaSaldo.Month = 12 And ldtmFechaSaldo.Day = 31 Then
                Return 0
            Else
                ldtmFechaDesde = DateSerial(ldtmFechaSaldo.Year, 1, 1)
            End If
        End If
        Dim lstrNombreTabla = ClsNovedad.SstrNombreTabla
        Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaDesde)
        Dim lstrFechaSaldo = ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaSaldo)
        MstbExpresionSql.Clear()
        With MstbExpresionSql
            .Append("SUM(").Append(ClsValor_NovDec.SstrNombreCampoBd).Append(")")
        End With
        Dim lstrTotal = MstbExpresionSql.ToString
        Dim lstrTipoDb = "'D' as Tipo", lstrFiltroCliente = String.Empty
        Dim lstrCamposSelectDb = {lstrTotal, lstrTipoDb}
        MstbExpresionSql.Clear()
        ' Filtro
        ' Filtro Cliente
        With MstbExpresionSql
            .Clear()
            .Append(" AND ").Append(ClsIdTercero_NovDbl.SstrNombreCampoBd).Append(" = ").Append(adblIdCiente)
            .Append(" AND ").Append(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd)
            .Append(" = '").Append(astrIdPredioAgrupador).Append("'")
        End With
        lstrFiltroCliente = MstbExpresionSql.ToString
        With MstbExpresionSql
            .Clear()
            .Append(ClsOrionCop.StrFiltroUbicacion)
            .Append(lstrFiltroCliente)
            .Append(" AND ").Append(ClsIdCuentaDb_NovStr.SstrNombreCampoBd)
            .Append(" = '").Append(astrIdCuentaCont).Append("' AND ").Append(ClsFechaNovedadDtm.SstrNombreCampoBd)
            .Append(" >= '").Append(lstrFechaDesde).Append("' AND ").Append(ClsFechaNovedadDtm.SstrNombreCampoBd)
            .Append(" <= '").Append(lstrFechaSaldo).Append("'")
        End With
        Dim lstrFiltro = MstbExpresionSql.ToString
        Dim lstrSqlDb = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelectDb,
                {{"", ""}}, lstrFiltro, Array.Empty(Of String)())
        MstbExpresionSql.Clear()
        Dim lstrTipoCr = "'C' as Tipo"
        Dim lstrCamposSelectCr = {lstrTotal, lstrTipoCr}
        MstbExpresionSql.Clear()
        With MstbExpresionSql
            .Append(ClsOrionCop.StrFiltroUbicacion)
            .Append(lstrFiltroCliente)
            .Append(" AND ").Append(ClsIdCuentaCr_NovStr.SstrNombreCampoBd)
            .Append(" = '").Append(astrIdCuentaCont).Append("' AND ").Append(ClsFechaNovedadDtm.SstrNombreCampoBd)
            .Append(" >= '").Append(lstrFechaDesde).Append("' AND ").Append(ClsFechaNovedadDtm.SstrNombreCampoBd)
            .Append(" <= '").Append(lstrFechaSaldo).Append("'")
        End With
        lstrFiltro = MstbExpresionSql.ToString
        Dim lstrSqlCr = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelectCr,
                {{"", ""}}, lstrFiltro, Array.Empty(Of String)())
        MstbExpresionSql.Clear()
        With MstbExpresionSql
            .Append("(").Append(lstrSqlDb).Append(") UNION ALL (").Append(lstrSqlCr).Append(")")
        End With
        Dim lstrSql = MstbExpresionSql.ToString
        Dim ldtbSaldo = ClsPanorama.FdtbDataTable(lstrSql)
        Dim ldrwTotales As DataRow() = ldtbSaldo.Select()
        Dim ldecSaldo = 0D
        For Each ldrwTotal As DataRow In ldrwTotales
            If ldrwTotal("Tipo") = "D" Then
                ldecSaldo += ClsPanorama.FobjValorCampo(ldrwTotal("Valor"), EnuTipoValor.enuDecimal)
            Else
                ldecSaldo -= ClsPanorama.FobjValorCampo(ldrwTotal("Valor"), EnuTipoValor.enuDecimal)
            End If
        Next
        Return ldecSaldo
    End Function
    Friend Function FdtbAuxiliarCon(ablnRefresque As Boolean) As DataTable
        If ablnRefresque Then MdsAuxiliarCon = Nothing
        SGenereDataSetAux()
        Return MdtbAuxiliarCon
    End Function
#End Region
#Region "Relacion Documentos"
    Private Sub SGenereDataSetRelDocs(adsRelDocs As DataSet)
        SRefresqueLogo()
        Dim lstrSqlEncabezado = FstrExpSqlEncabezado()
        Dim lstrSqlDetalle = FstrExpSqlDetalle()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlEncabezado)
        lcolExpresionesSql.Add(lstrSqlDetalle)
        lcolExpresionesSql.Add(FstrExpSqlTotalDocs)
        lcolNombresTablas.Add("OriDocumentos")
        lcolNombresTablas.Add("OriDetalles")
        lcolNombresTablas.Add("OriTotales")
        GobjPanDat.SdsDataSet(adsRelDocs, lcolExpresionesSql, lcolNombresTablas)
        SComplementeDtbServiciosRC(adsRelDocs)
        adsRelDocs.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "RelacionDocs" & ".XML"
        'adsRelDocs.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlEncabezado() As String
        Dim lstrOrden = " ORDER BY TipoDoc, PrefijoDocOrigen, IdDocOrigen"
        Dim lstrSql As String = "(" & FstrExpSqlRelFacturas() & ") UNION ALL (" &
                FstrExpSqlRelRecCaja() & ") UNION ALL (" &
                FstrExpSqlRelND() & ") UNION ALL (" &
                FstrExpSqlRelNCr() & ") UNION ALL (" & FstrExpSqlRelNCon() & ") UNION ALL (" &
                FstrExpSqlReiAnt() & ")" & lstrOrden
        Return lstrSql
    End Function
    Private Function FstrExpSqlDetalle() As String
        Dim lstrOrden = " ORDER BY TipoDoc, PrefijoDocOrigen, IdDocOrigen, PrefijoFac, IdFactura, IdCuentaDb"
        Dim lstrSql As String = FstrExpSqlRelServiciosFac() & " UNION ALL " &
                FstrExpSqlRelServiciosRecCaja() & " UNION ALL (" & FstrExpSqlDetalleND() & ") UNION ALL (" &
                FstrExpSqlDetalleNCr() & ") UNION ALL (" & FstrExpSqlDetalleNCon() & ") UNION ALL (" &
                FstrExpSqlDetalleReiAnt() & ")" & lstrOrden
        Return lstrSql
    End Function
    Private Function FstrExpSqlTotalDocs()
        Dim lstrFechaDesde = "'" & StrFechaDesde & "'"
        Dim lstrFechaHasta = "'" & StrFechaHasta & "'"
        Dim lstrUbicacion = " AND " & ClsOrionCop.StrFiltroUbicacion
        Dim lstrSqlFac = "SELECT '1 - RELACION FACTURAS' as TipoDoc, SUM(ValorFactura) AS GT " &
                "FROM OriFacturas WHERE fechafactura BETWEEN " & lstrFechaDesde & " AND " &
                lstrFechaHasta & lstrUbicacion & " GROUP BY TipoDoc"
        Dim lstrSqlRec = "SELECT '2 - RELACION RECIBOS DE CAJA' as TipoDoc, SUM(Valor) AS GT " &
                "FROM OriRecibosCaja WHERE FechaReciboCaja BETWEEN " & lstrFechaDesde & " AND " &
                lstrFechaHasta & lstrUbicacion & " GROUP BY TipoDoc"
        Dim ldtmFechaDesde As Date = CDate(StrFechaDesde).AddDays(1)
        Dim ldtmFechaHasta As Date = CDate(StrFechaHasta).AddDays(1)
        Dim lstrFechaDesdeNI = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaDesde) & "'"
        Dim lstrFechaHastaNI = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaHasta) & "'"
        Dim lstrSqlNdb = "SELECT '3 - RELACION NOTAS DE INTERESES' as TipoDoc, SUM(Valor) AS GT " &
                "FROM OriNotasDb WHERE FechaNota BETWEEN " & lstrFechaDesdeNI & " AND " &
                lstrFechaHastaNI & lstrUbicacion & " GROUP BY TipoDoc"
        Dim lstrSqlNCr = "SELECT '4 - RELACION NOTAS CREDITO' as TipoDoc, SUM(Valor) AS GT " &
                "FROM OriNotasCr WHERE FechaNota BETWEEN " & lstrFechaDesde & " AND " &
                lstrFechaHasta & lstrUbicacion & " GROUP BY TipoDoc"
        Dim lstrSqlNCon = "SELECT '5 - RELACION NOTAS APLICACION ANTICIPOS' as TipoDoc, SUM(Valor) AS GT " &
                "FROM OriNotasCon WHERE FechaNota BETWEEN " & lstrFechaDesde & " AND " &
                lstrFechaHasta & lstrUbicacion & " GROUP BY TipoDoc"
        Dim lstrSql = "(" & lstrSqlFac & ") UNION ALL (" & lstrSqlRec & ") UNION ALL (" &
                lstrSqlNdb & ") UNION ALL (" & lstrSqlNCr & ") UNION ALL (" & lstrSqlNCon & ")"
        Return lstrSql
    End Function
#Region "Relacion Facturas"
    Private Function FstrExpSqlRelFacturas() As String
        Dim lstrFechaDesde = "'" & StrFechaDesde & "'"
        Dim lstrFechaHasta = "'" & StrFechaHasta & "'"
        Dim lstrSql = "SELECT '1 - RELACION FACTURAS' as TipoDoc, fechafactura AS Fecha, " &
                "F.Prefijo AS PrefijoDocOrigen, F.IdFactura AS IdDocOrigen, IdPredioAgrupador, " &
                "F.IdTerceroCliente, C.NombreCompleto, ValorFactura AS Valor " &
                "FROM OriFacturas AS F INNER JOIN OriClientes AS C " &
                "ON F.IdTerceroCliente = C.IdTerceroCliente AND F.IdCarpeta = C.IdCarpeta AND " &
                " F.IdCentroUtil = C.IdCentroUtil " &
                "WHERE F.IdCarpeta = " & GshrIdCarpeta.ToString & " AND F.IdCentroUtil = " &
                GshrIdCentroUtil.ToString & " AND fechafactura BETWEEN " & lstrFechaDesde +
                " AND " & lstrFechaHasta
        Return lstrSql
    End Function
    Private Function FstrExpSqlRelServiciosFac() As String
        Dim lstrFechaDesde = "'" & StrFechaDesde & "'"
        Dim lstrFechaHasta = "'" & StrFechaHasta & "'"
        Dim lstrSql = "Select '1 - RELACION FACTURAS' as TipoDoc, PrefijoDocOrigen, IdDocOrigen, " &
            "PrefijoFac, IdFactura, RES.IdCuentaDb, RES.IdCuentaCr, IdTipoNovedad, Nombre, " &
            "SUM(Valor) AS Total FROM (SELECT PrefijoDocOrigen, IdDocOrigen, NOV.PrefijoFac, " &
            "NOV.IdFactura, NOV.IdItemFactura, IdTipoNovedad, IdCuentaDb, IdCuentaCr, Valor, IdAno, " &
            "IdServicio, Nombre FROM (SELECT PrefijoDocOrigen, IdDocOrigen, PrefijoFactura as PrefijoFac, IdFactura, " &
            "IdItemFactura, IdCuentaDb, IdCuentaCr, IdTipoNovedad, Valor FROM OriNovedades WHERE " &
            "IdCarpeta = " & GshrIdCarpeta & " AND idcentroutil = " & GshrIdCentroUtil &
            " AND IdTipoDocOrigen = 1 AND IdTipoNovedad <= 2 AND FechaNovedad BETWEEN " & lstrFechaDesde &
            " AND " & lstrFechaHasta & ") AS NOV INNER JOIN (SELECT Prefijo,IdFactura, IdItemFactura, " &
            "IdAno, IdServicio, Detalle as Nombre  FROM OriItemsFactura WHERE IdCarpeta = " & GshrIdCarpeta &
            " AND idcentroutil = " & GshrIdCentroUtil & ") AS ITEM ON NOV.PrefijoFac = ITEM.Prefijo " &
            "AND NOV.IdFactura = ITEM.IdFactura AND NOV.IdItemFactura = ITEM.IdItemFactura) As RES " &
            "INNER JOIN (SELECT IdAno, IdServicio, IdCuentaDb, IdCuentaCr FROM OriServicios WHERE " &
            "IdCarpeta = " & GshrIdCarpeta & " AND idcentroutil = " & GshrIdCentroUtil &
            ") AS S On RES.IdAno = S.IdAno AND RES.IdServicio = S.IdServicio GROUP BY PrefijoDocOrigen, " &
            "IdDocOrigen, PrefijoFac, IdFactura, IdCuentaDb, IdCuentaCr, Nombre"
        Return lstrSql
    End Function
#End Region
#Region "Relacion Recibos de Caja"
    Private Function FstrExpSqlRelRecCaja() As String
        Dim lstrFechaDesde = "'" & StrFechaDesde & "'"
        Dim lstrFechaHasta = "'" & StrFechaHasta & "'"
        Dim lstrSql = "SELECT '2 - RELACION RECIBOS DE CAJA' as TipoDoc, " &
                "FechaReciboCaja as Fecha, Prefijo, IdReciboCaja AS IdDocOrigen, IdPredioAgrupador, " &
                        "R.IdTerceroCliente, NombreCompleto, Valor FROM OriRecibosCaja AS R " &
                        "INNER JOIN OriClientes AS C ON R.IdTerceroCliente = C.IdTerceroCliente AND " &
                        "R.IdCarpeta = C.IdCarpeta AND R.IdCentroUtil = C.IdCentroUtil " &
                        "WHERE R.IdCarpeta = " & GshrIdCarpeta.ToString &
                        " AND R.IdCentroUtil = " & GshrIdCentroUtil.ToString &
                        " AND FechaReciboCaja BETWEEN " & lstrFechaDesde & " AND " & lstrFechaHasta
        Return lstrSql
    End Function
    Private Function FstrExpSqlRelServiciosRecCaja() As String
        Dim lstrPagosCap = FstrExpSqlPagosCap(False)
        Dim lstrPagosMor = FstrExpSqlPagosCap(True)
        Dim lstrDesctoRC = FstrExpSqlDescuentosRC()
        Dim lstrAntRecRC = FstrExpSqlAnticiposRC()
        Dim lstrSql = "(" & lstrPagosCap & ")" & " UNION ALL " & "(" & lstrPagosMor & ")" &
                " UNION ALL " & "(" & lstrDesctoRC & ")" &
                " UNION ALL " & "(" & lstrAntRecRC & ")"
        Return lstrSql
    End Function
    Private Function FstrExpSqlPagosCap(ablnPagoIntMora As Boolean) As String
        Dim lstrFechaDesde = "'" & StrFechaDesde & "'"
        Dim lstrFechaHasta = "'" & StrFechaHasta & "'"
        Dim lstrNombre = "Abono a "
        Dim lentIdTipoNov As Integer = EnuTipoNov.EnuCrPagoCap
        If ablnPagoIntMora Then
            lentIdTipoNov = EnuTipoNov.EnuCrPagoInt
            lstrNombre = "Abono a Int. Mora de "
        End If
        Dim lstrSql = "SELECT '2 - RELACION RECIBOS DE CAJA' as TipoDoc, PrefijoDocOrigen, IdDocOrigen, " &
                "PrefijoFactura, IdFactura, RES.IdCuentaDb, RES.IdCuentaCr, IdTipoNovedad, Nombre, " &
                "SUM(Valor) AS Total" &
                " FROM (SELECT PrefijoDocOrigen, IdDocOrigen, NOV.PrefijoFactura, NOV.IdFactura, " &
                "NOV.IdItemFactura, IdTipoNovedad, IdCuentaDb, IdCuentaCr, Valor, IdAno, IdServicio, Nombre " &
                "FROM (SELECT PrefijoDocOrigen, IdDocOrigen, PrefijoFactura, IdFactura, IdItemFactura, " &
                "IdCuentaDb, IdCuentaCr, IdTipoNovedad, Valor " &
                "FROM OriNovedades WHERE IdCarpeta = " & GshrIdCarpeta.ToString & " AND idcentroutil = " &
                GshrIdCentroUtil.ToString & " AND " &
                "IdTipoDocOrigen = 2 AND IdTipoNovedad = " & lentIdTipoNov.ToString &
                " AND FechaNovedad BETWEEN " & lstrFechaDesde &
                " AND " & lstrFechaHasta & ") AS NOV INNER JOIN " &
                "(SELECT Prefijo,IdFactura, IdItemFactura,CONCAT('" & lstrNombre & "' , Detalle) as Nombre, IdAno, IdServicio " &
                "FROM OriItemsFactura WHERE " &
                "IdCarpeta = " & GshrIdCarpeta.ToString & " AND IdCentroUtil = " & GshrIdCentroUtil.ToString &
                ") AS ITEM ON NOV.PrefijoFactura = ITEM.Prefijo AND  NOV.IdFactura = ITEM.IdFactura AND " &
                "NOV.IdItemFactura = ITEM.IdItemFactura) AS RES INNER JOIN " &
                "(SELECT IdAno, IdServicio, IdCuentaDb, IdCuentaCr FROM OriServicios " &
                "WHERE IdCarpeta = " & GshrIdCarpeta.ToString & " AND IdCentroUtil = " &
                GshrIdCentroUtil.ToString & ") AS S ON RES.IdAno = S.IdAno AND RES.IdServicio = S.IdServicio " &
                "GROUP BY PrefijoDocOrigen, IdDocOrigen, PrefijoFactura, IdFactura, IdCuentaDb, IdCuentaCr, Nombre "
        Return lstrSql
    End Function
    Private Function FstrExpSqlDescuentosRC() As String
        Dim lstrFechaDesde = "'" & StrFechaDesde & "'"
        Dim lstrFechaHasta = "'" & StrFechaHasta & "'"
        Dim lentIdTipoNovIni As Integer = EnuTipoNov.EnuCrDctoCap
        Dim lentIdTipoNovFin As Integer = EnuTipoNov.EnuCrAntRec
        Dim lstrIdTipoNov = "BETWEEN " & lentIdTipoNovIni.ToString & " AND " &
                lentIdTipoNovFin.ToString
        Dim lstrSql = "SELECT '2 - RELACION RECIBOS DE CAJA' AS TipoDoc, PrefijoDocOrigen, IdDocOrigen, " &
                "PrefijoFactura, IdFactura, RES.IdCuentaDb, RES.IdCuentaCr, IdTipoNovedad, Nombre, " &
                "SUM(Valor) AS Total" &
                " FROM (SELECT PrefijoDocOrigen, IdDocOrigen, NOV.PrefijoFactura, NOV.IdFactura, " &
                "NOV.IdItemFactura, IdTipoNovedad, IdCuentaDb, IdCuentaCr, Valor, IdAno, IdServicio, Nombre " &
                "FROM (SELECT PrefijoDocOrigen, IdDocOrigen,PrefijoFactura, IdFactura, IdItemFactura, " &
                "IdCuentaDb, IdCuentaCr, IdTipoNovedad, Valor " &
                "FROM OriNovedades WHERE IdCarpeta = " & GshrIdCarpeta.ToString & " AND idcentroutil = " &
                GshrIdCentroUtil.ToString & " AND " &
                "IdTipoDocOrigen = 2 AND IdTipoNovedad " & lstrIdTipoNov +
                " AND FechaNovedad BETWEEN " & lstrFechaDesde +
                " AND " & lstrFechaHasta & ") AS NOV INNER JOIN " &
                "(SELECT Prefijo,IdFactura, IdItemFactura, IdAno, IdServicio, Detalle AS Nombre FROM OriItemsFactura WHERE " &
                "IdCarpeta = " & GshrIdCarpeta.ToString & " AND IdCentroUtil = " & GshrIdCentroUtil.ToString +
                ") AS ITEM ON NOV.PrefijoFactura = ITEM.Prefijo AND  NOV.IdFactura = ITEM.IdFactura AND " &
                "NOV.IdItemFactura = ITEM.IdItemFactura) AS RES INNER JOIN " &
                "(SELECT IdAno, IdServicio, IdCuentaDb, IdCuentaCr FROM OriServicios " &
                "WHERE IdCarpeta = " & GshrIdCarpeta.ToString & " AND IdCentroUtil = " &
                GshrIdCentroUtil.ToString & ") AS S ON RES.IdAno = S.IdAno AND RES.IdServicio = S.IdServicio " &
                "GROUP BY PrefijoDocOrigen, IdDocOrigen, PrefijoFactura, IdFactura, IdCuentaDb, IdCuentaCr, Nombre "
        Return lstrSql
    End Function
    Private Function FstrExpSqlAnticiposRC() As String
        Dim lstrFechaDesde = "'" & StrFechaDesde & "'"
        Dim lstrFechaHasta = "'" & StrFechaHasta & "'"
        Dim lentIdTipoNovAnt As Integer = EnuTipoNov.EnuCrAntRec
        Dim lstrIdTipoNov = lentIdTipoNovAnt.ToString
        Dim lstrSql = "SELECT '2 - RELACION RECIBOS DE CAJA' AS TipoDoc, PrefijoDocOrigen, IdDocOrigen, " &
                "'' AS PrefijoFactura, 0 AS IdFactura, IdCuentaDb, IdCuentaCr, IdTipoNovedad, " &
                "'' AS Nombre, Valor AS Total FROM OriNovedadesAnt WHERE IdCarpeta = " &
                GshrIdCarpeta.ToString & " AND idcentroutil = " & GshrIdCentroUtil.ToString +
                " AND IdTipoDocOrigen = 2 AND IdTipoNovedad = " & lstrIdTipoNov +
                " AND FechaNovedad BETWEEN " & lstrFechaDesde & " AND " & lstrFechaHasta
        Return lstrSql
    End Function
    Private Shared Sub SComplementeDtbServiciosRC(adsRelDocs As DataSet)
        Dim ldtbSerRC = adsRelDocs.Tables("OriDetalles")
        For Each ldrwSer As DataRow In ldtbSerRC.Select("TipoDoc = '2 - RELACION RECIBOS DE CAJA'")
            Dim lentIdTipoNov As Integer = ClsPanorama.FobjValorCampo(
                    ldrwSer(ClsIdTipoNovedadByt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
            Dim lstrNombre As String = ClsPanorama.FobjValorCampo(
                    ldrwSer("Nombre"), EnuTipoValor.enuString)
            Select Case lentIdTipoNov
                Case EnuTipoNov.EnuCrDctoCap
                    lstrNombre = "Descuento " & lstrNombre
                Case EnuTipoNov.EnuCrDctoInt
                    lstrNombre = "Descuento Int."
                Case EnuTipoNov.EnuCrAntRec
                    lstrNombre = "Anticipo " & lstrNombre
                Case EnuTipoNov.EnuCrRetCre
                    lstrNombre = "ReteCree " & lstrNombre
                Case EnuTipoNov.EnuCrRetFte
                    lstrNombre = "ReteFuente " & lstrNombre
                Case EnuTipoNov.EnuCrRetIca
                    lstrNombre = "ReteIca " & lstrNombre
                Case EnuTipoNov.EnuCrRetIva
                    lstrNombre = "ReteIva " & lstrNombre
            End Select
            ldrwSer("Nombre") = lstrNombre
        Next
        For Each ldrwSer As DataRow In ldtbSerRC.Select("TipoDoc = '4 - RELACION NOTAS CREDITO'")
            Dim lentIdTipoNov As Integer = ClsPanorama.FobjValorCampo(
                    ldrwSer(ClsIdTipoNovedadByt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
            Dim lstrNombre As String = ClsPanorama.FobjValorCampo(
                    ldrwSer("Nombre"), EnuTipoValor.enuString)
            Select Case lentIdTipoNov
                Case EnuTipoNov.EnuCrDctoCap
                    lstrNombre = "Descuento a " & lstrNombre
                Case EnuTipoNov.EnuCrDctoInt
                    lstrNombre = "Dscto a Int. Mora " & lstrNombre
                Case EnuTipoNov.EnuCrRetCre
                    lstrNombre = "ReteCree a " & lstrNombre
                Case EnuTipoNov.EnuCrRetFte
                    lstrNombre = "ReteFuente a " & lstrNombre
                Case EnuTipoNov.EnuCrRetIca
                    lstrNombre = "ReteIca a " & lstrNombre
                Case EnuTipoNov.EnuCrRetIva
                    lstrNombre = "ReteIva a " & lstrNombre
            End Select
            ldrwSer("Nombre") = lstrNombre
        Next
    End Sub
#End Region
#Region "Notas de Intereses"
    Private Function FstrExpSqlRelND() As String
        Dim ldtmFechaDesde As Date = CDate(StrFechaDesde).AddDays(1)
        Dim ldtmFechaHasta As Date = CDate(StrFechaHasta).AddDays(1)
        Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaDesde) & "'"
        Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaHasta) & "'"
        Dim lstrSql = "SELECT '3 - RELACION NOTAS DE INTERESES' AS TipoDoc, FechaNota AS Fecha, " &
                "PrefijoNotaDb AS Prefijo, IdNotaDb AS IdDocOrigen, IdPredioAgrupador, " &
                "N.IdTerceroCliente, NombreCompleto, Valor FROM OriNotasDb AS N " &
                "INNER JOIN OriClientes AS C ON N.IdTerceroCliente = C.IdTerceroCliente  AND " &
                "N.IdCarpeta = C.IdCarpeta AND N.IdCentroUtil = C.IdCentroUtil " &
                "WHERE N.IdCarpeta = " & GshrIdCarpeta.ToString +
                " AND N.IdCentroUtil = " & GshrIdCentroUtil.ToString & " AND FechaNota BETWEEN " &
                lstrFechaDesde & " AND " & lstrFechaHasta
        Return lstrSql
    End Function
    Private Function FstrExpSqlDetalleND() As String
        Dim ldtmFechaDesde As Date = CDate(StrFechaDesde).AddDays(1)
        Dim ldtmFechaHasta As Date = CDate(StrFechaHasta).AddDays(1)
        Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaDesde) & "'"
        Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaHasta) & "'"
        Dim lstrSql = "SELECT '3 - RELACION NOTAS DE INTERESES' AS TipoDoc, PrefijoDocOrigen, " &
                "IdDocOrigen, PrefijoFactura, IdFactura, RES.IdCuentaDb, RES.IdCuentaCr, " &
                "IdTipoNovedad, Nombre, SUM(Valor) AS Total" &
                " FROM (SELECT PrefijoDocOrigen, IdDocOrigen, NOV.PrefijoFactura, NOV.IdFactura, " &
                "NOV.IdItemFactura, IdTipoNovedad, IdCuentaDb, IdCuentaCr, Valor, IdAno, IdServicio, 
                Nombre " &
                "FROM (SELECT PrefijoDocOrigen, IdDocOrigen, PrefijoFactura, IdFactura, " &
                "IdItemFactura, IdCuentaDb, IdCuentaCr, IdTipoNovedad, Valor " &
                "FROM OriNovedades WHERE IdCarpeta = " & GshrIdCarpeta.ToString &
                " AND idcentroutil = " & GshrIdCentroUtil.ToString & " AND " &
                "IdTipoDocOrigen = 4 AND (IdTipoNovedad = " & EnuTipoNov.EnuDbInt &
                " OR IdTipoNovedad = " & EnuTipoNov.EnuDbIvaInt &
                ") AND FechaNovedad BETWEEN " & lstrFechaDesde &
                " AND " & lstrFechaHasta & ") AS NOV INNER JOIN " &
                "(SELECT Prefijo,IdFactura, IdItemFactura, IdAno, IdServicio, " &
                "CONCAT('Int. de Mora a ',  Detalle) AS Nombre FROM OriItemsFactura WHERE " &
                "idCarpeta = " & GshrIdCarpeta.ToString & " AND IdCentroUtil = " &
                GshrIdCentroUtil.ToString & ") AS ITEM ON NOV.PrefijoFactura = ITEM.Prefijo AND  " &
                "NOV.IdFactura = ITEM.IdFactura AND " &
                "NOV.IdItemFactura = ITEM.IdItemFactura) AS RES INNER JOIN " &
                "(SELECT IdAno, IdServicio, IdCuentaDb, IdCuentaCr FROM OriServicios " &
                "WHERE IdCarpeta = " & GshrIdCarpeta.ToString & " AND IdCentroUtil = " &
                GshrIdCentroUtil.ToString & ") AS S ON RES.IdAno = S.IdAno AND RES.IdServicio = 
                S.IdServicio " &
                "GROUP BY PrefijoDocOrigen, IdDocOrigen, PrefijoFactura, IdFactura, IdCuentaDb, " &
                " IdCuentaCr, Nombre "
        Return lstrSql
    End Function
#End Region
#Region "Notas Credito"
    Private Function FstrExpSqlRelNCr() As String
        Dim lstrFechaDesde = "'" & StrFechaDesde & "'"
        Dim lstrFechaHasta = "'" & StrFechaHasta & "'"
        Dim lstrSql = "SELECT '4 - RELACION NOTAS CREDITO' AS TipoDoc, FechaNota AS Fecha, " &
                "PrefijoNotaCr AS Prefijo, IdNotaCr AS IdDocOrigen, IdPredioAgrupador, " &
                "N.IdTerceroCliente, NombreCompleto, Valor FROM OriNotasCr AS N " &
                "INNER JOIN OriClientes AS C ON N.IdTerceroCliente = C.IdTerceroCliente AND " &
                "N.IdCarpeta = C.IdCarpeta AND N.IdCentroUtil = C.IdCentroUtil " &
                "WHERE N.IdCarpeta = " & GshrIdCarpeta.ToString +
                " AND N.IdCentroUtil = " & GshrIdCentroUtil.ToString & " AND FechaNota BETWEEN " &
                lstrFechaDesde & " AND " & lstrFechaHasta
        Return lstrSql
    End Function
    Private Function FstrExpSqlDetalleNCr() As String
        Dim lstrFechaDesde = "'" & StrFechaDesde & "'"
        Dim lstrFechaHasta = "'" & StrFechaHasta & "'"
        Dim lstrSql = "SELECT '4 - RELACION NOTAS CREDITO' AS TipoDoc, PrefijoDocOrigen, IdDocOrigen, " &
                "PrefijoFactura, IdFactura, RES.IdCuentaDb, RES.IdCuentaCr, IdTipoNovedad, Nombre, " &
                "SUM(Valor) AS Total" &
                " FROM (SELECT PrefijoDocOrigen, IdDocOrigen, NOV.PrefijoFactura, NOV.IdFactura, " &
                "NOV.IdItemFactura, IdTipoNovedad, IdCuentaDb, IdCuentaCr, Valor, IdAno, IdServicio, Nombre " &
                "FROM (SELECT PrefijoDocOrigen, IdDocOrigen, PrefijoFactura, IdFactura, IdItemFactura, " &
                "IdCuentaDb, IdCuentaCr, IdTipoNovedad, Valor " &
                "FROM OriNovedades WHERE IdCarpeta = " & GshrIdCarpeta.ToString & " AND idcentroutil = " &
                GshrIdCentroUtil.ToString & " AND " &
                "IdTipoDocOrigen = 5 AND (IdTipoNovedad BETWEEN 8 AND 13 OR " &
                "IdTipoNovedad BETWEEN 17 AND 19)" &
                " AND FechaNovedad BETWEEN " & lstrFechaDesde &
                " AND " & lstrFechaHasta & ") AS NOV INNER JOIN " &
                "(SELECT Prefijo,IdFactura, IdItemFactura, IdAno, IdServicio, Detalle AS Nombre " &
                "FROM OriItemsFactura WHERE " &
                "idCarpeta = " & GshrIdCarpeta.ToString & " AND IdCentroUtil = " & GshrIdCentroUtil.ToString &
                ") AS ITEM ON NOV.PrefijoFactura = ITEM.Prefijo AND  NOV.IdFactura = ITEM.IdFactura AND " &
                "NOV.IdItemFactura = ITEM.IdItemFactura) AS RES INNER JOIN " &
                "(SELECT IdAno, IdServicio, IdCuentaDb, IdCuentaCr FROM OriServicios " &
                "WHERE IdCarpeta = " & GshrIdCarpeta.ToString & " AND IdCentroUtil = " &
                GshrIdCentroUtil.ToString & ") AS S ON RES.IdAno = S.IdAno AND RES.IdServicio = S.IdServicio " &
                "GROUP BY PrefijoDocOrigen, IdDocOrigen, PrefijoFactura, IdFactura, IdCuentaDb, " &
                " IdCuentaCr, Nombre "
        Return lstrSql
    End Function
#End Region
#Region "Notas Aplicación Anticipos (LA)"
    Private Function FstrExpSqlRelNCon() As String
        Dim lstrFechaDesde = "'" & StrFechaDesde & "'"
        Dim lstrFechaHasta = "'" & StrFechaHasta & "'"
        Dim lstrSql = "SELECT '5 - RELACION NOTAS APLICACION ANTICIPOS' AS TipoDoc, FechaNota AS Fecha, " &
                "PrefijoNotaCon AS Prefijo, IdNotaCon AS IdDocOrigen, IdPredioAgrupador, " &
                "N.IdTerceroCliente, NombreCompleto, Valor FROM OriNotasCon AS N " &
                "INNER JOIN OriClientes AS C ON N.IdTerceroCliente = C.IdTerceroCliente  AND " &
                "N.IdCarpeta = C.IdCarpeta AND N.IdCentroUtil = C.IdCentroUtil " &
                "WHERE N.IdCarpeta = " & GshrIdCarpeta.ToString &
                " AND N.IdCentroUtil = " & GshrIdCentroUtil.ToString & " AND FechaNota BETWEEN " &
                lstrFechaDesde & " AND " & lstrFechaHasta
        Return lstrSql
    End Function
    Private Function FstrExpSqlDetalleNCon() As String
        Dim lstrFechaDesde = "'" & StrFechaDesde & "'"
        Dim lstrFechaHasta = "'" & StrFechaHasta & "'"
        Dim lstrSql = "SELECT ' 5 - RELACION NOTAS APLICACION ANTICIPOS' AS TipoDoc, PrefijoDocOrigen, IdDocOrigen, " &
                "PrefijoFactura, IdFactura, RES.IdCuentaDb, RES.IdCuentaCr, IdTipoNovedad, Nombre, " &
                "SUM(Valor) AS Total" &
                " FROM (SELECT PrefijoDocOrigen, IdDocOrigen, NOV.PrefijoFactura, NOV.IdFactura, " &
                "NOV.IdItemFactura, IdTipoNovedad, IdCuentaDb, IdCuentaCr, Valor, IdAno, IdServicio, Nombre " &
                "FROM (SELECT PrefijoDocOrigen, IdDocOrigen, PrefijoFactura, IdFactura, IdItemFactura, " &
                "IdCuentaDb, IdCuentaCr, IdTipoNovedad, Valor " &
                "FROM OriNovedades WHERE IdCarpeta = " & GshrIdCarpeta.ToString & " AND IdCentroutil = " &
                GshrIdCentroUtil.ToString & " AND " &
                "IdTipoDocOrigen = 3 AND IdTipoNovedad BETWEEN 6 AND 7" &
                " AND FechaNovedad BETWEEN " & lstrFechaDesde &
                " AND " & lstrFechaHasta & ") AS NOV INNER JOIN " &
                "(SELECT Prefijo,IdFactura, IdItemFactura, IdAno, IdServicio, " &
                "CONCAT('Anticipo aplicado a ', Detalle) as Nombre FROM OriItemsFactura WHERE " &
                "idCarpeta = " & GshrIdCarpeta.ToString & " AND IdCentroUtil = " & GshrIdCentroUtil.ToString &
                ") AS ITEM ON NOV.PrefijoFactura = ITEM.Prefijo AND  NOV.IdFactura = ITEM.IdFactura AND " &
                "NOV.IdItemFactura = ITEM.IdItemFactura) AS RES INNER JOIN " &
                "(SELECT IdAno, IdServicio, IdCuentaDb, IdCuentaCr FROM OriServicios " &
                "WHERE IdCarpeta = " & GshrIdCarpeta.ToString & " AND IdCentroUtil = " &
                GshrIdCentroUtil.ToString & ") AS S ON RES.IdAno = S.IdAno AND RES.IdServicio = S.IdServicio " &
                "GROUP BY PrefijoDocOrigen, IdDocOrigen, PrefijoFactura, IdFactura, IdCuentaDb, " &
                " IdCuentaCr, Nombre "
        Return lstrSql
    End Function
#End Region
#Region "Reintegro Anticipos"
    Private Function FstrExpSqlReiAnt() As String
        Dim lstrFechaDesde = "'" & StrFechaDesde & "'"
        Dim lstrFechaHasta = "'" & StrFechaHasta & "'"
        Dim lstrSql = "SELECT '6 - RELACION REINTEGRO ANTICIPOS' AS TipoDoc, FechaNota AS Fecha, " &
                "PrefijoNota AS Prefijo, IdNotaDevAnt AS IdDocOrigen, IdPredioAgrupador, " &
                "N.IdTerceroCliente, NombreCompleto, Valor FROM OriNotasDevAnt AS N " &
                "INNER JOIN OriClientes AS C ON N.IdTerceroCliente = C.IdTerceroCliente  AND " &
                "N.IdCarpeta = C.IdCarpeta AND N.IdCentroUtil = C.IdCentroUtil " &
                "WHERE N.IdCarpeta = " & GshrIdCarpeta.ToString &
                " AND N.IdCentroUtil = " & GshrIdCentroUtil.ToString & " AND FechaNota BETWEEN " &
                lstrFechaDesde & " AND " & lstrFechaHasta
        Return lstrSql
    End Function
    Private Function FstrExpSqlDetalleReiAnt() As String
        Dim lstrFechaDesde = "'" & StrFechaDesde & "'"
        Dim lstrFechaHasta = "'" & StrFechaHasta & "'"
        Dim lstrSql = "SELECT '6 - RELACION REINTEGRO ANTICIPOS' AS TipoDoc, PrefijoDocOrigen, IdDocOrigen, " &
                "PrefijoFactura, IdFactura, RES.IdCuentaDb, RES.IdCuentaCr, IdTipoNovedad, Nombre, " &
                "SUM(Valor) AS Total" &
                " FROM (SELECT PrefijoDocOrigen, IdDocOrigen, NOV.PrefijoFactura, NOV.IdFactura, " &
                "NOV.IdItemFactura, IdTipoNovedad, IdCuentaDb, IdCuentaCr, Valor, IdAno, IdServicio " &
                "FROM (SELECT PrefijoDocOrigen, IdDocOrigen, '' AS PrefijoFactura, " &
                "0 AS IdFactura, 0 AS IdItemFactura, " &
                "IdCuentaDb, IdCuentaCr, IdTipoNovedad, Valor " &
                "FROM OriNovedadesAnt WHERE IdCarpeta = " & GshrIdCarpeta.ToString & " AND idcentroutil = " &
                GshrIdCentroUtil.ToString & " AND " &
                "IdTipoDocOrigen = 6 AND IdTipoNovedad =15" &
                " AND FechaNovedad BETWEEN " & lstrFechaDesde &
                " AND " & lstrFechaHasta & ") AS NOV INNER JOIN " &
                "(SELECT Prefijo,IdFactura, IdItemFactura, IdAno, IdServicio FROM OriItemsFactura WHERE " &
                "idCarpeta = " & GshrIdCarpeta.ToString & " AND IdCentroUtil = " & GshrIdCentroUtil.ToString &
                ") AS ITEM ON NOV.PrefijoFactura = ITEM.Prefijo AND  NOV.IdFactura = ITEM.IdFactura AND " &
                "NOV.IdItemFactura = ITEM.IdItemFactura) AS RES INNER JOIN " &
                "(SELECT IdAno, IdServicio, Nombre, IdCuentaDb, IdCuentaCr FROM OriServicios " &
                "WHERE IdCarpeta = " & GshrIdCarpeta.ToString & " AND IdCentroUtil = " &
                GshrIdCentroUtil.ToString & ") AS S ON RES.IdAno = S.IdAno AND RES.IdServicio = S.IdServicio " &
                "GROUP BY PrefijoDocOrigen, IdDocOrigen, PrefijoFactura, IdFactura, IdCuentaDb, " &
                " IdCuentaCr, Nombre "
        Return lstrSql
    End Function
#End Region
#End Region
#Region "Cartera detallada por predio agrupador, cliente, por servicio por factura"
    Private Sub SGenereDataSetCxCDetPorServicio(adsCxCDetPorServicio As DataSet)
        Dim lstrSqlCxCTotal = FstrSqlCxCTotal()
        Dim lstrSqlSaldosPorSer = FstrSqlSaldosPorSer(lstrSqlCxCTotal)
        Dim lstrSqlClientes = FstrSqlClientes()
        Dim lstrSqlCtasDbServicios = FstrCtasDbSer()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlCxCTotal)
        lcolExpresionesSql.Add(lstrSqlSaldosPorSer)
        lcolExpresionesSql.Add(lstrSqlClientes)
        lcolExpresionesSql.Add(lstrSqlCtasDbServicios)
        lcolNombresTablas.Add("OriCarteraXServicios")
        lcolNombresTablas.Add("OriSaldosXServicios")
        lcolNombresTablas.Add("OriClientes")
        lcolNombresTablas.Add("OriCtasDbSer")
        GobjPanDat.SdsDataSet(adsCxCDetPorServicio, lcolExpresionesSql, lcolNombresTablas)
        adsCxCDetPorServicio.Tables.Add(FdtbCentroUtilidad)
        SComplementeDtbCxCDetSer(adsCxCDetPorServicio)
        'Dim lstrNomArch = GstrTrayDatPrg & "CxCDetPorSer" & ".XML"
        'adsCxCDetPorServicio.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Shared Function FstrSqlSaldosPorSer(astrSqlCxCTotal As String) As String
        Dim lstrSql = "SELECT IdAno, IdServicio, Servicio, SUM(Saldo) FROM (" & astrSqlCxCTotal &
                ") AS RE GROUP BY IdAno, idservicio, Servicio"
        Return lstrSql
    End Function
    Private Function FstrSqlCxCTotal() As String
        Dim lstrSql = "SELECT F.IdPredioAgrupador, F.idtercerocliente, T.PrefijoFactura, " &
                "T.idfactura, F.FechaFactura, IdAno, IdServicio,'' AS Servicio, Saldo FROM (" &
                FstrSqlCxCServicios() & " UNION ALL " & FstrSqlCxCInteMora() & ") AS T " &
                "INNER JOIN " & ClsFactura.SstrNombreTabla & " AS F ON T." &
                ClsPrefijoFact_NovStr.SstrNombreCampoBd & " = F." & ClsPrefijo_FactStr.SstrNombreCampoBd &
                " AND T." & ClsIdFactura_NovEnt.SstrNombreCampoBd & " = F." &
                ClsIdFacturaEnt.SstrNombreCampoBd & " WHERE " & MstrCampoBdCarpeta & " = " &
                GshrIdCarpeta & " AND " & MstrCampoBdCentroUtil & " = " & GshrIdCentroUtil &
                " AND Saldo <> 0 ORDER BY " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd & ", " & ClsIdCliente_FactDbl.SstrNombreCampoBd &
                ", " & ClsIdAno_NovShr.SstrNombreCampoBd & ", " & ClsIdServicio_NovShr.SstrNombreCampoBd &
                ", " & ClsFechaFacturaDtm.SstrNombreCampoBd
        Return lstrSql
    End Function
    Private Function FstrSqlCxCServicios() As String
        With MstbExpresionSql
            .Clear()
            .Append("(SELECT ").Append(ClsPrefijoFact_NovStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_NovEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdAno_NovShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdServicio_NovShr.SstrNombreCampoBd).Append(", ")
            .Append("SUM(IF(").Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" BETWEEN 1 AND 2 Or ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 20 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 22 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 24 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" BETWEEN 26 AND 29 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 34").Append(", ")
            .Append(ClsValor_NovDec.SstrNombreCampoBd).Append(", ").Append(" 0) - IF(")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 4 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 6 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 8 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 33 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" BETWEEN 10 AND 13 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" BETWEEN 17 AND 18").Append(", ")
            .Append(ClsValor_NovDec.SstrNombreCampoBd).Append(", ").Append(" 0)) AS Saldo FROM ")
            .Append(ClsNovedad.SstrNombreTabla).Append(" WHERE ").Append(FstrFiltro())
            .Append(" GROUP BY ")
            .Append(ClsPrefijoFact_NovStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_NovEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdAno_NovShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdServicio_NovShr.SstrNombreCampoBd).Append(")")
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrSqlCxCInteMora() As String
        With MstbExpresionSql
            .Clear()
            .Append("(SELECT ").Append(ClsPrefijoFact_NovStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_NovEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdAno_NovShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdServicio_NovShr.SstrNombreCampoBd).Append(", ")
            .Append("SUM(IF(").Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 3 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 21 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 23 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 25 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 35").Append(", ")
            .Append(ClsValor_NovDec.SstrNombreCampoBd).Append(", ").Append("0) - IF(")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 5 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 7 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 9 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 19 OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = 36").Append(", ")
            .Append(ClsValor_NovDec.SstrNombreCampoBd).Append(", ").Append("0)) AS Saldo FROM ")
            .Append(ClsNovedad.SstrNombreTabla).Append(" WHERE ").Append(FstrFiltro())
            .Append(" GROUP BY ")
            .Append(ClsPrefijoFact_NovStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdFactura_NovEnt.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdAno_NovShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdServicio_NovShr.SstrNombreCampoBd).Append(")")
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrFiltro() As String
        Dim lblnEspredioAgr = False, lstrDestino = String.Empty
        If Not IsNothing(ObjParRepDocs) Then
            lblnEspredioAgr = Not String.IsNullOrEmpty(ObjParRepDocs.StrIdPredioAgr)
            If Not lblnEspredioAgr Then
                If ObjParRepDocs.DblIdTercero > 0 Then
                    lstrDestino = ObjParRepDocs.DblIdTercero.ToString
                End If
            Else
                lstrDestino = ObjParRepDocs.StrIdPredioAgr
            End If
        End If
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd &
                " <='" & StrFechaHasta & "'"
        If Not String.IsNullOrEmpty(lstrDestino) Then
            lstrFiltro &= " AND "
            If lblnEspredioAgr Then
                lstrFiltro &= ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd & " = '" & lstrDestino & "'"
            Else
                lstrFiltro &= ClsIdTercero_NovDbl.SstrNombreCampoBd & " = " & lstrDestino
            End If
        End If
        Return lstrFiltro
    End Function
    Private Shared Sub SComplementeDtbCxCDetSer(adsCxCDetPorSer As DataSet)
        Dim ldtbCxCDetSet As DataTable = adsCxCDetPorSer.Tables("OriCarteraXServicios")
        Dim ldtbSaldoPorSer As DataTable = adsCxCDetPorSer.Tables("OriSaldosXServicios")
        Dim lentIdServicio As Integer, lentIdAno As Integer, lstrServicio As String
        Dim lcolSerPer = GobjParametros.ColServiciosPer
        Dim lcolSerAno As Collection
        Dim lobjServicio As ClsServicio
        Dim lobjAno As ClsAno
        Dim lstrKey As String
        For Each ldrw As DataRow In ldtbCxCDetSet.Rows
            lentIdServicio = ClsPanorama.FobjValorCampo(ldrw("IdServicio"), EnuTipoValor.enuInteger)
            If lentIdServicio <> GCSHRIDMORA Then
                lentIdAno = ClsPanorama.FobjValorCampo(ldrw("IdAno"), EnuTipoValor.enuInteger)
                If lentIdAno = 0 Then
                    lstrKey = lentIdAno.ToString & "," & lentIdServicio.ToString
                    lobjServicio = lcolSerPer(lstrKey)
                Else
                    lstrKey = lentIdAno.ToString & "," & lentIdServicio.ToString
                    lobjAno = GobjParametros.ColAnos(lentIdAno.ToString)
                    lcolSerAno = lobjAno.ColServiciosAno
                    lobjServicio = lcolSerAno(lstrKey)
                End If
                lstrServicio = lobjServicio.ObjNombreServicioStr.ObjValorPro
            Else
                lstrServicio = "Intereses de Mora causados"
            End If
            ldrw("Servicio") = lstrServicio
        Next
        For Each ldrw As DataRow In ldtbSaldoPorSer.Rows
            lentIdServicio = ClsPanorama.FobjValorCampo(ldrw("IdServicio"), EnuTipoValor.enuInteger)
            If lentIdServicio <> GCSHRIDMORA Then
                lentIdAno = ClsPanorama.FobjValorCampo(ldrw("IdAno"), EnuTipoValor.enuInteger)
                If lentIdAno = 0 Then
                    lstrKey = lentIdAno.ToString & "," & lentIdServicio.ToString
                    lobjServicio = lcolSerPer(lstrKey)
                Else
                    lstrKey = lentIdAno.ToString & "," & lentIdServicio.ToString
                    lobjAno = GobjParametros.ColAnos(lentIdAno.ToString)
                    lcolSerAno = lobjAno.ColServiciosAno
                    lobjServicio = lcolSerAno(lstrKey)
                End If
                lstrServicio = lobjServicio.ObjNombreServicioStr.ObjValorPro
            Else
                lstrServicio = "Intereses de Mora causados"
            End If
            ldrw("Servicio") = lstrServicio
        Next
        Dim lstrIdCxCIntMora = GobjParametros.ObjIdCtaIntMoraDbStr.ToString()
        Dim ldtbCtasDb As DataTable = adsCxCDetPorSer.Tables("OriCtasDbSer")
        Dim ldrwNewCtaDb As DataRow = ldtbCtasDb.NewRow
        ldrwNewCtaDb(ClsIdAno_ServicioShr.SstrNombreCampoBd) = "000"
        ldrwNewCtaDb(ClsIdServicioShr.SstrNombreCampoBd) = "999"
        ldrwNewCtaDb(ClsCodigoCuentaDbStr.SstrNombreCampoBd) = lstrIdCxCIntMora
        ldtbCtasDb.Rows.Add(ldrwNewCtaDb)
    End Sub
    Private Function FstrCtasDbSer() As String
        Dim lstrTabla = ClsServicio.SstrNombreTabla
        Dim lstrCampSel As String() = {ClsIdAno_ServicioShr.SstrNombreCampoBd,
                ClsIdServicioShr.SstrNombreCampoBd, ClsCodigoCuentaDbStr.SstrNombreCampoBd}
        Dim lstrFilro = ClsOrionCop.StrFiltroUbicacion
        Dim lstrOrden As String(,) = {{ClsIdAno_ServicioShr.SstrNombreCampoBd, "ASC"},
                {ClsIdServicioShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCampSel,
                lstrOrden, lstrFilro, {})
        Return lstrExpSql
    End Function
#End Region
#Region "Resumen Movimiento Contable"
    Private Sub SGenereDataSetResumenMovCon(adsResumenMovCon As DataSet)
        SRefresqueLogo()
        Dim lstrSqlDb = FstrExpSqlResumenMovCon(True)
        Dim lstrSqlCr = FstrExpSqlResumenMovCon(False)
        Dim lstrSqlDbAnt = FstrExpSqlResumenMovConAnt(True)
        Dim lstrSqlCrAnt = FstrExpSqlResumenMovConAnt(False)
        MstbExpresionSql.Clear()
        ' Orden
        With MstbExpresionSql
            .Append(" ORDER BY IdCta")
        End With
        Dim lstrOrden = MstbExpresionSql.ToString
        MstbExpresionSql.Clear()
        ' Consulta Total
        With MstbExpresionSql
            .Append("(").Append(lstrSqlDb).Append(")").Append(" UNION ALL ")
            .Append("(").Append(lstrSqlDbAnt).Append(")").Append(" UNION ALL ")
            .Append("(").Append(lstrSqlCr).Append(")").Append(" UNION ALL ")
            .Append("(").Append(lstrSqlCrAnt).Append(")")
            .Append(lstrOrden)
        End With
        Dim lstrSqlResMovCont = MstbExpresionSql.ToString
        MstbExpresionSql.Clear()
        Dim lstrSqlCuentasCont = FstrExpresionCuentasCont()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlResMovCont)
        lcolExpresionesSql.Add(lstrSqlCuentasCont)
        lcolNombresTablas.Add("OriResumenMovCon")
        lcolNombresTablas.Add("PanCuentasCont")
        GobjPanDat.SdsDataSet(adsResumenMovCon, lcolExpresionesSql, lcolNombresTablas)
        adsResumenMovCon.Tables.Add(FdtbCentroUtilidad)
        SComplementeDsResMovCont(adsResumenMovCon)
        'Dim lstrNomArch = GstrTrayDatPrg & "ResumenMovCont" & ".XML"
        'adsResumenMovCon.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlResumenMovCon(ablnDb As Boolean) As String
        Dim lstrNombreTabla = ClsNovedad.SstrNombreTabla
        Dim lstrCta = String.Empty, lstrTipo = String.Empty, lstrDetalle = "'' AS Detalle"
        Dim lstrDebito = String.Empty, lstrCredito = String.Empty, lstrSaldo = "0 AS Saldo"
        Dim lstrIndice(,) = {{"", ""}}
        MstbExpresionSql.Clear()
        If ablnDb Then
            MstbExpresionSql.Append(ClsIdCuentaDb_NovStr.SstrNombreCampoBd).Append(" AS IdCta")
            lstrCta = MstbExpresionSql.ToString
            lstrTipo = ("'D' as Tipo")
            MstbExpresionSql.Clear()
            With MstbExpresionSql
                .Append("SUM(").Append(ClsValor_NovDec.SstrNombreCampoBd).Append(") AS Debito")
            End With
            lstrDebito = MstbExpresionSql.ToString
            lstrCredito = "0 AS Credito"
            MstbExpresionSql.Clear()
        Else
            MstbExpresionSql.Append(ClsIdCuentaCr_NovStr.SstrNombreCampoBd).Append(" AS IdCta")
            lstrCta = MstbExpresionSql.ToString
            lstrTipo = ("'C' as Tipo")
            lstrDebito = "0 AS Debito"
            MstbExpresionSql.Clear()
            With MstbExpresionSql
                .Append("SUM(").Append(ClsValor_NovDec.SstrNombreCampoBd).Append(") AS Credito")
            End With
            lstrCredito = MstbExpresionSql.ToString
            MstbExpresionSql.Clear()
        End If
        Dim lstrCamposSelect = {lstrDebito, lstrCredito, lstrSaldo,
                        lstrTipo, lstrDetalle, lstrCta}
        Dim lstrCamposAgrup = {"Saldo", "Tipo", "Detalle", "IdCta"}
        Dim lstrFiltro = FstrExpresionFiltro(ablnDb)
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, lstrCamposAgrup)
        Return lstrSql
    End Function
    Private Function FstrExpSqlResumenMovConAnt(ablnDb As Boolean) As String
        Dim lstrCta = String.Empty, lstrTipo = String.Empty, lstrDetalle = "'' AS Detalle"
        Dim lstrDebito = String.Empty, lstrCredito = String.Empty, lstrSaldo = "0 AS Saldo"
        DblIdCliente = 0.0
        MstbExpresionSql.Clear()
        If ablnDb Then
            MstbExpresionSql.Append(ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd).Append(" AS IdCta")
            lstrCta = MstbExpresionSql.ToString
            lstrTipo = ("'D' as Tipo")
            MstbExpresionSql.Clear()
            With MstbExpresionSql
                .Append("SUM(").Append(ClsValor_NovAntDec.SstrNombreCampoBd).Append(") AS Debito")
            End With
            lstrDebito = MstbExpresionSql.ToString
            lstrCredito = "0 AS Credito"
            MstbExpresionSql.Clear()
        Else
            MstbExpresionSql.Append(ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd).Append(" AS IdCta")
            lstrCta = MstbExpresionSql.ToString
            lstrTipo = ("'C' as Tipo")
            lstrDebito = "0 AS Debito"
            MstbExpresionSql.Clear()
            With MstbExpresionSql
                .Append("SUM(").Append(ClsValor_NovDec.SstrNombreCampoBd).Append(") AS Credito")
            End With
            lstrCredito = MstbExpresionSql.ToString
            MstbExpresionSql.Clear()
        End If
        ' Select
        With MstbExpresionSql
            .Append("SELECT ").Append(lstrDebito).Append(", ").Append(lstrCredito).Append(", ")
            .Append(lstrSaldo).Append(", ").Append(lstrTipo).Append(", ").Append(lstrDetalle)
            .Append(", ").Append(lstrCta)
        End With
        Dim lstrSelect = MstbExpresionSql.ToString
        ' From
        With MstbExpresionSql
            .Clear().Append(" FROM ").Append(ClsNovedadAnticipo.SstrNombreTabla)
        End With
        Dim lstrFrom = MstbExpresionSql.ToString
        ' Where
        Dim lstrFiltro = " WHERE " & FstrExpFiltroResMovConAnt(ablnDb)
        ' Group
        With MstbExpresionSql
            .Clear().Append(" GROUP BY ").Append("Saldo")
            .Append(", ").Append("Tipo").Append(", ").Append("Detalle")
            .Append(", ").Append("IdCta")
        End With
        Dim lstrGroup = MstbExpresionSql.ToString
        With MstbExpresionSql
            .Clear().Append(lstrSelect).Append(lstrFrom).Append(lstrFiltro).Append(lstrGroup)
        End With
        Dim lstrSql = MstbExpresionSql.ToString
        Return lstrSql
    End Function
    Private Function FstrExpFiltroResMovConAnt(ablnDebito As Boolean) 'Ok TipoNov
        Dim lstrNombreCampoCta = String.Empty
        MstbExpresionSql.Clear()
        With MstbExpresionSql
            .Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND ").Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCentroUtil).Append(" AND ")
            .Append(ClsFechaNovedadAntDtm.SstrNombreCampoBd).Append(" >= '").Append(StrFechaDesde)
            .Append("' AND ").Append(ClsFechaNovedadAntDtm.SstrNombreCampoBd).Append(" <='")
            .Append(StrFechaHasta).Append("'").Append(" AND (")
            .Append(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd).Append(" BETWEEN ")
            .Append(EnuTipoNov.EnuCrAntRec).Append(" AND ").Append(EnuTipoNov.EnuDbAntDev)
            .Append(" OR ").Append(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd).Append(" BETWEEN ")
            .Append(EnuTipoNov.EnuRCrAntRec).Append(" AND ")
            .Append(EnuTipoNov.EnuRDbAntDev).Append(")")
            If ablnDebito Then
                lstrNombreCampoCta = ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd
            Else
                lstrNombreCampoCta = ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd
            End If
            .Append(" AND ").Append(lstrNombreCampoCta).Append(" >= '").Append(StrIdCuentaContIni)
            .Append("' AND ").Append(lstrNombreCampoCta).Append(" <= '").Append(StrIdCuentaContFin).Append("'")
        End With
        Dim lstrFiltro = MstbExpresionSql.ToString
        Return lstrFiltro
    End Function
    Private Sub SComplementeDsResMovCont(adsResumenMovCon As DataSet)
        MdtbResMovCont = Nothing
        Dim ldtbCuentas = adsResumenMovCon.Tables("PanCuentasCont")
        If ldtbCuentas.Rows.Count > 0 Then
            For Each ldrwCuenta As DataRow In ldtbCuentas.Rows
                Dim lstrIdCta As String = ClsPanorama.FobjValorCampo(ldrwCuenta(
                        ClsIdCuentaContStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                SComplementeTablaResMovCont(adsResumenMovCon, lstrIdCta)
            Next
        End If
        adsResumenMovCon.Tables.Remove("OriResumenMovCon")
        adsResumenMovCon.Tables.Add(MdtbResMovCont)
    End Sub
    Private Sub SComplementeTablaResMovCont(adsResumenMovCon As DataSet, astrIdCta As String) 'Ok TipoNov
        Dim ldtbResMovCont = adsResumenMovCon.Tables("OriResumenMovCon")
        If IsNothing(MdtbResMovCont) Then
            MdtbResMovCont = ldtbResMovCont.Clone
            MdtbResMovCont.TableName = ldtbResMovCont.TableName
        End If
        Dim ldecSaldo As Decimal
        MstbExpresionSql.Clear()
        MstbExpresionSql.Append("IdCta = '").Append(astrIdCta).Append("'")
        Dim lstrFiltro = MstbExpresionSql.ToString
        Dim ldrwResumenMovCont = ldtbResMovCont.Select(lstrFiltro)
        If ldrwResumenMovCont.Length > 0 Then
            ldecSaldo = SInserteSaldoEnRes(astrIdCta)
            Dim ldecValorDb = 0D, ldecValorCr = 0D
            Dim lstrDetalle = "Movimiento entre Fechas"
            Dim lentCantReg = ldrwResumenMovCont.Length
            If lentCantReg > 0 Then
                Dim ldrwResMovCon As DataRow
                For i = 0 To lentCantReg - 1
                    ldrwResMovCon = ldrwResumenMovCont(i)
                    ldecValorDb += ClsPanorama.FobjValorCampo(ldrwResMovCon("Debito"),
                        EnuTipoValor.enuDecimal)
                    ldecValorCr += ClsPanorama.FobjValorCampo(ldrwResMovCon("Credito"),
                        EnuTipoValor.enuDecimal)
                Next
                ldrwResumenMovCont(0)("Debito") = ldecValorDb
                ldrwResumenMovCont(0)("Credito") = ldecValorCr
                For i = 1 To lentCantReg - 1
                    ldtbResMovCont.Rows.Remove(ldrwResumenMovCont(i))
                Next
            End If
            ldecSaldo += ldecValorDb - ldecValorCr
            ldrwResumenMovCont(0)("Detalle") = lstrDetalle
            ldrwResumenMovCont(0)("Saldo") = ldecSaldo
            SCopieNovedadesRes(ldrwResumenMovCont(0))
        End If
    End Sub
    ''' <summary>
    ''' Inserta un nuevo datarow en la  datatable mdtbResumenAuxCont que contiene el saldo al fecha "strFechaDesde" 
    ''' menos un dia para la cuenta pasada en el argumento "astrIdCta" y devuelve su valor.
    ''' </summary>
    ''' <remarks>Debe ser el primer DataRow del DataTable</remarks>
    Private Function SInserteSaldoEnRes(astrIdCta As String) As Decimal
        Dim ldtmFechaSaldo = CType(StrFechaDesde, Date).AddDays(-1)
        Dim ldrwSaldo = MdtbResMovCont.NewRow
        Dim ldecSaldo As Decimal
        Dim ldecSaldoNovAnt = FdecSaldoNovAnt(astrIdCta)
        ldecSaldo = FdecSaldoNov(astrIdCta)
        ldecSaldo += ldecSaldoNovAnt
        ldrwSaldo("Detalle") = "Saldo a la fecha: " & Format(ldtmFechaSaldo, "dd/MM/yyyy")
        If ldecSaldo < 0 Then
            ldrwSaldo("Credito") = ldecSaldo * -1
            ldrwSaldo("Debito") = 0
        Else
            ldrwSaldo("Credito") = 0
            ldrwSaldo("Debito") = ldecSaldo
        End If
        ldrwSaldo("Saldo") = ldrwSaldo("Debito") - ldrwSaldo("Credito")
        ldrwSaldo("IdCta") = astrIdCta
        MdtbResMovCont.Rows.Add(ldrwSaldo)
        Return ldecSaldo
    End Function
    ''' <summary>
    ''' Copia el DataRow pasado en el argumento a la DataTable mdtbResumenMovCont
    ''' </summary>
    ''' <param name="adrwResumenMovCont">DataRow que sera copiado</param>
    Private Sub SCopieNovedadesRes(adrwResumenMovCont As DataRow)
        Dim ldrwNuevoDataRow = MdtbResMovCont.NewRow
        For Each ldclCampo As DataColumn In adrwResumenMovCont.Table.Columns
            Dim lstrNomCam = ldclCampo.ColumnName
            ldrwNuevoDataRow(lstrNomCam) = adrwResumenMovCont(lstrNomCam)
        Next
        MdtbResMovCont.Rows.Add(ldrwNuevoDataRow)
    End Sub
    ''' <summary>
    ''' Devuelve el saldo en la tabla NovedadesAnt de la cuenta pasada en el argumento "astrIdCuentaCont" 
    ''' a la fecha dada por el campo "strFechaDesde" menos un dia.
    ''' </summary> 
    ''' <param name="astrIdCuentaCont">Id de la cuenta a la cual se le calcula el saldo.</param>
    Private Function FdecSaldoNovAnt(astrIdCuentaCont As String) As Decimal
        Dim ldtmFechaSaldo = CType(StrFechaDesde, Date).AddDays(-1)
        Dim ldtmFechaDesde = GCDTMFECHANULA
        Dim lstrIdCtaAntRec As String = GobjParametros.ObjIdCtaAnticiposRecibidosStr.ObjValorPro
        If Not astrIdCuentaCont = lstrIdCtaAntRec Then
            If ldtmFechaSaldo.Month = 12 AndAlso ldtmFechaSaldo.Day = 31 Then
                Return 0
            Else
                ldtmFechaDesde = DateSerial(ldtmFechaSaldo.Year, 1, 1)
            End If
        End If
        Dim lstrNombreTabla = ClsNovedadAnticipo.SstrNombreTabla
        Dim lstrFechaSaldo = ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaSaldo)
        Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaDesde)
        MstbExpresionSql.Clear()
        With MstbExpresionSql
            .Append("SUM(").Append(ClsValor_NovAntDec.SstrNombreCampoBd).Append(")")
        End With
        Dim lstrTotal = MstbExpresionSql.ToString
        Dim lstrTipoDb = "'D' as Tipo"
        Dim lstrCamposSelectDb = {lstrTotal, lstrTipoDb}
        MstbExpresionSql.Clear()
        ' Filtro
        With MstbExpresionSql
            .Clear()
            .Append(ClsOrionCop.StrFiltroUbicacion)
            .Append(" AND ").Append(ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd)
            .Append(" = '").Append(astrIdCuentaCont).Append("' AND ")
            .Append(ClsFechaNovedadAntDtm.SstrNombreCampoBd).Append(" >= '").Append(lstrFechaDesde)
            .Append("' AND ").Append(ClsFechaNovedadAntDtm.SstrNombreCampoBd).Append(" <= '")
            .Append(lstrFechaSaldo).Append("' AND (").Append(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd)
            .Append(" BETWEEN ").Append(EnuTipoNov.EnuDbAntDev).Append(" AND ")
            .Append(EnuTipoNov.EnuDbAntApl).Append(" OR ")
            .Append(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoNov.EnuRCrAntRec).Append(")")
        End With
        Dim lstrFiltro = MstbExpresionSql.ToString
        Dim lstrSqlDb = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelectDb,
                {{"", ""}}, lstrFiltro, Array.Empty(Of String)())
        MstbExpresionSql.Clear()
        Dim lstrTipoCr = "'C' as Tipo"
        Dim lstrCamposSelectCr = {lstrTotal, lstrTipoCr}
        MstbExpresionSql.Clear()
        With MstbExpresionSql
            .Append(ClsOrionCop.StrFiltroUbicacion)
            .Append(" AND ").Append(ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd)
            .Append(" = '").Append(astrIdCuentaCont).Append("' AND ")
            .Append(ClsFechaNovedadAntDtm.SstrNombreCampoBd).Append(" >= '").Append(lstrFechaDesde)
            .Append("' AND ").Append(ClsFechaNovedadAntDtm.SstrNombreCampoBd)
            .Append(" <= '").Append(lstrFechaSaldo).Append("' AND (")
            .Append(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd)
            .Append(" = ").Append(EnuTipoNov.EnuCrAntRec).Append(" OR ")
            .Append(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd).Append(" BETWEEN ")
            .Append(EnuTipoNov.EnuRDbAntDev).Append(" AND ")
            .Append(EnuTipoNov.EnuRDbAntApl).Append(")")
        End With
        lstrFiltro = MstbExpresionSql.ToString
        Dim lstrSqlCr = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelectCr,
                {{"", ""}}, lstrFiltro, Array.Empty(Of String)())
        MstbExpresionSql.Clear()
        With MstbExpresionSql
            .Append("(").Append(lstrSqlDb).Append(") UNION ALL (").Append(lstrSqlCr).Append(")")
        End With
        Dim lstrSql = MstbExpresionSql.ToString
        Dim ldtbSaldo = ClsPanorama.FdtbDataTable(lstrSql)
        Dim ldrwTotales As DataRow() = ldtbSaldo.Select()
        Dim ldecSaldo = 0D
        For Each ldrwTotal As DataRow In ldrwTotales
            If ldrwTotal("Tipo") = "D" Then
                ldecSaldo += ClsPanorama.FobjValorCampo(ldrwTotal("Valor"), EnuTipoValor.enuDecimal)
            Else
                ldecSaldo -= ClsPanorama.FobjValorCampo(ldrwTotal("Valor"), EnuTipoValor.enuDecimal)
            End If
        Next
        Return ldecSaldo
    End Function
#End Region
#Region "EFactura"
    Private Sub SGenereDataSetEFacNoReg(adsDocNoRegEFac As DataSet)
        SRefresqueLogo()
        Dim ldtmFechaInicio As Date = Date.Today.AddDays(-60)
        Dim lstrFechaInicio = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaInicio) & "'"
        Dim lstrSqlFac = FstrExpSqlFacNoReg()
        Dim lstrSqlNcr = FstrExpSqlNcrNoReg()
        Dim lstrSqlNdb = FstrExpSqlNdbNoReg()
        Dim lstrSqlNRcr = FstrExpSqlNotaRCrNoReg()
        Dim lstrSqlNConAju = FstrExpSqlNotaConAjuNoReg()
        MstbExpresionSql.Clear()
        ' Orden
        With MstbExpresionSql
            .Append(" ORDER BY " & "Doc, Prefijo, IdDocumento")
        End With
        Dim lstrOrden = MstbExpresionSql.ToString
        MstbExpresionSql.Clear()
        ' Consulta Total
        With MstbExpresionSql
            .Append("(").Append(lstrSqlFac).Append(")").Append(" UNION ALL ")
            .Append("(").Append(lstrSqlNdb).Append(")").Append(" UNION ALL ")
            .Append("(").Append(lstrSqlNcr).Append(")").Append(" UNION ALL ")
            .Append("(").Append(lstrSqlNRcr).Append(")").Append(" UNION ALL ")
            .Append("(").Append(lstrSqlNConAju).Append(")")
            .Append(lstrOrden)
        End With
        Dim lstrSqlEFacNoReg = MstbExpresionSql.ToString
        MstbExpresionSql.Clear()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlEFacNoReg)
        lcolNombresTablas.Add("OriDocsNoRegEFac")
        GobjPanDat.SdsDataSet(adsDocNoRegEFac, lcolExpresionesSql, lcolNombresTablas)
        adsDocNoRegEFac.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "DoscNoRegEFac" & ".XML"
        'adsDocNoRegEFac.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Shared Function FstrExpSqlFacNoReg() As String
        Dim lstrNombreTabla = ClsFactura.SstrNombreTabla
        Dim lstrIndice(,) = {{"", ""}}
        Dim lstrCamposSelect = {"'Factura' as Doc", ClsPrefijo_FactStr.SstrNombreCampoBd &
                " AS Prefijo", ClsIdFacturaEnt.SstrNombreCampoBd &
                " AS IdDocumento", ClsIdEstadoEDocEnt.SstrNombreCampoBd,
                ClsCUDocStr.SstrNombreCampoBd, ClsCUFEStr.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdEstadoEDocEnt.SstrNombreCampoBd & " < " & EnuEstadoEDoc.EnuEnviada
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrNombreTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, Array.Empty(Of String)())
        Return lstrSql
    End Function
    Private Shared Function FstrExpSqlNdbNoReg() As String
        Dim lstrTabla = ClsNotaDb.SstrNombreTabla
        Dim lstrCamposSelect = {"'Nota Int. Mora' as Doc",
                ClsPrefijo_NotaDbStr.SstrNombreCampoBd & " AS Prefijo",
                ClsIdNotaDbEnt.SstrNombreCampoBd & " AS IdDocumento",
                ClsIdEstadoEDocEnt.SstrNombreCampoBd, ClsCUDocStr.SstrNombreCampoBd,
                ClsCUDEStr.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdEstadoEDocEnt.SstrNombreCampoBd & " < " & EnuEstadoEDoc.EnuEnviada
        Dim lstrIndice(,) = {{"", ""}}
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, Array.Empty(Of String))
        Return lstrSql
    End Function
    Private Shared Function FstrExpSqlNcrNoReg() As String
        Dim lstrTabla = ClsNotaCr.SstrNombreTabla
        Dim lstrCamposSelect = {"'Nota Crédito' as Doc",
                ClsPrefijo_NotaCrStr.SstrNombreCampoBd & " AS Prefijo",
                ClsIdNotaCrEnt.SstrNombreCampoBd & " AS IdDocumento",
                ClsIdEstadoEDocEnt.SstrNombreCampoBd, ClsCUDocStr.SstrNombreCampoBd,
                ClsCUDEStr.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdEstadoEDocEnt.SstrNombreCampoBd & " < " & EnuEstadoEDoc.EnuEnviada & " AND " &
                ClsIdTipoNotaCrByt.SstrNombreCampoBd & " <> " & EnuTipoNotaCrDef.EnuRetenciones
        Dim lstrIndice(,) = {{"", ""}}
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, Array.Empty(Of String))
        Return lstrSql
    End Function
    Private Shared Function FstrExpSqlNotaRCrNoReg() As String
        Dim lstrTabla = ClsNotaReversionCr.SstrNombreTabla
        Dim lstrCamposSelect = {"'Nota Reversión Cr' as Doc",
                ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd & " AS Prefijo",
                ClsIdNotaReversaCrEnt.SstrNombreCampoBd & " AS IdDocumento",
                ClsIdEstadoEDocEnt.SstrNombreCampoBd, ClsCUDocStr.SstrNombreCampoBd,
                ClsCUDEStr.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdEstadoEDocEnt.SstrNombreCampoBd & " < " & EnuEstadoEDoc.EnuEnviada
        Dim lstrIndice(,) = {{"", ""}}
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, Array.Empty(Of String))
        Return lstrSql
    End Function
    Private Shared Function FstrExpSqlNotaConAjuNoReg() As String
        Dim lstrTabla = ClsNotaCon.SstrNombreTabla
        Dim lstrCamposSelect = {"'Nota Aplicación Ajuste' AS Doc",
                ClsPrefijo_NotaConStr.SstrNombreCampoBd & " AS Prefijo",
                ClsIdNotaConEnt.SstrNombreCampoBd & " AS IdDocumento",
                ClsIdEstadoEDocEnt.SstrNombreCampoBd, ClsCUDocStr.SstrNombreCampoBd,
                ClsCUDEStr.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdEstadoEDocEnt.SstrNombreCampoBd & " < " & EnuEstadoEDoc.EnuEnviada
        Dim lstrIndice(,) = {{"", ""}}
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, Array.Empty(Of String))
        Return lstrSql
    End Function
#End Region
#Region "Valores Facturados"
    ' A un Cliente y sus servicios
    Private Sub SGenereDataSetValFact(adsVlrsFacturados As DataSet)
        SRefresqueLogo()
        Dim lstrSqlVlrsFact = FstrExpSqlVlrsFacturados()
        Dim lstrSqlCliente = FstrExpSqlCliente()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlVlrsFact)
        lcolExpresionesSql.Add(lstrSqlCliente)
        lcolNombresTablas.Add("OriVlrsFact")
        lcolNombresTablas.Add("OriCliente")
        GobjPanDat.SdsDataSet(adsVlrsFacturados, lcolExpresionesSql, lcolNombresTablas)
        adsVlrsFacturados.Tables.Add(FdtbCentroUtilidad)
        SComplementeDSValFact(adsVlrsFacturados)
        'Dim lstrNomArch = GstrTrayDatPrg & "VlrsFacturados" & ".XML"
        'adsVlrsFacturados.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlVlrsFacturados() As String
        Dim lstrExpSqlCapitalFacturado As String = FstrExpSqlCapitalFacturado()
        Dim lstrExpSqlInteresesFacturados = FstrExpSqlIntFacturados()
        Dim lstrExpSqlIva = "(" & FstrExpSqlIva() & ") UNION ALL (" & FstrExpSqlIvaInt() & ")"
        Dim lstrExpSql = "(" & lstrExpSqlCapitalFacturado & ") UNION ALL (" &
                lstrExpSqlInteresesFacturados & ") UNION ALL " & lstrExpSqlIva &
                " ORDER BY " & ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd & ", " &
                ClsIdAno_NovShr.SstrNombreCampoBd & " DESC, " & ClsIdServicioShr.SstrNombreCampoBd
        Return lstrExpSql
    End Function
    Private Function FstrExpSqlCapitalFacturado() As String
        Dim lstrFechaIni = "'" & StrFechaDesde & "'"
        Dim lstrFechaFin = "'" & StrFechaHasta & "'"
        Dim lstrTablaPri = ClsNovedad.SstrNombreTabla
        Dim lstrTablaSec = ClsServicio.SstrNombreTabla
        Dim lstrNomCamTipNov = ClsIdTipoNovedadByt.SstrNombreCampoBd
        With MstbExpresionSql
            .Clear.Append("SELECT ").Append(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd)
            .Append(", N.").Append(ClsIdAno_NovShr.SstrNombreCampoBd).Append(", N.")
            .Append(ClsIdServicio_NovShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreServicioStr.SstrNombreCampoBd).Append(", 1 AS Tipo")
            .Append(", SUM(IF(").Append(lstrNomCamTipNov & " = ").Append(EnuTipoNov.EnuDbCap)
            .Append(" OR ").Append(lstrNomCamTipNov & " = ").Append(EnuTipoNov.EnuRCrDctoCap)
            .Append(", ").Append(ClsValor_NovDec.SstrNombreCampoBd).Append(", 0) - IF(")
            .Append(lstrNomCamTipNov).Append(" = ").Append(EnuTipoNov.EnuCrDctoCap & " OR ")
            .Append(lstrNomCamTipNov).Append(" = ").Append(EnuTipoNov.EnuRDbCap & ", ")
            .Append(ClsValor_NovDec.SstrNombreCampoBd & ", 0)) AS Total")
            .Append(" FROM ").Append(lstrTablaPri).Append(" AS N INNER JOIN ")
            .Append(lstrTablaSec).Append(" AS S ON N.")
            .Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = S.")
            .Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" AND N.")
            .Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = S.")
            .Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" AND N.")
            .Append(ClsIdAno_NovShr.SstrNombreCampoBd).Append(" = S.")
            .Append(ClsIdAno_ServicioShr.SstrNombreCampoBd).Append(" AND N.")
            .Append(ClsIdServicio_NovShr.SstrNombreCampoBd).Append(" = S.")
            .Append(ClsIdServicioShr.SstrNombreCampoBd)
        End With
        Dim lstrExpSql = MstbExpresionSql.ToString
        Dim lstrFiltro = " WHERE N." & PanL.ClsIdCarpetaShr.SstrNombreCampoBd &
                " = " & GshrIdCarpeta.ToString & " AND N." &
                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd & " = " &
                GshrIdCentroUtil.ToString & " AND " & ClsIdTercero_NovDbl.SstrNombreCampoBd & " = " &
                DblIdCliente.ToString & " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd & " BETWEEN " &
                lstrFechaIni & " AND " & lstrFechaFin & " AND (" & lstrNomCamTipNov & " = " &
                EnuTipoNov.EnuDbCap & " OR " & lstrNomCamTipNov & " = " & EnuTipoNov.EnuCrDctoCap &
                " OR " & lstrNomCamTipNov & " = " & EnuTipoNov.EnuRDbCap & " OR " &
                lstrNomCamTipNov & " = " & EnuTipoNov.EnuRCrDctoCap & ")"
        lstrExpSql &= lstrFiltro
        lstrExpSql &= " GROUP BY " & ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd & ", " &
                ClsIdAno_NovShr.SstrNombreCampoBd & ", " & ClsIdServicio_NovShr.SstrNombreCampoBd &
                ", Tipo"
        Return lstrExpSql
    End Function
    Private Function FstrExpSqlIntFacturados() As String
        Dim lstrFechaIni = "'" & StrFechaDesde & "'"
        Dim lstrFechaFin = "'" & StrFechaHasta & "'"
        Dim lstrTabla = ClsNovedad.SstrNombreTabla
        Dim lstrNomCamTipNov = ClsIdTipoNovedadByt.SstrNombreCampoBd
        Dim lstrVlr = "SUM(IF(" & lstrNomCamTipNov & " = " & EnuTipoNov.EnuDbInt &
                " OR " & lstrNomCamTipNov & " = " & EnuTipoNov.EnuRCrDctoInt &
                ", " & ClsValor_NovDec.SstrNombreCampoBd & ", 0) - IF(" &
                lstrNomCamTipNov & " = " & EnuTipoNov.EnuCrDctoInt & " OR " &
                lstrNomCamTipNov & " = " & EnuTipoNov.EnuRDbInt & ", " &
                ClsValor_NovDec.SstrNombreCampoBd & ", 0)) AS Total FROM " & lstrTabla
        With MstbExpresionSql
            .Clear.Append("SELECT ").Append(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd)
            .Append(", ").Append(ClsIdAno_NovShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdServicio_NovShr.SstrNombreCampoBd).Append(", ")
            .Append("'Interese de Mora' AS Nombre").Append(", 2 AS Tipo, ")
        End With
        Dim lstrFiltro = " WHERE " & PanL.ClsIdCarpetaShr.SstrNombreCampoBd &
                " = " & GshrIdCarpeta.ToString & " AND " &
                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd & " = " &
                GshrIdCentroUtil.ToString & " AND " & ClsIdTercero_NovDbl.SstrNombreCampoBd & " = " &
                DblIdCliente.ToString & " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd & " BETWEEN " &
                lstrFechaIni & " AND " & lstrFechaFin & " AND (" & lstrNomCamTipNov & " = " &
                EnuTipoNov.EnuDbInt & " OR " & lstrNomCamTipNov & " = " & EnuTipoNov.EnuCrDctoInt &
                " OR " & lstrNomCamTipNov & " = " & EnuTipoNov.EnuRDbInt &
                " OR " & lstrNomCamTipNov & " = " & EnuTipoNov.EnuRCrDctoInt & ")"
        Dim lstrCamGrou = " GROUP BY " & ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd & ", " &
                ClsIdAno_NovShr.SstrNombreCampoBd & ", " & ClsIdServicio_NovShr.SstrNombreCampoBd &
                ", Tipo"
        Dim lstrExpSql = MstbExpresionSql.ToString & lstrVlr & lstrFiltro & lstrCamGrou
        Return lstrExpSql
    End Function
    Private Function FstrExpSqlIva() As String
        Dim lstrFechaIni = "'" & StrFechaDesde & "'"
        Dim lstrFechaFin = "'" & StrFechaHasta & "'"
        Dim lstrTablaPri = ClsNovedad.SstrNombreTabla
        Dim lstrTablaSec = ClsServicio.SstrNombreTabla
        Dim lstrNomCamTipNov = ClsIdTipoNovedadByt.SstrNombreCampoBd
        With MstbExpresionSql
            .Clear.Append("SELECT ").Append(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd)
            .Append(", N.").Append(ClsIdAno_NovShr.SstrNombreCampoBd).Append(", N.")
            .Append(ClsIdServicio_NovShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreServicioStr.SstrNombreCampoBd).Append(", 3 AS Tipo")
            .Append(", SUM(IF(").Append(lstrNomCamTipNov & " = ").Append(EnuTipoNov.EnuDbIva)
            .Append(", ").Append(ClsValor_NovDec.SstrNombreCampoBd).Append(", 0) - IF(")
            .Append(lstrNomCamTipNov).Append(" = ").Append(EnuTipoNov.EnuRDbIva & ", ")
            .Append(ClsValor_NovDec.SstrNombreCampoBd & ", 0)) AS Total")
            .Append(" FROM ").Append(lstrTablaPri).Append(" AS N INNER JOIN ")
            .Append(lstrTablaSec).Append(" AS S ON N.")
            .Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = S.")
            .Append(PanL.ClsIdCarpetaShr.SstrNombreCampoBd).Append(" AND N.")
            .Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = S.")
            .Append(PanL.ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" AND N.")
            .Append(ClsIdAno_NovShr.SstrNombreCampoBd).Append(" = S.")
            .Append(ClsIdAno_ServicioShr.SstrNombreCampoBd).Append(" AND N.")
            .Append(ClsIdServicio_NovShr.SstrNombreCampoBd).Append(" = S.")
            .Append(ClsIdServicioShr.SstrNombreCampoBd)
        End With
        Dim lstrExpSql = MstbExpresionSql.ToString
        Dim lstrFiltro = " WHERE N." & PanL.ClsIdCarpetaShr.SstrNombreCampoBd &
                " = " & GshrIdCarpeta.ToString & " AND N." &
                PanL.ClsIdCentroUtilShr.SstrNombreCampoBd & " = " &
                GshrIdCentroUtil.ToString & " AND " & ClsIdTercero_NovDbl.SstrNombreCampoBd & " = " &
                DblIdCliente.ToString & " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd & " BETWEEN " &
                lstrFechaIni & " AND " & lstrFechaFin & " AND (" & lstrNomCamTipNov & " = " &
                EnuTipoNov.EnuDbIva & " OR " & lstrNomCamTipNov & " = " & EnuTipoNov.EnuRDbIva & ")"
        lstrExpSql &= lstrFiltro
        lstrExpSql &= " GROUP BY " & ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd & ", " &
                ClsIdAno_NovShr.SstrNombreCampoBd & ", " & ClsIdServicio_NovShr.SstrNombreCampoBd &
                ", Tipo"
        Return lstrExpSql
    End Function
    Private Function FstrExpSqlIvaInt() As String
        Dim lstrFechaIni = "'" & StrFechaDesde & "'"
        Dim lstrFechaFin = "'" & StrFechaHasta & "'"
        Dim lstrTablaPri = ClsNovedad.SstrNombreTabla
        Dim lstrNomCamTipNov = ClsIdTipoNovedadByt.SstrNombreCampoBd
        With MstbExpresionSql
            .Clear.Append("SELECT ").Append(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd)
            .Append(", ").Append(ClsIdAno_NovShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdServicio_NovShr.SstrNombreCampoBd).Append(", ")
            .Append("'IVA Intereses de Mora' AS Nombre").Append(", 4 AS Tipo")
            .Append(", SUM(IF(").Append(lstrNomCamTipNov & " = ").Append(EnuTipoNov.EnuDbIvaInt)
            .Append(", ").Append(ClsValor_NovDec.SstrNombreCampoBd).Append(", 0) - IF(")
            .Append(lstrNomCamTipNov).Append(" = ").Append(EnuTipoNov.EnuRDbIvaInt & ", ")
            .Append(ClsValor_NovDec.SstrNombreCampoBd & ", 0)) AS Total")
            .Append(" FROM ").Append(lstrTablaPri)
        End With
        Dim lstrExpSql = MstbExpresionSql.ToString
        Dim lstrFiltro = " WHERE " & PanL.ClsIdCarpetaShr.SstrNombreCampoBd & " = " &
                GshrIdCarpeta.ToString & " AND " & PanL.ClsIdCentroUtilShr.SstrNombreCampoBd &
                " = " & GshrIdCentroUtil.ToString & " AND " & ClsIdTercero_NovDbl.SstrNombreCampoBd &
                " = " & DblIdCliente.ToString & " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd &
                " BETWEEN " & lstrFechaIni & " AND " & lstrFechaFin & " AND (" & lstrNomCamTipNov &
                " = " & EnuTipoNov.EnuDbIvaInt & " OR " & lstrNomCamTipNov & " = " &
                EnuTipoNov.EnuRDbIvaInt & ")"
        lstrExpSql &= lstrFiltro
        lstrExpSql &= " GROUP BY " & ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd & ", " &
                ClsIdAno_NovShr.SstrNombreCampoBd & ", " & ClsIdServicio_NovShr.SstrNombreCampoBd &
                ", Tipo"
        Return lstrExpSql
    End Function
    Private Function FstrExpSqlCliente() As String
        Dim lstrTabla = ClsCliente.SstrNombreTabla
        Dim lstrCampSele = {ClsIdClienteDbl.SstrNombreCampoBd,
                ClsNombreCompletoStr.SstrNombreCampoBd, "0 AS VlrServicio",
                "0 AS ValorIva", "0 AS ValorInt", "0 AS ValorIvaInt", "'' AS TipoId"}
        Dim lstrOrden As String(,) = {{"", ""}}
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdClienteDbl.SstrNombreCampoBd & " = " & DblIdCliente
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCampSele,
                lstrOrden, lstrFiltro, {})
        Return lstrSql
    End Function
    Private Sub SComplementeDSValFact(adsVlrFact As DataSet)
        Dim ldtbValFac = adsVlrFact.Tables("OriVlrsFact")
        Dim ldecTotalServ = 0D, ldecTotalIva = 0D, ldecTotalInt = 0D, ldecTotalIvaInt = 0D
        Dim lentTipo As Integer, ldecVlr As String
        Dim lstrNombre As String, lshrAno As Short, lstrPredio As String
        For Each ldrwValFac As DataRow In ldtbValFac.Rows
            lstrPredio = ClsPanorama.FobjValorCampo(ldrwValFac(
                    ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
            If String.IsNullOrEmpty(lstrPredio) Then
                ldrwValFac(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd) = "Sin Predio"
            End If
            lentTipo = ClsPanorama.FobjValorCampo(ldrwValFac("Tipo"),
                    EnuTipoValor.enuInteger)
            ldecVlr = ClsPanorama.FobjValorCampo(ldrwValFac("Total"),
                    EnuTipoValor.enuDecimal)
            lstrNombre = ClsPanorama.FobjValorCampo(ldrwValFac(
                    ClsNombreServicioStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lshrAno = ClsPanorama.FobjValorCampo(ldrwValFac(
                    ClsIdAno_NovShr.SstrNombreCampoBd),
                    EnuTipoValor.enuShort)
            If lentTipo = 1 Then
                ldecTotalServ += ldecVlr
                If lshrAno > 0 Then
                    lstrNombre &= " del año " & lshrAno.ToString()
                    ldrwValFac(ClsNombreServicioStr.SstrNombreCampoBd) = lstrNombre
                End If
            ElseIf lentTipo = 2 Then
                ldecTotalInt += ldecVlr
            ElseIf lentTipo = 3 Then
                ldecTotalIva += ldecVlr
                lstrNombre = "IVA de " & lstrNombre
                ldrwValFac(ClsNombreServicioStr.SstrNombreCampoBd) = lstrNombre
            ElseIf lentTipo = 4 Then
                ldecTotalIvaInt += ldecVlr
            Else
                Throw New ErrorInesperadoPanLException("Tipo no esperado")
            End If
        Next
        Dim lstrTipoDoc = FstrNomTipoDocId(DblIdCliente)
        Dim ldtbCli As DataTable = adsVlrFact.Tables("OriCliente")
        ldtbCli(0)("VlrServicio") = ldecTotalServ
        ldtbCli(0)("ValorIva") = ldecTotalIva
        ldtbCli(0)("ValorInt") = ldecTotalInt
        ldtbCli(0)("ValorIvaInt") = ldecTotalIvaInt
        ldtbCli(0)("TipoId") = lstrTipoDoc
    End Sub
    ' A Todos los clientes entre fecha
    Private Sub SGenereDataSetValFactTodos(adsVlrsFacturados As DataSet)
        SRefresqueLogo()
        Dim lstrSqlVlrsFact = FstrExpSqlVlrsFacturadosTodos()
        Dim lstrSqlServicios = FstrExpSqlServiciosTodos()
        Dim lcolExpresionesSql As New Collection
        Dim lcolNombresTablas As New Collection
        lcolExpresionesSql.Add(lstrSqlVlrsFact)
        lcolExpresionesSql.Add(lstrSqlServicios)
        lcolNombresTablas.Add("OriVlrsFact")
        lcolNombresTablas.Add("OriServicios")
        GobjPanDat.SdsDataSet(adsVlrsFacturados, lcolExpresionesSql, lcolNombresTablas)
        adsVlrsFacturados.Tables.Add(FdtbCentroUtilidad)
        'Dim lstrNomArch = GstrTrayDatPrg & "VlrsFacturadosTodos" & ".XML"
        'adsVlrsFacturados.WriteXml(lstrNomArch, System.Data.XmlWriteMode.WriteSchema)
    End Sub
    Private Function FstrExpSqlVlrsFacturadosTodos() As String
        With MstbExpresionSql
            .Clear.Append("SELECT ").Append(ClsIdAno_ServicioShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdServicioShr.SstrNombreCampoBd).Append(", F." &
            ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd).Append(", SUM(" &
            ClsValor_ItemFactDec.SstrNombreCampoBd).Append(") AS Valor FROM ")
            .Append(ClsItemFactura.SstrNombreTabla).Append(" AS I INNER JOIN ")
            .Append(ClsFactura.SstrNombreTabla).Append(" AS F ON I.").Append(MstrCampoBdCarpeta)
            .Append(" = F.").Append(MstrCampoBdCarpeta).Append(" AND I.")
            .Append(MstrCampoBdCentroUtil).Append(" = F.").Append(MstrCampoBdCentroUtil)
            .Append(" AND I.").Append(ClsPrefijo_ItemFactStr.SstrNombreCampoBd).Append(" = F.")
            .Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(" AND I.")
            .Append(ClsIdFactura_ItemFactEnt.SstrNombreCampoBd).Append(" = F.")
            .Append(ClsIdFacturaEnt.SstrNombreCampoBd).Append(" INNER JOIN ")
            .Append(ClsCliente.SstrNombreTabla).Append(" AS C ON C.").Append(MstrCampoBdCarpeta)
            .Append(" = F.").Append(MstrCampoBdCarpeta).Append(" AND C.")
            .Append(MstrCampoBdCentroUtil).Append(" = F.").Append(MstrCampoBdCentroUtil)
            .Append(" AND C.").Append(ClsIdClienteDbl.SstrNombreCampoBd).Append(" = F.")
            .Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(" WHERE I.")
            .Append(MstrCampoBdCarpeta).Append(" = ").Append(GshrIdCarpeta).Append(" AND I.")
            .Append(MstrCampoBdCentroUtil).Append(" = ").Append(GshrIdCentroUtil).Append(" AND ")
            .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(" BETWEEN '").Append(StrFechaDesde)
            .Append("' AND '").Append(StrFechaHasta).Append("' AND F.")
            .Append(ClsAnuladoBln.SstrNombreCampoBd).Append(" = FALSE  GROUP BY ")
            .Append(ClsIdAno_ServicioShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdServicioShr.SstrNombreCampoBd).Append(", F.")
            .Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(", ")
            .Append(ClsNombreCompletoStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd).Append(" ORDER BY ")
            .Append(ClsIdAno_ServicioShr.SstrNombreCampoBd).Append(" DESC ").Append(", ")
            .Append(ClsIdServicioShr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd).Append(", ")
            .Append(ClsIdCliente_FactDbl.SstrNombreCampoBd)
        End With
        Return MstbExpresionSql.ToString
    End Function
    Private Function FstrExpSqlServiciosTodos() As String
        Dim lstrTabla = ClsServicio.SstrNombreTabla
        Dim lstrCampSel As String() = {ClsIdAno_ServicioShr.SstrNombreCampoBd,
                ClsIdServicioShr.SstrNombreCampoBd, ClsConceptoServicioStr.SstrNombreCampoBd}
        Dim lstrfiltro = ClsOrionCop.StrFiltroUbicacion
        Dim lstrOrden As String(,) = {{ClsIdAno_ServicioShr.SstrNombreCampoBd, "DESC"},
                {ClsIdServicioShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCampSel, lstrOrden,
                lstrfiltro, {})
        Return lstrExpSql
    End Function
#End Region
#Region "Copropiedad"
    Private Function FdtbCentroUtilidad() As DataTable
        Dim ldclIdCentrUtil As New DataColumn(ClsIdTerceroCentroUtilDbl.SstrNombreCampoBd,
                Type.GetType("System.Double"))
        Dim ldclNomCenUtil As New DataColumn(ClsNombreCentroUtilStr.SstrNombreCampoBd,
                Type.GetType("System.String"))
        Dim ldclDireccion1 As New DataColumn(ClsDireccionUnoStr.SstrNombreCampoBd,
                Type.GetType("System.String"))
        Dim ldclDireccion2 As New DataColumn(ClsDireccionDosStr.SstrNombreCampoBd,
                Type.GetType("System.String"))
        Dim ldclTel1 As New DataColumn(ClsTelefonoUnoStr.SstrNombreCampoBd,
                Type.GetType("System.String"))
        Dim ldclTel2 As New DataColumn(ClsTelefonoDosStr.SstrNombreCampoBd,
                Type.GetType("System.String"))
        Dim ldclPagWEb As New DataColumn(ClsPaginaWebStr.SstrNombreCampoBd,
                Type.GetType("System.String"))
        Dim ldclFechaDatos As New DataColumn("FechaDatos", Type.GetType("System.DateTime"))
        Dim ldclFechaDesde As New DataColumn("FechaDesde", Type.GetType("System.DateTime"))
        Dim ldclFechaHasta As New DataColumn("FechaHasta", Type.GetType("System.DateTime"))
        Dim ldclLogo As New DataColumn("Logo", Type.GetType("System.Byte[]"))
        Dim ldclFirma As New DataColumn("Firma", Type.GetType("System.Byte[]"))
        Dim ldclCodigoQR As New DataColumn("CodigoQR", Type.GetType("System.Byte[]"))
        Dim ldclNombreReLeg As New DataColumn("NombreRepLeg", Type.GetType("System.String"))
        Dim ldclIdRepLegal As New DataColumn("NroIdRepLeg", Type.GetType("System.Double"))
        Dim ldclTipoDocId As New DataColumn("TipoDocId", Type.GetType("System.String"))
        Dim ldclTelRepLeg As New DataColumn("TelRepLegal", Type.GetType("System.String"))
        Dim ldclCargo As New DataColumn("Cargo", Type.GetType("System.String"))
        Dim ldtbCentroUtilidad = New DataTable(ClsCentroUtilidad.SstrNombreTabla)
        ldtbCentroUtilidad.Columns.AddRange({ldclIdCentrUtil, ldclNomCenUtil, ldclDireccion1,
                ldclDireccion2, ldclFechaDatos, ldclFechaDesde, ldclFechaHasta, ldclLogo, ldclFirma,
                ldclCodigoQR, ldclNombreReLeg, ldclCargo, ldclTel1, ldclTel2, ldclPagWEb, ldclIdRepLegal,
                ldclTipoDocId, ldclTelRepLeg})
        Dim ldrwCenUtil = ldtbCentroUtilidad.NewRow
        With GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual
            ldrwCenUtil(ClsIdTerceroCentroUtilDbl.SstrNombreCampoBd) =
                    .ObjIdTerceroCentroUtilDbl.ObjValorPro
            ldrwCenUtil(ClsNombreCentroUtilStr.SstrNombreCampoBd) =
                    .ObjTerceroCentroUtilidad.FstrNombreCompleto()
            ldrwCenUtil(ClsDireccionUnoStr.SstrNombreCampoBd) =
                    .StrDireccion()
            ldrwCenUtil(ClsDireccionDosStr.SstrNombreCampoBd) =
                    .ObjTerceroCentroUtilidad.ObjDireccionDosStr.ObjValorPro
            ldrwCenUtil("FechaDatos") = Date.Now
            ldrwCenUtil("FechaDesde") = CDate(StrFechaDesde)
            ldrwCenUtil("FechaHasta") = CDate(StrFechaHasta)
            ldrwCenUtil("Logo") = MbytLogoCenUtilidad
            ldrwCenUtil("Firma") = .ObjTerRepLegal.FbytFirma
            ldrwCenUtil("CodigoQR") = MbytCodigoQR
            ldrwCenUtil("Cargo") = .ObjCargoStr.ToString
            ldrwCenUtil("NombreRepLeg") = .ObjIdTerceroRepLegalDbl.StrNombreRepLegal
            ldrwCenUtil("NroIdRepLeg") = .ObjIdTerceroRepLegalDbl.ObjValorPro
            ldrwCenUtil("TipoDocId") = FstrNomTipoDocId(.ObjIdTerceroRepLegalDbl.ObjValorPro)
            ldrwCenUtil("TipoDocId") = FstrNomTipoDocId(.ObjIdTerceroRepLegalDbl.ObjValorPro)
            ldrwCenUtil("TelRepLegal") = .ObjTerRepLegal.StrTelefono
            ldrwCenUtil(ClsTelefonoUnoStr.SstrNombreCampoBd) =
                    .ObjTerceroCentroUtilidad.StrTelefono
            ldrwCenUtil("PaginaWeb") = .ObjTerceroCentroUtilidad.ObjPaginaWebStr.ToString
        End With
        ldtbCentroUtilidad.Rows.Add(ldrwCenUtil)
        Return ldtbCentroUtilidad
    End Function
    Private Function FdtbCentroUtilidadSinFirma() As DataTable
        Dim ldclIdCentrUtil As New DataColumn(ClsIdTerceroCentroUtilDbl.SstrNombreCampoBd,
                    System.Type.GetType("System.Double"))
        Dim ldclNomCenUtil As New DataColumn(ClsNombreCentroUtilStr.SstrNombreCampoBd,
                System.Type.GetType("System.String"))
        Dim ldclDireccion1 As New DataColumn(ClsDireccionUnoStr.SstrNombreCampoBd,
                System.Type.GetType("System.String"))
        Dim ldclDireccion2 As New DataColumn(ClsDireccionDosStr.SstrNombreCampoBd,
                System.Type.GetType("System.String"))
        Dim ldclTel1 As New DataColumn(ClsTelefonoUnoStr.SstrNombreCampoBd,
                System.Type.GetType("System.String"))
        Dim ldclTel2 As New DataColumn(ClsTelefonoDosStr.SstrNombreCampoBd,
                System.Type.GetType("System.String"))
        Dim ldclPagWEb As New DataColumn(ClsPaginaWebStr.SstrNombreCampoBd,
                System.Type.GetType("System.String"))
        Dim ldclFechaDatos As New DataColumn("FechaDatos", System.Type.GetType("System.DateTime"))
        Dim ldclFechaDesde As New DataColumn("FechaDesde", System.Type.GetType("System.DateTime"))
        Dim ldclFechaHasta As New DataColumn("FechaHasta", System.Type.GetType("System.DateTime"))
        Dim ldclLogo As New DataColumn("Logo", System.Type.GetType("System.Byte[]"))
        Dim ldclFirma As New DataColumn("Firma", System.Type.GetType("System.Byte[]"))
        Dim ldclNombreReLeg As New DataColumn("NombreRepLeg", System.Type.GetType("System.String"))
        Dim ldclCargo As New DataColumn("Cargo", System.Type.GetType("System.String"))
        Dim ldtbCentroUtilidad = New DataTable(ClsCentroUtilidad.SstrNombreTabla)
        ldtbCentroUtilidad.Columns.AddRange({ldclIdCentrUtil, ldclNomCenUtil, ldclDireccion1,
                ldclDireccion2, ldclFechaDatos, ldclFechaDesde, ldclFechaHasta, ldclLogo, ldclFirma,
                ldclNombreReLeg, ldclCargo, ldclTel1, ldclTel2, ldclPagWEb})
        Dim ldrwCenUtil = ldtbCentroUtilidad.NewRow
        With GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual
            ldrwCenUtil(ClsIdTerceroCentroUtilDbl.SstrNombreCampoBd) = .ObjIdTerceroCentroUtilDbl.ObjValorPro
            ldrwCenUtil(ClsNombreCentroUtilStr.SstrNombreCampoBd) =
                    .ObjTerceroCentroUtilidad.FstrNombreCompleto()
            ldrwCenUtil(ClsDireccionUnoStr.SstrNombreCampoBd) =
                    .ObjTerceroCentroUtilidad.ObjDireccionUnoStr.ObjValorPro
            ldrwCenUtil(ClsDireccionDosStr.SstrNombreCampoBd) =
                    .ObjTerceroCentroUtilidad.ObjDireccionDosStr.ObjValorPro
            ldrwCenUtil("FechaDatos") = Date.Now
            ldrwCenUtil("FechaDesde") = CDate(StrFechaDesde)
            ldrwCenUtil("FechaHasta") = CDate(StrFechaHasta)
            ldrwCenUtil("Logo") = MbytLogoCenUtilidad
            ldrwCenUtil("Cargo") = .ObjCargoStr.ToString
            ldrwCenUtil("NombreRepLeg") = .ObjIdTerceroRepLegalDbl.StrNombreRepLegal
            ldrwCenUtil("Telefono1") = .ObjTerceroCentroUtilidad.ObjTelefonoUnoStr.ToString
            ldrwCenUtil("Telefono2") = .ObjTerceroCentroUtilidad.ObjCelularStr.ToString
            ldrwCenUtil("PaginaWeb") = .ObjTerceroCentroUtilidad.ObjPaginaWebStr.ToString
        End With
        ldtbCentroUtilidad.Rows.Add(ldrwCenUtil)
        Return ldtbCentroUtilidad
    End Function
    Private Sub SRefresqueLogo()
        Dim lobjTerCenUti As ClsTercero = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjTerceroCentroUtilidad
        lobjTerCenUti.SRefresqueObj()
        If Not IsNothing(lobjTerCenUti.ImgUltimaImagen) Then
            MbytLogoCenUtilidad =
                    lobjTerCenUti.ImgUltimaImagen.ObjPropiedadImagenImg.BytImagenGuardada
        Else
            MbytLogoCenUtilidad = Nothing
        End If
    End Sub
    Private Sub SRefresqueLogo(aobjTerceroCentroUtil As ClsTercero)
        aobjTerceroCentroUtil.SRefresqueObj()
        If Not IsNothing(aobjTerceroCentroUtil.ImgUltimaImagen) Then
            MbytLogoCenUtilidad = aobjTerceroCentroUtil.ImgUltimaImagen.ObjPropiedadImagenImg.BytImagenGuardada
        Else
            MbytLogoCenUtilidad = Nothing
        End If
    End Sub
    Private Sub SRefresqueCodigoQR()
        MbytCodigoQR = Nothing
        Dim ldtbCodigoQR = ClsOrionCop.FdtbImagenBancoQR(0)
        If ldtbCodigoQR.Rows.Count > 0 Then
            Dim lobjImagen As New ClsImagen(Me, ldtbCodigoQR.Rows(0))
            lobjImagen.SLeaValores(True)
            MbytCodigoQR = lobjImagen.ObjPropiedadImagenImg.BytImagenGuardada
        End If
    End Sub
#End Region
#End Region
#Region "Dispose"
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overridable Overloads Sub Dispose(ablnDisposing As Boolean)
        If Not MblnDisposed Then
            If ablnDisposing Then
                If Not IsNothing(MdsAuxiliarCon) Then
                    MdsAuxiliarCon.Dispose()
                End If
                If Not IsNothing(DsMovimiento) Then
                    DsMovimiento.Dispose()
                End If
                If Not IsNothing(MdtbAuxiliarCon) Then
                    MdtbAuxiliarCon.Dispose()
                End If
                If Not IsNothing(MdtbResMovCont) Then
                    MdtbResMovCont.Dispose()
                End If
                MblnDisposed = True
            End If
        End If
    End Sub
#End Region
End Class
Friend Class ClsParametrosReportesDocs
    Friend Property StrPrefijoDocsRep As String = String.Empty
    Friend Property EntIdDocInicial As Integer = 0
    Friend Property EntIdDocFinal As Integer = 0
    Friend Property StrIdPredioAgr As String = String.Empty
    Friend Property DblIdTercero As Double = 0
    Friend Property BlnExcluirFacEnvEmail As Boolean = False
    Friend Sub New(astrPrefDocs As String, aentIdDocIni As Integer,
                   aentIdDocFin As Integer)
        StrPrefijoDocsRep = astrPrefDocs
        EntIdDocInicial = aentIdDocIni
        EntIdDocFinal = aentIdDocFin
    End Sub
End Class