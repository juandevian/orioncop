Friend Class ClsCalculosServicios
    Friend Shared Function FblnCalculoItemsServicio(aobjServicio As ClsServicio) As Boolean
        Dim lblnNoHayError = False, lblnCalculo = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            SElimineItemsServicio(aobjServicio)
            SCalculeContribucionSectores(aobjServicio)
            lblnCalculo = FblnCreoItemsPrograma(aobjServicio)
            If lblnCalculo Then
                If Not aobjServicio.BlnEsCuotaAdministracion Then
                    aobjServicio.SActuEstaGeneradoProgramaFact(True)
                    Dim lstrAccion As String = "Calculo valores a cobrar Servicio No. " &
                            aobjServicio.ObjIdServicioShr.ObjValorPro & " / " &
                            aobjServicio.ObjNombreServicioStr.ObjValorPro
                    GobjPanorama.SRegistreAccionLogApp("clsCalculosServicios", lstrAccion)
                End If
            End If
            lblnNoHayError = True
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError AndAlso lblnCalculo Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
        Return lblnCalculo
    End Function
    ''' <summary>
    ''' Elimina los ítems del programa de facturación correspondientes al servicio pasado en
    ''' el argumento
    ''' </summary>
    ''' <param name="aobjServicio"></param>
    Friend Shared Sub SElimineItemsServicio(astrKeySer As String)
        Dim lobjServicio As ClsServicio = GobjParametros.FobjServicio(astrKeySer)
        SElimineItemsServicio(lobjServicio)
    End Sub
    Friend Shared Sub SElimineItemsServicio(aobjServicio As ClsServicio)
        Dim lstrTabla = ClsItemProgramaFact.SstrNombreTabla
        Dim lstrIdAno As String = aobjServicio.ObjIdAno_ServicioShr.ToString()
        Dim lstrIdServicio = aobjServicio.ObjIdServicioShr.ToString()
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & lstrIdAno &
                " AND " & ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " = " &
                lstrIdServicio
        Dim lstrExpSqlEli = ClsPanoramaDat.FstrConstruyaExpSqlEliminar(lstrTabla, lstrFiltro)
        Dim lentCantEli = GobjPanDat.SEjecuteSentenciaSql(lstrExpSqlEli)
        aobjServicio.SActuEstaGeneradoProgramaFact(False)
    End Sub
    Friend Shared Sub SLimpieItemsServicio(aobjServicio As ClsServicio)
        Dim lstrTabla = ClsItemProgramaFact.SstrNombreTabla
        Dim lstrIdAno As String = aobjServicio.ObjIdAno_ServicioShr.ToString()
        Dim lstrIdServicio = aobjServicio.ObjIdServicioShr.ToString()
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & lstrIdAno &
                " AND " & ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " = " &
                lstrIdServicio
        Dim lstrExpSqlAct = "UPDATE " & lstrTabla & " SET " &
                ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd & " = 0, " &
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd & " = 0 WHERE " &
                lstrFiltro
        Dim lentReg = GobjPanDat.SEjecuteSentenciaSql(lstrExpSqlAct)
        If lentReg > 0 Then
            aobjServicio.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
            aobjServicio.ObjEstaGenaradaProgramBln.ObjValorPro = False
            aobjServicio.SActualice(True)
        End If
    End Sub
    ''' <summary>
    ''' Calcula la contribucion de los Sectores que contribuyen con los Modulos que a su vez
    ''' contribuyen con el Servicio pasado en el Argumento
    ''' </summary>
    ''' <param name="aobjServicio">Servicio al cual contribuyen los sectores.</param>
    ''' <remarks></remarks>
    Friend Shared Sub SCalculeContribucionSectores(aobjServicio As ClsServicio)
        Dim lcolModulosServicio = aobjServicio.ColModulosServicio
        SElimineSectoresModuloServicio(aobjServicio)
        Dim lblnActualizaMod = False
        For Each lobjModuloServicio As ClsModuloServicio In lcolModulosServicio
            If lobjModuloServicio.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                lobjModuloServicio.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                lblnActualizaMod = True
            End If
            ' Calcula y actualiza la contribucion de cada Sector que contribuye al módulo
            SActualiceSectoresModuloServicio(lobjModuloServicio)
            If lblnActualizaMod Then
                lobjModuloServicio.SActualice(True)
            End If
        Next
    End Sub
    ''' <summary>
    ''' Calcula la contribucion de cada uno de los Sectores que contribuyen al Módulo del Servicio
    ''' pasado en el argumento y actualiza los objetos "SectorModuloServicio".
    ''' </summary>
    ''' <param name="aobjModuloServicio">Objeto que represente el Modulo que Contribuye al Servicio.</param>
    ''' <remarks></remarks>
    Friend Shared Sub SActualiceSectoresModuloServicio(aobjModuloServicio As ClsModuloServicio)
        Dim lobjServicio As ClsServicio = aobjModuloServicio.ObjMiServicio
        Dim lenuTipoBaseCalcSer As EnuTipoBaseCalculo =
                lobjServicio.ObjTipoBaseCalculoByt.ObjValorPro
        Dim lblnNuevo As Boolean, ldblTotalPar = 0
        Dim lcolSectores_ModuloCont As Collection =
                aobjModuloServicio.ObjMiModuloContribucion.ColSectoresModulo
        Dim lcolModulosContr = GobjParametros.ColModulos
        Dim lcolSectores_ModuloServicio = aobjModuloServicio.ColSectores_ModuloServicio
        Dim lobjModuloContribucion As ClsModuloContribucion =
                lcolModulosContr(aobjModuloServicio.ObjIdModulo_ModuloServicioShr.ToString)
        Dim ldecContribucionModulo As Decimal =
                aobjModuloServicio.ObjValorPres_ModuloServicioDec.ObjValorPro
        Dim lstrKey As String
        For Each lobjSectorModulo As ClsSectorModulo In lcolSectores_ModuloCont
            lblnNuevo = False
            Dim lstrIdSector = lobjSectorModulo.ObjIdSector_SectorModuloShr.ToString
            Dim ldblParticipacionSector As Double =
                    lobjModuloContribucion.FdblTasaParticipacionSector(lstrIdSector,
                            lenuTipoBaseCalcSer)
            ldblTotalPar += ldblParticipacionSector
            If ldblParticipacionSector > 0 Then
                Dim lblnModificoPermiso = False
                Dim ldecContribucionInicialSector As Decimal =
                        Math.Round(ldecContribucionModulo * ldblParticipacionSector, 2)
                lstrKey = lobjSectorModulo.ObjIdSector_SectorModuloShr.ToString
                Dim lobjSectorModuloServicio As ClsSectorModuloServicio = Nothing
                If lcolSectores_ModuloServicio.Contains(lstrKey) Then
                    lobjSectorModuloServicio = lcolSectores_ModuloServicio(lstrKey)
                    If Not CType(lobjSectorModuloServicio.EnuPermisosObj And EnuPermisosDef.enuModificar, Boolean) Then
                        lobjSectorModuloServicio.EnuPermisosObj += EnuPermisosDef.enuModificar
                        lblnModificoPermiso = True
                    End If
                    lobjSectorModuloServicio.SModifique()
                    If lblnModificoPermiso Then
                        lobjSectorModuloServicio.EnuPermisosObj -= EnuPermisosDef.enuModificar
                    End If
                Else
                    lblnNuevo = True
                    lobjSectorModuloServicio = aobjModuloServicio.ObjNuevoSectorModuloServicio
                    lobjSectorModuloServicio.ObjIdSector_SectorModuloServicioShr.ObjValorPro =
                            lobjSectorModulo.ObjIdSector_SectorModuloShr.ObjValorPro
                End If
                lobjSectorModuloServicio.ObjValor_SectorModuloServicioDec.ObjValorPro +=
                        ldecContribucionInicialSector
                If lblnNuevo Then
                    aobjModuloServicio.SAdicioneSector(lobjSectorModuloServicio)
                End If
            End If
        Next
    End Sub
    Private Shared Function FblnCreoItemsPrograma(aobjServicio As ClsServicio) As Boolean
        Dim lblnCreo = False
        Dim ldtbSecServicio = aobjServicio.FdtbSectoresServicio()
        Dim lstrFiltro As String
        For Each lobjSector As ClsSector In GobjParametros.ColSectores
            Dim lstrIdSector As String = lobjSector.ObjIdSectorShr.ToString
            lstrFiltro = ClsIdSector_SectorModuloServicioShr.SstrNombreCampoBd & " = " &
                    lstrIdSector
            If ldtbSecServicio.Select(lstrFiltro).Length > 0 Then
                lblnCreo = FblnActualizoItemsProgramaSector(aobjServicio,
                    ldtbSecServicio.Select(lstrFiltro)(0))
                If Not lblnCreo Then Exit For
            End If
        Next
        Return lblnCreo
    End Function
    ''' <summary>
    ''' Crea o actualiza los items del programa de facturacion para los predios de un sector que contribuyen con el
    ''' servicio pasado en el argumento
    ''' </summary>
    ''' <param name="aobjServicio">Servicio al cual contribuyen los predios del sector</param>
    ''' <param name="adrwSectorServicio">DataRow que contiene la informacion del sector al cual
    ''' pertenecen los predios que contribuyen con el Servicio.</param>
    ''' <remarks></remarks>
    Private Shared Function FblnActualizoItemsProgramaSector(aobjServicio As ClsServicio,
            adrwSectorServicio As DataRow) As Boolean
        Dim lblnActualizo = False
        Dim lenuBaseCalculo As EnuTipoBaseCalculo =
                aobjServicio.ObjTipoBaseCalculoByt.ObjValorPro
        Dim lshrIdSector As Short = ClsPanorama.FobjValorCampo(adrwSectorServicio(
                ClsIdSector_SectorModuloServicioShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
        Dim ldecValorSector As Decimal
        ldecValorSector = ClsPanorama.FobjValorCampo(adrwSectorServicio(
                ClsValor_SectorModuloServicioDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        Dim ldblBaseParticipacionSector As Double
        If ldecValorSector = 0 Then
            Return True
            Exit Function
        End If
        Dim lobjSector As ClsSector = GobjParametros.ColSectores(lshrIdSector.ToString)
        ldblBaseParticipacionSector = lobjSector.DblBaseParticipacionSector(lenuBaseCalculo)
        If ldblBaseParticipacionSector = 0 Then
            Return True
            Exit Function
        End If
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdSector_PredioShr.SstrNombreCampoBd & " = " & lshrIdSector.ToString
        Dim ldrwPrediosSector() = ClsPanorama.FdrwDataRow(ClsPredio.SstrNombreTabla,
         {"*"}, {{ClsIdPredioStr.SstrNombreCampoBd, "ASC"}}, lstrFiltro)
        Dim lstrIdAno = aobjServicio.ObjIdAno_ServicioShr.ToString
        Dim lstrIdServicio = aobjServicio.ObjIdServicioShr.ToString
        lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & lstrIdAno & " AND " &
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " = " & lstrIdServicio
        Dim ldtbItemsProgramaFact = ClsPanorama.FdtbDataTable(ClsItemProgramaFact.SstrNombreTabla,
                {"*"}, Nothing, lstrFiltro)
        ' Cantidad de periodos en los que se distribuirá el valor
        Dim lshrPeriodosDistribuir As Short
        If aobjServicio.BlnEsCuotaAdministracion Then
            lshrPeriodosDistribuir = 12
        Else
            lshrPeriodosDistribuir = aobjServicio.ObjCantPeriodos_ServicioShr.ObjValorPro
        End If
        For Each ldrwPredioSector As DataRow In ldrwPrediosSector
            Dim lobjPredio As New ClsPredio(ldrwPredioSector)
            lobjPredio.SLeaValores(True)
            Dim lstrIdPredio = lobjPredio.ObjIdPredioStr.ToString
            Dim ldrwItemProgramaFact As DataRow
            ldrwItemProgramaFact = ldtbItemsProgramaFact.NewRow
            lblnActualizo = FblnCreoItemProgramaFactPredio(aobjServicio, ldrwItemProgramaFact,
                        ldecValorSector, ldblBaseParticipacionSector, lshrPeriodosDistribuir,
                        lobjPredio)
            If Not lblnActualizo Then Exit For
        Next
        Return lblnActualizo
    End Function
    Private Shared Function FblnCreoItemProgramaFactPredio(aobjServicio As ClsServicio,
            adrwNewItemProgFact As DataRow, adecValorSectorAno As Decimal,
            adecBaseParticipacionSector As Decimal, ashrPeriodosDistribuir As Short,
            aobjPredio As ClsPredio) As Boolean
        Dim lblnCreo = False
        Dim lenuBaseCalculo As EnuTipoBaseCalculo =
                aobjServicio.ObjTipoBaseCalculoByt.ObjValorPro
        Dim ldecBaseParticipacionPredio As Decimal
        If lenuBaseCalculo = EnuTipoBaseCalculo.EnuUnidad Then
            ldecBaseParticipacionPredio = 1
        Else
            ldecBaseParticipacionPredio = aobjPredio.ObjAreaPredioDec.ObjValorPro *
                    aobjPredio.ObjFactorPonderaCPDbl.ObjValorPro
        End If
        ' Valor total a pagar para todos los periodos
        Dim ldecValorTotalPredio As Decimal = adecValorSectorAno *
                (ldecBaseParticipacionPredio / adecBaseParticipacionSector)
        ' Valora pagar por cada periodo
        Dim ldecValorPeriodoPredio = ldecValorTotalPredio / ashrPeriodosDistribuir
        ldecValorPeriodoPredio = ClsOrionCop.FdecValorRedondeado(ldecValorPeriodoPredio)
        Dim lobjItemProgramFact As New ClsItemProgramaFact(aobjPredio,
                EnuTipoDeudorDef.EnuPredio, adrwNewItemProgFact)
        Dim lblnCambioPermiso = False
        With lobjItemProgramFact
            If Not CType(.EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
                .EnuPermisosObj += EnuPermisosDef.enuCrear
                lblnCambioPermiso = True
            End If
            .SCreeObj(Nothing)
            .ObjIdCarpeta_ItemProgramaFactShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_ItemProgramaFactShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdAno_ItemProgramaFactShr.ObjValorPro = aobjServicio.ObjIdAno_ServicioShr.ObjValorPro
            .ObjIdServicio_ItemProgramaFactShr.ObjValorPro = aobjServicio.ObjIdServicioShr.ObjValorPro
            .ObjIdCliente_ItemProgramaFactDbl.ObjValorPro = 0
            .ObjIdPredio_ItemProgramaFactStr.ObjValorPro = aobjPredio.ObjIdPredioStr.ObjValorPro
            .ObjPeriodoIni_ItemProgStr.ObjValorPro = aobjServicio.ObjPeriodoInicioStr.ObjValorPro
            .ObjCantidadPeriodosShr.ObjValorPro = aobjServicio.ObjCantPeriodos_ServicioShr.ObjValorPro
            .ObjOrigen_ItemProgramaFacByt.ObjValorPro = EnuOrigenItemProgramaFactDef.EnuAplicacion
            .ObjValorPeriodo_ItemProgramaFactDec.ObjValorPro = ldecValorPeriodoPredio
            .SActualiceSaldoActual()
            .SNumereObj()
            If .FblnEstanTodosOk Then
                .SActualice(True)
                lblnCreo = True
                If lblnCambioPermiso Then
                    .EnuPermisosObj -= EnuPermisosDef.enuCrear
                End If
            Else
                Dim lstrMens = "Propiedad no valida. " & lobjItemProgramFact.HstrPropiedadNoValida
                Throw New ErrorInesperadoPanLException(lstrMens)
            End If
        End With
        Return lblnCreo
    End Function
    Private Shared Sub SElimineSectoresModuloServicio(aobjServicio As ClsServicio)
        Dim lstbFiltro As New System.Text.StringBuilder
        With lstbFiltro
            .Append(ClsOrionCop.StrFiltroUbicacion).Append(" And ")
            .Append(ClsIdAno_SectorModuloServicioShr.SstrNombreCampoBd).Append(" = ")
            .Append(aobjServicio.ObjIdAno_ServicioShr.ObjValorPro).Append(" And ")
            .Append(ClsIdServicio_SectorModuloServicioShr.SstrNombreCampoBd).Append(" = ")
            .Append(aobjServicio.ObjIdServicioShr.ObjValorPro)
        End With
        Dim lstrFiltro = lstbFiltro.ToString
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlEliminar(
                ClsSectorModuloServicio.SstrNombreTabla, lstrFiltro)
        GobjPanDat.SEjecuteSentenciaSql(lstrSql)
    End Sub
    Private Shared Sub SElimineSectoresServiciosAno(ashrIdAno As Short)
        Dim lstbFiltro As New System.Text.StringBuilder
        With lstbFiltro
            .Append(ClsOrionCop.StrFiltroUbicacion).Append(" AND ")
            .Append(ClsIdAno_SectorModuloServicioShr.SstrNombreCampoBd).Append(" = ")
            .Append(ashrIdAno)
        End With
        Dim lstrFiltro = lstbFiltro.ToString
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlEliminar(
                ClsSectorModuloServicio.SstrNombreTabla, lstrFiltro)
        GobjPanDat.SEjecuteSentenciaSql(lstrSql)
    End Sub
#Region "Calculos desde el año"
#Region "Con base en el coeficiente de propiedad o en la unidad"
    Friend Function FblnCalculoCuotasAdmin(aobjAno As ClsAno,
            ByRef astrMens As String) As Boolean
        Dim lblnCalculo = True
        Dim lenuTipoCalculo As EnuTipoBaseCalculo = aobjAno.ObjTipoCalculoCuotaByt.ObjValorPro
        If lenuTipoCalculo = EnuTipoBaseCalculo.EnuCoeficientePro OrElse lenuTipoCalculo =
                EnuTipoBaseCalculo.EnuUnidad Then
            Dim ldecPresupAplica = aobjAno.ObjValorPres_AnoDec.ObjValorPro
            SElimineItemsProgFact(aobjAno)
            SElimineSectoresServiciosAno(aobjAno.ObjIdAnoShr.ObjValorPro)
            If aobjAno.FblnCalcularCuotaAdminPorCP() AndAlso lenuTipoCalculo <>
                    EnuTipoBaseCalculo.EnuUnidad AndAlso
                    aobjAno.ObjModuloPorServicioBln.ObjValorPro Then
                SGenereCuotasAdminxCP(ldecPresupAplica, aobjAno.ObjIdAnoShr.ObjValorPro)
                SActualiceContribucionModulos(aobjAno)
            Else
                If aobjAno.ObjModuloPorServicioBln.ObjValorPro Then
                    ' En caso contrario el valor de cada modulo del servicio debe ser ingresado por IU
                    SCalculeModulosServicios(aobjAno, ldecPresupAplica)
                End If
                For Each lobjservicio As ClsServicio In aobjAno.ColServiciosAno
                    SCalculeContribucionSectores(lobjservicio)
                    If Not lobjservicio.ObjEsAjusteBln.ObjValorPro Then
                        lblnCalculo = FblnCalculoItemsServicio(lobjservicio)
                        If Not lblnCalculo Then Exit For
                        GobjPanorama.SRegistreAccionLogApp("clsCalculosServicios",
                                "Calculo Cuotas de Administración Servicio No. " &
                                lobjservicio.ObjIdServicioShr.ObjValorPro & "/" &
                                lobjservicio.ObjIdAno_ServicioShr.ObjValorPro)
                    End If
                Next
                If lblnCalculo Then
                    If Not aobjAno.ObjModuloPorServicioBln.ObjValorPro Then
                        aobjAno.SAjustePresupuesto()
                    End If
                End If
            End If
        ElseIf lenuTipoCalculo = EnuTipoBaseCalculo.EnuCuotaAnterior Then
            If GobjParametros.ColAnos.Count > 1 Then
                Dim lstrMens = String.Empty
                SElimineItemsProgFact(aobjAno)
                Dim ldblIncremento As Double, ldecPresAnoAnt As Decimal, ldecPresAno As Decimal
                Dim lshrIdAnoAnt = aobjAno.ObjIdAnoShr.ObjValorPro - 1
                Dim lobjAnoAnt As ClsAno = GobjParametros.ColAnos(lshrIdAnoAnt.ToString)
                ldecPresAnoAnt = lobjAnoAnt.ObjValorPres_AnoDec.ObjValorPro
                ldecPresAno = aobjAno.ObjValorPres_AnoDec.ObjValorPro
                If ldecPresAnoAnt > 0 Then
                    If ldecPresAno > 0 Then
                        ldblIncremento = Math.Round((ldecPresAno - ldecPresAnoAnt) /
                                ldecPresAnoAnt, 6)
                        SCreeCuotasBaseAnoAnt(lshrIdAnoAnt, ldblIncremento)
                        For Each lobjServicio As ClsServicio In aobjAno.ColServiciosAno
                            SActualiceSectoresACero(lobjServicio)
                        Next
                        SActualiceContribucionSectores(aobjAno, lobjAnoAnt)
                        SActualicePres(aobjAno)
                        lblnCalculo = String.IsNullOrEmpty(lstrMens)
                        If Not lblnCalculo Then
                            astrMens = lstrMens
                        Else
                            astrMens = "Proceso terminado exitosamente!"
                        End If
                    Else
                        astrMens = "Se requiere ingresar el Presupuesto para el presente Año!"
                    End If
                Else
                    astrMens = "Si en el año anterior no hubo presupuesto, solo se pueden calcular " &
                            "las cuotas con base en coeficiente de propiedad, o bien importarlas!"
                End If
            Else
                astrMens = "El tipo de Cálculo no es posible para el primer Año!"
            End If
        End If
        Return lblnCalculo
    End Function
    ''' <summary>
    ''' Calcula y actualiza en la base de datos el valor con que participa cada módulo
    ''' del servicio
    ''' </summary>
    ''' <param name="aobjAno">Objeto año para el cual se hace el calculo</param>
    ''' <param name="adecValorTotal">Valor total a distribuir</param>
    ''' <remarks>Aplica para instalaciones donde a cada módulo de contribución le 
    ''' corresponde un servicio </remarks>
    Private Shared Sub SCalculeModulosServicios(aobjAno As ClsAno,
            adecValorTotal As Decimal)
        For Each lobjServicio As ClsServicio In aobjAno.ColServiciosAno
            If Not lobjServicio.ObjEsAjusteBln.ObjValorPro Then
                lobjServicio.BlnCalculandoDesdeAno = True
                lobjServicio.SActuVlrModulo(adecValorTotal)
                lobjServicio.BlnCalculandoDesdeAno = False
            End If
        Next
    End Sub
    Private Shared Sub SActualicePres(aobjAno As ClsAno)
        Dim lentIdAno As Integer = aobjAno.ObjIdAnoShr.ObjValorPro
        Dim lentIdServicio As Integer
        For Each lobjServicio As ClsServicio In aobjAno.ColServiciosAno
            If Not lobjServicio.ObjEsAjusteBln.ObjValorPro Then
                lentIdServicio = lobjServicio.ObjIdServicioShr.ObjValorPro
                Dim lstrAccion As String = "Calculo Items Pro. Fac. Servicio No. " &
                    lobjServicio.ObjIdServicioShr.ObjValorPro & ", año " &
                    lobjServicio.ObjIdAno_ServicioShr.ObjValorPro
                GobjPanorama.SRegistreAccionLogApp("clsCalculosServicios", lstrAccion)
                lobjServicio.SActuEstaGeneradoProgramaFact(True)
            End If
        Next
    End Sub
    ''' <summary>
    ''' Elimina los itéms del programa de facturación para el año pasado en el argumento
    ''' </summary>
    ''' <param name="aobjAno">Objeto año al cual se le eliminarán los itéms</param>
    Friend Shared Sub SElimineItemsProgFact(aobjAno As ClsAno)
        For Each lobjServicio As ClsServicio In aobjAno.ColServiciosAno
            SElimineItemsProgFact(lobjServicio)
        Next
    End Sub
#Region "Con base en el año anterior"
    ''' <summary>
    ''' Crea Cuotas Admin con base en el presupuesto del año anterior y el nuevo presupuesto
    ''' </summary>
    ''' <param name="aobjAno">Objeto del año actual</param>
    ''' <param name="adblIncremento">Incremento entre el valor del presupuesto del año anterior
    ''' y el valor del presupuesto del año actual</param>
    ''' <remarks></remarks>
    Friend Shared Sub SCreeCuotasBaseAnoAnt(ashrIdAnoAnt As Short, adblIncremento As Double)
        Dim ldtbItemsAnoAnt As DataTable = FdtbItemsProFacAno(ashrIdAnoAnt)
        For Each ldrwItemProFac As DataRow In ldtbItemsAnoAnt.Rows
            SCreeItemProgramaFactPredio(ldrwItemProFac, adblIncremento)
        Next
    End Sub
    ''' <summary>
    ''' Crea los IPFs con base en el valor del año anterior mas el incremento entre el presupuesto
    ''' del año anterior y el presupuesto del año actual
    ''' </summary>
    ''' <param name="adrwItemProgFactAnoAnt"></param>
    ''' <param name="adblIncremento"></param>
    ''' <returns></returns>
    Private Shared Sub SCreeItemProgramaFactPredio(adrwItemProgFactAnoAnt As DataRow,
            adblIncremento As Double)
        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
        Dim lobjVlrLlave As Object()
        Dim lshrIdAno As Short = ClsPanorama.FobjValorCampo(adrwItemProgFactAnoAnt(
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
        Dim lshrIdServicio As Short = ClsPanorama.FobjValorCampo(adrwItemProgFactAnoAnt(
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
        Dim ldecVlrAnoAnt As Decimal = ClsPanorama.FobjValorCampo(adrwItemProgFactAnoAnt(
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
        Dim ldblIdCliente As Double = ClsPanorama.FobjValorCampo(adrwItemProgFactAnoAnt(
                ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd), EnuTipoValor.enuDouble)
        Dim lstrIdpredio As String = ClsPanorama.FobjValorCampo(adrwItemProgFactAnoAnt(
                ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        lobjVlrLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrIdpredio}
        lobjPredio.SAbra(lobjVlrLlave)
        Dim lstrPerIni = (lshrIdAno + 1).ToString & "01"
        Dim lentCantPerPorFac = 12
        Dim ldecValor = ClsOrionCop.FdecValorRedondeado(ldecVlrAnoAnt * (1 + adblIncremento))
        Dim ldecSaldo = ldecValor * lentCantPerPorFac
        Dim ldtbItemsAno As DataTable = FdtbItemsProFacAno(lshrIdAno)
        Dim ldrwNewIPF As DataRow = ldtbItemsAno.NewRow
        Dim lobjItemProgramFact As New ClsItemProgramaFact(lobjPredio, EnuTipoDeudorDef.EnuPredio,
                ldrwNewIPF)
        With lobjItemProgramFact
            If Not CType(.EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
                .EnuPermisosObj += EnuPermisosDef.enuCrear
            End If
            .SCreeObj(Nothing)
            .ObjIdCarpeta_ItemProgramaFactShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_ItemProgramaFactShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdAno_ItemProgramaFactShr.ObjValorPro = lshrIdAno + 1
            .ObjIdServicio_ItemProgramaFactShr.ObjValorPro = lshrIdServicio
            .ObjIdCliente_ItemProgramaFactDbl.ObjValorPro = 0
            .ObjIdPredio_ItemProgramaFactStr.ObjValorPro = lstrIdpredio
            .ObjPeriodoIni_ItemProgStr.ObjValorPro = lstrPerIni
            .ObjCantidadPeriodosShr.ObjValorPro = 12
            .ObjOrigen_ItemProgramaFacByt.ObjValorPro = EnuOrigenItemProgramaFactDef.EnuAplicacion
            .ObjValorPeriodo_ItemProgramaFactDec.ObjValorPro = ldecValor
            .ObjSaldo_ItemProgramaFactDec.ObjValorPro = ldecSaldo
            .SNumereObj()
            .SActualice(True)
        End With
    End Sub
#End Region
#End Region
#Region "Calculo cuota admin por CP"
    Private Sub SGenereCuotasAdminxCP(adecValorPres As Decimal, ashrIdAno As Short)
        Dim ldblCP As Double, ldecCuotaAdmAno As Decimal, ldecCuotaAdmMes As Decimal
        Dim lshrIdServicio As Short
        Dim lstrPerIni = (ashrIdAno).ToString & "01"
        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuNavegable)
        Dim lobjItemProgramFact As ClsItemProgramaFact
        Dim ldtbItemsAno As DataTable = FdtbItemsProFacAno(ashrIdAno)
        lobjPredio.SVayaAlPrimero()
        Do While lobjPredio.BlnExiste
            ldblCP = lobjPredio.ObjCoeficientePropiedadDec.ObjValorPro
            ldecCuotaAdmAno = adecValorPres * ldblCP
            lshrIdServicio = lobjPredio.FshrIdSerAdminContribuye
            ldecCuotaAdmMes = ClsOrionCop.FdecValorRedondeado(ldecCuotaAdmAno / 12)
            If ldecCuotaAdmMes > 0 Then
                Dim ldrwNewIPF As DataRow = ldtbItemsAno.NewRow
                lobjItemProgramFact = New ClsItemProgramaFact(lobjPredio, EnuTipoDeudorDef.EnuPredio,
                        ldrwNewIPF)
                With lobjItemProgramFact
                    If Not CType(.EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
                        .EnuPermisosObj += EnuPermisosDef.enuCrear
                    End If
                    .SCreeObj(Nothing)
                    .ObjIdCarpeta_ItemProgramaFactShr.ObjValorPro = GshrIdCarpeta
                    .ObjIdCentroUtil_ItemProgramaFactShr.ObjValorPro = GshrIdCentroUtil
                    .ObjIdAno_ItemProgramaFactShr.ObjValorPro = ashrIdAno
                    .ObjIdServicio_ItemProgramaFactShr.ObjValorPro = lshrIdServicio
                    .ObjIdCliente_ItemProgramaFactDbl.ObjValorPro = 0
                    .ObjIdPredio_ItemProgramaFactStr.ObjValorPro = lobjPredio.ObjIdPredioStr.ObjValorPro
                    .ObjPeriodoIni_ItemProgStr.ObjValorPro = lstrPerIni
                    .ObjCantidadPeriodosShr.ObjValorPro = 12
                    .ObjOrigen_ItemProgramaFacByt.ObjValorPro =
                            EnuOrigenItemProgramaFactDef.EnuAplicacion
                    .ObjValorPeriodo_ItemProgramaFactDec.ObjValorPro = ldecCuotaAdmMes
                    .SActualice(True)
                End With
            End If
            lobjPredio.SVayaAlSiguiente()
        Loop
    End Sub
    Private Sub SActualiceContribucionModulos(aobjAno As ClsAno)
        Dim ldecTotalContModulos = 0D, ldecVlrContrModulo As Decimal
        For Each lobjSerAno As ClsServicio In aobjAno.ColServiciosAno
            lobjSerAno.BlnCalculandoDesdeAno = True
            If lobjSerAno.ColModulosServicio.Count > 1 Then
                Throw New ErrorInesperadoPanLException("Al servicio contribuye más de un módulo")
            End If
            If Not lobjSerAno.ObjEsAjusteBln.ObjValorPro Then
                Dim lobjModSer As ClsModuloServicio = lobjSerAno.ColModulosServicio(1)
                SCalculeContribucionSectores(lobjModSer, aobjAno.ObjValorPres_AnoDec.ObjValorPro,
                        ldecVlrContrModulo)
                lobjModSer.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                lobjModSer.ObjValorPres_ModuloServicioDec.ObjValorPro = ldecVlrContrModulo
                lobjModSer.SActualice(True)
                ldecTotalContModulos += ldecVlrContrModulo
                ldecVlrContrModulo = 0
            End If
            lobjSerAno.BlnCalculandoDesdeAno = False
        Next
        If aobjAno.ObjValorPres_AnoDec.ObjValorPro <> ldecTotalContModulos AndAlso
                aobjAno.ObjValorPres_AnoDec.ObjValorPro <> Math.Round(ldecTotalContModulos, 0) Then
            Throw New ErrorInesperadoPanLException("Presupuesto del año descuadrado")
        End If
        aobjAno.SRefresqueObj()
    End Sub
    Private Sub SCalculeContribucionSectores(aobjModSer As ClsModuloServicio,
            adecValorPresAno As Decimal, ByRef adecTotalContMod As Decimal)
        Dim lshrIdSector As Short, ldblCoefPropSector As Double, ldecVlrContSectMod As Decimal
        Dim lobjSectorModSer As ClsSectorModuloServicio
        Dim ldtbContSecs = FdtbContribucionSectores()
        For Each ldrwContSec As DataRow In ldtbContSecs.Rows
            lshrIdSector = ClsPanorama.FobjValorCampo(ldrwContSec(
                    ClsIdSector_PredioShr.SstrNombreCampoBd), EnuTipoValor.enuString)
            If aobjModSer.ObjMiModuloContribucion.ColSectoresModulo.Contains(
                    lshrIdSector.ToString) Then
                ldblCoefPropSector = ClsPanorama.FobjValorCampo(ldrwContSec("TotCoeSec"),
                        EnuTipoValor.enuDecimal)
                If ldblCoefPropSector > 0 Then
                    ldecVlrContSectMod = adecValorPresAno * ldblCoefPropSector
                    lobjSectorModSer = aobjModSer.ObjNuevoSectorModuloServicio()
                    lobjSectorModSer.ObjIdSector_SectorModuloServicioShr.ObjValorPro = lshrIdSector
                    lobjSectorModSer.ObjValor_SectorModuloServicioDec.ObjValorPro = ldecVlrContSectMod
                    lobjSectorModSer.SActualice(True)
                    adecTotalContMod += ldecVlrContSectMod
                End If
            End If
        Next
        adecTotalContMod = Math.Round(adecTotalContMod, 2)
    End Sub
    Private Function FdtbContribucionSectores() As DataTable
        Dim lstrTabla = ClsPredio.SstrNombreTabla
        Dim lstrCampSele As String() = {ClsIdSector_PredioShr.SstrNombreCampoBd,
                "SUM(" & ClsCoeficientePropiedadDec.SstrNombreCampoBd & ") AS TotCoeSec"}
        Dim lstrFIltro = ClsOrionCop.StrFiltroUbicacion
        Dim lstrOrden As String(,) = {{ClsIdSector_PredioShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrGrupo As String() = {ClsIdSector_PredioShr.SstrNombreCampoBd}
        Dim ldtbContSec = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSele, lstrOrden, lstrFIltro,
                True, lstrGrupo)
        Return ldtbContSec
    End Function
#End Region
#Region "Calculos desde la Cuota anterior"
#Region "Crear Servicios Cuota Administracion Nuevo Año basado en Cuota Admin Año Actual"
    Private Shared Function FdtbItemsProFacAno(ashrIdAno As Short) As DataTable
        Dim lstrIndice(,) As String = {{ClsIdPredioStr.SstrNombreCampoBd, "ASC"},
                {ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd, "DESC"},
                {ClsIdServicioShr.SstrNombreCampoBd, "ASC"}}
        Dim lshrIdSerFin = 0S
        Dim lshrIdSerIni = FshrIdServicioIni(ashrIdAno, lshrIdSerFin)
        Dim lstrFIltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdAnoShr.SstrNombreCampoBd & " = " & ashrIdAno & " AND " &
                ClsIdServicioShr.SstrNombreCampoBd & " BETWEEN " &
                lshrIdSerIni & " AND " & lshrIdSerFin
        Dim ldtbItemsProFacAno = ClsPanorama.FdtbDataTable(
                ClsItemProgramaFact.SstrNombreTabla, {"*"}, lstrIndice, lstrFIltro)
        Return ldtbItemsProFacAno
    End Function
    Private Shared Function FshrIdServicioIni(ashrIdAno As Short,
                ByRef ashrIdServicioFin As Short) As Short
        Dim lshrIdSerIni = 0S, lshrIdSerFin = 0
        Dim lobjAno As ClsAno = GobjParametros.ColAnos(ashrIdAno.ToString)
        For Each lobjServico As ClsServicio In lobjAno.ColServiciosAno
            If Not lobjServico.ObjEsAjusteBln.ObjValorPro Then
                If lshrIdSerIni = 0 Then
                    lshrIdSerIni = lobjServico.ObjIdServicioShr.ObjValorPro
                ElseIf lobjServico.ObjIdServicioShr.ObjValorPro < lshrIdSerIni Then
                    lshrIdSerIni = lobjServico.ObjIdServicioShr.ObjValorPro
                End If
                If lshrIdSerFin = 0 Then
                    lshrIdSerFin = lobjServico.ObjIdServicioShr.ObjValorPro
                ElseIf lobjServico.ObjIdServicioShr.ObjValorPro > lshrIdSerFin Then
                    lshrIdSerFin = lobjServico.ObjIdServicioShr.ObjValorPro
                End If
            End If
        Next
        ashrIdServicioFin = lshrIdSerFin
        Return lshrIdSerIni
    End Function
    Private Sub SActualiceContribucionSectores(aobjAnoActual As ClsAno, aobjAnoAnt As ClsAno)
        Dim ldblIncr As Double = aobjAnoActual.ObjValorPres_AnoDec.ObjValorPro /
                aobjAnoAnt.ObjValorPres_AnoDec.ObjValorPro
        For Each lobjSerAnt As ClsServicio In aobjAnoAnt.ColServiciosAno
            SActualiceSectoresDelServicio(aobjAnoActual, lobjSerAnt, ldblIncr)
        Next
    End Sub
    Private Sub SActualiceSectoresDelServicio(aobjAnoActual As ClsAno,
            aobjServcicio As ClsServicio, adblInc As Double)
        For Each lobjModSer As ClsModuloServicio In aobjServcicio.ColModulosServicio
            SActuSectoresDelModuloDeServicio(aobjAnoActual, lobjModSer, adblInc)
        Next
    End Sub
    Private Sub SActuSectoresDelModuloDeServicio(aobjAnoActual As ClsAno,
             aobjModuloSer As ClsModuloServicio, adblInc As Double)
        Dim ldecVlrSecAnt As Decimal, ldecVlrSerAct As Decimal
        For Each lobjSectorMS As ClsSectorModuloServicio In aobjModuloSer.ColSectores_ModuloServicio
            ldecVlrSecAnt = lobjSectorMS.ObjValor_SectorModuloServicioDec.ObjValorPro
            ldecVlrSerAct = ldecVlrSecAnt * adblInc
            SActualiceSecModSerAct(aobjAnoActual, lobjSectorMS, ldecVlrSerAct)
        Next
    End Sub
    Private Sub SActualiceSecModSerAct(aobjAnoActual As ClsAno,
            aobjSecModSer As ClsSectorModuloServicio, adecValrSMS As Decimal)
        Dim lshrIdAnoAct As Short = aobjAnoActual.ObjIdAnoShr.ObjValorPro
        Dim lshrIdSer As Short = aobjSecModSer.ObjIdServicio_SectorModuloServicioShr.ObjValorPro
        Dim lstrKey = lshrIdAnoAct.ToString() & "," & lshrIdSer.ToString()
        Dim LobjSer As ClsServicio = aobjAnoActual.ColServiciosAno(lstrKey)
        Dim lshrIdMod As Short = aobjSecModSer.ObjIdModulo_SectorModuloServicioShr.ObjValorPro
        Dim lobjModSer As ClsModuloServicio = LobjSer.ColModulosServicio(lshrIdMod.ToString())
        Dim lshrIdSecModSer As Short = aobjSecModSer.ObjIdSector_SectorModuloServicioShr.ObjValorPro
        If lobjModSer.ColSectores_ModuloServicio.Contains(lshrIdSecModSer.ToString().Trim) Then
            Dim lobjSecModSerAc As ClsSectorModuloServicio =
                lobjModSer.ColSectores_ModuloServicio(lshrIdSecModSer.ToString().Trim)
            lobjSecModSerAc.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
            lobjSecModSerAc.ObjValor_SectorModuloServicioDec.ObjValorPro = adecValrSMS
            lobjSecModSerAc.SActualice(True)
        End If
    End Sub
    Private Function FdecValorModSerAnoAnt(aobjAnoAnt As ClsAno, ashrIdServicio As Short,
            ashrIdModuloSer As Short) As Decimal
        Dim lstrKey = aobjAnoAnt.ObjIdAnoShr.ToString() & "," & ashrIdServicio.ToString()
        Dim lobjServicio As ClsServicio = aobjAnoAnt.ColServiciosAno(lstrKey)
        Dim lobjModSer As ClsModuloServicio = lobjServicio.ColModulosServicio(
                ashrIdModuloSer.ToString())
        Return lobjModSer.ObjValorPres_ModuloServicioDec.ObjValorPro
    End Function
#End Region
#End Region
#Region "Calculo a partir de importacion de los IPF para Cuota de Administración"
    Friend Shared Sub SActualicePresAnoImportado(aobjAno As ClsAno)
        Dim lshrIdAno As Short = aobjAno.ObjIdAnoShr.ObjValorPro
        Dim lstrTabla = ClsItemProgramaFact.SstrNombreTabla
        Dim lstrCampSelec = {"SUM(" & ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd &
                ") AS Total"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & lshrIdAno
        Dim lstrFiltroServicio = String.Empty
        For Each lobjServicio As ClsServicio In aobjAno.ColServiciosAno
            If lobjServicio.ObjEsAjusteBln.ObjValorPro Then
                lstrFiltroServicio = " AND " & ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd &
                        " <> " & lobjServicio.ObjIdServicioShr.ObjValorPro
            End If
        Next
        lstrFiltro &= lstrFiltroServicio
        Dim ldtbPres = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSelec, {{"", ""}},
                lstrFiltro)
        Dim ldecVlrItems = ClsPanorama.FobjValorCampo(ldtbPres.Rows(0)("Total"),
                EnuTipoValor.enuDecimal)
        Dim ldecVlrPresAno = ldecVlrItems * 12
        aobjAno.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        aobjAno.ObjValorPres_AnoDec.ObjValorPro = ldecVlrPresAno
        aobjAno.ObjTipoCalculoCuotaByt.ObjValorPro = EnuTipoBaseCalculo.EnuImportadas
        aobjAno.SActualice(True)
        If aobjAno.ObjModuloPorServicioBln.ObjValorPro Then
            SActualicePresModulosAno(lshrIdAno)
            SElimineSectoresServiciosAno(lshrIdAno)
            SActualiceSectores(lshrIdAno, 0)
        Else
            For Each lobjServicio As ClsServicio In aobjAno.ColServiciosAno
                SActualiceModulosACero(lobjServicio)
                SActualiceSectoresACero(lobjServicio)
            Next
        End If
    End Sub
    ''' <summary>
    ''' Actualiza los modulos de los servicios del año cuando el los servicios son anuales
    ''' </summary>
    Friend Shared Sub SActualicePresModulosAno(ashrIdAno As Short)
        Dim lobjAno As ClsAno = GobjParametros.ColAnos(ashrIdAno.ToString)
        Dim lstrTabla_1 = ClsItemProgramaFact.SstrNombreTabla & " AS I "
        Dim lstrTabla_2 = ClsPredio.SstrNombreTabla & " AS P "
        Dim lstrTabla_3 = ClsSectorModulo.SstrNombreTabla & " AS S "
        Dim lstrCamSel As String = ClsIdModulo_SectorModuloShr.SstrNombreCampoBd & ", " &
                "SUM(" & ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd & " * " &
                ClsCantidadPeriodosShr.SstrNombreCampoBd & ") AS TOT "
        Dim lstrON_1 = "ON I." & StrCampoCarpeta & " = P." &
                StrCampoCarpeta & " AND I." & StrCampoCentroUtil &
                " = P." & StrCampoCentroUtil & " AND I." &
                ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd & " = P." &
                ClsIdPredioStr.SstrNombreCampoBd
        Dim lstrON_2 = "ON P." & StrCampoCarpeta & " = S." &
                StrCampoCarpeta & " AND P." & StrCampoCentroUtil &
                " = P." & StrCampoCentroUtil & " AND P." &
                ClsIdSector_PredioShr.SstrNombreCampoBd & " = S." &
                ClsIdSector_SectorModuloShr.SstrNombreCampoBd
        Dim lstrFilSer = String.Empty
        Dim lshrIdCar As Short = lobjAno.ObjIdCarpetaAnoShr.ObjValorPro
        Dim lshrIdCenUti As Short = lobjAno.ObjIdCentroUtilAnoShr.ObjValorPro
        For Each lobjServicio As ClsServicio In lobjAno.ColServiciosAno
            If lobjServicio.ObjEsAjusteBln.ObjValorPro Then
                lstrFilSer &= " AND " & ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd &
                        " <> " & lobjServicio.ObjIdServicioShr.ToString
            End If
        Next
        Dim lstrFiltro = "I." & StrCampoCarpeta & " = " & lshrIdCar & " AND I." &
                StrCampoCentroUtil & " = " & lshrIdCenUti & " AND I." &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & ashrIdAno
        Dim lstrGrupo = " GROUP BY " & ClsIdModulo_SectorModuloShr.SstrNombreCampoBd
        Dim lstrOrden = " ORDER BY " & ClsIdModulo_SectorModuloShr.SstrNombreCampoBd
        If Not String.IsNullOrEmpty(lstrFilSer) Then
            lstrFiltro &= lstrFilSer
        End If
        Dim lstrExp = "SELECT " & lstrCamSel & " FROM " & lstrTabla_1 & " INNER JOIN " &
                    lstrTabla_2 & lstrON_1 & " INNER JOIN " & lstrTabla_3 & lstrON_2 & " WHERE " &
                    lstrFiltro & lstrGrupo & lstrOrden
        Dim ldtbRes As DataTable = ClsPanorama.FdtbDataTable(lstrExp)
        Dim lshrIdMod As Short, ldecVlr As Decimal, lobjModSer As ClsModuloServicio
        For Each ldrwRes As DataRow In ldtbRes.Rows
            lshrIdMod = ClsPanorama.FobjValorCampo(ldrwRes(
                        ClsIdModulo_SectorModuloShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
            ldecVlr = ClsPanorama.FobjValorCampo(ldrwRes("TOT"), EnuTipoValor.enuDecimal)
            For Each lobjServicio As ClsServicio In lobjAno.ColServiciosAno
                If lobjServicio.ColModulosServicio.Contains(lshrIdMod.ToString) Then
                    lobjModSer = lobjServicio.ColModulosServicio(lshrIdMod.ToString)
                    lobjModSer.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                    lobjModSer.ObjValorPres_ModuloServicioDec.ObjValorPro = ldecVlr
                    lobjModSer.SActualice(True)
                    Exit For
                End If
            Next
        Next
    End Sub
    Private Shared Sub SCreeSectores(ashrIdCarpeta As Short, ashrIdCentroUtil As Short,
                ashrIdAno As Short, adtbSectores As DataTable)
        Dim lstrTabla = ClsSectorModuloServicio.SstrNombreTabla
        Dim lstrExp As String, lshrIdModCont As Short, lshrIdSec As Short,
                lshrIdSer As Short, ldecValor As Decimal
        For Each ldrwSect As DataRow In adtbSectores.Rows
            lshrIdModCont = ClsPanorama.FobjValorCampo(ldrwSect(
                    ClsIdModulo_SectorModuloShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
            lshrIdSec = ClsPanorama.FobjValorCampo(ldrwSect(
                    ClsIdSector_PredioShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
            lshrIdSer = ClsPanorama.FobjValorCampo(ldrwSect(
                    ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
            ldecValor = ClsPanorama.FobjValorCampo(ldrwSect("TOT"),
                    EnuTipoValor.enuDecimal)
            lstrExp = "INSERT INTO " & lstrTabla & " VALUES (" & ashrIdAno & ", " &
                    ashrIdCarpeta & ", " & ashrIdCentroUtil & ", " & lshrIdModCont & ", " &
                    lshrIdSec & ", " & lshrIdSer & ", " & ldecValor & ")"
            Dim lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExp)
        Next
    End Sub
#End Region
#Region "Calculo a partir de importacion de los IPF para servicios permanentes"
    ''' <summary>
    ''' Actualiza los modulos del servicio cuando este es permanente
    ''' </summary>
    Friend Shared Sub SActualicePresModulosServicio(aobjServicio As ClsServicio)
        If aobjServicio.ColModulosServicio.Count = 1 Then
            Dim lstrTabla_1 = ClsItemProgramaFact.SstrNombreTabla & " AS I "
            Dim lstrTabla_2 = ClsPredio.SstrNombreTabla & " AS P "
            Dim lstrTabla_3 = ClsSectorModulo.SstrNombreTabla & " AS S "
            Dim lstrCamSel As String = ClsIdModulo_SectorModuloShr.SstrNombreCampoBd & ", " &
                "SUM(" & ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd & " * " &
                ClsCantidadPeriodosShr.SstrNombreCampoBd & ") AS TOT "
            Dim lstrON_1 = "ON I." & StrCampoCarpeta & " = P." &
                StrCampoCarpeta & " AND I." & StrCampoCentroUtil &
                " = P." & StrCampoCentroUtil & " AND I." &
                ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd & " = P." &
                ClsIdPredioStr.SstrNombreCampoBd
            Dim lstrON_2 = "ON P." & StrCampoCarpeta & " = S." &
                StrCampoCarpeta & " AND P." & StrCampoCentroUtil &
                " = P." & StrCampoCentroUtil & " AND P." &
                ClsIdSector_PredioShr.SstrNombreCampoBd & " = S." &
                ClsIdSector_SectorModuloShr.SstrNombreCampoBd
            Dim lshrIdCar As Short = aobjServicio.ObjIdCarpeta_ServicioShr.ObjValorPro
            Dim lshrIdCenUti As Short = aobjServicio.ObjIdCentroUtil_ServicioShr.ObjValorPro
            Dim lstrFiltro = "I." & StrCampoCarpeta & " = " & lshrIdCar & " AND I." &
                StrCampoCentroUtil & " = " & lshrIdCenUti & " I." &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = 0"
            Dim lstrGrupo = " GROUP BY " & ClsIdModulo_SectorModuloShr.SstrNombreCampoBd
            Dim lstrOrden = " ORDER BY " & ClsIdModulo_SectorModuloShr.SstrNombreCampoBd
            Dim lstrExp = "SELECT " & lstrCamSel & " FROM " & lstrTabla_1 & " INNER JOIN " &
                lstrTabla_2 & lstrON_1 & " INNER JOIN " & lstrTabla_3 & lstrON_2 & " WHERE " &
                lstrFiltro & lstrGrupo & lstrOrden
            Dim ldtbRes As DataTable = ClsPanorama.FdtbDataTable(lstrExp)
            Dim lshrIdMod As Short, ldecVlr As Decimal, lobjModSer As ClsModuloServicio
            For Each ldrwRes As DataRow In ldtbRes.Rows
                lshrIdMod = ClsPanorama.FobjValorCampo(ldrwRes(
                        ClsIdModulo_SectorModuloShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
                ldecVlr = ClsPanorama.FobjValorCampo(ldrwRes("TOT"), EnuTipoValor.enuDecimal)
                If aobjServicio.ColModulosServicio.Contains(lshrIdMod.ToString) Then
                    lobjModSer = aobjServicio.ColModulosServicio(lshrIdMod.ToString)
                    lobjModSer.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                    lobjModSer.ObjValorPres_ModuloServicioDec.ObjValorPro = ldecVlr
                    lobjModSer.SActualice(True)
                End If
            Next
            SActualiceSectores(0, aobjServicio.ObjIdServicioShr.ObjValorPro)
        Else
            SActualiceModulosACero(aobjServicio)
            SActualiceSectoresACero(aobjServicio)
        End If
    End Sub
    ''' <summary>
    ''' Cuando a un servicio le contribuye mas de un módulo y el servicio es importado o 
    ''' calculado con base al año anterior, el valor de los módulos se establece en cero, 
    ''' lo mismo que el valor con que contribuyen los sectores
    ''' </summary>
    ''' <param name="aobjServicio"></param>
    Private Shared Sub SActualiceModulosACero(aobjServicio As ClsServicio)
        Dim lstrTabla = ClsModuloServicio.SstrNombreTabla
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdAno_ModuloServicioShr.SstrNombreCampoBd & " = " &
                aobjServicio.ObjIdAno_ServicioShr.ObjValorPro & " AND " &
                ClsIdServicio_ModuloServicioShr.SstrNombreCampoBd & " = " &
                aobjServicio.ObjIdServicioShr.ObjValorPro
        Dim lstrExpSql = "UPDATE " & lstrTabla & " SET " &
                ClsValorPres_ModuloServicioDec.SstrNombreCampoBd & " = 0.00 WHERE " &
                lstrFiltro
        Dim lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
    End Sub
    Friend Shared Sub SActualiceSectores(ashrIdAno As Short, ashrIdServicio As Short)
        Dim lobjAno As ClsAno = Nothing, lobjServicio As ClsServicio
        Dim lshrIdCar As Short, lshrIdCenUtil As Short
        If ashrIdAno = 0 Then
            lobjServicio = GobjParametros.FobjServicio("0," & ashrIdServicio)
            SElimineSectoresModuloServicio(lobjServicio)
            lshrIdCar = lobjServicio.ObjIdCarpeta_ServicioShr.ObjValorPro
            lshrIdCenUtil = lobjServicio.ObjIdCentroUtil_ServicioShr.ObjValorPro
        Else
            lobjAno = GobjParametros.ColAnos(ashrIdAno.ToString)
            lshrIdCar = lobjAno.ObjIdCarpetaAnoShr.ObjValorPro
            lshrIdCenUtil = lobjAno.ObjIdCentroUtilAnoShr.ObjValorPro
        End If
        Dim lstrTabla_1 = ClsItemProgramaFact.SstrNombreTabla & " AS I"
        Dim lstrTabla_2 = ClsPredio.SstrNombreTabla & " AS P"
        Dim lstrTabla_3 = ClsSectorModulo.SstrNombreTabla & " AS S"
        Dim lstrCampSel = "I." & ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & ", " &
                ClsIdModulo_SectorModuloShr.SstrNombreCampoBd & ", P." &
                ClsIdSector_PredioShr.SstrNombreCampoBd & ", SUM(" &
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd & " * " &
                ClsCantidadPeriodosShr.SstrNombreCampoBd & ") AS TOT "
        Dim Lstr_On_1 = " ON I." & StrCampoCarpeta & " = P." &
                StrCampoCarpeta & " AND I." & StrCampoCentroUtil &
                " = P." & StrCampoCentroUtil & " AND I." &
                ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd & " = P." &
                ClsIdPredioStr.SstrNombreCampoBd
        Dim Lstr_On_2 = " ON I." & StrCampoCarpeta & " = S." &
                StrCampoCarpeta & " AND I." & StrCampoCentroUtil &
                " = S." & StrCampoCentroUtil & " AND P." &
                ClsIdSector_PredioShr.SstrNombreCampoBd & " = S." &
                ClsIdSector_SectorModuloShr.SstrNombreCampoBd
        Dim lstrFiltro = " WHERE P." & StrCampoCarpeta & " = " & lshrIdCar &
                " AND P." & StrCampoCentroUtil & " = " & lshrIdCenUtil
        If ashrIdAno = 0 Then
            lstrFiltro &= " AND " & ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = 0 " &
                    " AND " & ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " = " &
                    ashrIdServicio
        Else
            Dim lstrFilSer = String.Empty
            lstrFiltro &= " AND " & ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " &
                    ashrIdAno
            For Each lobjSer As ClsServicio In lobjAno.ColServiciosAno
                If lobjSer.ObjEsAjusteBln.ObjValorPro Then
                    SElimineSectoresModuloServicio(lobjSer)
                    lstrFilSer &= " AND " & ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd &
                        " <> " & lobjSer.ObjIdServicioShr.ToString
                End If
            Next
            lstrFiltro &= lstrFilSer
        End If
        Dim lstrFin = " GROUP BY " & ClsIdSector_PredioShr.SstrNombreCampoBd & " ORDER BY " &
                ClsIdSector_PredioShr.SstrNombreCampoBd
        Dim lstrExp = "SELECT " & lstrCampSel & " FROM " & lstrTabla_1 & " INNER JOIN " &
                lstrTabla_2 & Lstr_On_1 & " INNER JOIN " & lstrTabla_3 & Lstr_On_2 &
                lstrFiltro & lstrFin
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrExp)
        ' Crear los sectores_modulo_servicio
        SCreeSectores(lshrIdCar, lshrIdCenUtil, ashrIdAno, ldtbRes)
    End Sub
    Private Shared Sub SActualiceSectoresACero(aobjServicio As ClsServicio)
        Dim lstrTabla = ClsSectorModuloServicio.SstrNombreTabla
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdAno_SectorModuloServicioShr.SstrNombreCampoBd & " = " &
                aobjServicio.ObjIdAno_ServicioShr.ObjValorPro
        Dim lstrExpSql = "UPDATE " & lstrTabla & " SET " &
                ClsValor_SectorModuloServicioDec.SstrNombreCampoBd & " = 0.00 WHERE " &
                lstrFiltro
        Dim lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
    End Sub
#End Region
#End Region
#Region "Métodos generales"
    Friend Shared Function FblnUnSectorContribuyeUnaVez(aobjServicio As ClsServicio) As Boolean
        Dim lblnOk = True, lshrIdSector As Short
        Dim larlSectores As New ArrayList
        For Each lobjModSer As ClsModuloServicio In aobjServicio.ColModulosServicio
            For Each lobjSectModSer As ClsSectorModuloServicio In
                    lobjModSer.ColSectores_ModuloServicio
                lshrIdSector = lobjSectModSer.ObjIdSector_SectorModuloServicioShr.ObjValorPro
                If Not larlSectores.Contains(lshrIdSector) Then
                    larlSectores.Add(lshrIdSector)
                Else
                    lblnOk = False
                    Exit For
                End If
                If Not lblnOk Then Exit For
            Next
        Next
        Return lblnOk
    End Function
    Friend Shared Sub SElimineItemsProgFact(aobjServicio As ClsServicio)
        Dim lstbFiltro As New Text.StringBuilder
        Dim lshrIdAno As Short = aobjServicio.ObjIdAno_ServicioShr.ObjValorPro
        Dim lshrIdServ As Short = aobjServicio.ObjIdServicioShr.ObjValorPro
        With lstbFiltro
            .Append(ClsOrionCop.StrFiltroUbicacion).Append(" AND ")
            .Append(ClsIdAno_ServicioShr.SstrNombreCampoBd).Append(" = ")
            .Append(lshrIdAno).Append(" AND ").Append(ClsIdServicioShr.SstrNombreCampoBd)
            .Append(" = ").Append(lshrIdServ)
        End With
        Dim lstrFiltro = lstbFiltro.ToString
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlEliminar(ClsItemProgramaFact.SstrNombreTabla,
                lstrFiltro)
        GobjPanDat.SEjecuteSentenciaSql(lstrSql)
    End Sub
#End Region
End Class