Friend Class ClsModuloServicio
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriModulosServicio"
    ' Variables de modulo
    Private MobjMiServicio As ClsServicio = Nothing
    Private McolSectores_ModuloServicio As Collection = Nothing
    Private MdtbSectores_ModuloServicio As DataTable = Nothing
#End Region
#Region "Constructores"
    Friend Sub New(aobjPadre As ClsServicio, aenuModoInstanciaObj As EnuModoInstanciaObjDef)
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = aobjPadre
        MobjMiServicio = aobjPadre
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
            lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                                ClsIdAno_ModuloServicioShr.SstrNombreCampoBd,
                                ClsIdServicio_ModuloServicioShr.SstrNombreCampoBd,
                                ClsIdModulo_ModuloServicioShr.SstrNombreCampoBd}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdAno_ModuloServicioShr.SstrNombreCampoBd &
                    " = " & MobjMiServicio.ObjIdAno_ServicioShr.ObjValorPro & " AND " &
                    ClsIdServicio_ModuloServicioShr.SstrNombreCampoBd & " = " &
                    MobjMiServicio.ObjIdServicioShr.ObjValorPro
            HcolFiltros.Add(lstrFiltro)
        Else
            HblnEsCreable = False
            HblnEsModificable = False
            HblnEsSuprimible = False
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        HblnEsAnulable = False
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsServicio, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        MobjMiServicio = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwObjeto
        DtbTablaColeccion = DrwRegistroActual.Table
    End Sub
#End Region
#Region "Propiedades"
#Region "Propiedades indentificadoras"
    Protected Overrides ReadOnly Property HstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Protected Overrides ReadOnly Property HenuIdClase As EnuIdClasesPanDef
        Get
            Return EnuIdClasesPanDef.enuModuloServicio
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Módulo Servicio"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & ObjIdModulo_ModuloServicioShr.ObjValorPro & " - " &
                    MobjMiServicio.ObjNombreServicioStr.ObjValorPro & Chr(34)
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdAno_ModuloServicioShr As New ClsIdAno_ModuloServicioShr(Me)
    Friend ReadOnly Property ObjIdCarpeta_ModuloServicioShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_ModuloServicioShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdModulo_ModuloServicioShr As New ClsIdModulo_ModuloServicioShr(Me)
    Friend ReadOnly Property ObjIdServicio_ModuloServicioShr As New ClsIdServicio_ModuloServicioShr(Me)
    Friend ReadOnly Property ObjValorPres_ModuloServicioDec As New ClsValorPres_ModuloServicioDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjIdAno_ModuloServicioShr)
                HcolPropiedades.Add(ObjIdCarpeta_ModuloServicioShr)
                HcolPropiedades.Add(ObjIdCentroUtil_ModuloServicioShr)
                HcolPropiedades.Add(ObjIdModulo_ModuloServicioShr)
                HcolPropiedades.Add(ObjIdServicio_ModuloServicioShr)
                HcolPropiedades.Add(ObjValorPres_ModuloServicioDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend ReadOnly Property ObjMiModuloContribucion As ClsModuloContribucion
        Get
            Dim lobjModContr As ClsModuloContribucion = Nothing
            If ObjIdModulo_ModuloServicioShr.BlnEsValido Then
                Dim lcolModulosCon As Collection = GobjParametros.ColModulos
                lobjModContr = lcolModulosCon(ObjIdModulo_ModuloServicioShr.ToString)
            End If
            Return lobjModContr
        End Get
    End Property
    Friend ReadOnly Property ObjMiServicio As ClsServicio
        Get
            If MobjMiServicio Is Nothing Then
                Dim lstrKey = ObjIdAno_ModuloServicioShr.ToString & "," & ObjIdServicio_ModuloServicioShr.ToString
                If ObjIdAno_ModuloServicioShr.ObjValorPro = 0 Then
                    MobjMiServicio = GobjParametros.ColServiciosPer(lstrKey)
                Else
                    If GobjParametros.ColAnos.Contains(ObjIdAno_ModuloServicioShr.ToString) Then
                        Dim lobjAno As ClsAno = GobjParametros.ColAnos(ObjIdAno_ModuloServicioShr.ToString)
                        If lobjAno.ColServiciosAno.Count > 0 AndAlso lobjAno.ColServiciosAno.Contains(lstrKey) Then
                            ObjMiServicio = lobjAno.ColServiciosAno(lstrKey)
                        End If
                    End If
                End If
            End If
            Return MobjMiServicio
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        McolSectores_ModuloServicio = Nothing
        MdtbSectores_ModuloServicio = Nothing
    End Sub
    Protected Overrides Sub SInicialiceObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            ObjIdCarpeta_ModuloServicioShr.ObjValorPro = GshrIdCarpeta
            ObjIdCentroUtil_ModuloServicioShr.ObjValorPro = GshrIdCentroUtil
            ObjIdServicio_ModuloServicioShr.ObjValorPro = MobjMiServicio.ObjIdServicioShr.ObjValorPro
            ObjIdAno_ModuloServicioShr.ObjValorPro = MobjMiServicio.ObjIdAno_ServicioShr.ObjValorPro
            ObjValorPres_ModuloServicioDec.ObjValorPro = 0
        End If
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False, lblnCambioVlr = False, lblnCuotasYaGeneradas = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            If EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                Dim lshrIdAno = 0S
                If ObjMiServicio.BlnEsCuotaAdministracion Then
                    lshrIdAno = ObjMiServicio.ObjMiAno.ObjIdAnoShr.ObjValorPro
                    If Not ObjMiServicio.ObjMiAno.ObjModuloPorServicioBln.ObjValorPro Then
                        lblnCambioVlr = ObjValorPres_ModuloServicioDec.BlnCambio
                        If lblnCambioVlr Then
                            lblnCuotasYaGeneradas = ClsOrionCop.FblnEstanCalcuCuotasAdmin(lshrIdAno)
                        End If
                    End If
                ElseIf ObjMiServicio.ObjGeneraProgramBln.ObjValorPro Then
                    lblnCambioVlr = ObjValorPres_ModuloServicioDec.BlnCambio
                    If lblnCambioVlr Then
                        lblnCuotasYaGeneradas = ClsOrionCop.FblnEstanCalcuCuotasAdmin(lshrIdAno)
                    End If
                End If
                ClsPanorama.SActualiceCol(ColSectores_ModuloServicio)
            End If
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                If lblnCambioVlr AndAlso Not GblnImportando Then
                    MobjMiServicio.SApliqueCambiosServicio(Not lblnCambioVlr,
                            Not lblnCuotasYaGeneradas)
                End If
            End If
            MyBase.SActualice(ablnExigeRequeridos)
            lblnNoHayError = True
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Protected Overrides Function FblnSuprimio() As Boolean
        Dim lblnSuprimio = FblnEsSuprimible()
        If lblnSuprimio Then
            Dim lblnNoHayError = False
            GobjPanDat.SControleProcesoObj(True)
            Try
                GobjPanDat.SInicialiceTransaccion()
                lblnSuprimio = ClsPanorama.FblnSuprimioCol(ColSectores_ModuloServicio)
                If lblnSuprimio Then
                    lblnSuprimio = MyBase.FblnSuprimio()
                End If
                If lblnSuprimio Then
                    GobjPanorama.SRegistreAccionLogApp(HstrNombreClase, "Suprimir Modulo No. " &
                                ObjIdModulo_ModuloServicioShr.ToString & " del Servicio No. " &
                                ObjIdServicio_ModuloServicioShr.ToString)
                    GobjPanDat.SConfirmeTransaccion()
                    MobjMiServicio.SActuEstaGeneradoProgramaFact(False)
                Else
                    GobjPanDat.SAborteTransaccion()
                End If
                lblnNoHayError = True
            Catch ex As ProveedorBdPanException
                Throw
            Catch ex As ArgumentOutOfRangeException
                Throw
            Catch ex As Exception
                Throw
            Finally
                If lblnNoHayError Then
                    GobjPanDat.SControleProcesoObj(False)
                Else
                    GobjPanDat.SAborteTransaccion()
                    GobjPanDat.SControleProcesoObj(False, True)
                End If
            End Try
        End If
        Return lblnSuprimio
    End Function
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Return FblnPermitidoSuprimir()
    End Function
    Private Function FblnServicioNoFacturado() As Boolean
        Dim lblnServicioNoFacturado = True
        Dim lstrAno = ObjIdAno_ModuloServicioShr.ToString
        Dim lstrIdServicio = ObjIdServicio_ModuloServicioShr.ToString
        Dim lstrNomTabla = ClsItemProgramaFact.SstrNombreTabla
        Dim lstrNomCampoCantPeri = ClsCantidadPeriodosShr.SstrNombreCampoBd
        Dim lstrNomCampoVlrPeri = ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd
        Dim lstrNomCampoSaldo = ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd
        Dim lstrCamposSelect() As String = {lstrNomCampoCantPeri, lstrNomCampoVlrPeri, lstrNomCampoSaldo}
        Dim lstrFiltro = StrCampoCarpeta & " = " & GshrIdCarpeta &
                " AND " & StrCampoCentroUtil & " = " & GshrIdCentroUtil & " AND " &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & lstrAno & " AND " &
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " = " & lstrIdServicio
        Dim ldtbItemsProFac = ClsPanorama.FdtbDataTable(lstrNomTabla, lstrCamposSelect, Nothing, lstrFiltro)
        Dim ldrwItemsProFac() = ldtbItemsProFac.Select()
        If ldrwItemsProFac.Length > 0 Then
            Dim lshrCantPerio As Short
            Dim ldecValorPer As Decimal
            Dim ldecSaldo As Decimal
            For Each ldrwItem As DataRow In ldrwItemsProFac
                lshrCantPerio = ClsPanorama.FobjValorCampo(ldrwItem(lstrNomCampoCantPeri),
                        EnuTipoValor.enuShort)
                ldecValorPer = ClsPanorama.FobjValorCampo(ldrwItem(lstrNomCampoVlrPeri),
                        EnuTipoValor.enuDecimal)
                ldecSaldo = ClsPanorama.FobjValorCampo(ldrwItem(lstrNomCampoSaldo),
                        EnuTipoValor.enuDecimal)
                If ldecSaldo <> lshrCantPerio * ldecValorPer Then
                    lblnServicioNoFacturado = False
                    Exit For
                End If
            Next
        End If
        Return lblnServicioNoFacturado
    End Function
#End Region
#Region "Procedimientos del objeto"
    ''' <summary>
    ''' Indica si el Sector identificado con el argumento 'ashrIdSector' Contribuye
    ''' con el modulo
    ''' </summary>
    ''' <param name="ashrIdSector">Identifica el Sector</param>
    ''' <returns>True si contribuye de lo contrario False</returns>
    ''' <remarks></remarks>
    Friend Function FblnSectorContribuyeModulo(ashrIdSector As Short) As Boolean
        Dim lobjModulo As ClsModuloContribucion
        Dim lcolModulos = GobjParametros.ColModulos
        If lcolModulos.Contains(ObjIdModulo_ModuloServicioShr.ToString) Then
            lobjModulo = lcolModulos(ObjIdModulo_ModuloServicioShr.ToString)
        Else
            Throw New ErrorInesperadoPanLException("El modulo no existe!")
        End If
        Dim lcolSectoresModulo As Collection = lobjModulo.ColSectoresModulo
        Return lcolSectoresModulo.Contains(ashrIdSector.ToString)
    End Function
    Friend Function FblnPuedeModificarPres(ByRef astrMens As String) As Boolean
        Dim lblnPuede = True
        If MobjMiServicio.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual Then
            If MobjMiServicio.ObjMiAno.ObjModuloPorServicioBln.ObjValorPro Then
                lblnPuede = False
                astrMens = "El Presupuesto solo puede ser modificado en el Año!"
            ElseIf GobjParametros.FblnPerActEsDicPrimerAno Then
                lblnPuede = False
                astrMens = "El Presupuesto no se puede cambiar si se está haciendo la " &
                        "instalación en Enero!"
            End If
            If lblnPuede Then
                lblnPuede = MobjMiServicio.ObjMiAno.FblnEsAnoActual OrElse
                        MobjMiServicio.ObjMiAno.ObjIdAnoShr.ObjValorPro >
                        GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
                If lblnPuede Then
                    Dim lobjPeriActu As ClsPeriodo = GobjParametros.ObjAnoActual.ObjPeriodoActual
                    lblnPuede = lobjPeriActu.ObjIdPeriodoShr.ObjValorPro < 12
                    If Not lblnPuede Then
                        lblnPuede = Not lobjPeriActu.BlnPeriodoFacturado
                    End If
                    If Not lblnPuede Then
                        astrMens = "El Presupuesto no puede ser modificado cuando " &
                                "diciembre ya está facturado!"
                    End If
                Else
                    astrMens = "El Valor del Presupuesto solo puede ser modificado " &
                            "en el Año Actual!"
                End If
            End If
        End If
        Return lblnPuede
    End Function
    ''' <summary>
    ''' Indica si todos los sectores que contribuyen con el módulo del servicio lo hacen con
    ''' el total del area
    ''' </summary>
    ''' <returns></returns>
    Friend Function FblnSectContriConTotalArea() As Boolean
        Dim lblnSi As Boolean
        For Each lobjSectMod As ClsSectorModulo In ObjMiModuloContribucion.ColSectoresModulo
            lblnSi = lobjSectMod.ObjTasaContribucionDbl.ObjValorPro = 1
            If Not lblnSi Then Exit For
        Next
        Return lblnSi
    End Function

#End Region
#Region "Manejo Sectores_ModuloServicio"
    Friend ReadOnly Property ColSectores_ModuloServicio As Collection
        Get
            If IsNothing(McolSectores_ModuloServicio) OrElse
                    McolSectores_ModuloServicio.Count = 0 Then
                McolSectores_ModuloServicio = New Collection
                SCargueDtbSectores_ModuloServicio()
                If Not IsNothing(MdtbSectores_ModuloServicio) AndAlso
                        MdtbSectores_ModuloServicio.Rows.Count > 0 Then
                    Dim ldrwSectoresModSer() As DataRow =
                            MdtbSectores_ModuloServicio.Select
                    For Each ldrwSectModSer In ldrwSectoresModSer
                        Dim lobjSecModSer As New ClsSectorModuloServicio(Me,
                                ldrwSectModSer)
                        lobjSecModSer.SLeaValores(True)
                        Dim lstrKey =
                                lobjSecModSer.ObjIdSector_SectorModuloServicioShr.ToString
                        McolSectores_ModuloServicio.Add(lobjSecModSer, lstrKey)
                    Next
                End If
            End If
            Return McolSectores_ModuloServicio
        End Get
    End Property
    Friend ReadOnly Property ObjNuevoSectorModuloServicio As ClsSectorModuloServicio
        Get
            SCargueDtbSectores_ModuloServicio()
            Dim lblnCambioPermiso = False
            Dim ldrwNuevo As DataRow = MdtbSectores_ModuloServicio.NewRow
            Dim lobjSectorModuloServicio As New ClsSectorModuloServicio(Me, ldrwNuevo)
            With lobjSectorModuloServicio
                If Not CType(.EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
                    .EnuPermisosObj += EnuPermisosDef.enuCrear
                    lblnCambioPermiso = True
                End If
                .SCreeObj(Nothing)
                .ObjIdCarpeta_SectorModuloServicioShr.ObjValorPro = GshrIdCarpeta
                .ObjIdCentroUtil_SectorModuloServicioShr.ObjValorPro = GshrIdCentroUtil
                .ObjIdAno_SectorModuloServicioShr.ObjValorPro = ObjIdAno_ModuloServicioShr.ObjValorPro
                .ObjIdServicio_SectorModuloServicioShr.ObjValorPro = ObjIdServicio_ModuloServicioShr.ObjValorPro
                .ObjIdModulo_SectorModuloServicioShr.ObjValorPro = ObjIdModulo_ModuloServicioShr.ObjValorPro
                If lblnCambioPermiso Then
                    .EnuPermisosObj -= EnuPermisosDef.enuCrear
                End If
            End With
            Return lobjSectorModuloServicio
        End Get
    End Property
    Friend Sub SAdicioneSector(aobjSectorModuloServicio As ClsSectorModuloServicio)
        If IsNothing(McolSectores_ModuloServicio) Then
            McolSectores_ModuloServicio = ColSectores_ModuloServicio
        End If
        Dim lstrKey = aobjSectorModuloServicio.ObjIdSector_SectorModuloServicioShr.ToString
        McolSectores_ModuloServicio.Add(aobjSectorModuloServicio, lstrKey)
    End Sub
    Friend Function FdtbSectoresModuloServicio() As DataTable
        Dim lstrTablaPri = ClsSector.SstrNombreTabla
        Dim lstrTablaSec = ClsSectorModuloServicio.SstrNombreTabla
        Dim lstrCamposTabPri = {ClsNombreSectorStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec = {ClsIdSector_SectorModuloServicioShr.SstrNombreCampoBd,
                 ClsValor_SectorModuloServicioDec.SstrNombreCampoBd}
        Dim lstrCamposPrimRel = {StrCampoCarpeta, StrCampoCentroUtil,
                                 ClsIdSectorShr.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta, StrCampoCentroUtil,
                                ClsIdSector_SectorModuloServicioShr.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsIdSector_SectorModuloServicioShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsIdAno_SectorModuloServicioShr.SstrNombreCampoBd & " = " &
                ObjIdAno_ModuloServicioShr.ToString & " AND " &
                ClsIdServicio_SectorModuloServicioShr.SstrNombreCampoBd & " = " &
                ObjIdServicio_ModuloServicioShr.ObjValorPro & " AND " &
                ClsIdModulo_SectorModuloServicioShr.SstrNombreCampoBd & " = " &
                ObjIdModulo_ModuloServicioShr.ObjValorPro
        Dim ldtbSecModSer As DataTable
        Try
            ldtbSecModSer = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamposTabPri,
                    lstrTablaSec, lstrCamposTabSec, lstrCamposPrimRel, lstrCamposRelSec,
                    lstrIndice, lstrFiltro, Array.Empty(Of String), True)
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        End Try
        Return ldtbSecModSer
    End Function
    Private Sub SCargueDtbSectores_ModuloServicio()
        If IsNothing(MdtbSectores_ModuloServicio) OrElse
                MdtbSectores_ModuloServicio.Rows.Count = 0 Then
            Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdAno_SectorModuloServicioShr.SstrNombreCampoBd & " = " &
                    ObjIdAno_ModuloServicioShr.ToString & " AND " &
                    ClsIdServicio_SectorModuloServicioShr.SstrNombreCampoBd & " = " &
                    ObjIdServicio_ModuloServicioShr.ObjValorPro & " AND " &
                    ClsIdModulo_SectorModuloServicioShr.SstrNombreCampoBd & " = " &
                    ObjIdModulo_ModuloServicioShr.ObjValorPro
            Dim lstrIndice(,) As String = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {ClsIdAno_SectorModuloServicioShr.SstrNombreCampoBd, "ASC"},
                              {ClsIdServicio_SectorModuloServicioShr.SstrNombreCampoBd, "ASC"},
                              {ClsIdModulo_SectorModuloServicioShr.SstrNombreCampoBd, "ASC"},
                              {ClsIdSector_SectorModuloServicioShr.SstrNombreCampoBd, "ASC"}}
            MdtbSectores_ModuloServicio = ClsPanorama.FdtbDataTable(
                    ClsSectorModuloServicio.SstrNombreTabla, {"*"}, lstrIndice,
                    lstrFiltro)
        End If
    End Sub
    Friend Function FdecValorParticipa(adecTotalPresupuesto As Decimal,
            aenuTipoBase As EnuTipoBaseCalculo) As Decimal
        Dim ldecValorParticipa = 0D
        Dim ldblFactorPart As Double
        Dim ldblTotalBasePart = GobjParametros.FdblTotalBasePart(aenuTipoBase)
        For Each lobjSectorMod As ClsSectorModulo In ObjMiModuloContribucion.ColSectoresModulo
            ldblFactorPart = lobjSectorMod.FdblBasePartiPonderada(aenuTipoBase) /
                    ldblTotalBasePart
            ldecValorParticipa += adecTotalPresupuesto * ldblFactorPart
        Next
        ldecValorParticipa = Math.Round(ldecValorParticipa, 0)
        Return ldecValorParticipa
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsIdAno_ModuloServicioShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAno"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdAno_ModuloServicio"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HstrOrdenIndice = "ASC"
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                GCSHRANOMAXIMO, BlnEsRequerido)
        If Not BlnLeyendoOrigen Then
            If lblnEsValido Then
                Dim lobjPadre As ClsModuloServicio = ObjPadre
                Dim lobjAbuelo As ClsServicio = lobjPadre.ObjPadre
                lblnEsValido = HobjValorNew = lobjAbuelo.ObjIdAno_ServicioShr.ObjValorPro
            End If
            If Not lblnEsValido Then
                HstrMens = "El Año ingresado no es valido!"
                SNotifiqueDatInv()
            End If
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdModulo_ModuloServicioShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdModuloContribucion"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdModuloContribucion"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 4
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1,
                       Short.MaxValue, BlnEsRequerido)
        If HblnEsValido Then
            Dim lcolModulos = GobjParametros.ColModulos
            HblnEsValido = lcolModulos.Contains(HobjValorNew.ToString)
            HstrMens = String.Empty
            If HblnEsValido Then
                If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    If Not BlnLeyendoOrigen Then
                        Dim lobjAbuelo As ClsServicio = ObjPadre.ObjPadre
                        Dim lobjLlavePrincipal() = {GshrIdCarpeta, GshrIdCentroUtil,
                                                    lobjAbuelo.ObjIdAno_ServicioShr.ObjValorPro,
                                                    lobjAbuelo.ObjIdServicioShr.ObjValorPro,
                                                    HobjValorNew}
                        If Not ObjPadre.FblnExisteLlave(lobjLlavePrincipal) Then
                            HstrMens = "La Id. del Módulo de Contribución ingresada, '" &
                                    lobjValorIng.ToString & "',  no existe!"
                            HblnEsValido = False
                        End If
                    End If
                End If
            Else
                HstrMens = "La Módulo de Contribución ingresado no existe!"
            End If
            If Not String.IsNullOrEmpty(HstrMens) Then
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdServicio_ModuloServicioShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdServicio"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdServicio_ModuloServicio"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1,
                       Short.MaxValue, BlnEsRequerido)
        Dim lobjValorIng = HobjValorNew
        If Not BlnLeyendoOrigen Then
            If HblnEsValido Then
                Dim lobjPadre As ClsModuloServicio = ObjPadre
                Dim lobjAbuelo As ClsServicio = lobjPadre.ObjPadre
                HblnEsValido = HobjValorNew = lobjAbuelo.ObjIdServicioShr.ObjValorPro
            End If
            If Not HblnEsValido Then
                HstrMens = "La Id. del Servicio ingresada, '" &
                        lobjValorIng.ToString & "',  no es valida!"
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsValorPres_ModuloServicioDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ValorPres"
    Private ReadOnly MobjPadre As ClsModuloServicio = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Valor Presupuesto Módulo"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                If HobjValorNew > 0 Then
                    HblnEsValido = Not GobjParametros.FblnPerActEsDicPrimerAno()
                End If
                If HblnEsValido Then
                    Dim lobjAbuelo As ClsServicio = ObjPadre.ObjPadre
                    If lobjAbuelo.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual Then
                        If GobjParametros.ObjAnoActual.ObjModuloPorServicioBln.ObjValorPro Then
                            HblnEsValido = If(lobjAbuelo.BlnCalculandoDesdeAno, True,
                                    HobjValorNew = 0 OrElse GblnImportando)
                            If Not HblnEsValido Then
                                HstrMens = "El valor de la contribución del módulo lo asigna el Año!"
                                SNotifiqueDatInv()
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
#End Region
