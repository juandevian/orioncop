Friend Class ClsEstadoCuenta
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriEstadosCuenta"
    ' Variables de modulo
    Private McolFacturasEstado As Collection = Nothing
    Private MdtbFacturasEstado As DataTable = Nothing
    Private McolServiciosEstado As Collection = Nothing
    Private MdtbServiciosEstado As DataTable = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Panorama.
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Indica si se instancia como un objeto navegable o como un Objeto único.</param>
    ''' <remarks>Si se instancia como un objeto navegable, se crea un datatable que contiene las columnas de
    ''' la llave con las llaves de todos los objetos y queda a la espera de que se indique que objeto abrir.
    ''' Si se instancia como un objeto único, queda a la espera de recibir el valor de los campos de la llave 
    ''' para abrir dicho objeto. </remarks>
    Public Sub New(aenuModoInstanciaObj As EnuModoInstanciaObjDef, ablnTodos As Boolean)
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
            Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion
            If Not ablnTodos Then
                lstrFiltro &= " AND " &
                   ClsIdFactura_EstadoEnt.SstrNombreCampoBd & " = 0 AND (" &
                   ClsDeudaCapitalDec.SstrNombreCampoBd & " > 0 OR " &
                   ClsDeudaIntMoraDec.SstrNombreCampoBd & " > 0)"
            End If
            HcolFiltros.Add(lstrFiltro)
            lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                   ClsIdEstadoCuentaEnt.SstrNombreCampoBd}
        Else
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        HblnEsModificable = False
        HblnEsAnulable = False
        HblnEsCreable = False
        HblnEsSuprimible = False
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto Estado de Cuenta como un objeto único 
    ''' </summary>
    ''' <remarks>Si se instancia como un objeto único, queda a la espera de recibir el valor de los 
    ''' campos de la llave para abrir dicho objeto. 
    ''' </remarks>
    Public Sub New(adrwEstadoCuenta As DataRow)
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        '
        DrwRegistroActual = adrwEstadoCuenta
        If Not IsNothing(drwRegistroActual) Then
            DtbTablaColeccion = DrwRegistroActual.Table
        End If
        HblnEsModificable = False
        HblnEsAnulable = False
        HblnEsSuprimible = False
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
            Return EnuIdClasesPanDef.enuEstadoCuenta
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Estado de Cuenta"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjAntPorAplDec As New ClsAntPorAplDec(Me)
    Friend ReadOnly Property ObjDeudaCapitalDec As New ClsDeudaCapitalDec(Me)
    Friend ReadOnly Property ObjDeudaIntMoraDec As New ClsDeudaIntMoraDec(Me)
    Friend ReadOnly Property ObjFechaEstadoDtm As New ClsFechaEstadoDtm(Me)
    Friend ReadOnly Property ObjIdCarpeta_EstadoShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_EstadoShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdCliente_EstadoDbl As New ClsIdCliente_EstadoDbl(Me)
    Friend ReadOnly Property ObjIdEstadoCuentaEnt As New ClsIdEstadoCuentaEnt(Me)
    Friend ReadOnly Property ObjIdFactura_EstadoEnt As New ClsIdFactura_EstadoEnt(Me)
    Friend ReadOnly Property ObjIdPredioAgr_EstadoStr As New ClsIdPredioAgr_EstadoStr(Me)
    Friend ReadOnly Property ObjPrefijoFac_EstadoStr As New ClsPrefijoFac_EstadoStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAntPorAplDec)
                HcolPropiedades.Add(ObjDeudaCapitalDec)
                HcolPropiedades.Add(ObjDeudaIntMoraDec)
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjFechaEstadoDtm)
                HcolPropiedades.Add(ObjIdCarpeta_EstadoShr)
                HcolPropiedades.Add(ObjIdCentroUtil_EstadoShr)
                HcolPropiedades.Add(ObjIdPredioAgr_EstadoStr)
                HcolPropiedades.Add(ObjIdCliente_EstadoDbl)
                HcolPropiedades.Add(ObjIdEstadoCuentaEnt)
                HcolPropiedades.Add(ObjIdFactura_EstadoEnt)
                HcolPropiedades.Add(ObjPrefijoFac_EstadoStr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend ReadOnly Property DecTotalDeuda As Decimal
        Get
            Dim ldecTotalDeuda As Decimal = ObjDeudaCapitalDec.ObjValorPro + ObjDeudaIntMoraDec.ObjValorPro
            Return ldecTotalDeuda
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        MdtbFacturasEstado = Nothing
        MdtbServiciosEstado = Nothing
        McolFacturasEstado = Nothing
        McolserviciosEstado = Nothing
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                SNumereObj()
                ObjFechaCreacionDtm.ObjValorPro = Now
                If Not IsNothing(McolFacturasEstado) AndAlso McolFacturasEstado.Count > 0 Then
                    ClsPanorama.SActualiceCol(McolFacturasEstado)
                End If
                MyBase.SActualice(ablnExigeRequeridos)
            Else
                MyBase.SActualice(ablnExigeRequeridos)
            End If
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
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Protected Overrides Function SAnuleEnObj() As Boolean
        Dim lblnAnulo = True
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        End If
        ObjIdFactura_EstadoEnt.ObjValorPro = 0
        ObjPrefijoFac_EstadoStr.ObjValorPro = String.Empty
        Return lblnAnulo
    End Function
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjPrefijoFac_EstadoStr.ToString & "," & ObjIdFactura_EstadoEnt.ObjValorPro
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lentIdEstado As Integer
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion
            lentIdEstado = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                    ClsIdEstadoCuentaEnt.SstrNombreCampoBd, ObjIdEstadoCuentaEnt.EnuTipoValor,
                    lstrFiltro) + 1
            ObjIdEstadoCuentaEnt.ObjValorPro = lentIdEstado
            If Not IsNothing(McolFacturasEstado) AndAlso McolFacturasEstado.Count > 0 Then
                For Each lobjFacturaEstado As ClsFacturaEstado In McolFacturasEstado
                    lobjFacturaEstado.ObjIdEstadoCta_FacEstadoShr.ObjValorPro = lentIdEstado
                Next
            End If
        End If
    End Sub
    Private Shared Function FdtbFacturasEstadoEstructura() As DataTable
        Dim ldtbFacturasEstado As New DataTable
        Dim ldclPrefFac As New DataColumn("PrefijoFactViva", System.Type.GetType("System.String"))
        Dim ldclIdFac As New DataColumn("IdFacturaViva", System.Type.GetType("System.Int32"))
        Dim ldclIdItemFac As New DataColumn("IdItemFactura", System.Type.GetType("System.Int32"))
        Dim ldclFecFac As New DataColumn("FechaFactura", System.Type.GetType("System.DateTime"))
        Dim ldclVlrFac As New DataColumn("ValorFactura", System.Type.GetType("System.Decimal"))
        Dim ldclDbFac As New DataColumn("DeudaCap", System.Type.GetType("System.Decimal"))
        Dim ldclCrFac As New DataColumn("DeudaInt", System.Type.GetType("System.Decimal"))
        Dim ldclSaldoFac As New DataColumn("SaldoFactura", System.Type.GetType("System.Decimal"))
        ldtbFacturasEstado.Columns.Add(ldclPrefFac)
        ldtbFacturasEstado.Columns.Add(ldclIdFac)
        ldtbFacturasEstado.Columns.Add(ldclIdItemFac)
        ldtbFacturasEstado.Columns.Add(ldclFecFac)
        ldtbFacturasEstado.Columns.Add(ldclVlrFac)
        ldtbFacturasEstado.Columns.Add(ldclDbFac)
        ldtbFacturasEstado.Columns.Add(ldclCrFac)
        ldtbFacturasEstado.Columns.Add(ldclSaldoFac)
        Return ldtbFacturasEstado
    End Function
#End Region
#Region "Manejo Items Facturas Estado"
    Friend Sub SAdicioneFacturaEstado(aobjFactura As ClsFactura, adtmFechaEstado As Date) ' Ok
        Dim lobjFacturaEstado As ClsFacturaEstado = Nothing
        If McolFacturasEstado Is Nothing Then
            McolFacturasEstado = New Collection
        End If
        For Each lobjItemFac As ClsItemFactura In aobjFactura.ColItemsFactura
            If lobjItemFac.DecDeuda > 0 Then
                Dim ldecMoraMes As Decimal = FdecIteresesMes(aobjFactura,
                        lobjItemFac.ObjIdItemFacturaShr.ObjValorPro, adtmFechaEstado)
                lobjFacturaEstado = FobjNuevaFactEstado()
                With lobjFacturaEstado
                    .ObjCreditos_ItFacEstadoDec.ObjValorPro = lobjItemFac.ObjCreditos_ItemFactDec.ObjValorPro
                    .ObjDebitos_ItFacEstadoDec.ObjValorPro = lobjItemFac.ObjDebitos_ItemFactDec.ObjValorPro
                    .ObjDetalleItemFac_EstadoStr.ObjValorPro = lobjItemFac.ObjDetalle_ItemFactStr.ObjValorPro
                    .ObjFecha_FacEstadoDtm.ObjValorPro = aobjFactura.ObjFechaFacturaDtm.ObjValorPro
                    .ObjIdAno_ItemFactEstadoShr.ObjValorPro = lobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro
                    .ObjPrefijoFacturaVivaStr.ObjValorPro = aobjFactura.ObjPrefijo_FactStr.ObjValorPro
                    .ObjIdFacturaVivaEnt.ObjValorPro = aobjFactura.ObjIdFacturaEnt.ObjValorPro
                    .ObjIdServicioItemFac_EstadoShr.ObjValorPro = lobjItemFac.ObjIdServicio_ItemFactShr.ObjValorPro
                    .ObjVlrItemFac_EstadoDec.ObjValorPro = lobjItemFac.ObjValor_ItemFactDec.ObjValorPro
                    .ObjDeudaCap_ItFacEstDec.ObjValorPro = lobjItemFac.FdecDeudaCapital
                    .ObjDeudaIntMes_ItFacEstDec.ObjValorPro = ldecMoraMes
                    .ObjDeudaIntMora_ItFacEstDec.ObjValorPro = lobjItemFac.FdecDeudaIntTotal
                    .ObjOrdinal_FacEstadoShr.ObjValorPro = McolFacturasEstado.Count + 1
                End With
                McolFacturasEstado.Add(lobjFacturaEstado)
                ObjDeudaCapitalDec.ObjValorPro += lobjItemFac.FdecDeudaCapital
                ObjDeudaIntMoraDec.ObjValorPro += lobjItemFac.FdecDeudaIntTotal
            End If
        Next
    End Sub
    Friend Sub SAdicioneFacturaEstado(aobjFactura As ClsFactura, adtmFechaEstado As Date,
                ashrIdServicio As Short)
        Dim lobjFacturaEstado As ClsFacturaEstado = Nothing, lblnAdicione As Boolean
        Dim ldecMoraMes = 0D
        If McolFacturasEstado Is Nothing Then
            McolFacturasEstado = New Collection
        End If
        For Each lobjItemFac As ClsItemFactura In aobjFactura.ColItemsFactura
            lblnAdicione = If(ashrIdServicio = 0,
                    lobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro > 0,
                    lobjItemFac.ObjIdServicio_ItemFactShr.ObjValorPro = ashrIdServicio)
            If lobjItemFac.DecDeuda > 0 AndAlso lblnAdicione Then
                ldecMoraMes = FdecIteresesMes(aobjFactura,
                        lobjItemFac.ObjIdItemFacturaShr.ObjValorPro, adtmFechaEstado)
                lobjFacturaEstado = FobjNuevaFactEstado()
                With lobjFacturaEstado
                    .ObjCreditos_ItFacEstadoDec.ObjValorPro = lobjItemFac.ObjCreditos_ItemFactDec.ObjValorPro
                    .ObjDebitos_ItFacEstadoDec.ObjValorPro = lobjItemFac.ObjDebitos_ItemFactDec.ObjValorPro
                    .ObjDetalleItemFac_EstadoStr.ObjValorPro = lobjItemFac.ObjDetalle_ItemFactStr.ObjValorPro
                    .ObjFecha_FacEstadoDtm.ObjValorPro = aobjFactura.ObjFechaFacturaDtm.ObjValorPro
                    .ObjIdAno_ItemFactEstadoShr.ObjValorPro = lobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro
                    .ObjPrefijoFacturaVivaStr.ObjValorPro = aobjFactura.ObjPrefijo_FactStr.ObjValorPro
                    .ObjIdFacturaVivaEnt.ObjValorPro = aobjFactura.ObjIdFacturaEnt.ObjValorPro
                    .ObjIdServicioItemFac_EstadoShr.ObjValorPro = lobjItemFac.ObjIdServicio_ItemFactShr.ObjValorPro
                    .ObjVlrItemFac_EstadoDec.ObjValorPro = lobjItemFac.ObjValor_ItemFactDec.ObjValorPro
                    .ObjDeudaCap_ItFacEstDec.ObjValorPro = lobjItemFac.FdecDeudaCapital
                    .ObjDeudaIntMes_ItFacEstDec.ObjValorPro = ldecMoraMes
                    .ObjDeudaIntMora_ItFacEstDec.ObjValorPro = lobjItemFac.FdecDeudaIntTotal
                    .ObjOrdinal_FacEstadoShr.ObjValorPro = McolFacturasEstado.Count + 1
                End With
                McolFacturasEstado.Add(lobjFacturaEstado)
                ObjDeudaCapitalDec.ObjValorPro += lobjItemFac.FdecDeudaCapital
                ObjDeudaIntMoraDec.ObjValorPro += lobjItemFac.FdecDeudaIntTotal
            End If
        Next
    End Sub
    Private Function FobjNuevaFactEstado() As ClsFacturaEstado
        SCargueDtbFacturasEstado()
        Dim ldrwNuevaFactEstado As DataRow = MdtbFacturasEstado.NewRow
        Dim lobjFactEstado As New ClsFacturaEstado(Me, ldrwNuevaFactEstado)
        Dim lblnModificoPermisos = False
        With lobjFactEstado
            If Not CType(.EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
                .EnuPermisosObj += EnuPermisosDef.enuCrear
                lblnModificoPermisos = True
            End If
            .SCreeObj(Nothing)
            .ObjIdCarpeta_FacEstadoShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_FacEstadoShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdPredioAgr_FacEstadoStr.ObjValorPro = ObjIdPredioAgr_EstadoStr.ObjValorPro
            .ObjIdCliente_FacEstadoDbl.ObjValorPro = ObjIdCliente_EstadoDbl.ObjValorPro
            If lblnModificoPermisos Then
                .EnuPermisosObj -= EnuPermisosDef.enuCrear
            End If
        End With
        Return lobjFactEstado
    End Function
    Friend ReadOnly Property ColFacturasEstado As Collection
        Get
            If IsNothing(McolFacturasEstado) Then
                McolFacturasEstado = New Collection
                If EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                    SCargueDtbFacturasEstado()
                    If Not IsNothing(MdtbFacturasEstado) AndAlso MdtbFacturasEstado.Rows.Count > 0 Then
                        Dim ldrwFacturasEstado() As DataRow = MdtbFacturasEstado.Select
                        For Each ldrwFacturaEstado As DataRow In ldrwFacturasEstado
                            Dim lobjFacturaEstado As New ClsFacturaEstado(Me, ldrwFacturaEstado)
                            lobjFacturaEstado.SLeaValores(True)
                            McolFacturasEstado.Add(lobjFacturaEstado)
                        Next
                    End If
                End If
            End If
            Return McolFacturasEstado
        End Get
    End Property
    Friend ReadOnly Property DtbFacturasEstado As DataTable
        Get
            SCargueDtbFacturasEstado()
            SComplementeDtb()
            Return MdtbFacturasEstado
        End Get
    End Property
    Friend Function FdtbFactEstadoResum() As DataTable
        Dim lstrIdEstado = "0"
        Dim lshrIdAno As Short = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
        If Not String.IsNullOrEmpty(ObjIdEstadoCuentaEnt.ToString) Then
            lstrIdEstado = ObjIdEstadoCuentaEnt.ToString
        End If
        Dim lstrTabla = ClsFacturaEstado.SstrNombreTabla
        Dim lstrCamposSele As String() =
                {"IF (" & ClsIdAno_ItemFactEstadoShr.SstrNombreCampoBd & " > 0," & lshrIdAno &
                ", 0) AS IdAnos", ClsIdServicioItemFac_EstadoShr.SstrNombreCampoBd,
                "SUM(" & ClsVlrItemFac_EstadoDec.SstrNombreCampoBd & ") AS Valor",
                "SUM(" & ClsDeudaCapitalDec.SstrNombreCampoBd & ") AS DeudaCapital",
                "SUM(" & ClsDeudaIntMoraDec.SstrNombreCampoBd & ") AS DeudaMora",
                "'' AS Servicio", "(SUM(" & ClsDeudaCapitalDec.SstrNombreCampoBd & ") + " &
                "SUM(" & ClsDeudaIntMoraDec.SstrNombreCampoBd & ")) AS Saldo"}
        Dim lstrCampGroup As String() = {"IdAnos", ClsIdServicioItemFac_EstadoShr.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{"IdAnos", "DESC"},
                {ClsIdServicioItemFac_EstadoShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdEstado_FactEstadoEnt.SstrNombreCampoBd & " = " & lstrIdEstado
        Dim ldtbFactEstRes = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSele, lstrOrden,
                lstrFiltro, False, lstrCampGroup)
        SComplementeDtbFacEstResum(ldtbFactEstRes)
        Return ldtbFactEstRes
    End Function
    Private Sub SComplementeDtb()
        If MdtbFacturasEstado.Rows.Count > 0 Then
            Dim ldrwFacturasEstado() As DataRow = MdtbFacturasEstado.Select
            For Each ldrwFacEst As DataRow In ldrwFacturasEstado
                Dim ldecDebitos As Decimal = ClsPanorama.FobjValorCampo(ldrwFacEst(
                            ClsDebitos_ItFacEstadoDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                Dim ldecCreditos As Decimal = ClsPanorama.FobjValorCampo(ldrwFacEst(
                            ClsCreditos_ItFacEstadoDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                Dim ldecSaldo = ldecDebitos - ldecCreditos
                ldrwFacEst("SaldoItemFac") = ldecSaldo
            Next
        End If
    End Sub
    Private Sub SCargueDtbFacturasEstado()
        If IsNothing(MdtbFacturasEstado) Then
            Dim lstrIdEstado = "0"
            If Not String.IsNullOrEmpty(ObjIdEstadoCuentaEnt.ToString) Then
                lstrIdEstado = ObjIdEstadoCuentaEnt.ToString
            End If
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {ClsIdEstado_FactEstadoEnt.SstrNombreCampoBd, "ASC"},
                              {ClsOrdinal_FacEstadoShr.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdEstado_FactEstadoEnt.SstrNombreCampoBd & " = " & lstrIdEstado
            Dim lstrCamposSelect() = {"*", "0 as SaldoItemFac"}
            MdtbFacturasEstado = ClsPanorama.FdtbDataTable(ClsFacturaEstado.SstrNombreTabla,
                    lstrCamposSelect, lstrIndice, lstrFiltro)
        End If
    End Sub
    Private Sub SComplementeDtbFacEstResum(adtbFacturasEstado As DataTable)
        Dim lshrIdAno As Short, lshrIdServicio As Short, lstrServicio As String
        For Each ldrwReg As DataRow In adtbFacturasEstado.Rows
            lshrIdAno = ClsPanorama.FobjValorCampo(ldrwReg("IdAnos"), EnuTipoValor.enuShort)
            lshrIdServicio = ClsPanorama.FobjValorCampo(ldrwReg(
                    ClsIdServicioItemFac_EstadoShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
            If lshrIdAno > 0 Then
                lstrServicio = "Cuota de Administración"
            Else
                lstrServicio = GobjParametros.FstrNombreServicio(lshrIdServicio)
            End If
            ldrwReg("Servicio") = lstrServicio
        Next
    End Sub
    Private Function FdecIteresesMes(aobjFactura As ClsFactura, ashrIdItemFac As Short,
        adtmFechaEstado As Date) As Decimal
        Dim ldecIntMes As Decimal, lstrPrefFac = aobjFactura.ObjPrefijo_FactStr.ObjValorPro,
                lentIdFact = aobjFactura.ObjIdFacturaEnt.ObjValorPro,
                ldblIdCliente = aobjFactura.ObjIdCliente_FactDbl.ObjValorPro,
                lstrIdPreAgr = aobjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
        Dim ldtmFecha As Date = adtmFechaEstado.AddDays(1)
        Dim lshrIdAno As Short = ldtmFecha.Year
        Dim lstrIdPer As String = Format(ldtmFecha.Month, "0#")
        If GobjParametros.ColAnos.Contains(lshrIdAno.ToString()) Then
            Dim lobjAno As ClsAno = GobjParametros.ColAnos(lshrIdAno.ToString)
            Dim lobjPer As ClsPeriodo = lobjAno.ColPeriodos(lstrIdPer)
            Dim lstrFecIni As String = ClsPanoramaDat.FstrFechaNormalizada(lobjPer.DtmFechaInicioPeriodo)
            Dim lstrFecFin As String = ClsPanoramaDat.FstrFechaNormalizada(ldtmFecha)
            Dim lstrExpSql = "SELECT SUM(I." & ClsValor_ItemNotaDbDec.SstrNombreCampoBd & ") FROM " &
                ClsItemNotaDb.SstrNombreTabla & " AS I INNER JOIN " & ClsNotaDb.SstrNombreTabla &
                " AS N ON I." & ClsIdCarpetaShr.SstrNombreCampoBd & " = N." &
                ClsIdCarpetaShr.SstrNombreCampoBd & " AND I." & ClsIdCentroUtilShr.SstrNombreCampoBd &
                " = N." & ClsIdCentroUtilShr.SstrNombreCampoBd & " AND I." &
                ClsPrefijo_NotaDbStr.SstrNombreCampoBd & " = N." & ClsPrefijo_NotaDbStr.
                SstrNombreCampoBd & " AND I." & ClsIdNotaDb_ItemNotaDbEnt.SstrNombreCampoBd &
                " = N." & ClsIdNotaDbEnt.SstrNombreCampoBd & " WHERE I." & ClsIdCarpetaShr.
                SstrNombreCampoBd & " = " & GshrIdCarpeta & " AND I." & ClsIdCentroUtilShr.
                SstrNombreCampoBd & " = " & GshrIdCentroUtil & " AND " & ClsPrefijoFact_ItemNotaDbStr.
                SstrNombreCampoBd & " = '" & lstrPrefFac & "' AND " & ClsIdFacturaEnt.SstrNombreCampoBd &
                " = " & lentIdFact & " AND " & ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " &
                ldblIdCliente & " AND " & ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " = '" &
                lstrIdPreAgr & "' AND " & ClsFecha_NotaDbDtm.SstrNombreCampoBd & " BETWEEN '" &
                lstrFecIni & "' AND '" & lstrFecFin & "' AND " &
                ClsIdItemFac_ItemNotaDbShr.SstrNombreCampoBd & " = " & ashrIdItemFac
            Dim ldtbIntMes = ClsPanorama.FdtbDataTable(lstrExpSql)
            ldecIntMes = ClsPanorama.FobjValorCampo(ldtbIntMes(0)(0), EnuTipoValor.enuDecimal)
        Else
            ldecIntMes = 0
        End If
        Return ldecIntMes
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsAntPorAplDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "AntPorApl"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DeudaCapital"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsEstadoCuenta = ObjPadre
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
        If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
Friend Class ClsDeudaCapitalDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DeudaCapital"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DeudaCapital"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsEstadoCuenta = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
        If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
Friend Class ClsDeudaIntMoraDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DeudaIntMora"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "MoraAnterior"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        Dim lobjPadre As ClsEstadoCuenta = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
        If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
Friend Class ClsFechaEstadoDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaEstado"
    Private ReadOnly MobjPadre As ClsEstadoCuenta = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha del Estado"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = GCDTMFECHANULA
        HobjValorPro = GCDTMFECHANULA
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin = GobjParametros.ObjAnoActual.DtmFechaInicioAno.AddDays(-1)
        Dim ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
        If ldtmFechaMax > Date.Today Then
            ldtmFechaMax = Now
        End If
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
        Else
            HblnEsValido = (HobjValorNew = HobjValorOriginal)
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, GCSTRFMTFECHASIMPLE)
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class
Friend Class ClsIdCliente_EstadoDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCliente"
    Private MobjCliente As ClsCliente = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdTerceroCliente_Estado"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng = HobjValorNew
        Dim lobjPadre As ClsEstadoCuenta = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC, BlnEsRequerido)
        If Not HblnEsValido Then
            HblnEsValido = Not String.IsNullOrEmpty(
                    lobjPadre.ObjIdPredioAgr_EstadoStr.ToString())
            If Not HblnEsValido Then
                Throw New ErrorInesperadoPanLException("La Id. del Cliente ingresada, '" &
                        lobjValorIng.ToString & "',  no es válida!")
            End If
        End If
    End Sub
    Friend ReadOnly Property ObjCliente As ClsCliente
        Get
            Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
            MobjCliente = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
            MobjCliente.SAbra(lobjValorLlave)
            Return MobjCliente
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdEstadoCuentaEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdEstadoCuenta"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdEstadoCuenta"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue,
                BlnEsRequerido, EnuTipoValor)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdFactura_EstadoEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdFactura"
    Private ReadOnly MobjPadre As ClsEstadoCuenta = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdFactura_Estado"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
Friend Class ClsIdPredioAgr_EstadoStr
    'Herencia
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredioAgrupador"
    Private ReadOnly MobjPadre As ClsCBObjetoPan = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = ObjPadre
        HstrNombre = "IdPredioAgrupador_Estado"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud,
                BlnEsRequerido)
        If HblnEsValido AndAlso HobjValorNew <> "" Then
            Dim lobjLlavePrincipal() As Object = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                lobjPredio.SAbra(lobjLlavePrincipal)
                HblnEsValido = lobjPredio.BlnExiste
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsPrefijoFac_EstadoStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoFact"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PrefijoFactura_Estado"
        HshrLongitud = 5
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then
            HobjValorNew = String.Empty
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
#End Region