Imports System.Text
Friend Class ClsActualizacionApl
#Region "Actualizacion Versión"
    Friend Shared Sub SActualiceVer(astrVersionAnt As String, ByRef ablnActualizoApl As Boolean)
        Dim lblnNoHayErrores = False
        GblnActualizandoApp = True
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            If astrVersionAnt < "14.30.380.0015" AndAlso astrVersionAnt >= "13.21.361.1068" Then
                SAct_14_30_380()
            End If
            If astrVersionAnt < "15.32.384.1208" Then
                SAct_15_32_384()
            End If
            If astrVersionAnt < "15.32.388.1225" Then
                SAct_15_32_388()
            End If
            If astrVersionAnt < "15.32.388.1226" Then
                SRepareNovedades()
            End If
            If astrVersionAnt < "15.32.389.1230" Then
                SAct_15_32_389()
            End If
            If astrVersionAnt < "16.33.391.1248" Then
                SAct_16_33_391()
            End If
            If astrVersionAnt < "16.34.396.1280" Then
                SAct_16_33_396()
            End If
            If astrVersionAnt < "16.35.400.1306" Then
                SAct_16_35_400()
            End If
            If astrVersionAnt < "17.36.405.1348" Then
                SAct_17_36_405()
            End If
            If astrVersionAnt < "17.38.409.1371" Then
                SAct_17_38_409()
            End If
            ablnActualizoApl = True
            lblnNoHayErrores = True
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayErrores Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
        GblnActualizandoApp = False
    End Sub
#Region "14.30.380"
    Private Shared Sub SAct_14_30_380()
        GobjPanDat.SControleProcesoObj(True)
        Dim ldtbFras = FdtbFacturas()
        Dim lshrIdCarpeta As Short, lshrIdCentroUtil As Short, lstrPref As String, lentIdFac As Integer
        Dim ldtmFechaCauso As Date
        For Each ldrwFac As DataRow In ldtbFras.Rows
            lshrIdCarpeta = ClsPanorama.FobjValorCampo(ldrwFac(StrCampoCarpeta),
                    EnuTipoValor.enuShort)
            lshrIdCentroUtil = ClsPanorama.FobjValorCampo(ldrwFac(StrCampoCentroUtil),
                    EnuTipoValor.enuShort)
            lstrPref = ClsPanorama.FobjValorCampo(ldrwFac(ClsPrefijo_FactStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lentIdFac = ClsPanorama.FobjValorCampo(ldrwFac(ClsIdFacturaEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            ldtmFechaCauso = ClsPanorama.FobjValorCampo(ldrwFac(ClsFechaCausoIntMora_Dtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate)
            SActuItems(lshrIdCarpeta, lshrIdCentroUtil, lstrPref, lentIdFac, ldtmFechaCauso)
        Next
        SActuBaseServ()
        SActuServ()
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Shared Sub SActuItems(ashrIdCarpeta As Short, ashrIdCentroUtil As Short,
            astrPreFac As String, aentIdFac As Integer, adtmFechaCauso As Date)
        Dim lcolNomCampos As New Collection From {
            ClsFechaCausoIntMora_Dtm.SstrNombreCampoBd
        }
        Dim lcolDatos As New Collection From {
            {adtmFechaCauso, ClsFechaCausoIntMora_Dtm.SstrNombreCampoBd}
        }
        Dim lcolCamRef As New Collection From {
            StrCampoCarpeta,
            StrCampoCentroUtil,
            ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
            ClsIdFactura_ItemFactEnt.SstrNombreCampoBd
        }
        Dim lcolDatosRef As New Collection From {
            ashrIdCarpeta, ashrIdCentroUtil, astrPreFac, aentIdFac
        }
        Dim lstrTabla = ClsItemFactura.SstrNombreTabla
        GobjPanDat.SActualiceRegistro(lstrTabla, lcolNomCampos, lcolDatos,
                lcolCamRef, lcolDatosRef)
    End Sub
    Private Shared Function FdtbFacturas() As DataTable
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCampSel = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_FactStr.SstrNombreCampoBd, ClsIdFacturaEnt.SstrNombreCampoBd,
                ClsFechaCausoIntMora_Dtm.SstrNombreCampoBd}
        Dim lstrFiltro = String.Empty
        Dim lstrOrden As String(,) = {{StrCampoCarpeta, "ASC"},
                {StrCampoCentroUtil, "ASC"},
                {ClsPrefijo_FactStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}}
        Dim ldtbFacs = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSel, lstrOrden, lstrFiltro)
        Return ldtbFacs
    End Function
    Private Shared Sub SActuBaseServ()
        Dim lstrTabla = ClsAno.SstrNombreTabla
        Dim lstbSql As New StringBuilder
        With lstbSql
            .Append("UPDATE ").Append(lstrTabla).Append(" SET ")
            .Append(ClsTipoCalculoCuotaByt.SstrNombreCampoBd).Append(" = ")
            .Append(ClsTipoCalculoCuotaByt.SstrNombreCampoBd).Append(" + 1 ")
            .Append(" WHERE ").Append(ClsTipoCalculoCuotaByt.SstrNombreCampoBd)
            .Append(" >= 2")
        End With
        Dim lstrSql = lstbSql.ToString()
        GobjPanDat.SEjecuteSentenciaSql(lstrSql)
        lstrTabla = ClsServicio.SstrNombreTabla
        With lstbSql
            .Clear.Append("UPDATE ").Append(lstrTabla).Append(" SET ")
            .Append(ClsTipoBaseCalculoByt.SstrNombreCampoBd).Append(" = ")
            .Append(ClsTipoBaseCalculoByt.SstrNombreCampoBd).Append(" + 1 ")
            .Append(" WHERE ").Append(ClsTipoBaseCalculoByt.SstrNombreCampoBd)
            .Append(" >= 2")
        End With
        lstrSql = lstbSql.ToString()
        GobjPanDat.SEjecuteSentenciaSql(lstrSql)
    End Sub
    Private Shared Sub SActuServ()
        Dim lstrTabla = ClsServicio.SstrNombreTabla
        Dim lstbSql As New StringBuilder
        With lstbSql
            .Append("UPDATE ").Append(lstrTabla).Append(" SET ")
            .Append(ClsEstaAjustadoBln.SstrNombreCampoBd).Append(" = TRUE ")
            .Append(" WHERE ").Append(ClsIdAno_ServicioShr.SstrNombreCampoBd)
            .Append(" > 0")
        End With
        Dim lstrSql = lstbSql.ToString()
        GobjPanDat.SEjecuteSentenciaSql(lstrSql)
    End Sub
#End Region
#Region "15_32_384"
    Private Shared Sub SAct_15_32_384()
        ' Actualiza nombre de predio con el idpredio
        Dim lstrTabla = ClsPredio.SstrNombreTabla
        Dim lstbSql As New StringBuilder
        With lstbSql
            .Append("UPDATE ").Append(lstrTabla).Append(" SET ")
            .Append(ClsNombrePredioStr.SstrNombreCampoBd).Append(" = ")
            .Append(ClsIdPredioStr.SstrNombreCampoBd)
        End With
        Dim lstrSql = lstbSql.ToString()
        Dim lentRegAfe = GobjPanDat.SEjecuteSentenciaSql(lstrSql)
        lstrTabla = ClsItemRecCaja.SstrNombreTabla
        With lstbSql
            .Clear.Append("UPDATE ").Append(lstrTabla).Append(" SET ")
            .Append(ClsIdItemFac_ItemRecShr.SstrNombreCampoBd).Append(" = 1 WHERE ")
            .Append(ClsIdItemFac_ItemRecShr.SstrNombreCampoBd).Append(" = 0 ")
        End With
        lstrSql = lstbSql.ToString()
        lentRegAfe = GobjPanDat.SEjecuteSentenciaSql(lstrSql)
    End Sub
#End Region
#Region "15_32_388"
    Private Shared Sub SAct_15_32_388()
        ' Actualiza fechas en los items de factura
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lentIdCarpeta As Integer, lentIdCentroUtil As Integer, lstrPrefFac As String
        Dim lentIdFact As Integer, lstrFiltro As String, lstrExpSql As String
        Dim lstrFecGra As String, lstrFecVen As String, ldtmFechaGra As Date, ldtmfechaVen As Date
        Dim lstrCamSel As String() = {StrCampoCarpeta,
            StrCampoCentroUtil, ClsPrefijo_FactStr.SstrNombreCampoBd,
            ClsIdFacturaEnt.SstrNombreCampoBd, ClsFechaGraciaDtm.SstrNombreCampoBd,
            ClsFechaVencimientoDtm.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{StrCampoCarpeta, "ASC"},
                {StrCampoCentroUtil, "ASC"},
                {ClsPrefijo_FactStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}}
        Dim ldtbFacs = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSel, lstrOrden, "")
        For Each ldrwFac As DataRow In ldtbFacs.Rows
            lentIdCarpeta = ClsPanorama.FobjValorCampo(ldrwFac(StrCampoCarpeta),
                    EnuTipoValor.enuInteger)
            lentIdCentroUtil = ClsPanorama.FobjValorCampo(ldrwFac(StrCampoCentroUtil),
                    EnuTipoValor.enuInteger)
            lstrPrefFac = ClsPanorama.FobjValorCampo(ldrwFac(ClsPrefijo_FactStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lentIdFact = ClsPanorama.FobjValorCampo(ldrwFac(ClsIdFacturaEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            ldtmFechaGra = ClsPanorama.FobjValorCampo(ldrwFac(ClsFechaGraciaDtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate)
            ldtmfechaVen = ClsPanorama.FobjValorCampo(ldrwFac(ClsFechaVencimientoDtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate)
            lstrFecGra = ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaGra)
            lstrFecVen = ClsPanoramaDat.FstrFechaNormalizada(ldtmfechaVen)
            lstrFiltro = StrCampoCarpeta & " = " & lentIdCarpeta & " AND " &
                    StrCampoCentroUtil & " = " & lentIdCentroUtil & " AND " &
                    ClsPrefijo_ItemFactStr.SstrNombreCampoBd & " = '" & lstrPrefFac & "' AND " &
                    ClsIdFactura_ItemFactEnt.SstrNombreCampoBd & " = " & lentIdFact
            lstrExpSql = "UPDATE " & ClsItemFactura.SstrNombreTabla & " SET " &
                    ClsFechaGraciaIFDtm.SstrNombreCampoBd & " = '" & lstrFecGra & "', " &
                    ClsFechaVencimientoIFDtm.SstrNombreCampoBd & " = '" & lstrFecVen & "' WHERE " &
                    lstrFiltro
            GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        Next
    End Sub
#End Region
#Region "Reparacion Novedades"
    Friend Shared Sub SRepareNovedades()
        Dim lstrTabla = ClsCarpeta.SstrNombreTabla
        Dim lstrTablaF = ClsNovedad.SstrNombreTabla
        Dim lstrCampSele = {StrCampoCarpeta}
        Dim lstrCampSeleF = {"DISTINCT " & ClsIdFactura_NovEnt.SstrNombreCampoBd}
        Dim lstrOrden = {{StrCampoCarpeta, "ASC"}}
        Dim lstrOrdenF = {{ClsIdFactura_NovEnt.SstrNombreCampoBd, "ASC"}}
        Dim ldtbCarpetas = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSele, lstrOrden, "")
        For Each ldrwCar As DataRow In ldtbCarpetas.Rows
            Dim lstrFiltro = StrCampoCarpeta & " = " &
                    ldrwCar(StrCampoCarpeta) & " AND " &
                    ClsIdServicio_NovShr.SstrNombreCampoBd & " = 0"
            Dim ldtbFacts = ClsPanorama.FdtbDataTable(lstrTablaF, lstrCampSeleF,
                    lstrOrdenF, lstrFiltro)
            Dim lentCanFilas As Integer = ldtbFacts.Rows.Count
            If lentCanFilas > 0 Then
                Dim lentIdFacIni As Integer = ClsPanorama.FobjValorCampo(ldtbFacts.Rows(0)(
                    ClsIdFactura_NovEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                Dim lentIdFacFin As Integer = ClsPanorama.FobjValorCampo(
                        ldtbFacts.Rows(lentCanFilas - 1)(ClsIdFactura_NovEnt.SstrNombreCampoBd),
                        EnuTipoValor.enuInteger)
                SReparaeNovCar(ldrwCar(StrCampoCarpeta), lentIdFacIni, lentIdFacFin)
            End If
        Next
    End Sub
    Private Shared Sub SReparaeNovCar(ashrIdCarpeta As Short, aentIdFacIni As Integer,
            aentIdFacFin As Integer)
        Dim lstrTabla1 = ClsNovedad.SstrNombreTabla, lentRegAfe As Integer
        Dim lstrTabla2 = ClsItemFactura.SstrNombreTabla
        Dim lstrCarpeta = StrCampoCarpeta, lentIdFacFin As Integer
        Dim lstbArr As New StringBuilder
        Do While lentIdFacFin < aentIdFacFin
            lentIdFacFin = If(lentIdFacFin = 0, aentIdFacIni + 200, lentIdFacFin + 200)
            With lstbArr
                lstbArr.Clear()
                .Append("UPDATE ").Append(lstrTabla1).Append(" AS N JOIN ").Append(lstrTabla2)
                .Append(" AS I ON N.").Append(lstrCarpeta).Append(" = I.").Append(lstrCarpeta)
                .Append(" AND N.").Append(ClsPrefijoFact_NovStr.SstrNombreCampoBd).Append(" = I.")
                .Append(ClsPrefijo_ItemFactStr.SstrNombreCampoBd).Append(" AND N.")
                .Append(ClsIdFactura_NovEnt.SstrNombreCampoBd).Append(" = I.")
                .Append(ClsIdFactura_ItemFactEnt.SstrNombreCampoBd).Append(" AND N.")
                .Append(ClsIdItemFacturaShr.SstrNombreCampoBd).Append(" = I.")
                .Append(ClsIdItemFacturaShr.SstrNombreCampoBd).Append(" SET N.")
                .Append(ClsIdAno_NovShr.SstrNombreCampoBd).Append(" = I.")
                .Append(ClsIdAno_NovShr.SstrNombreCampoBd).Append(", N.")
                .Append(ClsIdServicio_NovShr.SstrNombreCampoBd).Append(" = I.")
                .Append(ClsIdServicio_ItemFactShr.SstrNombreCampoBd).Append(" WHERE N.")
                .Append(lstrCarpeta).Append(" = ").Append(ashrIdCarpeta).Append(" AND N.")
                .Append(ClsIdServicio_NovShr.SstrNombreCampoBd).Append(" = 0").Append(" AND N.")
                .Append(ClsIdFactura_NovEnt.SstrNombreCampoBd).Append(" BETWEEN ")
                .Append(lentIdFacFin - 200).Append(" AND ").Append(lentIdFacFin)
            End With
            lentRegAfe = GobjPanDat.SEjecuteSentenciaSql(lstbArr.ToString())
            GobjPanDat.SEjecuteSentenciaSql("FLUSH TABLE " & lstrTabla1)
        Loop
    End Sub
#End Region
#Region "15_32_389"
    Private Shared Sub SAct_15_32_389()
        ' Actualiza el estado de documento electrónico en todos los documentos a 10 = No documento 
        ' electrónico
        SActualiceEstadoFacs()
        SActualiceEstadoNotas(ClsNotaDb.SstrNombreTabla)
        SActualiceEstadoNotas(ClsNotaCr.SstrNombreTabla)
        SActualiceEstadoNotas(ClsNotaCon.SstrNombreTabla)
        SActualiceEstadoNotas(ClsNotaReversionCr.SstrNombreTabla)
    End Sub
    Private Shared Sub SActualiceEstadoFacs()
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrExpSql = "UPDATE " & lstrTabla & " SET " & ClsIdEstadoEDocEnt.SstrNombreCampoBd &
                " = " & EnuEstadoEDoc.EnuNoEDoc & " WHERE " & ClsCUDocStr.SstrNombreCampoBd &
                " = '*****'"
        Dim lentRegAfe = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        lstrExpSql = "UPDATE " & lstrTabla & " SET " & ClsIdEstadoEDocEnt.SstrNombreCampoBd &
                " = " & EnuEstadoEDoc.EnuNoEDoc & " WHERE " & ClsCUDocStr.SstrNombreCampoBd &
                " <> '*****'" & " AND " & ClsCUFEStr.SstrNombreCampoBd & " = ''"
        lentRegAfe = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        lstrExpSql = "UPDATE " & lstrTabla & " SET " & ClsCUDocStr.SstrNombreCampoBd &
                " = '' WHERE " & ClsCUDocStr.SstrNombreCampoBd & " = '*****'"
        lentRegAfe = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
    End Sub
    Private Shared Sub SActualiceEstadoNotas(astrNomTabla As String)
        Dim lstrExpSql = "UPDATE " & astrNomTabla & " SET " & ClsIdEstadoEDocEnt.SstrNombreCampoBd &
                " = " & EnuEstadoEDoc.EnuNoEDoc & " WHERE " & ClsCUDocStr.SstrNombreCampoBd &
                " = '*****'"
        Dim lentRegAfe = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        lstrExpSql = "UPDATE " & astrNomTabla & " SET " & ClsCUDocStr.SstrNombreCampoBd &
                " = '' WHERE " & ClsCUDocStr.SstrNombreCampoBd & " = '*****'"
        lentRegAfe = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
    End Sub
#End Region
#Region "16_33_391"
    Private Shared Sub SAct_16_33_391()
        Dim lstrTablaPri = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec = ClsCliente.SstrNombreTabla
        Dim lstrCampPri As String() = {StrCampoCarpeta,
                StrCampoCentroUtil, ClsIdPredioStr.SstrNombreCampoBd,
                "IdTerceroPropietario"}
        Dim lstrCampSec As String() = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta,
                StrCampoCentroUtil, "IdTerceroPropietario"}
        Dim lstrCampRelSec As String() = {StrCampoCarpeta,
                StrCampoCentroUtil, ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{"P." & StrCampoCarpeta, "ASC"},
                {"P." & StrCampoCentroUtil, "ASC"},
                {ClsIdPredioStr.SstrNombreCampoBd, "ASC"}}
        Dim ldtpOri = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCampPri,
                lstrTablaSec, lstrCampSec, lstrCampRelPri, lstrCampRelSec,
                lstrOrden, False, String.Empty, {})
        Dim lshrIdCar As Short, lshrIdCenUtil As Short, lstrIdPredio As String
        Dim ldblIdProp As Double, lstrNombre As String
        Const ldblPorcientoProp = 1
        Dim lstrTablaDes = ClsPropietario.SstrNombreTabla
        Dim lcolNombresCamp As New Collection From {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdPredio_PropStr.SstrNombreCampoBd,
                ClsIdCliente_PropDbl.SstrNombreCampoBd,
                ClsPorcentajePartiDbl.SstrNombreCampoBd,
                ClsNombreCompleto_PropStr.SstrNombreCampoBd}
        Dim lcolDatos As New Collection
        Dim lentRes As Integer
        For Each ldrwOri As DataRow In ldtpOri.Rows
            lshrIdCar = ClsPanorama.FobjValorCampo(ldrwOri(
                    StrCampoCarpeta), EnuTipoValor.enuShort)
            lshrIdCenUtil = ClsPanorama.FobjValorCampo(ldrwOri(
                    StrCampoCentroUtil), EnuTipoValor.enuShort)
            lstrIdPredio = ClsPanorama.FobjValorCampo(ldrwOri(
                    ClsIdPredioStr.SstrNombreCampoBd), EnuTipoValor.enuString)
            ldblIdProp = ClsPanorama.FobjValorCampo(ldrwOri("IdTerceroPropietario"),
                    EnuTipoValor.enuDouble)
            lstrNombre = ClsPanorama.FobjValorCampo(ldrwOri(
                    ClsNombreCompletoStr.SstrNombreCampoBd), EnuTipoValor.enuString)
            lcolDatos.Add(lshrIdCar, StrCampoCarpeta)
            lcolDatos.Add(lshrIdCenUtil, StrCampoCentroUtil)
            lcolDatos.Add(lstrIdPredio, ClsIdPredio_PropStr.SstrNombreCampoBd)
            lcolDatos.Add(ldblIdProp, ClsIdCliente_PropDbl.SstrNombreCampoBd)
            lcolDatos.Add(ldblPorcientoProp, ClsPorcentajePartiDbl.SstrNombreCampoBd)
            lcolDatos.Add(lstrNombre, ClsNombreCompleto_PropStr.SstrNombreCampoBd)
            lcolDatos = New Collection From {
                    {lshrIdCar, StrCampoCarpeta},
                    {lshrIdCenUtil, StrCampoCentroUtil},
                    {lstrIdPredio, ClsIdPredio_PropStr.SstrNombreCampoBd},
                    {ldblIdProp, ClsIdCliente_PropDbl.SstrNombreCampoBd},
                    {ldblPorcientoProp, ClsPorcentajePartiDbl.SstrNombreCampoBd},
                    {lstrNombre, ClsNombreCompleto_PropStr.SstrNombreCampoBd}}
            lentRes = GobjPanDat.SInserteRegistro(lstrTablaDes, lcolNombresCamp,
                    lcolDatos)
            lcolDatos.Clear()
        Next
        Dim lstrExpSql = "ALTER TABLE OriPredios DROP COLUMN IdTerceroPropietario"
        GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        lstrExpSql = "ALTER TABLE OriPredios DROP COLUMN CantPropietarios"
        GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
    End Sub
#End Region
#Region "16_33_396()"
    Private Shared Sub SAct_16_33_396()
        Dim lstrExpSql = "UPDATE " & ClsPredio.SstrNombreTabla & " SET " &
                ClsFactorPonderaCPDbl.SstrNombreCampoBd & " = 1"
        Dim lentRegAfe = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
    End Sub
#End Region
#Region "16_35_400"
    Private Shared Sub SAct_16_35_400()
        SLimpieBaseDatos()
        '
        Dim larlIdCarpetas As ArrayList = FarlCarpetas(True), lblnModuloPorServicio As Boolean
        For Each lshrIdCar As Short In larlIdCarpetas
            Dim larlAnos As ArrayList = FarlAnosReparar(lshrIdCar)
            If larlAnos.Count > 2 Then
                larlAnos.RemoveAt(0)
            End If
            For Each lstrAno As String In larlAnos
                lblnModuloPorServicio = lstrAno.Split(",")(2)
                If lblnModuloPorServicio Then
                    SReparePresAnos(lshrIdCar, larlAnos)
                Else
                    SActualiceModulosACero(lshrIdCar, lstrAno)
                    SActualiceSectoresACero(lshrIdCar, lstrAno)
                End If
            Next
        Next
    End Sub
    Private Shared Function FblnExisteCol(astrNomTabla As String,
            astrNomColumna As String) As Boolean
        Dim lstrExpSql = "Select column_name FROM information_schema.columns WHERE " &
                "table_schema = 'Panorama_Net' AND table_name = '" & astrNomTabla &
                "' AND column_name = '" & astrNomColumna & "'"
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrExpSql)
        Return ldtbRes.Rows.Count > 0
    End Function
    Private Shared Sub SReparePresAnos(ashrIdCarpeta As Short, aarlAnos As ArrayList)
        Dim lshrIdCenUtil As Short, lshrIdAno As Short
        For Each lstrAno As String In aarlAnos
            lshrIdCenUtil = lstrAno.Split(",")(0)
            lshrIdAno = lstrAno.Split(",")(1)
            SReparePresAno(ashrIdCarpeta, lshrIdCenUtil, lshrIdAno)
            SRepareSectoresAno(ashrIdCarpeta, lshrIdCenUtil, lshrIdAno)
            SRepareModulosAno(ashrIdCarpeta, lshrIdCenUtil, lshrIdAno)
        Next
    End Sub
    Private Shared Sub SReparePresAno(ashrIdCarpeta As Short, ashrIdCentroUtil As Short,
                ashrIdAno As Short)
        Dim lstrTablaPri = ClsItemProgramaFact.SstrNombreTabla
        Dim lstrTablaSec = ClsServicio.SstrNombreTabla
        Dim lstrCamSelPri As String = "SUM(" &
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd & " * 12) AS TOT"
        Dim lstrFiltro = "P." & ClsIdCarpetaShr.SstrNombreCampoBd & " = " & ashrIdCarpeta &
                " AND P." & ClsIdCentroUtilShr.SstrNombreCampoBd & " = " & ashrIdCentroUtil &
                " AND P." & ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " &
                ashrIdAno & " AND " & ClsEsAjusteBln.SstrNombreCampoBd & " = FALSE"
        Dim lstrExpSql = "SELECT " & lstrCamSelPri & " FROM " & lstrTablaPri & " AS P INNER JOIN " &
                lstrTablaSec & " AS S ON P." & ClsIdCarpetaShr.SstrNombreCampoBd & " = S." &
                ClsIdCarpetaShr.SstrNombreCampoBd & " AND P." & ClsIdCentroUtilShr.SstrNombreCampoBd &
                " = S." & ClsIdCentroUtilShr.SstrNombreCampoBd & " AND P." &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = S." &
                ClsIdAno_ServicioShr.SstrNombreCampoBd & " WHERE " & lstrFiltro
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrExpSql)
        For Each ldrwPres As DataRow In ldtbRes.Rows
            SRepareAno(ashrIdCarpeta, ashrIdCentroUtil, ashrIdAno, ldrwPres)
        Next
    End Sub
    Private Shared Sub SRepareAno(ashrIdCarpeta As Short, ashrIdCentroUtil As Short,
                ashrIdAno As Short, adrwPres As DataRow)
        Dim ldecPres = ClsPanorama.FobjValorCampo(adrwPres("TOT"), EnuTipoValor.enuDecimal)
        Dim lstrSql = "UPDATE " & ClsAno.SstrNombreTabla & " SET " &
                ClsValorPres_AnoDec.SstrNombreCampoBd & " = " & ldecPres & " WHERE " &
                StrCampoCarpeta & " = " & ashrIdCarpeta & " AND " &
                StrCampoCentroUtil & " = " & ashrIdCentroUtil & " AND " &
                ClsIdAnoShr.SstrNombreCampoBd & " = " & ashrIdAno
        Dim lentCan = GobjPanDat.SEjecuteSentenciaSql(lstrSql)
    End Sub
    Private Shared Sub SRepareModulosAno(ashrIdCarpeta As Short, ashrIdCentroUtil As Short,
                ashrIdAno As Short)
        Dim lstrTabla = ClsSectorModuloServicio.SstrNombreTabla
        Dim lstrCampSel As String() = {ClsIdServicio_SectorModuloServicioShr.SstrNombreCampoBd,
                ClsIdModulo_SectorModuloServicioShr.SstrNombreCampoBd,
                "SUM(" & ClsValor_SectorModuloServicioDec.SstrNombreCampoBd & ") AS TOT"}
        Dim lstrOrden As String(,) = {{ClsIdModulo_SectorModuloServicioShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFIltro = StrCampoCarpeta & " = " & ashrIdCarpeta & " AND " &
                StrCampoCentroUtil & " = " & ashrIdCentroUtil & " AND " &
                ClsIdAno_SectorModuloServicioShr.SstrNombreCampoBd & " = " & ashrIdAno
        Dim lstrGroup As String() = {ClsIdServicio_SectorModuloServicioShr.SstrNombreCampoBd,
                ClsIdModulo_SectorModuloServicioShr.SstrNombreCampoBd}
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSel, lstrOrden, lstrFIltro,
                True, lstrGroup)
        Dim ldecPres As Decimal, lshrIdServicio As Short, lshrIdModulo As Short
        Dim lstrFiltroUp As String, lstrExpSql As String
        For Each ldrwRes As DataRow In ldtbRes.Rows
            lshrIdServicio = ClsPanorama.FobjValorCampo(ldrwRes(
                    ClsIdServicio_SectorModuloServicioShr.SstrNombreCampoBd),
                    EnuTipoValor.enuShort)
            lshrIdModulo = ClsPanorama.FobjValorCampo(ldrwRes(
                    ClsIdModulo_ModuloServicioShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
            ldecPres = ClsPanorama.FobjValorCampo(ldrwRes("TOT"), EnuTipoValor.enuDecimal)
            lstrFiltroUp = StrCampoCarpeta & " = " & ashrIdCarpeta &
                    " AND " & StrCampoCentroUtil & " = " & ashrIdCentroUtil &
                    " AND " & ClsIdAno_ModuloServicioShr.SstrNombreCampoBd & " = " & ashrIdAno &
                    " AND " & ClsIdServicio_ModuloServicioShr.SstrNombreCampoBd & " = " &
                     lshrIdServicio & " AND " & ClsIdModulo_ModuloServicioShr.SstrNombreCampoBd &
                    " = " & lshrIdModulo
            lstrExpSql = "UPDATE " & ClsModuloServicio.SstrNombreTabla & " SET " &
                ClsValorPres_ModuloServicioDec.SstrNombreCampoBd & " = " & ldecPres &
                " WHERE " & lstrFiltroUp
            Dim lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        Next
    End Sub
    Private Shared Sub SRepareSectoresAno(ashrIdCarpeta As Short, ashrIdCentroUtil As Short,
            ashrIdAno As Short)
        Dim lstrTablaPri = ClsItemProgramaFact.SstrNombreTabla
        Dim lstrTablaSec = ClsPredio.SstrNombreTabla
        Dim lstrTablaTer = ClsServicio.SstrNombreTabla
        Dim lstrExpSql = "SELECT S." & ClsIdSector_PredioShr.SstrNombreCampoBd & ", SUM(P." &
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd & " * P." &
                ClsCantidadPeriodosShr.SstrNombreCampoBd & ") AS TOT FROM " &
                lstrTablaPri & " AS P INNER JOIN " & lstrTablaSec & " AS S ON P." &
                StrCampoCarpeta & " = S." & StrCampoCarpeta &
                " AND P." & StrCampoCentroUtil & " = S." &
                StrCampoCentroUtil & " AND P." &
                ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd & " = S." &
                ClsIdPredioStr.SstrNombreCampoBd & " INNER JOIN " & lstrTablaTer &
                " AS SE ON P." & StrCampoCarpeta & " = SE." &
                StrCampoCarpeta & " AND P." &
                StrCampoCentroUtil & " = SE." &
                StrCampoCentroUtil & " AND P." &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = SE." &
                ClsIdAno_ServicioShr.SstrNombreCampoBd & " AND P." &
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " = SE." &
                ClsIdServicioShr.SstrNombreCampoBd & " WHERE P." &
                StrCampoCarpeta & " = " & ashrIdCarpeta & " AND P." &
                StrCampoCentroUtil & " = " & ashrIdCentroUtil & " AND P." &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & ashrIdAno & " AND " &
                ClsEsAjusteBln.SstrNombreCampoBd & " = FALSE GROUP BY " &
                ClsIdSector_PredioShr.SstrNombreCampoBd & " ORDER BY " &
                ClsIdSector_PredioShr.SstrNombreCampoBd & " ASC"
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrExpSql)
        Dim lstrTabla = ClsSectorModuloServicio.SstrNombreTabla
        Dim ldecVlrSector As Decimal, lstrIdSector As String, lstrFiltro As String
        For Each ldrwRes As DataRow In ldtbRes.Rows
            ldecVlrSector = ClsPanorama.FobjValorCampo(ldrwRes("TOT"), EnuTipoValor.enuDecimal)
            lstrIdSector = ClsPanorama.FobjValorCampo(ldrwRes(
                    ClsIdSector_PredioShr.SstrNombreCampoBd), EnuTipoValor.enuString)
            lstrFiltro = StrCampoCarpeta & " = " & ashrIdCarpeta &
                    " AND " & StrCampoCentroUtil & " = " & ashrIdCentroUtil &
                    " AND " & ClsIdAno_SectorModuloServicioShr.SstrNombreCampoBd & " = " &
                    ashrIdAno & " AND " & ClsIdSector_SectorModuloServicioShr.SstrNombreCampoBd &
                    " = " & lstrIdSector
            lstrExpSql = "UPDATE " & lstrTabla & " SET " &
                    ClsValor_SectorModuloServicioDec.SstrNombreCampoBd & " = " & ldecVlrSector &
                    " WHERE " & lstrFiltro
            Dim lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        Next
    End Sub
    Private Shared Function FarlAnosReparar(ashrIdCarpeta As Short) As ArrayList
        Dim larlAnos As New ArrayList
        Dim lstrtabla = ClsAno.SstrNombreTabla
        Dim lshrIdAnoIni As Short = Today.Year - 2
        Dim lstrCampSel As String() = {StrCampoCentroUtil,
                ClsIdAnoShr.SstrNombreCampoBd, ClsModuloPorServicioBln.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{StrCampoCentroUtil, "ASC"},
                {ClsIdAnoShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrCampoCarpeta & " = " & ashrIdCarpeta & " And " &
               ClsIdAnoShr.SstrNombreCampoBd & " >= " & lshrIdAnoIni & " And " &
               ClsTipoCalculoCuotaByt.SstrNombreCampoBd & " = " &
               EnuTipoBaseCalculo.EnuCuotaAnterior
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrtabla, lstrCampSel, lstrOrden, lstrFiltro)
        Dim lshrIdCenUti As Short, lshrIdAno As Short, lstrAno As String,
                lblnModPorServicio As Boolean
        For Each ldrwAno As DataRow In ldtbRes.Rows
            lshrIdCenUti = ClsPanorama.FobjValorCampo(ldrwAno(StrCampoCentroUtil),
                    EnuTipoValor.enuShort)
            lshrIdAno = ClsPanorama.FobjValorCampo(ldrwAno(ClsIdAnoShr.SstrNombreCampoBd),
                    EnuTipoValor.enuShort)
            lblnModPorServicio = ClsPanorama.FobjValorCampo(ldrwAno(
                    ClsModuloPorServicioBln.SstrNombreCampoBd), EnuTipoValor.enuBoolean)
            lstrAno = lshrIdCenUti & "," & lshrIdAno & "," & lblnModPorServicio
            If Not larlAnos.Contains(lstrAno) Then
                larlAnos.Add(lstrAno)
            End If
        Next
        Return larlAnos
    End Function
    Private Shared Function FarlCarpetas(ablnSoloActivas As Boolean) As ArrayList
        Dim larlCar As New ArrayList
        Dim lstrTabla = ClsCarpeta.SstrNombreTabla
        Dim lstrCampSel As String() = {"idcarpeta", "estaactiva"}
        Dim lstrOrden As String(,) = {{"idcarpeta", "ASC"}}
        Dim lstrFiltro = ""
        If ablnSoloActivas Then
            lstrFiltro = " estaactiva = True"
        End If
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSel, lstrOrden, lstrFiltro)
        Dim lshrIdCar As Short
        For Each ldrwRes As DataRow In ldtbRes.Rows
            lshrIdCar = ClsPanorama.FobjValorCampo(ldrwRes("idcarpeta"),
                    EnuTipoValor.enuShort)
            larlCar.Add(lshrIdCar)
        Next
        Return larlCar
    End Function
    Private Shared Sub SActualiceModulosACero(ashrIdCarpeta As Short, aarlAno As String)
        Dim lshrIdCenUtil = aarlAno.Split(",")(0)
        Dim lshrIdAno = aarlAno.Split(",")(1)
        Dim lstrTabla = ClsModuloServicio.SstrNombreTabla
        Dim lstrFiltro = " WHERE " & StrCampoCarpeta & " = " & ashrIdCarpeta &
                " AND " & StrCampoCentroUtil & " = " & lshrIdCenUtil &
                " AND " & ClsIdAno_ModuloServicioShr.SstrNombreCampoBd & " = " & lshrIdAno
        Dim lstrAcc = " SET " & ClsValorPres_ModuloServicioDec.SstrNombreCampoBd & " = 0"
        Dim lstrExpSql = "UPDATE " & lstrTabla & lstrAcc & lstrFiltro
        Dim lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
    End Sub
    Private Shared Sub SActualiceSectoresACero(ashrIdCarpeta As Short, aarlAno As String)
        Dim lshrIdCenUtil = aarlAno.Split(",")(0)
        Dim lshrIdAno = aarlAno.Split(",")(1)
        Dim lstrTabla = ClsSectorModuloServicio.SstrNombreTabla
        Dim lstrFiltro = " WHERE " & StrCampoCarpeta & " = " & ashrIdCarpeta &
                " AND " & StrCampoCentroUtil & " = " & lshrIdCenUtil &
                " AND " & ClsIdAno_ModuloServicioShr.SstrNombreCampoBd & " = " & lshrIdAno
        Dim lstrAcc = " SET " & ClsValor_SectorModuloServicioDec.SstrNombreCampoBd & " = 0"
        Dim lstrExpSql = "UPDATE " & lstrTabla & lstrAcc & lstrFiltro
        Dim lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
    End Sub
#End Region
#Region "16_35_401"
    Private Shared Sub SAct_17_36_402()
        Dim lentRes As Integer, lstrExpsql As String
        If FblnExisteCol("panaplicaciones", "cantidadlicenciada") Then
            lstrExpsql = "ALTER TABLE panaplicaciones DROP COLUMN cantidadlicenciada"
            lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpsql)
        End If
        If FblnExisteCol("panaplicaciones", "FechaLicenciamiento") Then
            lstrExpsql = "ALTER TABLE panaplicaciones DROP COLUMN FechaLicenciamiento"
            lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpsql)
        End If
    End Sub
#End Region
#Region "17_36_405"
    Private Shared Sub SAct_17_36_405()
        Dim ldtbCenUtil = FdtbCentrosUtil()
        Dim lblnCausaRC As Boolean = False
        For Each ldrwCenUtil As DataRow In ldtbCenUtil.Rows
            SActuServiciosCenUtil(ldrwCenUtil, lblnCausaRC)
        Next
        SLimpieBaseDatos()
        SReduzcaLog()
    End Sub
    Private Shared Sub SActuServiciosCenUtil(adrwCenUtil As DataRow, ablnCausaRC As Boolean)
        Dim lstrTabla = ClsServicio.SstrNombreTabla
        Dim lstrCamSel As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsIdAno_ServicioShr.SstrNombreCampoBd, ClsIdServicioShr.SstrNombreCampoBd,
                "CausaMora", "CausaMesCompleto", "CausaUltimoDia"}
        Dim lstrOrden As String(,) = {{StrCampoCarpeta, "ASC"}, {StrCampoCentroUtil, "ASC"},
                {ClsIdAno_ServicioShr.SstrNombreCampoBd, "ASC"},
                {ClsIdServicioShr.SstrNombreCampoBd, "ASC"}}
        Dim lshrIdCarpeta As Short = ClsPanorama.FobjValorCampo(adrwCenUtil(StrCampoCarpeta),
                EnuTipoValor.enuShort)
        Dim lshrIdCenUtil As Short = ClsPanorama.FobjValorCampo(adrwCenUtil(StrCampoCentroUtil),
                EnuTipoValor.enuShort)
        Dim lstrFiltro = StrCampoCarpeta & " = " & lshrIdCarpeta & " AND " & StrCampoCentroUtil &
                " = " & lshrIdCenUtil
        Dim lshrIdAno As Short, lshrIdSer As Short, lblnCM As Boolean, lblnCFM As Boolean,
                lblnCUD As Boolean
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSel, lstrOrden, lstrFiltro)
        For Each ldrwRes As DataRow In ldtbRes.Rows
            lshrIdAno = ClsPanorama.FobjValorCampo(ldrwRes(
                ClsIdAno_ServicioShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
            lshrIdSer = ClsPanorama.FobjValorCampo(ldrwRes(
                ClsIdServicioShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
            lblnCM = ClsPanorama.FobjValorCampo(ldrwRes("CausaMora"), EnuTipoValor.enuBoolean)
            lblnCFM = ClsPanorama.FobjValorCampo(ldrwRes("CausaMesCompleto"),
                    EnuTipoValor.enuBoolean)
            lblnCUD = ClsPanorama.FobjValorCampo(ldrwRes("CausaUltimoDia"), EnuTipoValor.enuBoolean)
            SActualiceSer(lshrIdCarpeta, lshrIdCenUtil, lshrIdAno, lshrIdSer, lblnCM, lblnCFM,
                    lblnCUD, ablnCausaRC)
        Next
    End Sub
    Private Shared Sub SActualiceSer(ashrIdCar As Short, ashrIdcenutil As Short, ashrIdAno As Short,
            ashrIdSer As Short, ablnCausa As Boolean, ablnCausaFM As Boolean, ablnCausaUD As Boolean,
            ablnCausaRC As Boolean)
        Dim lenuModoCM As EnuModoCausaMora
        If ablnCausa Then
            If ablnCausaFM Then
                lenuModoCM = EnuModoCausaMora.EnuFinMes
            ElseIf ablnCausaUD Then
                lenuModoCM = EnuModoCausaMora.EnuUltimoDia
            ElseIf ablnCausaRC Then
                lenuModoCM = EnuModoCausaMora.EnuAlReciboCaja
            Else
                lenuModoCM = EnuModoCausaMora.EnuEnFecha
            End If
        Else
            lenuModoCM = EnuModoCausaMora.EnuNoCausa
        End If
        Dim lstrTabla = ClsServicio.SstrNombreTabla
        Dim lstrFiltro = StrCampoCarpeta & " = " & ashrIdCar & " AND " & StrCampoCentroUtil & " = " &
                ashrIdcenutil & " AND " & ClsIdAno_ServicioShr.SstrNombreCampoBd & " = " &
                ashrIdAno & " AND " & ClsIdServicioShr.SstrNombreCampoBd & " = " & ashrIdSer
        Dim lstrExpSql = "UPDATE " & lstrTabla & " SET " &
                ClsModoCausaInteresesByt.SstrNombreCampoBd & " = " & lenuModoCM & " WHERE " &
                lstrFiltro
        Dim lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
    End Sub
    Private Shared Sub SReduzcaLog()
        Dim lstrTabla = ClsLogApp.SstrNombreTabla
        Dim ldtmFechaDesde As Date = Date.Today.AddYears(-2)
        Dim lstrFecha = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaDesde.ToString()) & "'"
        Dim lstrExpSql = "DELETE FROM " & lstrTabla & " WHERE " &
                ClsFechaCreacionDtm.SstrNombreCampoBd & " < " & lstrFecha
        Dim lentNoUsado = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
    End Sub
    Private Shared Function FdtbCentrosUtil() As DataTable
        Dim lstrTabla = ClsCentroUtilidad.SstrNombreTabla
        Dim lstrCampSel As String() = {ClsIdCarpetaCenUtilShr.SstrNombreCampoBd,
                ClsIdCentroUtilShr.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{ClsIdCarpetaCenUtilShr.SstrNombreCampoBd, "ASC"},
                {ClsIdCentroUtilShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ""
        Dim ldtbCenUtil = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSel, lstrOrden, lstrFiltro)
        Return ldtbCenUtil
    End Function
#End Region
#Region "17_36_406"
    Private Shared Sub SAct_17_38_409()
        SLimpieBaseDatos()
        SElimineModulosSer
    End Sub
#End Region
#End Region
#Region "Revisar integridad novedades y corregir novedades"
    Friend Shared Sub SCorrijaProblemasNovedades(adtmFechaFin As Date)
        Dim lstrPref As String, lentIdFac As Integer, lobjFac As New ClsFactura()
        Dim lblnHayErrInt = False
        Dim lobjValorLlave As Object()
        SInicialiceArchivos()
        Dim ldtbFacs = FdtbFacsAnalizar(adtmFechaFin)
        For Each ldrwFac As DataRow In ldtbFacs.Rows
            lstrPref = ClsPanorama.FobjValorCampo(ldrwFac(
                    ClsPrefijo_FactStr.SstrNombreCampoBd), EnuTipoValor.enuString)
            lentIdFac = ClsPanorama.FobjValorCampo(ldrwFac(
                    ClsIdFacturaEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFac}
            lobjFac.SAbra(lobjValorLlave)
            SCompruebeNovsMora(lblnHayErrInt, lobjFac)
            SCompruebeNovsCap(lblnHayErrInt, lobjFac)
        Next
        If lblnHayErrInt Then
            Process.Start("notepad.exe", GstrTrayDatPrg & "ReporteIntegridad.txt")
        End If
    End Sub
    ''' <summary>
    ''' Devuelve un datatable con la identificación de las facturas canceladas según el encabezado.
    ''' Es decir que los debitos i los creditos en la tabla OriFactura son iguales!
    ''' </summary>
    ''' <param name="adtmFechaFin"></param>
    ''' <returns></returns>
    Private Shared Function FdtbFacsAnalizar(adtmFechaFin As Date) As DataTable
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrFecha = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmFechaFin.ToString()) & "'"
        Dim lstrCampSele As String() = {ClsPrefijo_FactStr.SstrNombreCampoBd,
                ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsFechaFacturaDtm.SstrNombreCampoBd & " < " & lstrFecha & " AND " &
                ClsDebitos_FactDec.SstrNombreCampoBd & " = " & ClsCreditos_FactDec.SstrNombreCampoBd
        Dim lstrOrden As String(,) = {{ClsPrefijo_FactStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}}
        Dim ldtbFacs = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSele, lstrOrden, lstrFiltro)
        Return ldtbFacs
    End Function
    Private Shared Sub SCompruebeNovsMora(ByRef ablnHayErr As Boolean, aobjFac As ClsFactura)
        Dim ldecVlrDb As Decimal, ldecVlrCr As Decimal, ldecDif As Decimal
        Dim lstrDif As String() = {}, i = -1, lstrError As String, lblnHayError = False
        For Each lobjItemFac As ClsItemFactura In aobjFac.ColItemsFactura
            ldecVlrDb = FDecDebitosMora(lobjItemFac)
            ldecVlrCr = FDecCreditosMora(lobjItemFac)
            ldecDif = ldecVlrDb - ldecVlrCr
            If ldecDif <> 0 Then
                lstrError = "Factura nro " & lobjItemFac.ObjMiFactura.StrIdObjeto &
                        " Item Factura = " & lobjItemFac.ObjIdItemFacturaShr.ObjValorPro & " Db = " &
                        ldecVlrDb.ToString & " Cr = " & ldecVlrCr.ToString &
                        " Diferencia = " & ldecDif.ToString()
                SAdicioneErrorIntegridad(lstrError)
                i += 1
                ReDim Preserve lstrDif(i)
                lstrDif(i) = lobjItemFac.ObjIdItemFacturaShr.ToString() & "," & ldecDif.ToString
            End If
        Next
        If lstrDif.Length > 0 Then
            SRepareNovsMora(aobjFac, lstrDif)
            lblnHayError = True
        End If
        ablnHayErr = lblnHayError
    End Sub
    Private Shared Sub SCompruebeNovsCap(ByRef ablnHayErr As Boolean, aobjFac As ClsFactura)
        Dim ldecVlrDb As Decimal, ldecVlrCr As Decimal, ldecDif As Decimal
        Dim lstrDif As String() = {}, i = -1, lstrError As String, lblnHayError = False
        For Each lobjItemFac As ClsItemFactura In aobjFac.ColItemsFactura
            ldecVlrDb = FDecDebitosCap(lobjItemFac)
            ldecVlrCr = FDecCreditosCap(lobjItemFac)
            ldecDif = ldecVlrDb - ldecVlrCr
            If ldecDif <> 0 Then
                lstrError = "Factura nro " & lobjItemFac.ObjMiFactura.StrIdObjeto &
                        " Item Factura = " & lobjItemFac.ObjIdItemFacturaShr.ObjValorPro & " Db = " &
                        ldecVlrDb.ToString & " Cr = " & ldecVlrCr.ToString &
                        " Diferencia capital = " & ldecDif.ToString()
                SAdicioneErrorIntegridad(lstrError)
                i += 1
                ReDim Preserve lstrDif(i)
                lstrDif(i) = lobjItemFac.ObjIdItemFacturaShr.ToString() & "," & ldecDif.ToString
            End If
        Next
        If lstrDif.Length > 0 Then
            SRepareNovsCap(aobjFac, lstrDif)
            lblnHayError = True
        End If
        ablnHayErr = lblnHayError
    End Sub
    Private Shared Function FDecDebitosMora(aobjItemFac As ClsItemFactura) As Decimal
        Dim lenuTipoNov As EnuTipoNov, ldecVlr = 0D
        Dim lshrIdAno As Short = aobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro
        Dim lshrIdSer As Short = aobjItemFac.ObjIdServicio_ItemFactShr.ObjValorPro
        For Each lobjNov As ClsNovedad In aobjItemFac.ColNovedades
            If lobjNov.ObjIdAno_NovShr.ObjValorPro = lshrIdAno AndAlso
                    lobjNov.ObjIdServicio_NovShr.ObjValorPro = lshrIdSer Then
                lenuTipoNov = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
                If lenuTipoNov = EnuTipoNov.EnuDbInt OrElse
                        lenuTipoNov = EnuTipoNov.EnuDbIvaInt OrElse
                        lenuTipoNov = EnuTipoNov.EnuRCrPagoInt OrElse
                        lenuTipoNov = EnuTipoNov.EnuRCrAnApInt OrElse
                        lenuTipoNov = EnuTipoNov.EnuRCrDctoInt OrElse
                        lenuTipoNov = EnuTipoNov.EnuRCrDctoInt Then
                    ldecVlr += lobjNov.ObjValor_NovDec.ObjValorPro
                End If
            End If
        Next
        Return ldecVlr
    End Function
    Private Shared Function FDecCreditosMora(aobjItemFac As ClsItemFactura) As Decimal
        Dim lenuTipoNov As EnuTipoNov, ldecVlr = 0D
        Dim lshrIdAno As Short = aobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro
        Dim lshrIdSer As Short = aobjItemFac.ObjIdServicio_ItemFactShr.ObjValorPro
        For Each lobjNov As ClsNovedad In aobjItemFac.ColNovedades
            If lobjNov.ObjIdAno_NovShr.ObjValorPro = lshrIdAno AndAlso
                    lobjNov.ObjIdServicio_NovShr.ObjValorPro = lshrIdSer Then
                lenuTipoNov = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
                If lenuTipoNov = EnuTipoNov.EnuCrPagoInt OrElse
                        lenuTipoNov = EnuTipoNov.EnuRDbIvaInt OrElse
                        lenuTipoNov = EnuTipoNov.EnuCrAnApInt OrElse
                        lenuTipoNov = EnuTipoNov.EnuCrDctoInt OrElse
                        lenuTipoNov = EnuTipoNov.EnuRDbInt Then
                    ldecVlr += lobjNov.ObjValor_NovDec.ObjValorPro
                End If
            End If
        Next
        Return ldecVlr
    End Function
    Private Shared Function FDecDebitosCap(aobjItemFac As ClsItemFactura) As Decimal
        Dim lenuTipoNov As EnuTipoNov, ldecVlr = 0D
        Dim lshrIdAno As Short = aobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro
        Dim lshrIdSer As Short = aobjItemFac.ObjIdServicio_ItemFactShr.ObjValorPro
        For Each lobjNov As ClsNovedad In aobjItemFac.ColNovedades
            If lobjNov.ObjIdAno_NovShr.ObjValorPro = lshrIdAno AndAlso
                    lobjNov.ObjIdServicio_NovShr.ObjValorPro = lshrIdSer Then
                lenuTipoNov = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
                If lenuTipoNov = EnuTipoNov.EnuDbCap OrElse
                    lenuTipoNov = EnuTipoNov.EnuDbIva OrElse
                    lenuTipoNov = EnuTipoNov.EnuRCrPagoCap OrElse
                    lenuTipoNov = EnuTipoNov.EnuRCrAnApCap OrElse
                    lenuTipoNov = EnuTipoNov.EnuRCrDctoCap Then
                    ldecVlr += lobjNov.ObjValor_NovDec.ObjValorPro
                End If
            End If
        Next
        Return ldecVlr
    End Function
    Private Shared Function FDecCreditosCap(aobjItemFac As ClsItemFactura) As Decimal
        Dim lenuTipoNov As EnuTipoNov, ldecVlr = 0D
        Dim lshrIdAno As Short = aobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro
        Dim lshrIdSer As Short = aobjItemFac.ObjIdServicio_ItemFactShr.ObjValorPro
        For Each lobjNov As ClsNovedad In aobjItemFac.ColNovedades
            If lobjNov.ObjIdAno_NovShr.ObjValorPro = lshrIdAno AndAlso
                    lobjNov.ObjIdServicio_NovShr.ObjValorPro = lshrIdSer Then
                lenuTipoNov = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
                If lenuTipoNov = EnuTipoNov.EnuCrPagoCap OrElse
                        lenuTipoNov = EnuTipoNov.EnuCrAnApCap OrElse
                        lenuTipoNov = EnuTipoNov.EnuCrDctoCap OrElse
                        lenuTipoNov = EnuTipoNov.EnuCrRetFte OrElse
                        lenuTipoNov = EnuTipoNov.EnuCrRetIca OrElse
                        lenuTipoNov = EnuTipoNov.EnuCrRetIva OrElse
                        lenuTipoNov = EnuTipoNov.EnuCrRetCre OrElse
                        lenuTipoNov = EnuTipoNov.EnuRDbCap Then
                    ldecVlr += lobjNov.ObjValor_NovDec.ObjValorPro
                End If
            End If
        Next
        Return ldecVlr
    End Function
    Private Shared Sub SRepareNovsMora(aobjFac As ClsFactura, astrDif() As String)
        Dim lstrMens As String
        If astrDif.Length <> 2 Then
            lstrMens = "Factura " & aobjFac.ObjIdFacturaEnt.ToString() & " con " &
                    astrDif.Length.ToString() & " ítems descuadrados!"
            SAdicioneErrorIntegridad(lstrMens)
            Exit Sub
        End If
        Dim lobjItemFac As ClsItemFactura
        Dim lshrIdItemFac1 = CShort(astrDif(0).Split(",")(0))
        Dim lshrIdItemFac2 = CShort(astrDif(1).Split(",")(0))
        Dim ldecVlr1 = CDec(astrDif(0).Split(",")(1))
        Dim ldecVlr2 = CDec(astrDif(1).Split(",")(1))
        If ldecVlr1 + ldecVlr2 = 0 Then
            lobjItemFac = aobjFac.ColItemsFactura(lshrIdItemFac1)
            SAjusteNovItemMora(lobjItemFac, ldecVlr1)
            lobjItemFac = aobjFac.ColItemsFactura(lshrIdItemFac2)
            SAjusteNovItemMora(lobjItemFac, ldecVlr2)
        End If
    End Sub
    Private Shared Sub SAjusteNovItemMora(aobjItemFac As ClsItemFactura, adecVal As Decimal)
        For Each lobjNov As ClsNovedad In aobjItemFac.ColNovedades
            If lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuCrPagoInt OrElse
                    lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuCrAnApInt OrElse
                    lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuCrDctoInt Then
                lobjNov.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                lobjNov.ObjValor_NovDec.ObjValorPro += adecVal
                lobjNov.SActualice(False)
                aobjItemFac.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                aobjItemFac.ObjCreditos_ItemFactDec.ObjValorPro += adecVal
                aobjItemFac.SActualice(False)
                Exit For
            End If
        Next
    End Sub
    Private Shared Sub SRepareNovsCap(aobjFac As ClsFactura, astrDif() As String)
        Dim lstrMens As String
        If astrDif.Length <> 2 Then
            lstrMens = "Factura " & aobjFac.ObjIdFacturaEnt.ToString() & " con " &
                    astrDif.Length.ToString() & " ítems descuadrados!"
            SAdicioneErrorIntegridad(lstrMens)
            Exit Sub
        End If
        Dim lobjItemFac As ClsItemFactura
        Dim lshrIdItemFac1 = CShort(astrDif(0).Split(",")(0))
        Dim lshrIdItemFac2 = CShort(astrDif(1).Split(",")(0))
        Dim ldecVlr1 = CDec(astrDif(0).Split(",")(1))
        Dim ldecVlr2 = CDec(astrDif(1).Split(",")(1))
        If ldecVlr1 + ldecVlr2 = 0 Then
            lobjItemFac = aobjFac.ColItemsFactura(lshrIdItemFac1)
            SAjusteNovItemCap(lobjItemFac, ldecVlr1)
            lobjItemFac = aobjFac.ColItemsFactura(lshrIdItemFac2)
            SAjusteNovItemCap(lobjItemFac, ldecVlr2)
        End If
    End Sub
    Private Shared Sub SAjusteNovItemCap(aobjItemFac As ClsItemFactura, adecVal As Decimal)
        For Each lobjNov As ClsNovedad In aobjItemFac.ColNovedades
            If lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuCrPagoCap OrElse
                    lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuCrAnApCap OrElse
                    lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuCrDctoCap Then
                lobjNov.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                lobjNov.ObjValor_NovDec.ObjValorPro += adecVal
                lobjNov.SActualice(False)
                aobjItemFac.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                aobjItemFac.ObjCreditos_ItemFactDec.ObjValorPro += adecVal
                aobjItemFac.SActualice(False)
                Exit For
            End If
        Next
    End Sub
    Private Shared Sub SAdicioneErrorIntegridad(astLinea As String)
        Dim lswMapImpo As StreamWriter
        lswMapImpo = ClsPanorama.FswStreamWriterAppend(GstrTrayDatPrg & "ReporteIntegridad.txt")
        lswMapImpo.WriteLine(astLinea)
        lswMapImpo.Close()
    End Sub
    Private Shared Sub SInicialiceArchivos()
        Dim lswMapImpo As StreamWriter
        lswMapImpo = ClsPanorama.FswStreamWriter(GstrTrayDatPrg & "ReporteIntegridad.txt")
        lswMapImpo.Close()
    End Sub
    Private Shared Sub SLimpieBaseDatos()
        Dim lstrExpSql = "DROP TABLE IF EXISTS anos"
        Dim lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        lstrExpSql = "DROP TABLE IF EXISTS anticipos"
        lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        lstrExpSql = "DROP TABLE IF EXISTS clients"
        lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        lstrExpSql = "DROP TABLE IF EXISTS serviciosestadocta"
        lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        If FblnExisteCol("orianticipos", "IdAgrupadorServicios") Then
            lstrExpSql = "ALTER TABLE orianticipos DROP COLUMN IdAgrupadorServicios"
            lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        End If
        If FblnExisteCol("oriestadoscuenta", "IdAgrupadorServicios") Then
            lstrExpSql = "ALTER TABLE oriestadoscuenta DROP COLUMN IdAgrupadorServicios"
            lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        End If
        If FblnExisteCol("orifacturas", "IdAgrupadorServicios") Then
            lstrExpSql = "ALTER TABLE orifacturas DROP COLUMN IdAgrupadorServicios"
            lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        End If
        If FblnExisteCol("oriservicios", "IdAgrupadorServicios") Then
            lstrExpSql = "ALTER TABLE oriservicios DROP COLUMN IdAgrupadorServicios"
            lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        End If
        If FblnExisteCol("oriCentrosUtilidadOriCop", "CausaIntCierreMes") Then
            lstrExpSql = "ALTER TABLE oriCentrosUtilidadOriCop DROP COLUMN CausaIntCierreMes"
            lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        End If
        If FblnExisteCol("oriCentrosUtilidadOriCop", "CausaIntUltimoDia") Then
            lstrExpSql = "ALTER TABLE oriCentrosUtilidadOriCop DROP COLUMN CausaIntUltimoDia"
            lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        End If
        If FblnExisteCol("oriCentrosUtilidadOriCop", "CausaIntAlRecCaja") Then
            lstrExpSql = "ALTER TABLE oriCentrosUtilidadOriCop DROP COLUMN CausaIntAlRecCaja"
            lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        End If
        If FblnExisteCol("oriServicios", "CausaMora") Then
            lstrExpSql = "ALTER TABLE oriServicios DROP COLUMN CausaMora"
            lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        End If
        If FblnExisteCol("oriServicios", "CausaMesCompleto") Then
            lstrExpSql = "ALTER TABLE oriServicios DROP COLUMN CausaMesCompleto"
            lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        End If
        If FblnExisteCol("oriServicios", "CausaUltimoDia") Then
            lstrExpSql = "ALTER TABLE oriServicios DROP COLUMN CausaUltimoDia"
            lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        End If
    End Sub
    Private Shared Sub SElimineModulosSer()
        Dim lstrTabla = ClsModuloServicio.SstrNombreTabla
        Dim lstrFiltro = ClsIdAno_ModuloServicioShr.SstrNombreCampoBd & " = 0 AND " &
                ClsValorPres_ModuloServicioDec.SstrNombreCampoBd & " = 0 "
        Dim lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlEliminar(lstrTabla, lstrFiltro)
        Dim lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
    End Sub
#End Region
End Class