Imports System.Text
Partial Friend Class ClsCliente
#Region "Definiciones"
    Private MdtbIdFrasVivas As DataTable = Nothing
    Private MdtbFras As DataTable = Nothing
    Private MdtbFrasVivas As DataTable = Nothing
    Private McolFacturas As Collection = Nothing
    Private McolFacturasVivas As Collection = Nothing
#End Region
#Region "Tercero del Cliente"
    Private MobjTerceroCliente As ClsTercero = Nothing
    Friend ReadOnly Property ObjTerceroCliente As ClsTercero
        Get
            SAbraTerceroCliente()
            Return MobjTerceroCliente
        End Get
    End Property
    Private Sub SAbraTerceroCliente()
        Dim ldblIdTercero As Double = ObjIdClienteDbl.ObjValorPro
        If MobjTerceroCliente Is Nothing Then
            MobjTerceroCliente = New ClsTercero(EnuTipoObjeto)
        End If
        If ldblIdTercero > 0 AndAlso ObjIdClienteDbl.BlnEsValido Then
            If MobjTerceroCliente.ObjIdTerceroDbl.ObjValorPro <> ldblIdTercero Then
                MobjTerceroCliente.SAbra({ldblIdTercero})
                MobjTerceroCliente.SLeaValores(True)
            End If
        End If
    End Sub
#End Region
#Region "Propiedades del tercero"
    Friend Property ObjApellidoPrimeroStr As ClsApellidoPrimeroStr = ObjTerceroCliente.ObjApellidoPrimeroStr
    Friend Property ObjApellidoSegundoStr As ClsApellidoSegundoStr = ObjTerceroCliente.ObjApellidoSegundoStr
    Friend Property ObjDireccionUnoStr As ClsDireccionUnoStr = ObjTerceroCliente.ObjDireccionUnoStr
    Friend Property ObjDireccionDosStr As ClsDireccionDosStr = ObjTerceroCliente.ObjDireccionDosStr
    Friend Property ObjEmailStr As ClsEmailStr = ObjTerceroCliente.ObjEmailStr
    Friend Property ObjPaginaWebStr As ClsPaginaWebStr = ObjTerceroCliente.ObjPaginaWebStr
    Friend Property ObjCiudadDirShr As ClsCiudadDirShr = ObjTerceroCliente.ObjCiudadDirShr
    Friend Property ObjDepartamentoDirByt As ClsDepartamentoDirByt = ObjTerceroCliente.ObjDepartamentoDirByt
    Friend Property ObjPaisDirStr As ClsPaisDirStr = ObjTerceroCliente.ObjPaisDirStr
    Friend Property ObjTipoDocIdentidadByt As ClsTipoDocIdentidadByt = ObjTerceroCliente.ObjTipoDocIdentidadByt
    Friend Property ObjIdTerceroDbl As ClsIdTerceroDbl = ObjTerceroCliente.ObjIdTerceroDbl
    Friend Property ObjTipoTerceroByt As ClsTipoTerceroByt = ObjTerceroCliente.ObjTipoTerceroByt
    Friend Property ObjNombrePrimeroStr As ClsNombrePrimeroStr = ObjTerceroCliente.ObjNombrePrimeroStr
    Friend Property ObjNombreSegundoStr As ClsNombreSegundoStr = ObjTerceroCliente.ObjNombreSegundoStr
    Friend Property ObjRazonSocialStr As ClsRazonSocialStr = ObjTerceroCliente.ObjRazonSocialStr
    Friend Property ObjTelefonoUnoStr As ClsTelefonoUnoStr = ObjTerceroCliente.ObjTelefonoUnoStr
    Friend Property ObjTelefonoDosStr As ClsTelefonoDosStr = ObjTerceroCliente.ObjTelefonoDosStr
    Friend Property ObjCelularStr As ClsCelularStr = ObjTerceroCliente.ObjCelularStr
    Friend Property ObjCelular2Str As ClsCelular2Str = ObjTerceroCliente.ObjCelular2Str
    Friend Property ObjCodigoPostalStr As ClsCodigoPostalStr = ObjTerceroCliente.ObjCodigoPostalStr
    Friend Function FblnEsValidoEmail() As Boolean
        Dim lblnEsValido = True
        If EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If GobjParametros.BlnEFacAutorizado Then
                lblnEsValido = Not IsNothing(ObjEmailStr.ObjValorPro) AndAlso
                        ObjEmailStr.BlnEsValido AndAlso
                        (ObjEmailStr.ToString.Length > 0)
            End If
        End If
        Return lblnEsValido
    End Function
#End Region
#Region "Nuevo manejo de facturas"
    Private Sub SVacieCompl()
        MdtbIdFrasVivas = Nothing
        MdtbFras = Nothing
        MdtbFrasVivas = Nothing
        McolFacturas = Nothing
        McolFacturasVivas = Nothing
    End Sub
    ''' <summary>
    ''' Devuelve la colección de facturas del cliente de acuerdo a los parámetros recibidos
    ''' </summary>
    ''' <param name="astrIdPrediosAgr">Array de los Predios Agrupadores del Cliente que se tendran
    ''' en cuenta.</param>
    ''' <param name="ashrIdAgrupadorSer">Servicios de las Facturas a tener en cuenta</param>
    ''' <param name="ablnSoloVivas">Indica si solo se tienen en cuenta las facturas cob saldo</param>
    ''' <returns></returns>
    Friend Function FcolFacturas(astrIdPrediosAgr As String(), astrServicios As String(),
            ablnSoloVivas As Boolean) As Collection
        Static lstrIdPreAgrs As String() = Nothing
        Static lstrServicios As String() = Nothing
        Dim lblnConstruir As Boolean
        If ablnSoloVivas Then
            lblnConstruir = McolFacturasVivas Is Nothing OrElse McolFacturasVivas.Count = 0
        Else
            lblnConstruir = McolFacturas Is Nothing
        End If
        If Not lblnConstruir Then
            lblnConstruir = Not FblnStrArrayIguales(lstrIdPreAgrs, astrIdPrediosAgr)
            If Not lblnConstruir Then
                lblnConstruir = FblnStrArrayIguales(lstrServicios, astrServicios)
            End If
        End If
        If lblnConstruir Then
            If ablnSoloVivas Then
                McolFacturasVivas = FcolFact(astrIdPrediosAgr, astrServicios, True)
            Else
                McolFacturas = FcolFact(astrIdPrediosAgr, astrServicios, False)
            End If
            lstrIdPreAgrs = astrIdPrediosAgr
            lstrServicios = astrServicios
        End If
        If ablnSoloVivas Then
            Return McolFacturasVivas
        Else
            Return McolFacturas
        End If
    End Function
    Private Function FcolFact(astrIdPrediosAgr As String(), astrServicios As String(),
             ablnSoloVivas As Boolean) As Collection
        Dim lstrPref As String, lentIdFac As Integer
        Dim lcolFacturas As New Collection, lobjFact As ClsFactura
        If astrIdPrediosAgr.Length > 0 Then
            Dim ldtbIdFacts = FdtbIdFras(astrIdPrediosAgr, astrServicios, ablnSoloVivas)
            For Each ldrwFact As DataRow In ldtbIdFacts.Rows
                lstrPref = ClsPanorama.FobjValorCampo(ldrwFact(
                    ClsPrefijo_ItemFactStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                lentIdFac = ClsPanorama.FobjValorCampo(ldrwFact(
                    ClsIdFactura_ItemFactEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                lobjFact = New ClsFactura()
                lobjFact.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFac})
                Dim lstrKey As String = lobjFact.StrNumeroFactura
                lcolFacturas.Add(lobjFact, lstrKey)
            Next
        End If
        Return lcolFacturas
    End Function
    Friend Function FcolFacturasAuto(astrIdpredAgru As String, adtmFecha As Date) As Collection
        Dim lcolFacAuto As New Collection
        Dim lstrFecha = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmFecha) & "'"
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCamSele = {ClsPrefijo_FactStr.SstrNombreCampoBd, ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ObjValorPro &
                " AND " & ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " = '" &
                astrIdpredAgru & "' AND " & ClsFechaFacturaDtm.SstrNombreCampoBd & " = " &
                lstrFecha & " AND " & ClsIdModoFacturacionByt.SstrNombreCampoBd & " = " &
                EnuModoFacturacionDef.EnuSistema
        Dim lstrOrden = {{"", ""}}
        Dim ldtbFac = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSele, lstrOrden, lstrFiltro)
        Dim lstrPref As String, lentIdFac As Integer
        Dim lobjFac As New ClsFactura()
        For Each ldrwFac As DataRow In ldtbFac.Rows
            lstrPref = ClsPanorama.FobjValorCampo(ldtbFac.Rows(0)(
                ClsPrefijo_FactStr.SstrNombreCampoBd), EnuTipoValor.enuString)
            lentIdFac = ClsPanorama.FobjValorCampo(ldtbFac.Rows(0)(
                ClsIdFacturaEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
            lobjFac.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFac})
            If lobjFac.BlnExiste Then
                lcolFacAuto.Add(lobjFac)
            End If
        Next
        Return lcolFacAuto
    End Function
#Region "Lectura de datos en BD"
    ''' <summary>
    ''' Devuelve una string con el sql que devuelve la identificación de las facturas 
    ''' correspondientes a los argumentos pasado 
    ''' </summary>
    ''' <param name="astrIdPrediosAgru"></param>
    ''' <param name="astrServicios">Array de los servicios tenidos en cuenta. Si es vacio,
    ''' se toman todos los servicios</param>
    ''' <param name="ablnSoloVivas"></param>
    ''' <returns></returns>
    Private Function FdtbIdFras(astrIdPrediosAgru As String(),
            astrServicios As String(), ablnSoloVivas As Boolean) As DataTable
        Dim lstrFiltroSer = FstrFiltroServi(astrServicios)
        Dim lstrFiltro As String
        Dim lstrTablaPri = ClsItemFactura.SstrNombreTabla
        Dim lstrTablaSec = ClsFactura.SstrNombreTabla
        Dim lstrCampPri As String() = {"DISTINCT " & ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
                ClsIdFactura_ItemFactEnt.SstrNombreCampoBd}
        Dim lstrCamSec As String() = {ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd,
                ClsFechaFacturaDtm.SstrNombreCampoBd}
        Dim lstrCamRelPri = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_ItemFactStr.SstrNombreCampoBd, ClsIdFactura_ItemFactEnt.SstrNombreCampoBd}
        Dim lstrCamRelSec = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_FactStr.SstrNombreCampoBd, ClsIdFacturaEnt.SstrNombreCampoBd}
        lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri
        If Not String.IsNullOrEmpty(lstrFiltroSer) Then
            lstrFiltro &= " AND " & lstrFiltroSer
        End If
        lstrFiltro &= " AND " & FstrFiltroPredios(astrIdPrediosAgru) &
                " AND " & ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " &
                ObjIdClienteDbl.ObjValorPro
        If ablnSoloVivas Then
            lstrFiltro &= " AND P." & ClsDebitos_ItemFactDec.SstrNombreCampoBd & " <> P." &
                    ClsCreditos_ItemFactDec.SstrNombreCampoBd
        End If
        Dim lstrOrden = {{ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd, "ASC"},
                {ClsFechaFacturaDtm.SstrNombreCampoBd, "ASC"},
                {ClsPrefijo_ItemFactStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFactura_ItemFactEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri, lstrCampPri,
                lstrTablaSec, lstrCamSec, lstrCamRelPri, lstrCamRelSec, lstrOrden, lstrFiltro,
                Array.Empty(Of String))
        Dim ldtbIdFras = ClsPanorama.FdtbDataTable(lstrSql)
        Return ldtbIdFras
    End Function
    Private Function FstrFiltroPredios(astrIdpreAgr As String()) As String
        Dim lstrFiltro = "(", i As Integer, lblnEsUltimo As Boolean
        For Each lstrIdPreAgr As String In astrIdpreAgr
            i += 1
            If lstrIdPreAgr = GCSTRSINPA Then
                lstrIdPreAgr = ""
            End If
            lstrFiltro &= ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " = '" &
                lstrIdPreAgr & "'"
            lblnEsUltimo = (i = astrIdpreAgr.Length)
            If Not lblnEsUltimo Then
                lstrFiltro &= " OR "
            Else
                lstrFiltro &= ")"
            End If
        Next
        Return lstrFiltro
    End Function
    Private Function FstrFiltroServi(astrServicios As String()) As String
        Dim lstrFiltroSer = String.Empty, lstrFiltro As String
        Dim lstrFiltroAno = String.Empty
        If astrServicios.Length > 0 AndAlso Not astrServicios.Contains("A") Then
            If astrServicios.Contains("0") Then
                lstrFiltroAno = "(" & ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd & " <> 0)"
            End If
            For Each lstrSer As String In astrServicios
                If lstrSer <> "0" AndAlso lstrSer <> "A" Then
                    If String.IsNullOrEmpty(lstrFiltroSer) Then
                        lstrFiltroSer &= "(" & ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd &
                                " = 0 AND ("
                    End If
                    lstrFiltroSer &= ClsIdServicio_ItemFactShr.SstrNombreCampoBd & " = " &
                                lstrSer & " OR "
                End If
            Next
            If lstrFiltroSer.EndsWith(" OR ") Then
                lstrFiltroSer = lstrFiltroSer.Substring(0, lstrFiltroSer.Length - 4) & "))"
            End If
        End If
        If Not String.IsNullOrEmpty(lstrFiltroAno) AndAlso Not String.IsNullOrEmpty(lstrFiltroSer) Then
            lstrFiltro = "(" & lstrFiltroAno & " OR " & lstrFiltroSer & ")"
        Else
            lstrFiltro = lstrFiltroAno & lstrFiltroSer
        End If
        Return lstrFiltro
    End Function
#End Region
#End Region
End Class
