Imports System.IO
Imports System.Text
Friend Class ClsInterfazApolo
    Inherits ClsCBInterfazContableOri
#Region "Definiciones"
    'Variables
    Private MstrNombreArchivo As String = String.Empty
    Private MentIdComprobante As Integer = 0, mentConsecutivo As Integer = 0
    Private ReadOnly MstbLinea As New StringBuilder
    Private MdtbInterfaz As DataTable = Nothing
    Private MdtmFechaFinInterfazAnterior As Date = GCDTMFECHANULA
    Private MentIdUltimoComprobanteInterfazAnterior As Integer = 0, mstrNroInterfaz As String = String.Empty
    Private MstrTipoComprobante As String = String.Empty, mblnLeidoUltimoRegistro = False
#End Region
#Region "Constructores"
    Friend Sub New(aobjRegistro As Object)
        MyBase.New(aobjRegistro)
    End Sub
#End Region
    Friend Overrides Sub SGenerereInterfazContable(ablnFinMes As Boolean, ByRef astrMens As String)
        Dim lblnNohayError As Boolean
        Try
            GobjPanDat.SControleProcesoObj(True)
            If Not FblnEstanTodosOkInterfaz() Then
                Throw New ValorArgumentoInvalidoException("Hay argumentos no validos")
            End If
            Dim lenuTipoInterfaz As EnuTipoInterfazDef = GobjParametros.ObjTipoInterfazByt.ObjValorPro
            Dim lstrFechaDesde = ClsPanorama.FstrFechayyyymmdd(DtmFechaDesde)
            Dim lstrFechaHasta = ClsPanorama.FstrFechayyyymmdd(DtmFechaHasta)
            If ablnFinMes Then
                MstrNombreArchivo = "InterfazApolo_FM_" & lstrFechaDesde & "_" & lstrFechaHasta
            Else
                MstrNombreArchivo = "InterfazApolo_" & lstrFechaDesde & "_" & lstrFechaHasta
            End If
            HstrArchivoSalidaInterfaz = GstrTrayInterfContable & "\" & MstrNombreArchivo
            mstrNroInterfaz = Date.Now.ToString
            If lenuTipoInterfaz = EnuTipoInterfazDef.EnuPorComprobante Then
                MstrTipoComprobante = ClsCBInterfazContableOri.StrTipoComprobante
                MentIdComprobante = EntIdComprobanteInicial - 1
                If GobjParametros.ObjIdAppContableByt.ObjValorPro = EnuAppConta.EnuApoloAP Then
                    SGenereIntConApoloAPCom(astrMens)
                ElseIf GobjParametros.ObjIdAppContableByt.ObjValorPro = EnuAppConta.EnuApoloBD Then
                    SGenereIntConApoloBDCom(astrMens)
                End If
            Else
                If GobjParametros.ObjIdAppContableByt.ObjValorPro = EnuAppConta.EnuApoloAP Then
                    SGenereIntConApoloAPDoc()
                ElseIf GobjParametros.ObjIdAppContableByt.ObjValorPro = EnuAppConta.EnuApoloBD Then
                    SGenereIntConApoloBDDoc()
                End If
            End If
            lblnNohayError = True
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As ArgumentException
            Throw
        Catch ex As PathTooLongException
            Throw
        Catch ex As DirectoryNotFoundException
            Throw
        Catch ex As NotSupportedException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNohayError Then
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
#Region "Por Documento"
#Region "Documento Archivo Plano"
    Private Sub SGenereIntConApoloAPDoc()
        Dim ldtbMovi = FdtbMovimiento()
        Using lswInterfaz = ClsPanorama.FswStreamWriter(HstrArchivoSalidaInterfaz & ".txt")
            SProceseDoc(ldtbMovi, EnuIdDocumentoDef.enuFacturaVenta, lswInterfaz)
            SProceseDoc(ldtbMovi, EnuIdDocumentoDef.enuNotaIntMora, lswInterfaz)
            SProceseRecibosCaja(lswInterfaz)
            SProceseDoc(ldtbMovi, EnuIdDocumentoDef.enuNotaAplicacionAnt, lswInterfaz)
            SProceseDoc(ldtbMovi, EnuIdDocumentoDef.enuNotaCr, lswInterfaz)
            SProceseNotaR(lswInterfaz, EnuIdDocumentoDef.enuNotaReintegroAnt)
            SProceseDoc(ldtbMovi, EnuIdDocumentoDef.enuNotaReversaCr, lswInterfaz)
            SProceseNotaR(lswInterfaz, EnuIdDocumentoDef.enuNotaReversaCr)
            SProceseNotasAjuste(lswInterfaz)
        End Using
    End Sub
    Private Sub SProceseDoc(adtbMovi As DataTable, aenuIdDocumento As EnuIdDocumentoDef,
                            aswInterfaz As StreamWriter) 'Ok
        Dim ldecValor As Decimal
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(aenuIdDocumento)
        Dim lstrFiltro = ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " & aenuIdDocumento
        Dim ldrwNovedades As DataRow() = adtbMovi.Select(lstrFiltro)
        Dim lobjFactura As New ClsFactura()
        For Each ldrwNov As DataRow In ldrwNovedades
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            SEscribaMovDocDbAP(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz)
            If aenuIdDocumento = EnuIdDocumentoDef.EnuFacturaVenta Then
                SEscribaCrFacAP(ldrwNov, lstrTipoDoc, ldecValor, lobjFactura, aswInterfaz)
            Else
                SEscribaMovDocCrAP(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz)
            End If
        Next
    End Sub
    Private Sub SProceseRecibosCaja(aswInterfaz As StreamWriter)
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.enuReciboCaja)
        Dim lstrPrefRC = String.Empty, lentIdRC = 0, lstrPrefRCNov As String, lentIdRCNov As Integer
        Dim lstrFiltro As String, lstrIdCuentaDb = String.Empty, lstrIdCuentaDbNov As String
        Dim ldrwMovRC As DataRow()
        Dim ldtbMovRC = FdtbMoviRC()
        For Each ldrwNov As DataRow In ldtbMovRC.Rows
            lstrPrefRCNov = ClsPanorama.FobjValorCampo(ldrwNov(ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lentIdRCNov = ClsPanorama.FobjValorCampo(ldrwNov(ClsIdDocOrigenEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            lstrIdCuentaDbNov = ClsPanorama.FobjValorCampo(ldrwNov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            If lstrPrefRC <> lstrPrefRCNov OrElse lentIdRC <> lentIdRCNov OrElse
                    lstrIdCuentaDb <> lstrIdCuentaDbNov Then
                lstrFiltro = ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" & lstrPrefRCNov &
                        "' AND " & ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " & lentIdRCNov &
                        " AND " & ClsIdCuentaDb_NovStr.SstrNombreCampoBd & " = '" &
                        lstrIdCuentaDbNov & "'"
                ldrwMovRC = ldtbMovRC.Select(lstrFiltro)
                SProceseRecCajaAP(ldrwMovRC, lstrTipoDoc, aswInterfaz)
                lstrPrefRC = lstrPrefRCNov
                lentIdRC = lentIdRCNov
                lstrIdCuentaDb = lstrIdCuentaDbNov
            End If
        Next
    End Sub
    Private Sub SProceseRecCajaAP(adrwMoviRC As DataRow(), astrTipoDoc As String,
                aswInterfaz As StreamWriter)
        Dim ldecValor = 0D, ldecValorNov As Decimal
        ' Calcular valor debito 
        For Each ldrwMovNov As DataRow In adrwMoviRC
            ldecValorNov = ClsPanorama.FobjValorCampo(ldrwMovNov(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            ldecValor += ldecValorNov
        Next
        SEscribaMovDocDbAP(adrwMoviRC(0), astrTipoDoc, ldecValor, aswInterfaz)
        For Each ldrwMovNov As DataRow In adrwMoviRC
            ldecValorNov = ClsPanorama.FobjValorCampo(ldrwMovNov(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            SEscribaMovDocCrAP(ldrwMovNov, astrTipoDoc, ldecValorNov, aswInterfaz)
        Next
    End Sub
    Private Sub SProceseNotaR(aswInterfaz As StreamWriter,
            aenuIdDocOrigen As EnuIdDocumentoDef)
        Dim ldecValor As Decimal
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(aenuIdDocOrigen)
        Dim ldtbMovAntRei = FdtbMoviAnt(aenuIdDocOrigen)
        For Each ldrwNov As DataRow In ldtbMovAntRei.Rows
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            SEscribaMovDocDbAP(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz)
            SEscribaMovDocCrAP(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz)
        Next
    End Sub
    Private Sub SProceseNotasAjuste(aswInterfaz As StreamWriter)
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.enuNotaAjuste)
        Dim ldecValor As Decimal
        Dim ldtbMovNA = FdtbMoviNotaAjuste()
        For Each ldrwNov As DataRow In ldtbMovNA.Rows
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                        EnuTipoValor.enuDecimal)
            SEscribaMovDocDbAP(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz)
            SEscribaMovDocCrAP(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz)
        Next
    End Sub
    Private Sub SEscribaMovDocDbAP(adrwMov As DataRow, astrTipoDoc As String,
        adecValor As Decimal, aswInterfaz As StreamWriter)   'Ok
        Dim lstrPrefijoDoc As String = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString), String)
        If IsNothing(lstrPrefijoDoc) Then lstrPrefijoDoc = String.Empty
        Dim lentIdDoc As Integer = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
        Dim lstrFechaMov As String = CType(ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
            EnuTipoValor.enuDate).Date, String)
        Dim lstrIdCtaDb As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim ldecBase As Decimal = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        Dim lstrIdTercero = String.Empty
        Dim lstrNombreCliente = String.Empty
        Dim lblnEsCajaBanco = (lstrIdCtaDb = GobjParametros.ObjIdCtaCajaStr.ObjValorPro) OrElse
                GobjParametros.FblnEsCuentaBanco(lstrIdCtaDb)
        If lblnEsCajaBanco Then
            lstrIdTercero = FstrIdTerceroCajaBancos(adrwMov)
            lstrNombreCliente = FstrNombreTerceroCajaBancos(adrwMov)
        Else
            lstrIdTercero = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lstrNombreCliente = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        End If
        ' Db
        With MstbLinea
            .Clear().Append(astrTipoDoc).Append(CHSTRCOMA).Append(lstrPrefijoDoc).Append(CHSTRCOMA)
            .Append(lentIdDoc).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(lstrFechaMov).Append(CHSTRCOMA).Append(lstrIdCtaDb).Append(CHSTRCOMA)
            If BlnIdTerceroStr Then
                .Append(lstrIdTercero.PadLeft(12, "0"))
            Else
                .Append(lstrIdTercero)
            End If
            .Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(FstrDetalle(adrwMov)).Append(CHSTRCOMA)
            .Append(adecValor).Append(CHSTRCOMA).Append(0).Append(CHSTRCOMA).Append(ldecBase)
            .Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append("Supervisor").Append(CHSTRCOMA).Append("D").Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append(CHSTRCOMA).Append(lstrNombreCliente).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA)
            .Append(mstrNroInterfaz)
        End With
        aswInterfaz.WriteLine(MstbLinea.ToString)
    End Sub
    Private Sub SEscribaCrFacAP(adrwMov As DataRow, astrTipoDoc As String,
            adecValor As Decimal, aobjFactura As ClsFactura,
            aswInterfaz As StreamWriter) 'Ok
        Dim lstrPrefijoFac As String = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString), String)
        If IsNothing(lstrPrefijoFac) Then lstrPrefijoFac = String.Empty
        Dim lentIdFac As Integer = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
        Dim lshrIdItemFac As Short = ClsPanorama.FobjValorCampo(adrwMov(ClsIdItemFacturaShr.SstrNombreCampoBd),
                EnuTipoValor.enuShort)
        Dim lobjValorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefijoFac, lentIdFac}
        aobjFactura.SAbra(lobjValorLlave)
        Dim lstrIdTerceroCtaCr = aobjFactura.FstrIdTerceroCtaCr(lshrIdItemFac)
        Dim lstrFechaMov As String = CType(ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
            EnuTipoValor.enuDate).Date, String)
        Dim lstrIdCtaCr As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim ldecBase As Decimal = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        Dim lstrNombreCliente = String.Empty
        If String.IsNullOrEmpty(lstrIdTerceroCtaCr) OrElse lstrIdTerceroCtaCr = "0" Then
            lstrIdTerceroCtaCr = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
            lstrNombreCliente = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        Else
            If IsNumeric(lstrIdTerceroCtaCr) Then
                lstrNombreCliente = ClsPanorama.FstrNombreTercero(CType(lstrIdTerceroCtaCr, Double))
            End If
        End If
        ' Cr
        With MstbLinea
            .Clear().Append(astrTipoDoc).Append(CHSTRCOMA).Append(lstrPrefijoFac).Append(CHSTRCOMA)
            .Append(lentIdFac).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(lstrFechaMov).Append(CHSTRCOMA).Append(lstrIdCtaCr).Append(CHSTRCOMA)
            If BlnIdTerceroStr Then
                .Append(lstrIdTerceroCtaCr.PadLeft(12, "0"))
            Else
                .Append(lstrIdTerceroCtaCr)
            End If
            .Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(FstrDetalle(adrwMov)).Append(CHSTRCOMA)
            .Append(0).Append(CHSTRCOMA).Append(adecValor).Append(CHSTRCOMA).Append(ldecBase)
            .Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append("Supervisor").Append(CHSTRCOMA).Append("C").Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append(CHSTRCOMA).Append(lstrNombreCliente).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA)
            .Append(mstrNroInterfaz)
        End With
        aswInterfaz.WriteLine(MstbLinea.ToString)
    End Sub
    Private Sub SEscribaMovDocCrAP(adrwMov As DataRow, astrTipoDoc As String,
            adecValor As Decimal, aswInterfaz As StreamWriter) 'Ok
        Dim lstrPrefijoDoc As String = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString), String)
        If IsNothing(lstrPrefijoDoc) Then lstrPrefijoDoc = String.Empty
        Dim lentIdDoc As Integer = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
        Dim lstrFechaMov As String = CType(ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
            EnuTipoValor.enuDate).Date, String)
        Dim lstrIdCtaCr As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim ldecBase As Decimal = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        Dim lstrIdTercero = String.Empty
        Dim lstrNombreCliente = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                EnuTipoValor.enuString)
        Dim lblnEsCajaBanco = (lstrIdCtaCr = GobjParametros.ObjIdCtaCajaStr.ObjValorPro) OrElse
                GobjParametros.FblnEsCuentaBanco(lstrIdCtaCr)
        If lblnEsCajaBanco Then
            lstrIdTercero = FstrIdTerceroCajaBancos(adrwMov)
            lstrNombreCliente = FstrNombreTerceroCajaBancos(adrwMov)
        Else
            lstrIdTercero = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lstrNombreCliente = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        End If
        ' Cr
        With MstbLinea
            .Clear().Append(astrTipoDoc).Append(CHSTRCOMA).Append(lstrPrefijoDoc).Append(CHSTRCOMA)
            .Append(lentIdDoc).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(lstrFechaMov).Append(CHSTRCOMA).Append(lstrIdCtaCr).Append(CHSTRCOMA)
            If BlnIdTerceroStr Then
                .Append(lstrIdTercero.ToString.PadLeft(12, "0"))
            Else
                .Append(lstrIdTercero)
            End If
            .Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(FstrDetalle(adrwMov)).Append(CHSTRCOMA)
            .Append(0).Append(CHSTRCOMA).Append(adecValor).Append(CHSTRCOMA).Append(ldecBase)
            .Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append("Supervisor").Append(CHSTRCOMA).Append("C").Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append(CHSTRCOMA).Append(lstrNombreCliente).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA)
            .Append(mstrNroInterfaz)
        End With
        aswInterfaz.WriteLine(MstbLinea.ToString)
    End Sub
#End Region
#Region "Documento BD"
    Private Sub SGenereIntConApoloBDDoc()
        Dim ldtbMovi = FdtbMovimiento()
        SCreeBdApolo(HstrArchivoSalidaInterfaz & ".mdb")
        ' Crea la conección y la usa
        Using lcnnApolo = FcnnApoloMdb(MstrNombreArchivo & ".mdb")
            ' Limpia la tabla Apolo
            GobjPanDat.SLimpieTabla(lcnnApolo, EnuProveedorBD.enuOleDb, "Apolo")
            SProceseDoc(ldtbMovi, EnuIdDocumentoDef.enuFacturaVenta, lcnnApolo)
            SProceseDoc(ldtbMovi, EnuIdDocumentoDef.enuNotaIntMora, lcnnApolo)
            SProceseRecibosCaja(lcnnApolo)
            SProceseDoc(ldtbMovi, EnuIdDocumentoDef.enuNotaAplicacionAnt, lcnnApolo)
            SProceseDoc(ldtbMovi, EnuIdDocumentoDef.enuNotaCr, lcnnApolo)
            SProceseNotaR(lcnnApolo, EnuIdDocumentoDef.enuNotaReintegroAnt)
            SProceseDoc(ldtbMovi, EnuIdDocumentoDef.enuNotaReversaCr, lcnnApolo)
            SProceseNotaR(lcnnApolo, EnuIdDocumentoDef.enuNotaReversaCr)
            SProceseNotasAjuste(lcnnApolo)
        End Using
    End Sub
    Private Sub SProceseDoc(adtbMovi As DataTable, aenuIdDocumento As EnuIdDocumentoDef,
                            acnnConexion As Object)
        Dim ldecValor As Decimal
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(aenuIdDocumento)
        Dim lstrFiltro = ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " & aenuIdDocumento
        Dim ldrwNovedades As DataRow() = adtbMovi.Select(lstrFiltro)
        Dim lobjFactura As New ClsFactura()
        For Each ldrwNov As DataRow In ldrwNovedades
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            SEscribaMovDocDbBD(ldrwNov, lstrTipoDoc, ldecValor, acnnConexion)
            If aenuIdDocumento = EnuIdDocumentoDef.EnuFacturaVenta Then
                SEscribaCrFacBD(ldrwNov, lstrTipoDoc, ldecValor, lobjFactura, acnnConexion)
            Else
                SEscribaMovDocCrBD(ldrwNov, lstrTipoDoc, ldecValor, acnnConexion)
            End If
        Next
    End Sub
    Private Sub SProceseRecibosCaja(acnnConexion As Object)
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.enuReciboCaja)
        Dim lstrPrefRC = String.Empty, lentIdRC = 0, lstrPrefRCNov As String, lentIdRCNov As Integer
        Dim lstrFiltro As String, lstrIdCuentaDb = String.Empty, lstrIdCuentaDbNov As String
        Dim ldtbMovRC = FdtbMoviRC()
        For Each ldrwNov As DataRow In ldtbMovRC.Rows
            lstrPrefRCNov = ClsPanorama.FobjValorCampo(ldrwNov(ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lentIdRCNov = ClsPanorama.FobjValorCampo(ldrwNov(ClsIdDocOrigenEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            lstrIdCuentaDbNov = ClsPanorama.FobjValorCampo(ldrwNov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            If lstrPrefRC <> lstrPrefRCNov OrElse lentIdRC <> lentIdRCNov OrElse
                    lstrIdCuentaDb <> lstrIdCuentaDbNov Then
                lstrFiltro = ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" & lstrPrefRCNov &
                        "' AND " & ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " & lentIdRCNov &
                        " AND " & ClsIdCuentaDb_NovStr.SstrNombreCampoBd & " = '" &
                        lstrIdCuentaDbNov & "'"
                Dim ldrwMovRC As DataRow() = ldtbMovRC.Select(lstrFiltro)
                SProceseRecCajaBD(ldrwMovRC, lstrTipoDoc, acnnConexion)
                lstrPrefRC = lstrPrefRCNov
                lentIdRC = lentIdRCNov
                lstrIdCuentaDb = lstrIdCuentaDbNov
            End If
        Next
    End Sub
    Private Sub SProceseRecCajaBD(adrwMoviRC As DataRow(), astrTipoDoc As String,
                acnnConexion As Object)
        Dim ldecValor = 0D, ldecValorNov As Decimal
        ' Calcular valor debito 
        For Each ldrwMovNov As DataRow In adrwMoviRC
            ldecValorNov = ClsPanorama.FobjValorCampo(ldrwMovNov(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            ldecValor += ldecValorNov
        Next
        SEscribaMovDocDbBD(adrwMoviRC(0), astrTipoDoc, ldecValor, acnnConexion)
        For Each ldrwMovNov As DataRow In adrwMoviRC
            ldecValorNov = ClsPanorama.FobjValorCampo(ldrwMovNov(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            SEscribaMovDocCrBD(ldrwMovNov, astrTipoDoc, ldecValorNov, acnnConexion)
        Next
    End Sub
    Private Sub SProceseNotaR(acnnConexion As Object,
                aenuIdDocumento As EnuIdDocumentoDef)
        Dim ldecValor As Decimal
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(aenuIdDocumento)
        Dim ldtbMovAntR = FdtbMoviAnt(aenuIdDocumento)
        For Each ldrwNov As DataRow In ldtbMovAntR.Rows
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            SEscribaMovDocDbBD(ldrwNov, lstrTipoDoc, ldecValor, acnnConexion)
            SEscribaMovDocCrBD(ldrwNov, lstrTipoDoc, ldecValor, acnnConexion)
        Next
    End Sub
    Private Sub SProceseNotasAjuste(acnnConexion As Object)
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.enuNotaAjuste)
        Dim ldecValor As Decimal
        Dim ldtbMovNA = FdtbMoviNotaAjuste()
        For Each ldrwNov As DataRow In ldtbMovNA.Rows
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                        EnuTipoValor.enuDecimal)
            SEscribaMovDocDbBD(ldrwNov, lstrTipoDoc, ldecValor, acnnConexion)
            SEscribaMovDocCrBD(ldrwNov, lstrTipoDoc, ldecValor, acnnConexion)
        Next
    End Sub
    Private Sub SEscribaMovDocDbBD(adrwMov As DataRow, astrTipoDoc As String,
        adecValor As Decimal, acnnConexion As Object)
        Dim lstrPrefijoDoc As String = ClsPanorama.FobjValorCampo(adrwMov(
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        If IsNothing(lstrPrefijoDoc) Then lstrPrefijoDoc = String.Empty
        Dim lentIdDoc As Integer = ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
        Dim lcolCampos As Collection = FcolCamposDoc(), lcolDatos As New Collection
        Dim ldtmFechaMov As Date = ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate)
        Dim lstrIdCtaDb As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim lstrIdTercero As String, lstrNombreCliente = String.Empty
        Dim lblnCajaBancos = (lstrIdCtaDb = GobjParametros.ObjIdCtaCajaStr.ObjValorPro) OrElse
                GobjParametros.FblnEsCuentaBanco(lstrIdCtaDb)
        If lblnCajaBancos Then
            lstrIdTercero = FstrIdTerceroCajaBancos(adrwMov)
            lstrNombreCliente = FstrNombreTerceroCajaBancos(adrwMov)
        Else
            lstrIdTercero = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
            lstrNombreCliente = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        End If
        Dim ldecBase As Decimal = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        With lcolDatos
            .Add(astrTipoDoc)
            .Add(lstrPrefijoDoc)
            .Add(lentIdDoc)
            .Add(ldtmFechaMov)
            .Add(lstrIdCtaDb)
            If BlnIdTerceroStr Then
                .Add(lstrIdTercero.PadLeft(12, "0"))
            Else
                .Add(lstrIdTercero)
            End If
            .Add(String.Empty)
            .Add(FstrDetalle(adrwMov))
            .Add(adecValor)
            .Add(0)
            .Add(ldecBase)
            .Add("Supervisor")
            .Add("D")
            .Add(lstrNombreCliente)
            .Add(String.Empty)
            .Add(mstrNroInterfaz)
        End With
        GobjPanDat.SInserteRegistro(acnnConexion, EnuProveedorBD.enuOleDb, "Apolo", lcolCampos, lcolDatos)
    End Sub
    Private Sub SEscribaMovDocCrBD(adrwMov As DataRow, astrTipoDoc As String,
            adecValor As Decimal, acnnConexion As Object)
        Dim lstrPrefijoDoc As String = CType(ClsPanorama.FobjValorCampo(adrwMov(
        ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString), String)
        If IsNothing(lstrPrefijoDoc) Then lstrPrefijoDoc = String.Empty
        Dim lentIdDoc As Integer = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
        Dim lcolCampos As Collection = FcolCamposDoc(), lcolDatos As New Collection
        Dim ldtmFechaMov As Date = ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate)
        Dim lstrIdCtaCr As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim ldecBase = 0D, lstrIdTercero = String.Empty
        Dim lstrNombreCliente = String.Empty
        Dim lblnCajaBancos = (lstrIdCtaCr = GobjParametros.ObjIdCtaCajaStr.ObjValorPro) OrElse
                GobjParametros.FblnEsCuentaBanco(lstrIdCtaCr)
        If lblnCajaBancos Then
            lstrIdTercero = FstrIdTerceroCajaBancos(adrwMov)
            lstrNombreCliente = FstrNombreTerceroCajaBancos(adrwMov)
        Else
            lstrIdTercero = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
            lstrNombreCliente = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        End If
        ldecBase = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        With lcolDatos
            .Add(astrTipoDoc)
            .Add(lstrPrefijoDoc)
            .Add(lentIdDoc)
            .Add(ldtmFechaMov)
            .Add(lstrIdCtaCr)
            If BlnIdTerceroStr Then
                .Add(lstrIdTercero.PadLeft(12, "0"))
            Else
                .Add(lstrIdTercero)
            End If
            .Add(String.Empty)
            .Add(FstrDetalle(adrwMov))
            .Add(0)
            .Add(adecValor)
            .Add(ldecBase)
            .Add("Supervisor")
            .Add("C")
            .Add(lstrNombreCliente)
            .Add(String.Empty)
            .Add(mstrNroInterfaz)
        End With
        GobjPanDat.SInserteRegistro(acnnConexion, EnuProveedorBD.enuOleDb, "Apolo", lcolCampos, lcolDatos)
    End Sub
    Private Sub SEscribaCrFacBD(adrwMov As DataRow, astrTipoDoc As String,
            adecValor As Decimal, aobjFactura As ClsFactura,
            acnnConexion As Object)
        Dim lstrPrefijoFac As String = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString), String)
        If IsNothing(lstrPrefijoFac) Then lstrPrefijoFac = String.Empty
        Dim lentIdFac As Integer = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
        Dim lshrIdItemFac As Short = ClsPanorama.FobjValorCampo(adrwMov(ClsIdItemFacturaShr.SstrNombreCampoBd),
                EnuTipoValor.enuShort)
        Dim lobjValorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefijoFac, lentIdFac}
        aobjFactura.SAbra(lobjValorLlave)
        Dim lstrIdTerceroCtaCr = aobjFactura.FstrIdTerceroCtaCr(lshrIdItemFac)
        Dim lstrNombreCliente = String.Empty
        If String.IsNullOrEmpty(lstrIdTerceroCtaCr) Then
            lstrIdTerceroCtaCr = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
            lstrNombreCliente = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        Else
            lstrNombreCliente = ClsPanorama.FstrNombreTercero(lstrIdTerceroCtaCr)
        End If
        Dim lcolCampos As Collection = FcolCamposDoc(), lcolDatos As New Collection
        Dim ldtmFechaMov As Date = ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate)
        Dim lstrIdCtaCr As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim ldecBase = 0D
        ldecBase = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        With lcolDatos
            .Add(astrTipoDoc)
            .Add(lstrPrefijoFac)
            .Add(lentIdFac)
            .Add(ldtmFechaMov)
            .Add(lstrIdCtaCr)
            If BlnIdTerceroStr Then
                .Add(lstrIdTerceroCtaCr.PadLeft(12, "0"))
            Else
                .Add(lstrIdTerceroCtaCr)
            End If
            .Add(String.Empty)
            .Add(FstrDetalle(adrwMov))
            .Add(0)
            .Add(adecValor)
            .Add(ldecBase)
            .Add("Supervisor")
            .Add("C")
            .Add(lstrNombreCliente)
            .Add(String.Empty)
            .Add(mstrNroInterfaz)
        End With
        GobjPanDat.SInserteRegistro(acnnConexion, EnuProveedorBD.enuOleDb, "Apolo", lcolCampos, lcolDatos)
    End Sub
    Private Shared Function FcolCamposDoc() As Collection
        Dim lcolCampos As New Collection
        With lcolCampos
            .Add("Tipo")
            .Add("Prefijo")
            .Add("Numero")
            .Add("Fecha")
            .Add("Cuenta")
            .Add("Tercero")
            .Add("Centro")
            .Add("Detalle")
            .Add("Debito")
            .Add("Credito")
            .Add("Base")
            .Add("Usuario")
            .Add("Signo")
            .Add("NombreTercero")
            .Add("NombreCentro")
            .Add("InterfaceNo")
        End With
        Return lcolCampos
    End Function
#End Region
#End Region
#Region "Por Comprobante"
#Region "Comprobante Archivo Plano"
    Private Sub SGenereIntConApoloAPCom(ByRef astrMens As String)
        Dim lblnHayMovimiento = False
        Dim ldtbMovi = FdtbMovimiento()
        Dim ldtbCajaBancos = FdtbCajaBancos()
        Dim ldtbAnticipos = FdtbAnticipos()
        Dim ldtmFechaMov = DtmFechaDesde, lstrFiltro As String
        Dim ldrwMovimientos() As DataRow
        Dim ldecValorMov As Decimal, lstrIdCtaDb As String, lstrIdCtaCr As String
        Dim lblnHayMovFechaRC = False, lblnHayMovAnt = False, lblnHayMovFecha As Boolean
        Dim lstrCtaCaja = GobjParametros.ObjIdCtaCajaStr.ObjValorPro
        Dim lobjFactura As New ClsFactura()
        Dim lstrArchivoPlanoInterfaz = HstrArchivoSalidaInterfaz & ".txt"
        Using lswInterfaz = ClsPanorama.FswStreamWriter(lstrArchivoPlanoInterfaz)
            Do While ldtmFechaMov <= DtmFechaHasta
                lblnHayMovFecha = False
                lblnHayMovFechaRC = False
                lblnHayMovAnt = False
                SProceseCajaBancos(ldtbCajaBancos, ldtmFechaMov, lswInterfaz, lblnHayMovFechaRC)
                SProceseAnticipos(ldtbAnticipos, ldtmFechaMov, lswInterfaz, lblnHayMovFechaRC, lblnHayMovAnt)
                lblnHayMovFecha = lblnHayMovFechaRC OrElse lblnHayMovAnt
                lstrFiltro = ClsFechaNovedadDtm.SstrNombreCampoBd & " = '" & ldtmFechaMov & "'"
                ldrwMovimientos = ldtbMovi.Select(lstrFiltro)
                If Not lblnHayMovimiento Then
                    lblnHayMovimiento = lblnHayMovFecha
                End If
                If ldrwMovimientos.Length > 0 Then
                    If Not lblnHayMovFecha Then
                        mentConsecutivo = 0
                        MentIdComprobante += 1
                    End If
                    If Not lblnHayMovimiento Then
                        lblnHayMovimiento = True
                    End If
                    For Each ldrwMov As DataRow In ldrwMovimientos
                        lstrIdCtaDb = ClsPanorama.FobjValorCampo(ldrwMov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                                EnuTipoValor.enuString)
                        lstrIdCtaCr = ClsPanorama.FobjValorCampo(ldrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                                EnuTipoValor.enuString)
                        ldecValorMov = ClsPanorama.FobjValorCampo(ldrwMov(ClsValor_NovDec.SstrNombreCampoBd),
                                EnuTipoValor.enuDecimal)
                        SProceseMovAP(ldrwMov, ldecValorMov, lstrIdCtaDb, lstrIdCtaCr, lstrCtaCaja,
                                lobjFactura, lswInterfaz)
                    Next
                End If
                ldtmFechaMov = ldtmFechaMov.AddDays(1)
            Loop
        End Using
        If Not lblnHayMovimiento Then
            astrMens = "No hay Movimiento entre las Fechas seleccionadas!"
        End If
    End Sub

    Private Sub SProceseMovAP(adrwMov As DataRow, adecVlr As Decimal,
            astrIdCtaDb As String, astrIdCtaCr As String, astrIdCtaCaja As String,
            aobjFactura As ClsFactura, aswInterfaz As StreamWriter)
        Dim lenuTipoDoc As EnuTipoDocOri
        If astrIdCtaDb = astrIdCtaCaja OrElse
                GobjParametros.FblnEsCuentaBanco(astrIdCtaDb) Then
            SEscribaAPMovCr(adrwMov, adecVlr, aswInterfaz, False)
        Else
            SEscribaAPMovDb(adrwMov, adecVlr, aswInterfaz, False)
            If astrIdCtaCr = astrIdCtaCaja OrElse
                    GobjParametros.FblnEsCuentaBanco(astrIdCtaCr) Then
                SEscribaAPMovCr(adrwMov, adecVlr, aswInterfaz, True)
            Else
                lenuTipoDoc = ClsPanorama.FobjValorCampo(adrwMov(ClsIdTipoDocOrigenByt.SstrNombreCampoBd),
                        EnuTipoValor.enuByte)
                If lenuTipoDoc = EnuTipoDocOri.enuFactura Then
                    SEscribaAPFacCr(adrwMov, adecVlr, aswInterfaz, aobjFactura)
                Else
                    SEscribaAPMovCr(adrwMov, adecVlr, aswInterfaz, False)
                End If
            End If
        End If
    End Sub

    Private Sub SProceseCajaBancos(adtbCajaBancos As DataTable, adtmFecha As Date,
                aswInterfaz As StreamWriter, ByRef ablnHayMovFecha As Boolean) 'Ok
        Dim lblnHayMov = False, lentIndiceDatRow As Integer
        Dim ldecValor As Decimal, lblnDataRowAsignado As Boolean
        Dim lstrFiltro = ClsFechaNovedadDtm.SstrNombreCampoBd & " = '" & adtmFecha & "'"
        Dim ldrwCajaBancos = adtbCajaBancos.Select(lstrFiltro)
        If ldrwCajaBancos.Length > 0 Then
            Dim ldrwCajaBanco As DataRow, lstrIdCtaDb As String
            MentIdComprobante += 1
            mentConsecutivo = 0
            lblnHayMov = True
            For i = 0 To ldrwCajaBancos.Length - 1
                lblnDataRowAsignado = False
                ldecValor = 0
                lentIndiceDatRow = 0
                ldrwCajaBanco = ldrwCajaBancos(i)
                Dim lstrPrefijoDoc As String = CType(ClsPanorama.FobjValorCampo(ldrwCajaBanco(
                        ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString), String)
                If IsNothing(lstrPrefijoDoc) Then lstrPrefijoDoc = String.Empty
                Dim lentIdDoc As Integer = CType(ClsPanorama.FobjValorCampo(ldrwCajaBanco(
                        ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
                lstrIdCtaDb = ldrwCajaBanco(ClsIdCuentaDb_NovStr.SstrNombreCampoBd)
                Do While ClsPanorama.FobjValorCampo(ldrwCajaBanco(ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString) = lstrPrefijoDoc AndAlso
                        ClsPanorama.FobjValorCampo(ldrwCajaBanco(ClsIdDocOrigenEnt.SstrNombreCampoBd),
                        EnuTipoValor.enuInteger) = lentIdDoc AndAlso
                        ldrwCajaBanco(ClsIdCuentaDb_NovStr.SstrNombreCampoBd) = lstrIdCtaDb
                    If Not lblnDataRowAsignado Then
                        lblnDataRowAsignado = True
                        lentIndiceDatRow = i
                    End If
                    ldecValor += ClsPanorama.FobjValorCampo(ldrwCajaBanco(ClsValor_NovDec.SstrNombreCampoBd),
                                                EnuTipoValor.enuDecimal)
                    i += 1
                    If i < ldrwCajaBancos.Length Then
                        ldrwCajaBanco = ldrwCajaBancos(i)
                    Else
                        Exit Do
                    End If
                Loop
                i -= 1
                ldrwCajaBanco = ldrwCajaBancos(lentIndiceDatRow)
                SEscribaAPMovDb(ldrwCajaBanco, ldecValor, aswInterfaz, True)
            Next
        End If
        ablnHayMovFecha = lblnHayMov
    End Sub

    Private Sub SProceseAnticipos(adtbAnticipos As DataTable, adtmFecha As Date,
                    aswInterfaz As StreamWriter, ablnHayMovRC As Boolean, ByRef ablnHayMovAnt As Boolean) 'Ok TipoNov
        Dim lentIdAnticipo As Integer, lobjAnticipo As ClsAnticipo
        Dim lenuTipoDocOrigenAnt As EnuTipoDocOri
        Dim ldecValor As Decimal, lenuIdTipoNovAnt As EnuTipoNov
        Dim lstrFiltro = ClsFechaNovedadDtm.SstrNombreCampoBd & " = '" & adtmFecha & "'"
        Dim ldrwAnticipos = adtbAnticipos.Select(lstrFiltro)
        Dim lblnHayMov = False
        If ldrwAnticipos.Length > 0 Then
            If Not ablnHayMovRC Then
                MentIdComprobante += 1
                mentConsecutivo = 0
            End If
            lblnHayMov = True
            For Each ldrwAnt As DataRow In ldrwAnticipos
                lentIdAnticipo = ClsPanorama.FobjValorCampo(ldrwAnt(ClsIdAnticipo_NovEnt.SstrNombreCampoBd),
                            EnuTipoValor.enuLong)
                lobjAnticipo = New ClsAnticipo(EnuModoInstanciaObjDef.enuUnico)
                lobjAnticipo.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lentIdAnticipo})
                lenuTipoDocOrigenAnt = lobjAnticipo.ObjIdTipoDocOrigen_AntByt.ObjValorPro
                ldecValor = ClsPanorama.FobjValorCampo(ldrwAnt(ClsValor_NovDec.SstrNombreCampoBd),
                            EnuTipoValor.enuDecimal)
                lenuIdTipoNovAnt = ClsPanorama.FobjValorCampo(ldrwAnt(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd),
                            EnuTipoValor.enuByte)
                If lenuIdTipoNovAnt = EnuTipoNov.EnuDbAntDev OrElse
                        lenuIdTipoNovAnt = EnuTipoNov.EnuRDbAntDev OrElse
                        lenuIdTipoNovAnt = EnuTipoNov.EnuRCrAntRec Then
                    SEscribaAPMovDbAnt(ldrwAnt, ldecValor, aswInterfaz, lenuTipoDocOrigenAnt)
                    SEscribaAPMovCrAnt(ldrwAnt, ldecValor, aswInterfaz, lenuTipoDocOrigenAnt)
                ElseIf lenuIdTipoNovAnt = EnuTipoNov.EnuCrAntRec Then
                    If lenuTipoDocOrigenAnt = EnuTipoDocOri.EnuNotaAjuste Then
                        SEscribaAPMovDbAnt(ldrwAnt, ldecValor, aswInterfaz, lenuTipoDocOrigenAnt)
                    End If
                    SEscribaAPMovCrAnt(ldrwAnt, ldecValor, aswInterfaz, lenuTipoDocOrigenAnt)
                Else
                    Throw New ErrorInesperadoPanLException("Tipo de la Novedad del Anticipo no esperada!")
                End If
            Next
        End If
        ablnHayMovAnt = lblnHayMov
    End Sub

    Private Function FdtbCajaBancos() As DataTable
        Dim lstrIndice = ClsFechaNovedadDtm.SstrNombreCampoBd & " ASC, " &
                ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd & " ASC, " &
                ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd & " ASC, " &
                ClsIdCuentaDb_NovStr.SstrNombreCampoBd & " ASC"
        Dim lstrSqlCajaBancosNov = FstrExpSqlCajaBancosNovedades()
        Dim lstrSqlCajaBancosNovAnt = FstrExpSqlCajaBancosNovedadesAnt()
        Dim lstrsqlCajaBancos = "(" & lstrSqlCajaBancosNov & ") UNION ALL (" &
                lstrSqlCajaBancosNovAnt & ") ORDER BY " & lstrIndice
        Dim ldtbCajaBancos = ClsPanorama.FdtbDataTable(lstrsqlCajaBancos)
        Return ldtbCajaBancos
    End Function

    Private Function FdtbAnticipos() As DataTable 'Ok TipoNov
        Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaDesde.ToString) & "'"
        Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaHasta.ToString) & "'"
        Dim lstrNombreTablaT1 = ClsNovedadAnticipo.SstrNombreTabla
        Dim lstrNombreTablaT2 = ClsCliente.SstrNombreTabla
        Dim lstrCamposSelectT1 = {ClsFechaNovedadAntDtm.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd,
                ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd,
                "'' AS " & ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd,
                ClsIdAnticipo_NovEnt.SstrNombreCampoBd, ClsIdNovedadAntShr.SstrNombreCampoBd,
                ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd, ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd,
                ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd, ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd,
                ClsAliasCont_NovAntStr.SstrNombreCampoBd, ClsIdTercero_NovAntDbl.SstrNombreCampoBd,
                "SUM(" & ClsValor_NovAntDec.SstrNombreCampoBd & ")",
                ClsIdTipoNovedadByt.SstrNombreCampoBd}
        Dim lstrCamposSelectT2 = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdTercero_NovAntDbl.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsFechaNovedadAntDtm.SstrNombreCampoBd, "ASC"},
                {ClsIdAnticipo_NovEnt.SstrNombreCampoBd, "ASC"},
                {ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd, "ASC"},
                {ClsIdNovedadShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta.ToString &
                " AND P." & StrCampoCentroUtil & " = " & GshrIdCentroUtil.ToString &
                " AND " & ClsFechaNovedadAntDtm.SstrNombreCampoBd & " >= " & lstrFechaDesde & " AND " &
                ClsFechaNovedadAntDtm.SstrNombreCampoBd & " <= " & lstrFechaHasta & " AND (" &
                ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd & " <= " & EnuTipoNov.EnuDbAntDev &
                " OR " & ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd &
                " BETWEEN " & EnuTipoNov.EnuRCrAntRec &
                " AND " & EnuTipoNov.EnuRDbAntDev & ") AND " &
                ClsValor_NovDec.SstrNombreCampoBd & " > 0"
        Dim lstrCamposGrupo = {ClsFechaNovedadAntDtm.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd,
                ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd,
                ClsIdAnticipo_NovEnt.SstrNombreCampoBd, ClsIdNovedadAntShr.SstrNombreCampoBd,
                ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd, ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd,
                ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd, ClsAliasCont_NovAntStr.SstrNombreCampoBd,
                ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim ldtbMovConTer = ClsPanorama.FdtbDataTable(lstrNombreTablaT1, lstrCamposSelectT1,
                lstrNombreTablaT2, lstrCamposSelectT2, lstrCamposRelPri, lstrCamposRelSec,
                lstrIndice, lstrFiltro, lstrCamposGrupo, True)
        Return ldtbMovConTer
    End Function

    ''' <summary>
    ''' Devuelve una expresio SQL que seleccionar los registros de la tabla Novedades relacionadas con los
    ''' ingresos, por Recibos de Caja, a Caja y Bancos.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FstrExpSqlCajaBancosNovedades() As String
        Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaDesde.ToString) & "'"
        Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaHasta.ToString) & "'"
        Dim lstrTablaPri = ClsNovedad.SstrNombreTabla
        Dim lstrTablaSec = ClsCliente.SstrNombreTabla
        Dim lstrCamposTablaPri = {ClsIdTipoDocOrigenByt.SstrNombreCampoBd,
                ClsFechaNovedadDtm.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd,
                ClsIdDocOrigenEnt.SstrNombreCampoBd,
                ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd,
                ClsIdCuentaDb_NovStr.SstrNombreCampoBd,
                ClsAliasCont_NovStr.SstrNombreCampoBd, ClsIdTercero_NovDbl.SstrNombreCampoBd,
                "SUM(" & ClsBaseDec.SstrNombreCampoBd & ")",
                "SUM(" & ClsValor_NovDec.SstrNombreCampoBd & ")",
                ClsIdTipoNovedadByt.SstrNombreCampoBd}
        Dim lstrCamposTablaSec = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdTercero_NovDbl.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamposIndice = {{"", ""}}
        Dim lstrFiltro = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta.ToString &
                " AND P." & StrCampoCentroUtil & " = " & GshrIdCentroUtil.ToString &
                " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd & " >= " & lstrFechaDesde & " AND " &
                ClsFechaNovedadDtm.SstrNombreCampoBd & " <= " & lstrFechaHasta & " AND " &
                FstrFiltroCajaBancoNov() & " AND " & ClsValor_NovDec.SstrNombreCampoBd & " > 0"
        Dim lstrCamposgrupo = {ClsFechaNovedadDtm.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd, ClsIdDocOrigenEnt.SstrNombreCampoBd,
                ClsIdCuentaDb_NovStr.SstrNombreCampoBd, ClsIdTipoDocOrigenByt.SstrNombreCampoBd,
                ClsAliasCont_NovStr.SstrNombreCampoBd, ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri, lstrCamposTablaPri, lstrTablaSec,
                lstrCamposTablaSec, lstrCamposRelPri, lstrCamposRelSec, lstrCamposIndice, lstrFiltro, lstrCamposgrupo)
        Return lstrSql
    End Function

    ''' <summary>
    ''' Devuelve una expresio SQL que seleccionar los registros de la tabla NovedadesAnt relacionadas con los
    ''' ingresos, por Recibos de Caja, a Caja y Bancos.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FstrExpSqlCajaBancosNovedadesAnt() As String
        Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaDesde.ToString) & "'"
        Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(DtmFechaHasta.ToString) & "'"
        Dim lstrTablaPri = ClsNovedadAnticipo.SstrNombreTabla
        Dim lstrTablaSec = ClsCliente.SstrNombreTabla
        Dim lstrCamposTablaPri = {ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd,
                ClsFechaNovedadAntDtm.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd,
                ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd,
                "'' AS " & ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd,
                ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd,
                ClsAliasCont_NovAntStr.SstrNombreCampoBd, ClsIdTercero_NovAntDbl.SstrNombreCampoBd,
                "0 AS " & ClsBaseDec.SstrNombreCampoBd,
                "SUM(" & ClsValor_NovAntDec.SstrNombreCampoBd & ")",
                ClsIdTipoNovedadByt.SstrNombreCampoBd}
        Dim lstrCamposTablaSec = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdTercero_NovAntDbl.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamposIndice = {{"", ""}}
        Dim lstrFiltro = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta.ToString &
                " AND P." & StrCampoCentroUtil & " = " & GshrIdCentroUtil.ToString &
                " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd & " >= " & lstrFechaDesde & " AND " &
                ClsFechaNovedadDtm.SstrNombreCampoBd & " <= " & lstrFechaHasta & " AND " &
                FstrFiltroCajaBancoNovAnt() & " AND " & ClsValor_NovDec.SstrNombreCampoBd & " > 0"
        Dim lstrCamposgrupo = {ClsFechaNovedadDtm.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd, ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd,
                ClsIdCuentaDb_NovStr.SstrNombreCampoBd, ClsIdTipoDocOrigenByt.SstrNombreCampoBd,
                ClsAliasCont_NovStr.SstrNombreCampoBd, ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri, lstrCamposTablaPri, lstrTablaSec,
                lstrCamposTablaSec, lstrCamposRelPri, lstrCamposRelSec, lstrCamposIndice, lstrFiltro, lstrCamposgrupo)
        Return lstrSql
    End Function
#End Region
#Region "Comprobante BD"
    Private Sub SGenereIntConApoloBDCom(ByRef astrMens As String)
        Dim lblnHayMovimiento = False
        Dim ldtbMovi = FdtbMovimiento()
        Dim ldtbCajaBancos = FdtbCajaBancos()
        Dim ldtbAnticipos = FdtbAnticipos()
        Dim ldtmFechaMov = DtmFechaDesde, lstrFiltro As String
        Dim ldrwMovimientos() As DataRow
        Dim ldecValorMov As Decimal, lstrIdCtaDb As String, lstrIdCtaCr As String
        Dim lblnHayMovFechaRC As Boolean
        Dim lblnHayMovAnt As Boolean, lblnHayMovFecha As Boolean
        Dim lstrCtaCaja = GobjParametros.ObjIdCtaCajaStr.ObjValorPro
        Dim lobjFactura As New ClsFactura()
        SCreeBdApolo(HstrArchivoSalidaInterfaz & ".mdb")
        ' Crea la conección y la usa
        Using lcnnApolo = FcnnApoloMdb(MstrNombreArchivo & ".mdb")
                ' Limpia la tabla Apolo
                GobjPanDat.SLimpieTabla(lcnnApolo, EnuProveedorBD.enuOleDb, "Apolo")
                Do While ldtmFechaMov <= DtmFechaHasta
                    lblnHayMovFecha = False
                    lblnHayMovFechaRC = False
                    lblnHayMovAnt = False
                    SProceseCajaBancosBD(ldtbCajaBancos, ldtmFechaMov, lcnnApolo, lblnHayMovFechaRC)
                    SProceseAnticiposBD(ldtbAnticipos, ldtmFechaMov, lcnnApolo, lblnHayMovFechaRC, lblnHayMovAnt)
                    lblnHayMovFecha = lblnHayMovFechaRC OrElse lblnHayMovAnt
                    lstrFiltro = ClsFechaNovedadDtm.SstrNombreCampoBd & " = '" & ldtmFechaMov & "'"
                    ldrwMovimientos = ldtbMovi.Select(lstrFiltro)
                    If Not lblnHayMovimiento Then
                        lblnHayMovimiento = lblnHayMovFecha
                    End If
                    If ldrwMovimientos.Length > 0 Then
                        If Not lblnHayMovFecha Then
                            mentConsecutivo = 0
                            MentIdComprobante += 1
                        End If
                        If Not lblnHayMovimiento Then
                            lblnHayMovimiento = True
                        End If
                        For Each ldrwMov As DataRow In ldrwMovimientos
                            lstrIdCtaDb = ClsPanorama.FobjValorCampo(ldrwMov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                                    EnuTipoValor.enuString)
                            lstrIdCtaCr = ClsPanorama.FobjValorCampo(ldrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                                    EnuTipoValor.enuString)
                            ldecValorMov = ClsPanorama.FobjValorCampo(ldrwMov(ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                            If lstrIdCtaDb = lstrCtaCaja OrElse
                                    GobjParametros.FblnEsCuentaBanco(lstrIdCtaDb) Then
                                SEscribaBdMovCr(ldrwMov, ldecValorMov, lcnnApolo, False)
                            Else
                                SProceseMovBD(ldrwMov, ldecValorMov, lstrIdCtaCr, lstrCtaCaja,
                                        lobjFactura, lcnnApolo)
                            End If
                        Next
                    End If
                    ldtmFechaMov = ldtmFechaMov.AddDays(1)
                Loop
            End Using
        If Not lblnHayMovimiento Then
            astrMens = "No hay Movimiento entre las Fechas dadas!"
        End If
    End Sub

    Private Sub SProceseMovBD(adrwMov As DataRow, ldecVlr As Decimal,
                astrIdCtaCr As String, astrCtaCaja As String,
                aobjFactura As ClsFactura, acnnApolo As Object)
        Dim lenuTipoDoc As EnuTipoDocOri
        SEscribaBdMovDb(adrwMov, ldecVlr, acnnApolo, False)
        If astrIdCtaCr = astrCtaCaja OrElse
                GobjParametros.FblnEsCuentaBanco(astrIdCtaCr) Then
            SEscribaBdMovCr(adrwMov, ldecVlr, acnnApolo, True)
        Else
            lenuTipoDoc = ClsPanorama.FobjValorCampo(adrwMov(ClsIdTipoDocOrigenByt.SstrNombreCampoBd),
                    EnuTipoValor.enuByte)
            If lenuTipoDoc = EnuTipoDocOri.enuFactura Then
                SEscribaBdFacCr(adrwMov, ldecVlr, acnnApolo,
                    aobjFactura)
            Else
                SEscribaBdMovCr(adrwMov, ldecVlr, acnnApolo, False)
            End If
        End If
    End Sub

    Private Sub SProceseCajaBancosBD(adtbCajaBancos As DataTable, adtmFecha As Date,
                acnnConeccion As Object, ByRef ablnHayMovFecha As Boolean) 'Ok
        Dim lblnHayMov = False
        Dim ldecValor As Decimal
        Dim lstrFiltro = ClsFechaNovedadDtm.SstrNombreCampoBd & " = '" & adtmFecha & "'"
        Dim ldrwCajaBancos = adtbCajaBancos.Select(lstrFiltro)
        If ldrwCajaBancos.Length > 0 Then
            Dim ldrwCajaBanco As DataRow, lstrIdCtaDb As String
            Dim lentIndiceDatRow As Integer, lblnDataRowAsignado As Boolean
            lblnHayMov = True
            MentIdComprobante += 1
            mentConsecutivo = 0
            For i = 0 To ldrwCajaBancos.Length - 1
                lblnDataRowAsignado = False
                ldecValor = 0
                lentIndiceDatRow = 0
                ldrwCajaBanco = ldrwCajaBancos(i)
                Dim lstrPrefijoDoc As String = CType(ClsPanorama.FobjValorCampo(ldrwCajaBanco(
                        ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString), String)
                If IsNothing(lstrPrefijoDoc) Then lstrPrefijoDoc = String.Empty
                Dim lentIdDoc As Integer = CType(ClsPanorama.FobjValorCampo(ldrwCajaBanco(
                        ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
                lstrIdCtaDb = ldrwCajaBanco(ClsIdCuentaDb_NovStr.SstrNombreCampoBd)
                Do While ClsPanorama.FobjValorCampo(ldrwCajaBanco(ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString) = lstrPrefijoDoc AndAlso
                        ClsPanorama.FobjValorCampo(ldrwCajaBanco(ClsIdDocOrigenEnt.SstrNombreCampoBd),
                        EnuTipoValor.enuInteger) = lentIdDoc AndAlso
                        ldrwCajaBanco(ClsIdCuentaDb_NovStr.SstrNombreCampoBd) = lstrIdCtaDb
                    If Not lblnDataRowAsignado Then
                        lblnDataRowAsignado = True
                        lentIndiceDatRow = i
                    End If
                    ldecValor += ClsPanorama.FobjValorCampo(ldrwCajaBanco(ClsValor_NovDec.SstrNombreCampoBd),
                            EnuTipoValor.enuDecimal)
                    i += 1
                    If i < ldrwCajaBancos.Length Then
                        ldrwCajaBanco = ldrwCajaBancos(i)
                    Else
                        Exit Do
                    End If
                Loop
                i -= 1
                ldrwCajaBanco = ldrwCajaBancos(lentIndiceDatRow)
                SEscribaBdMovDb(ldrwCajaBanco, ldecValor, acnnConeccion, True)
            Next
        End If
        ablnHayMovFecha = lblnHayMov
    End Sub

    Private Sub SProceseAnticiposBD(adtbAnticipos As DataTable, adtmFecha As Date,
                acnnConeccion As Object, ablnHayMovRC As Boolean, ByRef ablnHayMovAnt As Boolean) 'Ok TipoNov
        Dim lentIdAnticipo As Integer, lobjAnticipo As ClsAnticipo
        Dim lenuTipoDocOrigen As EnuTipoDocOri
        Dim ldecValor As Decimal, lenuIdTipoNovAnt As EnuTipoNov
        Dim lstrFiltro = ClsFechaNovedadDtm.SstrNombreCampoBd & " = '" & adtmFecha & "'"
        Dim ldrwAnticipos = adtbAnticipos.Select(lstrFiltro)
        Dim lblnHayMov = False
        If ldrwAnticipos.Length > 0 Then
            For Each ldrwAnt As DataRow In ldrwAnticipos
                lentIdAnticipo = ClsPanorama.FobjValorCampo(ldrwAnt(ClsIdAnticipo_NovEnt.SstrNombreCampoBd),
                        EnuTipoValor.enuLong)
                lobjAnticipo = New ClsAnticipo(EnuModoInstanciaObjDef.enuUnico)
                lobjAnticipo.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lentIdAnticipo})
                lenuTipoDocOrigen = lobjAnticipo.ObjIdTipoDocOrigen_AntByt.ObjValorPro
                If Not ablnHayMovRC Then
                    MentIdComprobante += 1
                    mentConsecutivo = 0
                End If
                lblnHayMov = True
                ldecValor = ClsPanorama.FobjValorCampo(ldrwAnt(ClsValor_NovAntDec.SstrNombreCampoBd),
                                            EnuTipoValor.enuDecimal)
                lenuIdTipoNovAnt = ClsPanorama.FobjValorCampo(ldrwAnt(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd),
                            EnuTipoValor.enuByte)
                If lenuIdTipoNovAnt = EnuTipoNov.EnuDbAntDev OrElse
                        lenuIdTipoNovAnt = EnuTipoNov.EnuRDbAntDev OrElse
                        lenuIdTipoNovAnt = EnuTipoNov.EnuRCrAntRec Then
                    SEscribaBdMovDbAnt(ldrwAnt, ldecValor, acnnConeccion, lenuTipoDocOrigen)
                    SEscribaBdMovCrAnt(ldrwAnt, ldecValor, acnnConeccion, lenuTipoDocOrigen)
                ElseIf lenuIdTipoNovAnt = EnuTipoNov.EnuCrAntRec Then
                    If lenuTipoDocOrigen = EnuTipoDocOri.EnuNotaAjuste Then
                        SEscribaBdMovDbAnt(ldrwAnt, ldecValor, acnnConeccion, lenuTipoDocOrigen)
                    End If
                    SEscribaBdMovCrAnt(ldrwAnt, ldecValor, acnnConeccion, lenuTipoDocOrigen)
                Else
                    Throw New ErrorInesperadoPanLException("Tipo de la Novedad del Anticipo no esperada!")
                End If
            Next
        End If
        ablnHayMovAnt = lblnHayMov
    End Sub
#End Region
#End Region
#Region "Comunes AP"
    Private Sub SLeaUltimoRegInteApoloAnterAP()
        If Not mblnLeidoUltimoRegistro Then
            Dim lstrUltimoArch = FstrUltimoArchInt("InterfazApolo*.txt")
            If Not String.IsNullOrEmpty(lstrUltimoArch) Then
                Dim lstrLinea = String.Empty
                Using lsrInterfaz = ClsPanorama.FsrStreamReader(lstrUltimoArch)
                    lstrLinea = lsrInterfaz.ReadLine
                    Do While Not IsNothing(lstrLinea)
                        If lsrInterfaz.EndOfStream Then
                            Exit Do
                        End If
                        lstrLinea = lsrInterfaz.ReadLine
                    Loop
                    If Not String.IsNullOrEmpty(lstrLinea) AndAlso lstrLinea.Contains(",") Then
                        Dim lstrPartes = lstrLinea.Split(",")
                        MdtmFechaFinInterfazAnterior = CType(lstrPartes(4), Date)
                        MentIdUltimoComprobanteInterfazAnterior = CType(lstrPartes(2), Integer)
                    End If
                End Using
                mblnLeidoUltimoRegistro = True
            End If
        End If
    End Sub
    Private Sub SEscribaAPMovDb(adrwMov As DataRow, adecValorMov As Decimal,
                   aswInterfaz As StreamWriter, ablnCajaBancos As Boolean)
        Dim lstrFechaMov As String = CType(ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate).Date, String)
        Dim lstrIdCtaDb As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim lstrIdTercero = String.Empty, lstrNombreCliente = String.Empty
        Dim ldecBase = 0D
        If ablnCajaBancos Then
            lstrIdTercero = FstrIdTerceroCajaBancos(adrwMov)
            lstrNombreCliente = FstrNombreTerceroCajaBancos(adrwMov)
        Else
            lstrIdTercero = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lstrNombreCliente = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        End If
        ldecBase = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        mentConsecutivo += 1
        With MstbLinea
            .Clear().Append(MstrTipoComprobante).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(MentIdComprobante).Append(CHSTRCOMA).Append(mentConsecutivo).Append(CHSTRCOMA)
            .Append(lstrFechaMov).Append(CHSTRCOMA).Append(lstrIdCtaDb).Append(CHSTRCOMA)
            If Not String.IsNullOrEmpty(lstrIdTercero) Then
                If BlnIdTerceroStr Then
                    .Append(lstrIdTercero.PadLeft(12, "0"))
                Else
                    .Append(lstrIdTercero)
                End If
            Else
                .Append(String.Empty)
            End If
            .Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(FstrDetalle(adrwMov)).Append(CHSTRCOMA).Append(adecValorMov).Append(CHSTRCOMA)
            .Append(0).Append(CHSTRCOMA).Append(ldecBase).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append(CHSTRCOMA).Append(0).Append(CHSTRCOMA).Append("Supervisor").Append(CHSTRCOMA)
            .Append("D").Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA).Append(lstrNombreCliente).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA)
            .Append(mstrNroInterfaz)
        End With
        aswInterfaz.WriteLine(MstbLinea.ToString)
    End Sub
    Private Sub SEscribaAPMovDbAnt(adrwMov As DataRow, adecValorMov As Decimal,
                aswInterfaz As StreamWriter, aenuTipoDocOri As EnuTipoDocOri)
        Dim lstrFechaMov As String = CType(ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate).Date, String)
        Dim lstrIdCtaDb As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim lstrIdTercero As String = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                EnuTipoValor.enuString)
        Dim lstrNombreCliente As String = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                EnuTipoValor.enuString)
        Dim ldecBase As Decimal = 0
        Dim lstrDetalle = String.Empty
        mentConsecutivo += 1
        With MstbLinea
            .Clear().Append(MstrTipoComprobante).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(MentIdComprobante).Append(CHSTRCOMA).Append(mentConsecutivo).Append(CHSTRCOMA)
            .Append(lstrFechaMov).Append(CHSTRCOMA).Append(lstrIdCtaDb).Append(CHSTRCOMA)
            If BlnIdTerceroStr Then
                .Append(lstrIdTercero.PadLeft(12, "0"))
            Else
                .Append(lstrIdTercero)
            End If
            .Append(CHSTRCOMA).Append(CHSTRCOMA)
            Select Case aenuTipoDocOri
                Case EnuTipoDocOri.enuReciboCaja
                    lstrDetalle = FstrDetalle(adrwMov)
                Case EnuTipoDocOri.enuNotaAjuste
                    lstrDetalle = "Anticipo generado por Ajuste Cuotas de Administración."
            End Select
            .Append(lstrDetalle).Append(CHSTRCOMA).Append(adecValorMov).Append(CHSTRCOMA)
            .Append(0).Append(CHSTRCOMA).Append(ldecBase).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append(CHSTRCOMA).Append(0).Append(CHSTRCOMA).Append("Supervisor").Append(CHSTRCOMA)
            .Append("D").Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA).Append(lstrNombreCliente).Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append(mstrNroInterfaz)
        End With
        aswInterfaz.WriteLine(MstbLinea.ToString)
    End Sub
    Private Sub SEscribaAPMovCr(adrwMov As DataRow, adecValorMov As Decimal,
                aswInterfaz As StreamWriter, ablnCajaBancos As Boolean)
        Dim lstrFechaMov As String = CType(ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                            EnuTipoValor.enuDate).Date, String)
        Dim lstrIdCtaCr As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim lstrIdTercero = String.Empty, lstrNombreCliente = String.Empty
        Dim ldecBase = 0D
        If ablnCajaBancos Then
            lstrIdTercero = FstrIdTerceroCajaBancos(adrwMov)
            lstrNombreCliente = FstrNombreTerceroCajaBancos(adrwMov)
        Else
            lstrIdTercero = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
            lstrNombreCliente = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        End If
        ldecBase = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        mentConsecutivo += 1
        With MstbLinea
            .Clear().Append(MstrTipoComprobante).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(MentIdComprobante).Append(CHSTRCOMA).Append(mentConsecutivo).Append(CHSTRCOMA)
            .Append(lstrFechaMov).Append(CHSTRCOMA).Append(lstrIdCtaCr).Append(CHSTRCOMA)
            If BlnIdTerceroStr Then
                .Append(lstrIdTercero.PadLeft(12, "0"))
            Else
                .Append(lstrIdTercero)
            End If
            .Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(FstrDetalle(adrwMov)).Append(CHSTRCOMA).Append(0).Append(CHSTRCOMA)
            .Append(adecValorMov).Append(CHSTRCOMA).Append(ldecBase).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append(CHSTRCOMA).Append(0).Append(CHSTRCOMA).Append("Supervisor").Append(CHSTRCOMA)
            .Append("C").Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA).Append(lstrNombreCliente).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA)
            .Append(mstrNroInterfaz)
        End With
        aswInterfaz.WriteLine(MstbLinea.ToString)
    End Sub
    Private Sub SEscribaAPFacCr(adrwMov As DataRow, adecValorMov As Decimal,
                aswInterfaz As StreamWriter, aobjFactura As ClsFactura)
        Dim lstrFechaMov As String = CType(ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                            EnuTipoValor.enuDate).Date, String)
        Dim lstrIdCtaCr As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim lstrNombreCliente = String.Empty
        Dim ldecBase = 0D
        Dim lstrPrefijoFac As String = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString), String)
        If IsNothing(lstrPrefijoFac) Then lstrPrefijoFac = String.Empty
        Dim lentIdFac As Integer = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
        Dim lshrIdItemFac As Short = ClsPanorama.FobjValorCampo(adrwMov(ClsIdItemFacturaShr.SstrNombreCampoBd),
                EnuTipoValor.enuShort)
        Dim lobjValorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefijoFac, lentIdFac}
        aobjFactura.SAbra(lobjValorLlave)
        Dim lstrIdTerceroCtaCr = aobjFactura.FstrIdTerceroCtaCr(lshrIdItemFac)
        If String.IsNullOrEmpty(lstrIdTerceroCtaCr) Then
            lstrIdTerceroCtaCr = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                            EnuTipoValor.enuString)
            lstrNombreCliente = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        Else
            lstrNombreCliente = ClsPanorama.FstrNombreTercero(lstrIdTerceroCtaCr)
        End If
        ldecBase = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        mentConsecutivo += 1
        With MstbLinea
            .Clear().Append(MstrTipoComprobante).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(MentIdComprobante).Append(CHSTRCOMA).Append(mentConsecutivo).Append(CHSTRCOMA)
            .Append(lstrFechaMov).Append(CHSTRCOMA).Append(lstrIdCtaCr).Append(CHSTRCOMA)
            If BlnIdTerceroStr Then
                .Append(lstrIdTerceroCtaCr.PadLeft(12, "0"))
            Else
                .Append(lstrIdTerceroCtaCr)
            End If
            .Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(FstrDetalle(adrwMov)).Append(CHSTRCOMA).Append(0).Append(CHSTRCOMA)
            .Append(adecValorMov).Append(CHSTRCOMA).Append(ldecBase).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append(CHSTRCOMA).Append(0).Append(CHSTRCOMA).Append("Supervisor").Append(CHSTRCOMA)
            .Append("C").Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA).Append(lstrNombreCliente).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA)
            .Append(mstrNroInterfaz)
        End With
        aswInterfaz.WriteLine(MstbLinea.ToString)
    End Sub
    Private Sub SEscribaAPMovCrAnt(adrwMov As DataRow, adecValorMov As Decimal,
                aswInterfaz As StreamWriter, aenuTipoDocOri As EnuTipoDocOri)
        Dim lstrFechaMov As String = CType(ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadAntDtm.SstrNombreCampoBd),
                            EnuTipoValor.enuDate).Date, String)
        Dim lstrIdCtaCr As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim lstrIdTercero = String.Empty
        Dim lstrNombreCliente = String.Empty
        Dim lblnEsBancCaja = (lstrIdCtaCr = GobjParametros.ObjIdCtaCajaStr.ObjValorPro) OrElse
                GobjParametros.FblnEsCuentaBanco(lstrIdCtaCr)
        If lblnEsBancCaja Then
            lstrIdTercero = FstrIdTerceroCajaBancos(adrwMov)
            lstrNombreCliente = FstrNombreTerceroCajaBancos(adrwMov)
        Else
            lstrIdTercero = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovAntStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lstrNombreCliente = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        End If
        Dim lstrDetalle = String.Empty
        Dim ldecBase As Decimal = 0
        mentConsecutivo += 1
        With MstbLinea
            .Clear().Append(MstrTipoComprobante).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(MentIdComprobante).Append(CHSTRCOMA).Append(mentConsecutivo).Append(CHSTRCOMA)
            .Append(lstrFechaMov).Append(CHSTRCOMA).Append(lstrIdCtaCr).Append(CHSTRCOMA)
            If BlnIdTerceroStr Then
                .Append(lstrIdTercero.PadLeft(12, "0"))
            Else
                .Append(lstrIdTercero)
            End If
            .Append(CHSTRCOMA).Append(CHSTRCOMA)
            Select Case aenuTipoDocOri
                Case EnuTipoDocOri.enuReciboCaja
                    lstrDetalle = FstrDetalle(adrwMov)
                Case EnuTipoDocOri.enuNotaAjuste
                    lstrDetalle = "Anticipo generado por Ajuste Cuotas de Administración."
            End Select
            .Append(lstrDetalle).Append(CHSTRCOMA).Append(0).Append(CHSTRCOMA)
            .Append(adecValorMov).Append(CHSTRCOMA).Append(ldecBase).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append(CHSTRCOMA).Append(0).Append(CHSTRCOMA).Append("Supervisor").Append(CHSTRCOMA)
            .Append("C").Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA).Append(lstrNombreCliente).Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append(mstrNroInterfaz)
        End With
        aswInterfaz.WriteLine(MstbLinea.ToString)
    End Sub
#End Region
#Region "Comunes BD"
    Private Sub SLeaUltimoRegInteApoloAnterBD()
        If Not mblnLeidoUltimoRegistro Then
            Dim lstrUltimoArch = FstrUltimoArchInt("InterfazApolo*.mdb")
            Dim lstrSql = "SELECT Tipo, Numero, Secuencia, Fecha FROM Apolo ORDER BY Numero ASC, Secuencia ASC"
            If Not String.IsNullOrEmpty(lstrUltimoArch) Then
                Using ldstDataSet As New DataSet
                    GobjPanDat.SdstDataSetAccess(ldstDataSet, lstrUltimoArch, "", lstrSql)
                    MdtbInterfaz = ldstDataSet.Tables(0)
                    If MdtbInterfaz.Rows.Count > 0 Then
                        Dim ldrwUltimoReg As DataRow = MdtbInterfaz.Rows(MdtbInterfaz.Rows.Count - 1)
                        MentIdUltimoComprobanteInterfazAnterior = ClsPanorama.FobjValorCampo(ldrwUltimoReg(1),
                                EnuTipoValor.enuInteger)
                        MdtmFechaFinInterfazAnterior = ClsPanorama.FobjValorCampo(ldrwUltimoReg(3),
                                EnuTipoValor.enuDate)
                    End If
                End Using
                mblnLeidoUltimoRegistro = True
            End If
        End If
    End Sub
    Private Sub SEscribaBdMovDb(adrwMov As DataRow, adecValorMov As Decimal,
                acnnConexion As Object, ablnCajaBancos As Boolean)
        Dim lcolCampos As Collection = FcolCampos(), lcolDatos As New Collection
        Dim ldtmFechaMov As Date = ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate)
        Dim lstrIdCtaDb As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim lstrIdTercero = String.Empty, lstrNombreCliente = String.Empty, ldecBase = 0D
        If ablnCajaBancos Then
            lstrIdTercero = FstrIdTerceroCajaBancos(adrwMov)
            lstrNombreCliente = FstrNombreTerceroCajaBancos(adrwMov)
        Else
            lstrIdTercero = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lstrNombreCliente = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        End If
        ldecBase = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        mentConsecutivo += 1
        With lcolDatos
            .Add(MstrTipoComprobante)
            .Add("")
            .Add(MentIdComprobante)
            .Add(mentConsecutivo)
            .Add(ldtmFechaMov)
            .Add(lstrIdCtaDb)
            If BlnIdTerceroStr Then
                .Add(lstrIdTercero.ToString.PadLeft(12, "0"))
            Else
                .Add(lstrIdTercero)
            End If
            .Add(String.Empty)
            .Add(FstrDetalle(adrwMov))
            .Add(adecValorMov)
            .Add(0)
            .Add(ldecBase)
            .Add("Supervisor")
            .Add("D")
            .Add(lstrNombreCliente)
            .Add(String.Empty)
            .Add(mstrNroInterfaz)
        End With
        GobjPanDat.SInserteRegistro(acnnConexion, EnuProveedorBD.enuOleDb, "Apolo", lcolCampos, lcolDatos)
    End Sub
    Private Sub SEscribaBdMovDbAnt(adrwMov As DataRow, adecValorMov As Decimal,
                acnnConexion As Object, aenuOrigenAnticipo As EnuTipoDocOri)
        Dim lcolCampos As Collection = FcolCampos(), lcolDatos As New Collection
        Dim ldtmFechaMov As Date = ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate)
        Dim lstrIdCtaDb As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim lstrIdTercero As String = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovAntStr.SstrNombreCampoBd),
                EnuTipoValor.enuString)
        Dim lstrNombreCliente As String = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                EnuTipoValor.enuString)
        Dim lstrDetalle = String.Empty
        mentConsecutivo += 1
        With lcolDatos
            .Add(MstrTipoComprobante)
            .Add("")
            .Add(MentIdComprobante)
            .Add(mentConsecutivo)
            .Add(ldtmFechaMov)
            .Add(lstrIdCtaDb)
            If BlnIdTerceroStr Then
                .Add(lstrIdTercero.PadLeft(12, "0"))
            Else
                .Add(lstrIdTercero)
            End If
            .Add(String.Empty)
            Select Case aenuOrigenAnticipo
                Case EnuTipoDocOri.enuReciboCaja
                    lstrDetalle = FstrDetalle(adrwMov)
                Case EnuTipoDocOri.enuNotaAjuste
                    lstrDetalle = "Anticipo generado por Ajuste Cuotas de Administración."
            End Select
            .Add(lstrDetalle)
            .Add(adecValorMov)
            .Add(0)
            .Add(0)
            .Add("Supervisor")
            .Add("D")
            .Add(lstrNombreCliente)
            .Add("")
            .Add(mstrNroInterfaz)
        End With
        GobjPanDat.SInserteRegistro(acnnConexion, EnuProveedorBD.enuOleDb, "Apolo", lcolCampos, lcolDatos)
    End Sub
    Private Sub SEscribaBdMovCr(adrwMov As DataRow, adecValorMov As Decimal,
                acnnConexion As Object, ablnCajaBancos As Boolean)
        Dim lcolCampos As Collection = FcolCampos(), lcolDatos As New Collection
        Dim ldtmFechaMov As Date = ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                            EnuTipoValor.enuDate)
        Dim lstrIdCtaCr As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim lstrIdTercero = String.Empty, lstrNombreCliente = String.Empty, ldecBase = 0D
        If ablnCajaBancos Then
            lstrIdTercero = FstrIdTerceroCajaBancos(adrwMov)
            lstrNombreCliente = FstrNombreTerceroCajaBancos(adrwMov)
        Else
            lstrIdTercero = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
            lstrNombreCliente = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                EnuTipoValor.enuString)
        End If
        ldecBase = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
            EnuTipoValor.enuDecimal)
        mentConsecutivo += 1
        With lcolDatos
            .Add(MstrTipoComprobante)
            .Add("")
            .Add(MentIdComprobante)
            .Add(mentConsecutivo)
            .Add(ldtmFechaMov)
            .Add(lstrIdCtaCr)
            If BlnIdTerceroStr Then
                .Add(lstrIdTercero.PadLeft(12, "0"))
            Else
                .Add(lstrIdTercero)
            End If
            .Add(String.Empty)
            .Add(FstrDetalle(adrwMov))
            .Add(0)
            .Add(adecValorMov)
            .Add(ldecBase)
            .Add("Supervisor")
            .Add("C")
            .Add(lstrNombreCliente)
            .Add(String.Empty)
            .Add(mstrNroInterfaz)
        End With
        GobjPanDat.SInserteRegistro(acnnConexion, EnuProveedorBD.enuOleDb, "Apolo", lcolCampos, lcolDatos)
    End Sub
    Private Sub SEscribaBdFacCr(adrwMov As DataRow, adecValorMov As Decimal,
                acnnConexion As Object, aobjFactura As ClsFactura)
        Dim lcolCampos As Collection = FcolCampos(), lcolDatos As New Collection
        Dim ldtmFechaMov As Date = ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                            EnuTipoValor.enuDate)
        Dim lstrIdCtaCr As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim lstrNombreCliente = String.Empty, ldecBase = 0D
        Dim lstrPrefijoFac As String = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString), String)
        If IsNothing(lstrPrefijoFac) Then lstrPrefijoFac = String.Empty
        Dim lentIdFac As Integer = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
        Dim lshrIdItemFac As Short = ClsPanorama.FobjValorCampo(adrwMov(ClsIdItemFacturaShr.SstrNombreCampoBd),
                EnuTipoValor.enuShort)
        Dim lobjValorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefijoFac, lentIdFac}
        aobjFactura.SAbra(lobjValorLlave)
        Dim lstrIdTerceroCtaCr = aobjFactura.FstrIdTerceroCtaCr(lshrIdItemFac)
        If String.IsNullOrEmpty(lstrIdTerceroCtaCr) Then
            lstrIdTerceroCtaCr = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                            EnuTipoValor.enuString)
            lstrNombreCliente = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        Else
            lstrNombreCliente = ClsPanorama.FstrNombreTercero(lstrIdTerceroCtaCr)
        End If
        ldecBase = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
            EnuTipoValor.enuDecimal)
        mentConsecutivo += 1
        With lcolDatos
            .Add(MstrTipoComprobante)
            .Add("")
            .Add(MentIdComprobante)
            .Add(mentConsecutivo)
            .Add(ldtmFechaMov)
            .Add(lstrIdCtaCr)
            If BlnIdTerceroStr Then
                .Add(lstrIdTerceroCtaCr.PadLeft(12, "0"))
            Else
                .Add(lstrIdTerceroCtaCr)
            End If
            .Add(String.Empty)
            .Add(FstrDetalle(adrwMov))
            .Add(0)
            .Add(adecValorMov)
            .Add(ldecBase)
            .Add("Supervisor")
            .Add("C")
            .Add(lstrNombreCliente)
            .Add(String.Empty)
            .Add(mstrNroInterfaz)
        End With
        GobjPanDat.SInserteRegistro(acnnConexion, EnuProveedorBD.enuOleDb, "Apolo", lcolCampos, lcolDatos)
    End Sub
    Private Sub SEscribaBdMovCrAnt(adrwMov As DataRow, adecValorMov As Decimal,
                acnnConexion As Object, aenuOrigenAnticipo As EnuTipoDocOri)
        Dim lcolCampos As Collection = FcolCampos(), lcolDatos As New Collection
        Dim ldtmFechaMov As Date = ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                            EnuTipoValor.enuDate)
        Dim lstrIdCtaCr As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim lstrIdTercero = String.Empty
        Dim lstrNombreCliente = String.Empty
        Dim lblnEsBancCaja = (lstrIdCtaCr = GobjParametros.ObjIdCtaCajaStr.ObjValorPro) OrElse
                GobjParametros.FblnEsCuentaBanco(lstrIdCtaCr)
        If lblnEsBancCaja Then
            lstrIdTercero = FstrIdTerceroCajaBancos(adrwMov)
            lstrNombreCliente = FstrNombreTerceroCajaBancos(adrwMov)
        Else
            lstrIdTercero = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovAntStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lstrNombreCliente = ClsPanorama.FobjValorCampo(adrwMov(ClsNombreCompletoStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        End If
        Dim lstrDetalle = String.Empty
        mentConsecutivo += 1
        With lcolDatos
            .Add(MstrTipoComprobante)
            .Add("")
            .Add(MentIdComprobante)
            .Add(mentConsecutivo)
            .Add(ldtmFechaMov)
            .Add(lstrIdCtaCr)
            If BlnIdTerceroStr Then
                .Add(lstrIdTercero.PadLeft(12, "0"))
            Else
                .Add(lstrIdTercero)
            End If
            .Add("")
            Select Case aenuOrigenAnticipo
                Case EnuTipoDocOri.enuReciboCaja
                    lstrDetalle = FstrDetalle(adrwMov)
                Case EnuTipoDocOri.enuNotaAjuste
                    lstrDetalle = "Anticipo generado por Ajuste Cuotas de Administración."
            End Select
            .Add(lstrDetalle)
            .Add(0)
            .Add(adecValorMov)
            .Add(0)
            .Add("Supervisor")
            .Add("C")
            .Add(lstrNombreCliente)
            .Add("")
            .Add(mstrNroInterfaz)
        End With
        GobjPanDat.SInserteRegistro(acnnConexion, EnuProveedorBD.enuOleDb, "Apolo", lcolCampos, lcolDatos)
    End Sub
    Private Shared Function FcolCampos() As Collection
        Dim lcolCampos As New Collection
        With lcolCampos
            .Add("Tipo")
            .Add("Prefijo")
            .Add("Numero")
            .Add("Secuencia")
            .Add("Fecha")
            .Add("Cuenta")
            .Add("Tercero")
            .Add("Centro")
            .Add("Detalle")
            .Add("Debito")
            .Add("Credito")
            .Add("Base")
            .Add("Usuario")
            .Add("Signo")
            .Add("NombreTercero")
            .Add("NombreCentro")
            .Add("InterfaceNo")
        End With
        Return lcolCampos
    End Function
    Private Shared Sub SCreeBdApolo(astrNombreBd As String)
        If Not My.Computer.FileSystem.FileExists(astrNombreBd) Then
            Dim lstrArchiOri = GstrTrayInterfContable & "\Apolo.mdb"
            If My.Computer.FileSystem.FileExists(lstrArchiOri) Then
                My.Computer.FileSystem.CopyFile(lstrArchiOri, astrNombreBd)
            Else
                Throw New ErrorInesperadoPanLException("No existe el archivo Apolo.mdb")
            End If
        End If
    End Sub
    Private Shared Function FcnnApoloMdb(astrNombreBd As String) As Object
        Dim lobjCnnDat As New ClsCnnDat
        Dim lcnnApolo = lobjCnnDat.FcnnConexion(EnuProveedorBD.enuOleDb, EnuTipoAutenticacion.None, "",
                GstrTrayInterfContable, astrNombreBd, "", "", "", False, False)
        lcnnApolo.Open()
        Return lcnnApolo
    End Function
#End Region
#Region "Generales"
    Friend Overrides ReadOnly Property DtmFechaFinInterfazAnterior() As Date
        Get
            If Not mblnLeidoUltimoRegistro Then
                Select Case GobjParametros.ObjIdAppContableByt.ObjValorPro
                    Case EnuAppConta.enuApoloAP
                        SLeaUltimoRegInteApoloAnterAP()
                    Case EnuAppConta.enuApoloBD
                        SLeaUltimoRegInteApoloAnterBD()
                    Case Else
                        Throw New ErrorInesperadoPanLException("Tipo de App Contable no esperado")
                End Select
            End If
            Return MdtmFechaFinInterfazAnterior
        End Get
    End Property
    Friend Overrides Function FentIdUltimoComprobanteInterfazAnterior() As Integer
        If Not mblnLeidoUltimoRegistro Then
            Select Case GobjParametros.ObjIdAppContableByt.ObjValorPro
                Case EnuAppConta.enuApoloAP
                    SLeaUltimoRegInteApoloAnterAP()
                Case EnuAppConta.enuApoloBD
                    SLeaUltimoRegInteApoloAnterBD()
                Case Else
                    Throw New ErrorInesperadoPanLException("Tipo de App Contable no esperado")
            End Select
        End If
        Return MentIdUltimoComprobanteInterfazAnterior
    End Function
    Private Shared Function FstrFiltroCajaBancoNov() As String
        Dim lstrCtaCaja As String = "'" & GobjParametros.ObjIdCtaCajaStr.ObjValorPro & "'"
        Dim lstrFiltro = "(" & ClsIdCuentaDb_NovStr.SstrNombreCampoBd & " = " & lstrCtaCaja
        Dim lcolCuentasBanco = GobjParametros.FcolCuentasBanco
        For Each lobjCtaBanco As ClsCuentaBanco In lcolCuentasBanco
            Dim lstrCtaBanco = "'" & lobjCtaBanco.ObjIdCtaContabilidadStr.ObjValorPro & "'"
            lstrFiltro &= " OR " & ClsIdCuentaDb_NovStr.SstrNombreCampoBd & " = " & lstrCtaBanco
        Next
        lstrFiltro &= ")"
        Return lstrFiltro
    End Function
    Private Shared Function FstrFiltroCajaBancoNovAnt() As String
        Dim lstrCtaCaja As String = "'" & GobjParametros.ObjIdCtaCajaStr.ObjValorPro & "'"
        Dim lstrFiltro = "(" & ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd & " = " & lstrCtaCaja
        Dim lcolCuentasBanco = GobjParametros.FcolCuentasBanco
        For Each lobjCtaBanco As ClsCuentaBanco In lcolCuentasBanco
            Dim lstrCtaBanco = "'" & lobjCtaBanco.ObjIdCtaContabilidadStr.ObjValorPro & "'"
            lstrFiltro &= " OR " & ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd & " = " & lstrCtaBanco
        Next
        lstrFiltro &= ")"
        Return lstrFiltro
    End Function
#End Region
End Class
