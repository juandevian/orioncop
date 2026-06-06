Imports System.IO
Imports System.Text
Friend Class ClsInterfazPodium
#Region "Definiciones"
    Inherits ClsCBInterfazContableOri
    'Constantes
    'Variables
    Private MstrNombreArchivo As String = String.Empty
    Private ReadOnly MstbLinea As New StringBuilder
    Private MblnLeidoUltimoRegistro As Boolean = False
    Private MdtmFechaFinInterfazAnterior As Date = GCDTMFECHANULA
#End Region
#Region "Constructores"
    Friend Sub New(aobjRegistro As Object)
        MyBase.New(aobjRegistro)
    End Sub
#End Region
#Region "Interfaz"
    Friend Overrides Sub SGenerereInterfazContable(ablnFinMes As Boolean, ByRef astrMens As String)
        Dim lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            If Not FblnEstanTodosOkInterfaz() Then
                Throw New ValorArgumentoInvalidoException("Hay argumentos no validos")
            End If
            If ablnFinMes Then
                MstrNombreArchivo = "InterfazPodium_FM_" & ClsPanorama.FstrFechayyyymmdd(
                        DtmFechaDesde) & "_" & ClsPanorama.FstrFechayyyymmdd(DtmFechaHasta)
            Else
                MstrNombreArchivo = "InterfazPodium_" & ClsPanorama.FstrFechayyyymmdd(
                        DtmFechaDesde) & "_" & ClsPanorama.FstrFechayyyymmdd(DtmFechaHasta)
            End If
            HstrArchivoSalidaInterfaz = GstrTrayInterfContable & "\" & MstrNombreArchivo
            SGenereIntPodiumDoc()
            MdtmFechaFinInterfazAnterior = GCDTMFECHANULA
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PathTooLongException
            Throw
        Catch ex As DirectoryNotFoundException
            Throw
        Catch ex As IOException
            astrMens = "No se generó la interfaz debido a que el archivo de destino está abierto!"
        Catch ex As ArgumentException
            Throw
        Catch ex As NotSupportedException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
                astrMens = "La Interfaz Contable se generó exitosamente!"
            Else
                GobjPanDat.SControleProcesoObj(False, True)
            End If
            If Not String.IsNullOrEmpty(astrMens) Then
                MsgBox(astrMens, vbOKOnly, "Información")
            End If
        End Try
    End Sub
    Private Sub SGenereIntPodiumDoc()
        Dim ldtbMovi = FdtbMovimiento()
        Using lswInterfaz = ClsPanorama.FswStreamWriter(HstrArchivoSalidaInterfaz & ".txt")
            SProceseDoc(ldtbMovi, EnuIdDocumentoDef.EnuFacturaVenta, lswInterfaz)
            SProceseDoc(ldtbMovi, EnuIdDocumentoDef.EnuNotaIntMora, lswInterfaz)
            SProceseRecibosCaja(lswInterfaz)
            SProceseDoc(ldtbMovi, EnuIdDocumentoDef.EnuNotaAplicacionAnt, lswInterfaz)
            SProceseDoc(ldtbMovi, EnuIdDocumentoDef.EnuNotaCr, lswInterfaz)
            SProceseNotaR(EnuIdDocumentoDef.EnuNotaReintegroAnt, lswInterfaz)
            SProceseDoc(ldtbMovi, EnuIdDocumentoDef.EnuNotaReversaCr, lswInterfaz)
            SProceseNotaR(EnuIdDocumentoDef.EnuNotaReversaCr, lswInterfaz)
            SProceseNotasAjuste(lswInterfaz)
        End Using
    End Sub
    Private Sub SProceseDoc(adtbMovi As DataTable, aenuIdDocumento As EnuIdDocumentoDef,
                            aswInterfaz As StreamWriter)
        Dim ldecValor As Decimal
        Dim lstrPref As String, lentIdFra As Integer, lstrNroFac As String, lstrNroFacAnt = String.Empty
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(aenuIdDocumento)
        Dim lstrFiltro = ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " & aenuIdDocumento
        Dim ldrwNovedades As DataRow() = adtbMovi.Select(lstrFiltro)
        For Each ldrwNov As DataRow In ldrwNovedades
            If aenuIdDocumento = EnuIdDocumentoDef.EnuFacturaVenta Then
                lstrPref = CType(ClsPanorama.FobjValorCampo(ldrwNov(
                    ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString), String)
                If IsNothing(lstrPref) Then lstrPref = String.Empty
                lentIdFra = CType(ClsPanorama.FobjValorCampo(ldrwNov(
                        ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
                lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPref, lentIdFra)
                If lstrNroFacAnt <> lstrNroFac Then
                    SProceseFactura(adtbMovi, lstrPref, lentIdFra, aswInterfaz)
                    lstrNroFacAnt = lstrNroFac
                End If
            Else
                ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                        EnuTipoValor.enuDecimal)
                SEscribaMovDocDb(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz)
                SEscribaMovDocCr(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz)
            End If
        Next
    End Sub
    Private Sub SProceseFactura(adtbMovi As DataTable, astrPref As String,
            aentIdFact As Integer, aswInterfaz As StreamWriter)
        Dim lstrFiltro = ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " &
                EnuIdDocumentoDef.EnuFacturaVenta & " AND " &
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" & astrPref &
                "' AND " & ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " & aentIdFact
        Dim ldrwNovsFra As DataRow() = adtbMovi.Select(lstrFiltro), ldrwNov As DataRow
        Dim ldecValorFra = 0D, ldecValor As Decimal, lstrIdCtaDb = String.Empty, i = 0
        Dim lshrIdItemFac As Short, lstrIdTerceroCtaCr As String
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.EnuFacturaVenta)
        Dim ldrwNovDbFra As DataRow = Nothing
        Dim lenuTipoNovDbFra As EnuTipoNov
        Dim lobjFactura As New ClsFactura()
        Dim lobjValorLlave As Object()
        Do
            ldrwNov = ldrwNovsFra(i)
            lenuTipoNovDbFra = ClsPanorama.FobjValorCampo(ldrwNov(
                    ClsIdTipoNovedadByt.SstrNombreCampoBd), EnuTipoValor.enuByte)
            If lstrIdCtaDb <> ClsPanorama.FobjValorCampo(ldrwNov(
                     ClsIdCuentaDb_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString) Then
                If Not String.IsNullOrEmpty(lstrIdCtaDb) Then
                    SEscribaMovDocDb(ldrwNovsFra(i - 1), lstrTipoDoc, ldecValorFra, aswInterfaz)
                End If
                lstrIdCtaDb = ClsPanorama.FobjValorCampo(ldrwNov(
                        ClsIdCuentaDb_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                ldecValorFra = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                            EnuTipoValor.enuDecimal)
            Else
                ldecValorFra += ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                            EnuTipoValor.enuDecimal)
            End If
            If lenuTipoNovDbFra = EnuTipoNov.EnuDbCap Then
                ldrwNovDbFra = ldrwNov
            End If
            If i = ldrwNovsFra.Length - 1 Then
                SEscribaMovDocDb(ldrwNovDbFra, lstrTipoDoc, ldecValorFra, aswInterfaz)
            End If
            i += 1
        Loop While i < ldrwNovsFra.Length
        For Each ldrwNovFac As DataRow In ldrwNovsFra
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNovFac(ClsValor_NovDec.SstrNombreCampoBd),
                        EnuTipoValor.enuDecimal)
            lshrIdItemFac = ClsPanorama.FobjValorCampo(ldrwNovFac(ClsIdItemFacturaShr.SstrNombreCampoBd),
                    EnuTipoValor.enuShort)
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, astrPref, aentIdFact}
            lobjFactura.SAbra(lobjValorLlave)
            lstrIdTerceroCtaCr = lobjFactura.FstrIdTerceroCtaCr(lshrIdItemFac)
            If String.IsNullOrEmpty(lstrIdTerceroCtaCr) Then
                SEscribaMovDocCr(ldrwNovFac, lstrTipoDoc, ldecValor, aswInterfaz)
            Else
                SEscribaMovCrFac(ldrwNovFac, lstrTipoDoc, ldecValor, aswInterfaz, lstrIdTerceroCtaCr)
            End If
        Next
    End Sub
    Private Sub SProceseRecibosCaja(aswInterfaz As StreamWriter)
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.EnuReciboCaja)
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
        SEscribaMovDocDb(adrwMoviRC(0), astrTipoDoc, ldecValor, aswInterfaz)
        For Each ldrwMovNov As DataRow In adrwMoviRC
            ldecValorNov = ClsPanorama.FobjValorCampo(ldrwMovNov(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            SEscribaMovDocCr(ldrwMovNov, astrTipoDoc, ldecValorNov, aswInterfaz)
        Next
    End Sub
    Private Sub SProceseNotaR(aenuIdDocumento As EnuIdDocumentoDef, aswInterfaz As StreamWriter)
        Dim ldecValor As Decimal
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(aenuIdDocumento)
        Dim ldtbMovAntR = FdtbMoviAnt(aenuIdDocumento)
        For Each ldrwNov As DataRow In ldtbMovAntR.Rows
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            SEscribaMovDocDb(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz)
            SEscribaMovDocCr(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz)
        Next
    End Sub
    Private Sub SProceseNotasAjuste(aswInterfaz As StreamWriter)
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.EnuNotaAjuste)
        Dim ldecValor As Decimal
        Dim ldtbMovNA = FdtbMoviNotaAjuste()
        For Each ldrwNov As DataRow In ldtbMovNA.Rows
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                        EnuTipoValor.enuDecimal)
            SEscribaMovDocDb(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz)
            SEscribaMovDocCr(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz)
        Next
    End Sub
    Private Sub SEscribaMovDocDb(adrwMov As DataRow, astrTipoDoc As String,
            adecValor As Decimal, aswInterfaz As StreamWriter)
        Dim lstrPrefFact As String = String.Empty
        Dim lentIdFactura As Integer = ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdFactura_NovEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
        Dim lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPrefFact, lentIdFactura)
        Dim lstrPrefijoDoc As String = String.Empty
        Dim lentIdDoc As Integer = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
        Dim lstrNroDoc As String = ClsPanorama.FstrNumeroDcto(lstrPrefijoDoc, lentIdDoc)
        Dim lstrFechaMov As String = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsFechaNovedadDtm.SstrNombreCampoBd), EnuTipoValor.enuDate).Date, String)
        Dim lstrIdCtaDb As String = ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdCuentaDb_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        Dim ldecBase As Decimal = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        Dim lstrIdTercero As String = ClsPanorama.FobjValorCampo(adrwMov(
                ClsAliasCont_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        Dim lstrPredioAgr As String = ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        Dim lstrRef As String
        If String.IsNullOrEmpty(lstrPredioAgr) Then
            lstrRef = ClsPanorama.FobjValorCampo(adrwMov(
                ClsNombreCompletoStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        Else
            lstrRef = lstrPredioAgr
        End If
        lstrRef = FstrRef(lstrRef)
        If astrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.EnuFacturaVenta) Then
            ldecBase = 0
        End If
        Dim lblnEsCajaBanco = (lstrIdCtaDb = GobjParametros.ObjIdCtaCajaStr.ObjValorPro) OrElse
                GobjParametros.FblnEsCuentaBanco(lstrIdCtaDb)
        Dim lblnDbAPasivo = (lstrIdCtaDb.StartsWith("2"))
        If lblnEsCajaBanco Then
            lstrIdTercero = FstrIdTerceroCajaBancos(adrwMov)
        End If
        If Not String.IsNullOrEmpty(lstrIdTercero) Then
            If BlnIdTerceroStr Then
                lstrIdTercero = lstrIdTercero.PadLeft(12, "0")
            End If
        End If
        ' Db
        With MstbLinea
            .Clear().Append(astrTipoDoc).Append(CHSTRCOMA).Append(lstrNroDoc).Append(CHSTRCOMA)
            .Append(lstrFechaMov).Append(CHSTRCOMA).Append(lstrIdCtaDb).Append(CHSTRCOMA)
            .Append(FstrDetalle(adrwMov)).Append(CHSTRCOMA).Append("1").Append(CHSTRCOMA)
            .Append(adecValor).Append(CHSTRCOMA).Append(lstrIdTercero).Append(CHSTRCOMA)
            .Append(ldecBase).Append(CHSTRCOMA).Append(CHSTRCOMA).Append(lstrRef)
        End With
        aswInterfaz.WriteLine(MstbLinea.ToString)
    End Sub
    Private Sub SEscribaMovDocCr(adrwMov As DataRow, astrTipoDoc As String,
            adecValor As Decimal, aswInterfaz As StreamWriter)
        Dim lstrPrefFact As String = String.Empty
        Dim lentIdFactura As Integer = ClsPanorama.FobjValorCampo(adrwMov(ClsIdFactura_NovEnt.SstrNombreCampoBd),
                EnuTipoValor.enuInteger)
        Dim lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPrefFact, lentIdFactura)
        Dim lstrPrefijoDoc As String = String.Empty
        Dim lentIdDoc As Integer = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
        Dim lstrNroDoc As String = ClsPanorama.FstrNumeroDcto(lstrPrefijoDoc, lentIdDoc)
        Dim lstrFechaMov As String = CType(ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
            EnuTipoValor.enuDate).Date, String)
        Dim lstrIdCtaCr As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim ldecBase As Decimal = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        Dim lstrPredioAgr As String = ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        Dim lstrRef As String
        If String.IsNullOrEmpty(lstrPredioAgr) Then
            lstrRef = ClsPanorama.FobjValorCampo(adrwMov(
                ClsNombreCompletoStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        Else
            lstrRef = lstrPredioAgr
        End If
        lstrRef = FstrRef(lstrRef)
        Dim lstrIdTercero As String
        Dim lblnEsCajaBanco = (lstrIdCtaCr = GobjParametros.ObjIdCtaCajaStr.ObjValorPro) OrElse
                GobjParametros.FblnEsCuentaBanco(lstrIdCtaCr)
        If lblnEsCajaBanco Then
            lstrIdTercero = FstrIdTerceroCajaBancos(adrwMov)
        Else
            lstrIdTercero = ClsPanorama.FobjValorCampo(adrwMov(
                    ClsAliasCont_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        End If
        If Not String.IsNullOrEmpty(lstrIdTercero) Then
            If BlnIdTerceroStr Then
                lstrIdTercero = lstrIdTercero.PadLeft(12, "0")
            End If
        End If
        ' Cr
        With MstbLinea
            .Clear().Append(astrTipoDoc).Append(CHSTRCOMA).Append(lstrNroDoc).Append(CHSTRCOMA)
            .Append(lstrFechaMov).Append(CHSTRCOMA).Append(lstrIdCtaCr).Append(CHSTRCOMA)
            .Append(FstrDetalle(adrwMov)).Append(CHSTRCOMA).Append("2").Append(CHSTRCOMA)
            .Append(adecValor).Append(CHSTRCOMA).Append(lstrIdTercero).Append(CHSTRCOMA)
            .Append(ldecBase).Append(CHSTRCOMA).Append(CHSTRCOMA).Append(lstrRef)
        End With
        aswInterfaz.WriteLine(MstbLinea.ToString)
    End Sub
    Private Sub SEscribaMovCrFac(adrwMov As DataRow, astrTipoDoc As String,
                adecValor As Decimal, aswInterfaz As StreamWriter, astrIdTercero As String)
        Dim lstrPrefijoDoc As String = String.Empty
        Dim lentIdDoc As Integer = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
        Dim lstrNroDoc As String = ClsPanorama.FstrNumeroDcto(lstrPrefijoDoc, lentIdDoc)
        Dim lstrFechaMov As String = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsFechaNovedadDtm.SstrNombreCampoBd), EnuTipoValor.enuDate).Date, String)
        Dim lstrIdCtaCr As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                EnuTipoValor.enuString)
        Dim ldecBase As Decimal = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        Dim lstrPredioAgr As String = ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        Dim lstrRef As String
        If String.IsNullOrEmpty(lstrPredioAgr) Then
            lstrRef = ClsPanorama.FobjValorCampo(adrwMov(
                ClsNombreCompletoStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        Else
            lstrRef = lstrPredioAgr
        End If
        lstrRef = FstrRef(lstrRef)
        If Not String.IsNullOrEmpty(astrIdTercero) Then
            If BlnIdTerceroStr Then
                astrIdTercero = astrIdTercero.PadLeft(12, "0")
            End If
        End If
        Dim lblnCrAPasivo = (lstrIdCtaCr.ToString.StartsWith("2"))
        ' Cr
        With MstbLinea
            .Clear().Append(astrTipoDoc).Append(CHSTRCOMA).Append(lstrNroDoc).Append(CHSTRCOMA)
            .Append(lstrFechaMov).Append(CHSTRCOMA).Append(lstrIdCtaCr).Append(CHSTRCOMA)
            .Append(FstrDetalle(adrwMov)).Append(CHSTRCOMA).Append("2").Append(CHSTRCOMA)
            .Append(adecValor).Append(CHSTRCOMA).Append(astrIdTercero).Append(CHSTRCOMA)
            .Append(ldecBase).Append(CHSTRCOMA).Append(CHSTRCOMA).Append(lstrRef)
        End With
        aswInterfaz.WriteLine(MstbLinea.ToString)
    End Sub
#End Region
#Region "Generales"
    Friend Overrides ReadOnly Property DtmFechaFinInterfazAnterior() As Date
        Get
            If MdtmFechaFinInterfazAnterior = GCDTMFECHANULA Then
                If Not MblnLeidoUltimoRegistro Then
                    Select Case GobjParametros.ObjIdAppContableByt.ObjValorPro
                        Case EnuAppConta.EnuContaPyme
                            SLeaUltimoRegInteAnterAP()
                        Case Else
                            Throw New ErrorInesperadoPanLException("Tipo de App Contable no esperado")
                    End Select
                End If
            End If
            Return MdtmFechaFinInterfazAnterior
        End Get
    End Property
    Private Sub SLeaUltimoRegInteAnterAP()
        If Not MblnLeidoUltimoRegistro Then
            Dim lstrUltimoArch = FstrUltimoArchInt("InterfazPodium*.txt")
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
                        MdtmFechaFinInterfazAnterior = CType(lstrPartes(1), Date)
                    End If
                End Using
                MblnLeidoUltimoRegistro = True
            End If
        End If
    End Sub
    Private Function FstrRef(astrRef As String) As String
        Dim lstrRef = astrRef
        Dim lentLargoRef = lstrRef.Length
        If lentLargoRef < 5 Then
            For i = 0 To 5 - lentLargoRef
                lstrRef = "#" & lstrRef
                i += 1
            Next
        End If
        Return lstrRef
    End Function
#End Region
End Class
