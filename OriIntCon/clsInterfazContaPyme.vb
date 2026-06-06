Imports System.IO
Imports System.Text
Friend Class ClsInterfazContaPyme
#Region "Definiciones"
    Inherits ClsCBInterfazContableOri
    'Constantes
    'Variables
    Private ReadOnly MshrCodigoEmp As Short = GobjParametros.ObjCodigoEmpShr.ObjValorPro
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
                MstrNombreArchivo = "InterfazContaPyme_FM_" & ClsPanorama.FstrFechayyyymmdd(
                        DtmFechaDesde) & "_" & ClsPanorama.FstrFechayyyymmdd(DtmFechaHasta)
            Else
                MstrNombreArchivo = "InterfazContaPyme_" & ClsPanorama.FstrFechayyyymmdd(
                        DtmFechaDesde) & "_" & ClsPanorama.FstrFechayyyymmdd(DtmFechaHasta)
            End If
            HstrArchivoSalidaInterfaz = GstrTrayInterfContable & "\" & MstrNombreArchivo
            SGenereIntContaPymeDoc()
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
    Private Sub SGenereIntContaPymeDoc()
        Dim ldtbMovi = FdtbMovimiento()
        Using lswInterfaz = ClsPanorama.FswStreamWriter(HstrArchivoSalidaInterfaz & ".csv")
            SEscribaEncabezado(lswInterfaz)
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
        Dim lblnDbAfectaCxC = False, lblnCrAfectaCxC = False
        Select Case aenuIdDocumento
            Case EnuIdDocumentoDef.EnuFacturaVenta, EnuIdDocumentoDef.EnuNotaIntMora
                lblnDbAfectaCxC = True
                lblnCrAfectaCxC = False
            Case EnuIdDocumentoDef.EnuNotaAjuste, EnuIdDocumentoDef.EnuNotaReintegroAnt
                lblnDbAfectaCxC = False
                lblnCrAfectaCxC = False
            Case EnuIdDocumentoDef.EnuNotaAplicacionAnt, EnuIdDocumentoDef.EnuNotaCr,
                    EnuIdDocumentoDef.EnuReciboCaja
                lblnDbAfectaCxC = False
                lblnCrAfectaCxC = True
            Case EnuIdDocumentoDef.EnuNotaReversaCr
                lblnDbAfectaCxC = True
                lblnCrAfectaCxC = False
        End Select
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
                        EnuTipoValor.EnuDecimal)
                SEscribaMovDocDb(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz, lblnDbAfectaCxC)
                SEscribaMovDocCr(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz, lblnCrAfectaCxC)
            End If
        Next
    End Sub
    Private Sub SProceseFactura(adtbMovi As DataTable, astrPref As String,
            aentIdFact As Integer, aswInterfaz As StreamWriter)
        Dim lstrFiltro = ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " & EnuIdDocumentoDef.EnuFacturaVenta &
                " AND " & ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" & astrPref & "' AND " &
                ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " & aentIdFact
        Dim ldrwNovsFra As DataRow() = adtbMovi.Select(lstrFiltro), ldrwNov As DataRow
        Dim ldecValorFra = 0D, ldecValor As Decimal, lstrIdCtaDb = String.Empty, i = 0, lshrIdItemFac As Short
        Dim lstrIdTerceroCtaCr As String
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.EnuFacturaVenta)
        Dim ldrwNovDbFra As DataRow = Nothing
        Dim lenuTipoNovDbFra As EnuTipoNov
        Dim lobjFactura As New ClsFactura()
        Dim lobjValorLlave As Object()
        Do
            ldrwNov = ldrwNovsFra(i)
            lenuTipoNovDbFra = ClsPanorama.FobjValorCampo(ldrwNov(ClsIdTipoNovedadByt.SstrNombreCampoBd),
                    EnuTipoValor.EnuByte)
            If lstrIdCtaDb <> ClsPanorama.FobjValorCampo(ldrwNov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.EnuString) Then
                If Not String.IsNullOrEmpty(lstrIdCtaDb) Then
                    SEscribaMovDocDb(ldrwNovsFra(i - 1), lstrTipoDoc, ldecValorFra, aswInterfaz, True)
                End If
                lstrIdCtaDb = ClsPanorama.FobjValorCampo(ldrwNov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.EnuString)
                ldecValorFra = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                            EnuTipoValor.EnuDecimal)
            Else
                ldecValorFra += ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                            EnuTipoValor.EnuDecimal)
            End If
            If lenuTipoNovDbFra = EnuTipoNov.EnuDbCap Then
                ldrwNovDbFra = ldrwNov
            End If
            If i = ldrwNovsFra.Length - 1 Then
                SEscribaMovDocDb(ldrwNovDbFra, lstrTipoDoc, ldecValorFra, aswInterfaz, True)
            End If
            i += 1
        Loop While i < ldrwNovsFra.Length
        For Each ldrwNovFac As DataRow In ldrwNovsFra
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNovFac(ClsValor_NovDec.SstrNombreCampoBd),
                        EnuTipoValor.EnuDecimal)
            lshrIdItemFac = ClsPanorama.FobjValorCampo(ldrwNovFac(ClsIdItemFacturaShr.SstrNombreCampoBd),
                    EnuTipoValor.EnuShort)
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, astrPref, aentIdFact}
            lobjFactura.SAbra(lobjValorLlave)
            lstrIdTerceroCtaCr = lobjFactura.FstrIdTerceroCtaCr(lshrIdItemFac)
            If String.IsNullOrEmpty(lstrIdTerceroCtaCr) Then
                SEscribaMovDocCr(ldrwNovFac, lstrTipoDoc, ldecValor, aswInterfaz, False)
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
                    EnuTipoValor.EnuString)
            lentIdRCNov = ClsPanorama.FobjValorCampo(ldrwNov(ClsIdDocOrigenEnt.SstrNombreCampoBd),
                    EnuTipoValor.EnuInteger)
            lstrIdCuentaDbNov = ClsPanorama.FobjValorCampo(ldrwNov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.EnuString)
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
                    EnuTipoValor.EnuDecimal)
            ldecValor += ldecValorNov
        Next
        SEscribaMovDocDb(adrwMoviRC(0), astrTipoDoc, ldecValor, aswInterfaz, False)
        For Each ldrwMovNov As DataRow In adrwMoviRC
            ldecValorNov = ClsPanorama.FobjValorCampo(ldrwMovNov(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.EnuDecimal)
            SEscribaMovDocCr(ldrwMovNov, astrTipoDoc, ldecValorNov, aswInterfaz, True)
        Next
    End Sub
    Private Sub SProceseNotaR(aenuIdDocumento As EnuIdDocumentoDef, aswInterfaz As StreamWriter)
        Dim ldecValor As Decimal
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(aenuIdDocumento)
        Dim ldtbMovAntR = FdtbMoviAnt(aenuIdDocumento)
        For Each ldrwNov As DataRow In ldtbMovAntR.Rows
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.EnuDecimal)
            SEscribaMovDocDb(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz, False)
            SEscribaMovDocCr(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz, False)
        Next
    End Sub
    Private Sub SProceseNotasAjuste(aswInterfaz As StreamWriter)
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.EnuNotaAjuste)
        Dim ldecValor As Decimal
        Dim ldtbMovNA = FdtbMoviNotaAjuste()
        For Each ldrwNov As DataRow In ldtbMovNA.Rows
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                        EnuTipoValor.EnuDecimal)
            SEscribaMovDocDb(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz, False)
            SEscribaMovDocCr(ldrwNov, lstrTipoDoc, ldecValor, aswInterfaz, False)
        Next
    End Sub
    Private Sub SEscribaMovDocDb(adrwMov As DataRow, astrTipoDoc As String,
            adecValor As Decimal, aswInterfaz As StreamWriter, ablnDbAfectaCxC As Boolean)
        Dim lstrPrefFact = String.Empty
        Dim lentIdFactura = 0
        Dim lstrNroFac = String.Empty
        If ablnDbAfectaCxC Then
            lstrPrefFact = ClsPanorama.FobjValorCampo(adrwMov(ClsPrefijoFact_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.EnuString)
            lentIdFactura = ClsPanorama.FobjValorCampo(adrwMov(ClsIdFactura_NovEnt.SstrNombreCampoBd),
                    EnuTipoValor.EnuInteger)
            lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPrefFact, lentIdFactura)
        End If
        Dim lstrPrefijoDoc As String = CType(ClsPanorama.FobjValorCampo(adrwMov(
                        ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.EnuString), String)
        If IsNothing(lstrPrefijoDoc) Then lstrPrefijoDoc = String.Empty
        Dim lentIdDoc As Integer = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger), Integer)
        Dim lstrNroDoc As String = ClsPanorama.FstrNumeroDcto(lstrPrefijoDoc, lentIdDoc)
        Dim lstrFechaMov As String = CType(ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                EnuTipoValor.EnuDate).Date, String)
        Dim lstrIdCtaDb As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.EnuString)
        Dim ldecBase As Decimal = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.EnuDecimal)
        Dim lstrIdTercero = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.EnuString)
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
            .Clear().Append(MshrCodigoEmp).Append(CHSTRCOMA).Append(lstrFechaMov).Append(CHSTRCOMA)
            .Append(astrTipoDoc).Append(CHSTRCOMA).Append(lstrNroDoc).Append(CHSTRCOMA)
            .Append(lstrIdCtaDb).Append(CHSTRCOMA).Append("1").Append(CHSTRCOMA)
            .Append(FstrDetalle(adrwMov)).Append(CHSTRCOMA)
            .Append(adecValor).Append(CHSTRCOMA).Append(CHSTRCOMA).Append(ldecBase).Append(CHSTRCOMA)
            .Append(lstrIdTercero).Append(CHSTRCOMA)
            If ablnDbAfectaCxC OrElse lblnDbAPasivo Then
                .Append(lstrIdTercero)
            End If
            .Append(CHSTRCOMA)
            If lblnDbAPasivo Then
                .Append(lstrFechaMov)
            End If
            .Append(CHSTRCOMA).Append(CHSTRCOMA)
            If ablnDbAfectaCxC Then
                .Append(lstrNroFac)
            ElseIf lblnDbAPasivo Then
                .Append(lstrNroDoc)
            End If
            .Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA)
        End With
        aswInterfaz.WriteLine(MstbLinea.ToString)
    End Sub
    Private Sub SEscribaMovDocCr(adrwMov As DataRow, astrTipoDoc As String,
            adecValor As Decimal, aswInterfaz As StreamWriter, ablnCrAfectaCxC As Boolean)
        Dim lstrPrefijoDoc As String = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString), String)
        If IsNothing(lstrPrefijoDoc) Then lstrPrefijoDoc = String.Empty
        Dim lstrPrefFact = String.Empty
        Dim lentIdFactura = 0
        Dim lstrNroFac = String.Empty
        If ablnCrAfectaCxC Then
            lstrPrefFact = ClsPanorama.FobjValorCampo(adrwMov(ClsPrefijoFact_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.EnuString)
            lentIdFactura = ClsPanorama.FobjValorCampo(adrwMov(ClsIdFactura_NovEnt.SstrNombreCampoBd),
                    EnuTipoValor.EnuInteger)
            lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPrefFact, lentIdFactura)
        End If
        Dim lentIdDoc As Integer = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
        Dim lstrNroDoc As String = ClsPanorama.FstrNumeroDcto(lstrPrefijoDoc, lentIdDoc)
        Dim lstrFechaMov As String = CType(ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
            EnuTipoValor.enuDate).Date, String)
        Dim lstrIdCtaCr As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim ldecBase As Decimal = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        Dim lstrIdTercero = String.Empty
        Dim blnEsAnticipo = False
        Dim lblnEsAnticipo As Boolean =
                GobjParametros.ObjIdCtaAnticiposRecibidosStr.ObjValorPro = lstrIdCtaCr
        If lblnEsAnticipo Then ablnCrAfectaCxC = False
        Dim lblnEsCajaBanco = (lstrIdCtaCr = GobjParametros.ObjIdCtaCajaStr.ObjValorPro) OrElse
          GobjParametros.FblnEsCuentaBanco(lstrIdCtaCr)
        If lblnEsCajaBanco Then
            lstrIdTercero = FstrIdTerceroCajaBancos(adrwMov)
        Else
            lstrIdTercero = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        End If
        If Not String.IsNullOrEmpty(lstrIdTercero) Then
            If BlnIdTerceroStr Then
                lstrIdTercero = lstrIdTercero.PadLeft(12, "0")
            End If
        End If
        ' Cr
        With MstbLinea
            .Clear().Append(MshrCodigoEmp).Append(CHSTRCOMA).Append(lstrFechaMov).Append(CHSTRCOMA)
            .Append(astrTipoDoc).Append(CHSTRCOMA).Append(lstrNroDoc).Append(CHSTRCOMA)
            .Append(lstrIdCtaCr).Append(CHSTRCOMA).Append("1").Append(CHSTRCOMA)
            .Append(FstrDetalle(adrwMov)).Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append(adecValor).Append(CHSTRCOMA).Append(ldecBase).Append(CHSTRCOMA)
            .Append(lstrIdTercero)
            .Append(CHSTRCOMA)
            If ablnCrAfectaCxC Then
                .Append(lstrIdTercero).Append(CHSTRCOMA).Append(lstrFechaMov).Append(CHSTRCOMA)
            ElseIf lblnEsAnticipo Then
                .Append(lstrIdTercero).Append(CHSTRCOMA).Append(CHSTRCOMA)
            Else
                .Append(CHSTRCOMA).Append(CHSTRCOMA)
            End If
            .Append(CHSTRCOMA)
            If ablnCrAfectaCxC Then
                .Append(lstrNroFac)
            ElseIf lblnEsAnticipo Then
                .Append(lstrNroDoc)
            End If
            .Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append(CHSTRCOMA)
        End With
        aswInterfaz.WriteLine(MstbLinea.ToString)
    End Sub
    Private Sub SEscribaMovCrFac(adrwMov As DataRow, astrTipoDoc As String,
                adecValor As Decimal, aswInterfaz As StreamWriter,
                astrIdTercero As String)
        Dim lstrPrefijoDoc As String = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString), String)
        If IsNothing(lstrPrefijoDoc) Then lstrPrefijoDoc = String.Empty
        Dim lentIdDoc As Integer = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
        Dim lstrNroDoc As String = ClsPanorama.FstrNumeroDcto(lstrPrefijoDoc, lentIdDoc)
        Dim lstrFechaMov As String = CType(ClsPanorama.FobjValorCampo(adrwMov(
                ClsFechaNovedadDtm.SstrNombreCampoBd), EnuTipoValor.enuDate).Date, String)
        Dim lstrIdCtaCr As String = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        Dim ldecBase As Decimal = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        If Not String.IsNullOrEmpty(astrIdTercero) Then
            If BlnIdTerceroStr Then
                astrIdTercero = astrIdTercero.PadLeft(12, "0")
            End If
        End If
        Dim lblnCrAPasivo = (lstrIdCtaCr.ToString.StartsWith("2"))
        ' Cr
        With MstbLinea
            .Clear().Append(MshrCodigoEmp).Append(CHSTRCOMA).Append(lstrFechaMov).Append(CHSTRCOMA)
            .Append(astrTipoDoc).Append(CHSTRCOMA).Append(lstrNroDoc).Append(CHSTRCOMA)
            .Append(lstrIdCtaCr).Append(CHSTRCOMA).Append("1").Append(CHSTRCOMA)
            .Append(FstrDetalle(adrwMov)).Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append(adecValor).Append(CHSTRCOMA).Append(ldecBase).Append(CHSTRCOMA)
            .Append(astrIdTercero)
            .Append(CHSTRCOMA)
            If lblnCrAPasivo Then
                .Append(astrIdTercero)
            End If
            .Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA)
            If lblnCrAPasivo Then
                .Append(lstrNroDoc)
            End If
            .Append(CHSTRCOMA).Append(CHSTRCOMA).Append(CHSTRCOMA)
            .Append(CHSTRCOMA).Append(CHSTRCOMA)
        End With
        aswInterfaz.WriteLine(MstbLinea.ToString)
    End Sub
#End Region
#Region "Comunes AP"
    Private Sub SEscribaEncabezado(aswInterfaz As StreamWriter)
        With MstbLinea
            .Clear().Append("IEMP").Append(CHSTRCOMA).Append("FSOPORT").Append(CHSTRCOMA)
            .Append("ITDSOP").Append(CHSTRCOMA).Append("INUMSOP").Append(CHSTRCOMA)
            .Append("ICUENTA").Append(CHSTRCOMA).Append("ICCSUBCC").Append(CHSTRCOMA)
            .Append("TDETALLE").Append(CHSTRCOMA).Append("MDEBITO").Append(CHSTRCOMA)
            .Append("MCREDITO").Append(CHSTRCOMA).Append("MVRBASE").Append(CHSTRCOMA)
            .Append("INIT").Append(CHSTRCOMA).Append("INITCXX").Append(CHSTRCOMA)
            .Append("FPAGOCXX").Append(CHSTRCOMA).Append("IFLUJOEFEC").Append(CHSTRCOMA)
            .Append("REFERENCIA").Append(CHSTRCOMA).Append("IBANCO").Append(CHSTRCOMA)
            .Append("ICHEQUE").Append(CHSTRCOMA).Append("IACTIVO").Append(CHSTRCOMA)
            .Append("MVROTRAMON").Append(CHSTRCOMA).Append("SCOMANDOS")
        End With
        aswInterfaz.WriteLine(MstbLinea.ToString)
    End Sub
#End Region
#Region "Generales"
    Friend Overrides ReadOnly Property DtmFechaFinInterfazAnterior() As Date
        Get
            If mdtmFechaFinInterfazAnterior = GCDTMFECHANULA Then
                If Not mblnLeidoUltimoRegistro Then
                    Select Case GobjParametros.objIdAppContableByt.objValorPro
                        Case EnuAppConta.EnuContaPyme
                            SLeaUltimoRegInteAnterAP()
                        Case Else
                            Throw New ErrorInesperadoPanLException("Tipo de App Contable no esperado")
                    End Select
                End If
            End If
            Return mdtmFechaFinInterfazAnterior
        End Get
    End Property
    Private Sub SLeaUltimoRegInteAnterAP()
        If Not MblnLeidoUltimoRegistro Then
            Dim lstrUltimoArch = FstrUltimoArchInt("InterfazContaPyme*.csv")
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
#End Region
End Class