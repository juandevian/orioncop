Imports Microsoft.Office.Interop
Public Class ClsDocsWord
    Private MdblIdCliente As Double = 0, mshrIdAno As Integer = 0
    Private MdtbPagos As System.Data.DataTable = Nothing
    Private MobjAno As ClsAno = Nothing
    Private MobjTercero As ClsTercero
    Private wrdApp As Word.Application = Nothing
    Private wrdDoc As Word._Document = Nothing
    Private MentIdLineaTablaConsPago As Integer = 0
    Private ReadOnly MobjCenUtil As ClsCentroUtilidad = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual

    Friend Sub SImprimaDoc(ByVal aenuDoc As EnuDocumentosWordDef, ByVal adblIdCliente As Double,
                           ByVal ashrIdAno As Integer)
        MdblIdCliente = adblIdCliente
        MobjTercero = New ClsTercero(EnuModoInstanciaObjDef.enuUnico)
        MobjTercero.SAbra({MdblIdCliente})
        mshrIdAno = ashrIdAno
        If GobjCentroUtilOriCop.ColAnos.Contains(mshrIdAno.ToString) Then
            MobjAno = GobjCentroUtilOriCop.ColAnos(mshrIdAno.ToString)
            Select Case aenuDoc
                Case EnuDocumentosWordDef.enuConstanciaPago
                    Try
                        SImprimaConstanciaPago()
                    Catch ex As Exception
                        Throw
                    End Try
                Case EnuDocumentosWordDef.enuCartaCobro
            End Select
        End If
    End Sub

    Friend Sub SImprimaPazySalvo(ByVal astrIdPredio As String)
        Dim lstrArchPlanPazYSalvo = GstrTrayAppDat & "PlantillaPazYSalvo.dot"
        If My.Computer.FileSystem.FileExists(lstrArchPlanPazYSalvo) Then
            Dim lstrArchPazYSalvo = GstrTrayReportes
            Dim lstrFecArch = ClsPanorama.FstrFechayyyymmdd(Date.Today)
            lstrArchPazYSalvo &= "\" & lstrFecArch & "_PazYSalvo_" & astrIdPredio & ".doc"
            Dim lstrFecha = Format(Date.Today, "dd \de MMMM \de yyyy")
            Dim lstrIdProp = String.Empty, lstrTipoDocId = String.Empty
            Dim lstrNombrePropietario = FstrNombreProp(astrIdPredio, lstrIdProp, lstrTipoDocId)
            wrdApp = CreateObject("Word.Application")
            wrdDoc = wrdApp.Documents.Add(lstrArchPlanPazYSalvo)
            With wrdApp
                .ActiveDocument.Bookmarks("fecha").Select()
                .Selection.Text = lstrFecha
                .ActiveDocument.Bookmarks("nomrepleg1").Select()
                .Selection.Text = FstrNombreRepLegal()
                .ActiveDocument.Bookmarks("idreplegal").Select()
                .Selection.Text = FstrIdRepLegal()
                .ActiveDocument.Bookmarks("nomcoprop").Select()
                .Selection.Text = FstrNombreCopropiedad()
                .ActiveDocument.Bookmarks("nitcopro").Select()
                .Selection.Text = FstrIdCopropiedad()
                .ActiveDocument.Bookmarks("nomprop").Select()
                .Selection.Text = lstrNombrePropietario
                .ActiveDocument.Bookmarks("tipoidpro").Select()
                .Selection.Text = lstrTipoDocId
                .ActiveDocument.Bookmarks("idprop").Select()
                .Selection.Text = lstrIdProp
                .ActiveDocument.Bookmarks("idpredio").Select()
                .Selection.Text = astrIdPredio
                .ActiveDocument.Bookmarks("nomrepleg").Select()
                .Selection.Text = FstrNombreRepLegal()
                .ActiveDocument.Bookmarks("cargoadmin").Select()
                .Selection.Text = FstrCargoRepLegal()
            End With
            wrdApp.Visible = True
            wrdDoc.SaveAs2(lstrArchPazYSalvo)
        Else
            GobjMensaje.SRegistreMensaje("No existe la plantilla. Por Favor comuniquese con Soporte!",
                    EnuSeveridadMen.enuInformacion)
        End If
    End Sub

    Private Sub SImprimaConstanciaPago()
        Dim lstrArchPlanConsPago = GstrTrayAppDat & "PlantillaConstanciaPagos.dot"
        If My.Computer.FileSystem.FileExists(lstrArchPlanConsPago) Then
            MentIdLineaTablaConsPago = 0
            Dim lstrArchConsPago = GstrTrayReportes
            lstrArchConsPago &= "\ConstanciaPagos_" & MobjAno.ObjIdAnoShr.ObjValorPro & "_" &
                    MdblIdCliente.ToString & ".doc"
            Dim lstrFecha = Format(Date.Today, "dd \de MMMM \de yyyy")
            wrdApp = CreateObject("Word.Application")
            wrdDoc = wrdApp.Documents.Add(lstrArchPlanConsPago)
            With wrdApp
                .ActiveDocument.Bookmarks("fecha").Select()
                .Selection.Text = lstrFecha

                .ActiveDocument.Bookmarks("origen").Select()
                .Selection.Text = FstrNombreCopropiedad()

                .ActiveDocument.Bookmarks("idorigen").Select()
                .Selection.Text = FstrIdCopropiedad()

                .ActiveDocument.Bookmarks("cliente").Select()
                .Selection.Text = FstrNombreCliente()

                .ActiveDocument.Bookmarks("doc").Select()
                .Selection.Text = FstrTipoDoc()

                .ActiveDocument.Bookmarks("idcliente").Select()
                .Selection.Text = Format(MdblIdCliente, GCSTRFMTIDTERCERO)

                .ActiveDocument.Bookmarks("año").Select()
                .Selection.Text = mshrIdAno.ToString
            End With
            Dim lwrdRange = wrdDoc.Range
            lwrdRange.InsertParagraphAfter()
            lwrdRange.InsertParagraphAfter()
            SInserteTabla()
            wrdApp.Visible = True
            wrdDoc.SaveAs2(lstrArchConsPago)
        Else
            GobjMensaje.SRegistreMensaje("No existe la plantilla. Por Favor comuniquese con Soporte!",
                EnuSeveridadMen.enuInformacion)
        End If
    End Sub

    Private Sub SPuebleDtbValoresPagados()
        If IsNothing(MdtbPagos) Then
            Dim lstrFechaIni = "'" & ClsPanoramaDat.FstrFechaNormalizada(MobjAno.DtmFechaInicioAno) & "'"
            Dim lstrFechaFin = "'" & ClsPanoramaDat.FstrFechaNormalizada(MobjAno.DtmFechaFinAno) & "'"
            Dim lstrTablaPri = ClsNovedad.SstrNombreTabla
            Dim lstrTablaSec = ClsItemFactura.SstrNombreTabla
            Dim lstrCamposSelectPri = {ClsIdTercero_NovDbl.SstrNombreCampoBd,
                                       ClsIdTipoNovedadByt.SstrNombreCampoBd,
                                       "SUM(" & ClsValor_NovDec.SstrNombreCampoBd & ")"}
            Dim lstrCamposSelectSec = {ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd,
                                       ClsIdServicio_ItemFactShr.SstrNombreCampoBd,
                                       ClsIdPredio_ItemFactStr.SstrNombreCampoBd}
            Dim lstrCamposRelPri = {OPT.OrionP.PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                                    OPT.OrionP.PanL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                                    ClsPrefijoFact_NovStr.SstrNombreCampoBd,
                                    ClsIdFactura_NovEnt.SstrNombreCampoBd,
                                    ClsIdItemFacturaShr.SstrNombreCampoBd}
            Dim lstrCamposRelSec = {OPT.OrionP.PanL.ClsIdCarpetaShr.SstrNombreCampoBd,
                                    OPT.OrionP.PanL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                                    ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
                                    ClsIdFactura_ItemFactEnt.SstrNombreCampoBd,
                                    ClsIdItemFacturaShr.SstrNombreCampoBd}
            Dim lstrCamposGrup = {ClsIdTercero_NovDbl.SstrNombreCampoBd,
                                  ClsIdPredio_ItemFactStr.SstrNombreCampoBd,
                                  ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd,
                                  ClsIdServicio_ItemFactShr.SstrNombreCampoBd,
                                  ClsIdTipoNovedadByt.SstrNombreCampoBd}
            Dim lstrIndice = {{ClsIdPredio_ItemFactStr.SstrNombreCampoBd, "ASC"},
                             {ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd, "DESC"},
                             {ClsIdServicio_ItemFactShr.SstrNombreCampoBd, "ASC"},
                             {ClsIdTipoNovedadByt.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = "P." & OPT.OrionP.PanL.ClsIdCarpetaShr.SstrNombreCampoBd & " = " & GshrIdCarpeta &
                    " AND " & "P." & OPT.OrionP.PanL.ClsIdCentroUtilShr.SstrNombreCampoBd & " = " & GshrIdCentroUtil &
                    " AND " & ClsIdTercero_NovDbl.SstrNombreCampoBd & " = " & MdblIdCliente & " AND (" &
                    ClsIdTipoNovedadByt.SstrNombreCampoBd & " = " & EnuTipoNovedadDef.enuDbCap & " OR " &
                    ClsIdTipoNovedadByt.SstrNombreCampoBd & " = " & EnuTipoNovedadDef.enuCrDctoCap & ") AND " &
                    ClsFechaNovedadDtm.SstrNombreCampoBd & " >= " & lstrFechaIni & " AND " &
                    ClsFechaNovedadDtm.SstrNombreCampoBd & " <= " & lstrFechaFin
            MdtbPagos = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamposSelectPri, lstrTablaSec,
                                                      lstrCamposSelectSec, lstrCamposRelPri, lstrCamposRelSec,
                                                      lstrIndice, lstrFiltro, lstrCamposGrup)
        End If
    End Sub

    Private Function FstrNombreCliente() As String
        Dim lobjCliente As ClsCliente = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
        lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, MdblIdCliente})
        Return lobjCliente.ObjNombreCompletoStr.ObjValorPro
    End Function

    Private Function FdecValorTotal() As Decimal
        Dim lenuIdTipoNovedad As EnuTipoNovedadDef
        Dim ldecVlrTotal = 0D, ldecValor As Decimal
        SPuebleDtbValoresPagados()
        If MdtbPagos.Rows.Count > 0 Then
            For Each ldrwPago In MdtbPagos.Rows
                lenuIdTipoNovedad = ClsPanorama.FobjValorCampo(ldrwPago(ClsIdTipoNovedadByt.SstrNombreCampoBd),
                        EnuTipoValorDef.enuInteger)
                ldecValor = ClsPanorama.FobjValorCampo(ldrwPago("Valor"), EnuTipoValorDef.enuDecimal)
                If lenuIdTipoNovedad = EnuTipoNovedadDef.enuDbCap Then
                    ldecVlrTotal += ldecValor
                Else
                    ldecVlrTotal -= ldecValor
                End If
            Next
        End If
        Return ldecVlrTotal
    End Function

    Private Sub SInserteTabla()
        SPuebleDtbValoresPagados()
        Dim lwrdRango = wrdDoc.Range(wrdDoc.Paragraphs(wrdDoc.Paragraphs.Count).Range.Start)
        wrdDoc.Tables.Add(lwrdRango, MdtbPagos.Rows.Count + 2, 3)
        Dim lwrdTable = wrdDoc.Tables(1)
        Dim lwrdCelda As Word.Cell
        Dim lstrIdPredio = "***", lstrFiltro As String
        Dim ldrwItemsPredio() As DataRow
        lwrdTable.Columns(1).Width = 275
        lwrdTable.Columns(2).Width = 65
        lwrdTable.Columns(3).Width = 90
        For Each ldrwPago As DataRow In MdtbPagos.Rows
            If lstrIdPredio <> ClsPanorama.FobjValorCampo(ldrwPago(ClsIdPredio_ItemFactStr.SstrNombreCampoBd),
                EnuTipoValorDef.enuString) Then
                lstrIdPredio = ClsPanorama.FobjValorCampo(ldrwPago(ClsIdPredio_ItemFactStr.SstrNombreCampoBd),
                        EnuTipoValorDef.enuString)
                lstrFiltro = ClsIdPredio_ItemFactStr.SstrNombreCampoBd & " = '" & lstrIdPredio & "'"
                ldrwItemsPredio = MdtbPagos.Select(lstrFiltro)
                SProcesePredio(lwrdTable, ldrwItemsPredio)
            End If
        Next
        MentIdLineaTablaConsPago += 2
        lwrdCelda = lwrdTable.Cell(MentIdLineaTablaConsPago, 1)
        lwrdCelda.Range.Text = "TOTAL"
        lwrdCelda = lwrdTable.Cell(MentIdLineaTablaConsPago, 3)
        lwrdCelda.Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        lwrdCelda.Range.Text = Format(FdecValorTotal, "c")
    End Sub

    Private Sub SProcesePredio(ByVal awrdTabla As Word.Table, ByVal adrwItemsPredio() As DataRow)
        Dim lshrIdAno = 0S, lshrIdServicio = 0S, ldecValor = 0, ldecValorItem As Decimal
        Dim lstrNombreSer As String
        Dim lenuIdTipoNovedad As EnuTipoNovedadDef
        Dim lstrPredio As String = ClsPanorama.FobjValorCampo(adrwItemsPredio(0) _
                (ClsIdPredio_ItemFactStr.SstrNombreCampoBd), EnuTipoValorDef.enuString)
        For Each ldrwItem As DataRow In adrwItemsPredio
            If lshrIdAno <> ClsPanorama.FobjValorCampo(ldrwItem(ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd),
                    EnuTipoValorDef.enuShort) OrElse lshrIdServicio <>
                    ClsPanorama.FobjValorCampo(ldrwItem(ClsIdServicio_ItemFactShr.SstrNombreCampoBd),
                    EnuTipoValorDef.enuShort) Then
                If ldecValor <> 0 Then
                    lstrNombreSer = GobjCentroUtilOriCop.FstrNombreServicio(lshrIdAno, lshrIdServicio)
                    SInserteCelda(awrdTabla, lstrNombreSer, lstrPredio, ldecValor)
                    ldecValor = 0
                End If
                lshrIdAno = ClsPanorama.FobjValorCampo(ldrwItem(ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd),
                                       EnuTipoValorDef.enuShort)
                lshrIdServicio = ClsPanorama.FobjValorCampo(ldrwItem(ClsIdServicio_ItemFactShr.SstrNombreCampoBd),
                        EnuTipoValorDef.enuShort)
            End If
            lenuIdTipoNovedad = ClsPanorama.FobjValorCampo(ldrwItem(ClsIdTipoNovedadByt.SstrNombreCampoBd),
                    EnuTipoValorDef.enuInteger)
            ldecValorItem = ClsPanorama.FobjValorCampo(ldrwItem("Valor"), EnuTipoValorDef.enuDecimal)
            If lenuIdTipoNovedad = EnuTipoNovedadDef.enuDbCap Then
                ldecValor += ldecValorItem
            Else
                ldecValor -= ldecValorItem
            End If
        Next
        If ldecValor <> 0 Then
            lstrNombreSer = GobjCentroUtilOriCop.FstrNombreServicio(lshrIdAno, lshrIdServicio)
            SInserteCelda(awrdTabla, lstrNombreSer, lstrPredio, ldecValor)
        End If
    End Sub

    Private Sub SInserteCelda(ByVal awrdTabla As Word.Table, ByVal astrNombreServicio As String,
            ByVal astrIdPredio As String, ByVal adecValor As Decimal)
        Dim lwrdCelda As Word.Cell
        MentIdLineaTablaConsPago += 1
        lwrdCelda = awrdTabla.Cell(MentIdLineaTablaConsPago, 1)
        lwrdCelda.Range.Text = astrNombreServicio
        lwrdCelda = awrdTabla.Cell(MentIdLineaTablaConsPago, 2)
        lwrdCelda.Range.Text = astrIdPredio
        lwrdCelda = awrdTabla.Cell(MentIdLineaTablaConsPago, 3)
        lwrdCelda.Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        lwrdCelda.Range.Text = Format(adecValor, "c")
    End Sub

    Private Function FstrTipoDoc() As String
        Return MobjTercero.ObjTipoDocIdentidadByt.ToString()
    End Function

    Private Function FstrNombreCopropiedad() As String
        Return MobjCenUtil.ObjIdTerceroCentroUtilDbl.StrNombreTerceroCentroutil
    End Function

    Private Function FstrIdCopropiedad() As String
        Dim lstrIdOri = Format(MobjCenUtil.ObjTerceroCentroUtilidad.ObjIdTerceroDbl.ObjValorPro,
                GCSTRFMTIDTERCERO) & "-" &
                MobjCenUtil.ObjTerceroCentroUtilidad.ObjIdTerceroDbl.SbyDigitoVerificacion
        Return lstrIdOri
    End Function

    Private Function FstrNombreRepLegal() As String
        Return MobjCenUtil.ObjIdTerceroRepLegalDbl.StrNombreRepLegal
    End Function

    Private Function FstrIdRepLegal() As String
        Return MobjCenUtil.ObjIdTerceroRepLegalDbl.ToString()
    End Function

    Private Function FstrCargoRepLegal()
        Return MobjCenUtil.ObjCargoStr.ObjValorPro
    End Function

    Private Shared Function FstrNombreProp(ByVal astrIdPredio As String, ByRef astrIdProp As String,
                                    ByRef astrTipoDocId As String)
        Dim lobjPredio As ClsPredio = New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
        Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, astrIdPredio}
        lobjPredio.SAbra(lobjValorLlave)
        If Not lobjPredio.BlnExiste Then
            Throw New ValorArgumentoInvalidoException("astrIdPredio")
        End If
        astrIdProp = lobjPredio.ObjIdClientePropietarioDbl.ToString
        Dim lobjClientePro As ClsCliente = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
        lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lobjPredio.ObjIdClientePropietarioDbl.ObjValorPro}
        lobjClientePro.SAbra(lobjValorLlave)
        If Not lobjClientePro.BlnExiste Then
            Throw New ErrorInesperadoPanLException("Id Cliente no existe!")
        End If
        Dim lbytIdTipoDocId As Byte = lobjClientePro.ObjTerceroCliente.ObjTipoDocIdentidadByt.ObjValorPro
        If lbytIdTipoDocId = EnuTipoDocIdDef.enuNit Then
            astrTipoDocId = "NIT"
        Else
            astrTipoDocId = "CC"
        End If
        Return lobjPredio.ObjNombrePropietarioStr.ObjValorPro
    End Function
End Class
