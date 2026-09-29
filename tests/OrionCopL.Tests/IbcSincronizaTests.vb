Imports OPT.OrionP.OrionCopL
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class IbcSincronizaTests
    Private Class ProveedorFalso
        Implements IIbcProveedor
        Friend Lista As New List(Of StcIbcCertificado)
        Friend Falla As Boolean = False
        Friend Llamadas As Integer = 0
        Public Function FlstConsulteCertificados(aentCantidad As Integer) As List(Of StcIbcCertificado) _
                Implements IIbcProveedor.FlstConsulteCertificados
            Llamadas += 1
            If Falla Then Throw New ErrorIbcApiException("API caída", 503)
            Return Lista
        End Function
    End Class

    Private Class AlmacenFalso
        Implements IIbcAlmacen
        Friend Lista As New List(Of StcIbcCertificado)
        Public Function FlstCertificados() As List(Of StcIbcCertificado) Implements IIbcAlmacen.FlstCertificados
            Return Lista
        End Function
        Public Sub SGuarde(alstCertificados As IEnumerable(Of StcIbcCertificado)) Implements IIbcAlmacen.SGuarde
            For Each lstc In alstCertificados
                If Not Lista.Any(Function(x) x.StrIdfile = lstc.StrIdfile) Then Lista.Add(lstc)
            Next
        End Sub
    End Class

    Private Class TasasFalsas
        Implements IIbcTasasMora
        Friend Vigente As Double = 0
        Friend UltimaDesde As Date = #2026-01-01#
        Friend Registros As New List(Of Tuple(Of Date, Double))
        ' Si no es Nothing, SRegistreTasa la lanza (sin registrar nada)
        Friend ErrorAlRegistrar As Exception = Nothing
        Public Function FdblTasaVigente(adtmFecha As Date) As Double Implements IIbcTasasMora.FdblTasaVigente
            Return Vigente
        End Function
        Public Function FdtmFechaDesdeUltima() As Date Implements IIbcTasasMora.FdtmFechaDesdeUltima
            Return UltimaDesde
        End Function
        Public Sub SRegistreTasa(adtmFechaDesde As Date, adblTasaAnual As Double) Implements IIbcTasasMora.SRegistreTasa
            If Not IsNothing(ErrorAlRegistrar) Then Throw ErrorAlRegistrar
            Registros.Add(Tuple.Create(adtmFechaDesde, adblTasaAnual))
            Vigente = adblTasaAnual
            UltimaDesde = adtmFechaDesde
        End Sub
    End Class

    Private Shared Function FstcCert(astrId As String, adtmDesde As Date, adtmHasta As Date,
            adblIbc As Double) As StcIbcCertificado
        Dim lstc As New StcIbcCertificado
        lstc.StrIdfile = astrId : lstc.DtmFechaDesde = adtmDesde : lstc.DtmFechaHasta = adtmHasta
        lstc.DblIbc = adblIbc : lstc.DtmFechaCertificado = adtmDesde.AddDays(-3)
        Return lstc
    End Function

    Private Shared Function FobjSinc(aobjP As ProveedorFalso, aobjA As AlmacenFalso, aobjT As TasasFalsas) As ClsIbcSincroniza
        Return New ClsIbcSincroniza(aobjP, aobjA, aobjT)
    End Function

    ' ErrorInesperadoPanLException es Friend en PanL (sin InternalsVisibleTo hacia las pruebas): se crea por reflexión
    Private Shared Function FexcInesperadoPanL(astrMensaje As String) As Exception
        Dim ltypError = Type.GetType("OPT.OrionP.PanL.ErrorInesperadoPanLException, PanL", True)
        Return CType(Activator.CreateInstance(ltypError, astrMensaje), Exception)
    End Function

    ' Fecha de "hoy" para las reglas de FechaDesde: se inyecta como parámetro de la función
    Private ReadOnly MdtmHoy As Date = #2026-09-28#

    <TestMethod>
    Public Sub Modo_none_no_consulta_ni_escribe_y_devuelve_true()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        Dim lstrMens As String = String.Empty
        Assert.IsTrue(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.None, 0, 0,
                False, MdtmHoy, lstrMens))
        Assert.AreEqual(0, lobjP.Llamadas)
        Assert.AreEqual(0, lobjT.Registros.Count)
        Assert.AreEqual(String.Empty, lstrMens)
    End Sub

    <TestMethod>
    Public Sub Variable_con_dato_local_vigente_no_llama_a_la_api_y_escribe_ibc_por_factor()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        Dim lstrMens As String = String.Empty
        Assert.IsTrue(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 1.3,
                False, MdtmHoy, lstrMens))
        Assert.AreEqual(0, lobjP.Llamadas)
        Assert.AreEqual(1, lobjT.Registros.Count)
        Assert.AreEqual(0.25194, lobjT.Registros(0).Item2, 0.0000001)
        Assert.AreEqual(#2026-09-01#, lobjT.Registros(0).Item1)
    End Sub

    <TestMethod>
    Public Sub Fijo_baja_al_tope_y_luego_vuelve_a_la_deseada()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.14))     ' tope 21 %
        lobjA.Lista.Add(FstcCert("B", #2026-10-01#, #2026-10-31#, 0.1938))   ' tope 29.07 %
        Dim lobjS = FobjSinc(lobjP, lobjA, lobjT)
        Dim lstrMens As String = String.Empty
        Assert.IsTrue(lobjS.FblnSincronice(#2026-09-28#, EnuModoInteres.EnuFijo, 0.24, 0, False, #2026-09-28#, lstrMens))
        Assert.AreEqual(0.21, lobjT.Vigente, 0.0000001)
        Assert.IsTrue(lobjS.FblnSincronice(#2026-10-15#, EnuModoInteres.EnuFijo, 0.24, 0, False, #2026-10-15#, lstrMens))
        Assert.AreEqual(0.24, lobjT.Vigente, 0.0000001)
    End Sub

    <TestMethod>
    Public Sub No_escribe_si_la_tasa_vigente_ya_es_la_que_corresponde()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        lobjT.Vigente = 0.24
        Dim lstrMens As String = String.Empty
        Assert.IsTrue(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuFijo, 0.24, 0,
                False, MdtmHoy, lstrMens))
        Assert.AreEqual(0, lobjT.Registros.Count)
    End Sub

    <TestMethod>
    Public Sub Api_caida_y_sin_dato_local_devuelve_false_con_mensaje_y_no_escribe()
        Dim lobjP As New ProveedorFalso With {.Falla = True}
        Dim lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        Dim lstrMens As String = String.Empty
        Assert.IsFalse(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 1.3,
                False, MdtmHoy, lstrMens))
        StringAssert.Contains(lstrMens, "API caída")
        Assert.AreEqual(0, lobjT.Registros.Count)
    End Sub

    <TestMethod>
    Public Sub Api_caida_pero_con_dato_local_vigente_continua()
        Dim lobjP As New ProveedorFalso With {.Falla = True}
        Dim lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        Dim lstrMens As String = String.Empty
        Assert.IsTrue(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 1.3,
                True, MdtmHoy, lstrMens))    ' forzar consulta falla, pero hay dato local vigente
        Assert.AreEqual(1, lobjP.Llamadas)
        Assert.AreEqual(1, lobjT.Registros.Count)
    End Sub

    <TestMethod>
    Public Sub Sin_certificado_que_cubra_la_fecha_tras_consultar_devuelve_false()
        Dim lobjP As New ProveedorFalso
        lobjP.Lista.Add(FstcCert("FUT", #2026-10-01#, #2026-10-31#, 0.2))
        Dim lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        Dim lstrMens As String = String.Empty
        Assert.IsFalse(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 1.3,
                False, MdtmHoy, lstrMens))
        Assert.AreEqual(1, lobjA.Lista.Count)   ' el certificado consultado se guardó igual
        Assert.AreEqual(0, lobjT.Registros.Count)
        Assert.IsFalse(String.IsNullOrEmpty(lstrMens))
    End Sub

    <TestMethod>
    Public Sub No_escribe_si_la_fecha_desde_calculada_supera_hoy()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        lobjT.UltimaDesde = MdtmHoy        ' última fila es de hoy => FechaDesde nueva = mañana > hoy
        Dim lstrMens As String = String.Empty
        Assert.IsFalse(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 1.3,
                False, MdtmHoy, lstrMens))
        Assert.AreEqual(0, lobjT.Registros.Count)
        Assert.IsFalse(String.IsNullOrEmpty(lstrMens))
    End Sub

    <TestMethod>
    Public Sub Factor_invalido_devuelve_false()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        Dim lstrMens As String = String.Empty
        Assert.IsFalse(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 1.8,
                False, MdtmHoy, lstrMens))
        Assert.AreEqual(0, lobjT.Registros.Count)
    End Sub

    <TestMethod>
    Public Sub Factor_cero_en_modo_variable_devuelve_false()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        Dim lstrMens As String = String.Empty
        Assert.IsFalse(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 0,
                False, MdtmHoy, lstrMens))
        Assert.AreEqual(0, lobjT.Registros.Count)
        Assert.IsFalse(String.IsNullOrEmpty(lstrMens))
    End Sub

    <TestMethod>
    Public Sub Modo_desconocido_devuelve_false_y_no_escribe()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        Dim lstrMens As String = String.Empty
        Assert.IsFalse(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, CType(7, EnuModoInteres), 0.24, 1.3,
                False, MdtmHoy, lstrMens))
        Assert.AreEqual(0, lobjP.Llamadas)
        Assert.AreEqual(0, lobjT.Registros.Count)
        Assert.IsFalse(String.IsNullOrEmpty(lstrMens))
    End Sub

    ' Causación del día F lee la tasa de F - 1: el 1 de octubre se causa septiembre con el certificado de septiembre
    <TestMethod>
    Public Sub Frontera_de_mes_usa_el_certificado_vigente_el_dia_anterior()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("SEP", #2026-09-01#, #2026-09-30#, 0.14))
        lobjA.Lista.Add(FstcCert("OCT", #2026-10-01#, #2026-10-31#, 0.1938))
        Dim lobjS = FobjSinc(lobjP, lobjA, lobjT)
        Dim lstrMens As String = String.Empty
        Assert.IsTrue(lobjS.FblnSincronice(#2026-10-01#, EnuModoInteres.EnuVariable, 0, 1.0, False, #2026-10-01#,
                lstrMens), lstrMens)
        Assert.AreEqual(1, lobjT.Registros.Count)
        Assert.AreEqual(#2026-09-01#, lobjT.Registros(0).Item1)
        Assert.AreEqual(0.14, lobjT.Registros(0).Item2, 0.0000001)
        Assert.IsTrue(lobjS.FblnSincronice(#2026-11-01#, EnuModoInteres.EnuVariable, 0, 1.0, False, #2026-11-01#,
                lstrMens), lstrMens)
        Assert.AreEqual(2, lobjT.Registros.Count)
        Assert.AreEqual(#2026-10-01#, lobjT.Registros(1).Item1)
        Assert.AreEqual(0.1938, lobjT.Registros(1).Item2, 0.0000001)
        Assert.AreEqual(0, lobjP.Llamadas)
    End Sub

    <TestMethod>
    Public Sub Frontera_de_mes_sin_certificado_del_mes_anterior_consulta_y_no_usa_el_del_mes_nuevo()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("OCT", #2026-10-01#, #2026-10-31#, 0.1938))
        Dim lstrMens As String = String.Empty
        Assert.IsFalse(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(#2026-10-01#, EnuModoInteres.EnuVariable, 0, 1.0,
                False, #2026-10-01#, lstrMens))
        Assert.AreEqual(1, lobjP.Llamadas)
        Assert.AreEqual(0, lobjT.Registros.Count)
    End Sub

    ' La fila nueva debe regir en F - 1 (la fecha que lee la causación); si su FechaDesde queda después, no sirve
    <TestMethod>
    Public Sub No_escribe_si_la_fecha_desde_no_rige_el_dia_que_lee_la_causacion()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        lobjT.UltimaDesde = #2026-09-27#   ' FechaDesde nueva = 28/09 > F - 1 = 27/09 (aunque no supera hoy)
        Dim lstrMens As String = String.Empty
        Assert.IsFalse(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(#2026-09-28#, EnuModoInteres.EnuVariable, 0, 1.3,
                False, #2026-09-30#, lstrMens))
        Assert.AreEqual(0, lobjT.Registros.Count)
        StringAssert.Contains(lstrMens, "2026-09-27")
        StringAssert.Contains(lstrMens, "no regiría")
    End Sub

    <TestMethod>
    Public Sub No_escribe_si_la_fecha_desde_supera_hoy_aunque_rija_para_la_causacion()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        lobjT.UltimaDesde = MdtmHoy        ' FechaDesde nueva = 29/09 <= F - 1 = 30/09 pero > hoy
        Dim lstrMens As String = String.Empty
        Assert.IsFalse(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(#2026-10-01#, EnuModoInteres.EnuVariable, 0, 1.3,
                False, MdtmHoy, lstrMens))
        Assert.AreEqual(0, lobjT.Registros.Count)
        StringAssert.Contains(lstrMens, "posterior a hoy")
    End Sub

    <TestMethod>
    Public Sub Tasa_guardada_pero_inconsistente_devuelve_false_con_el_mensaje_de_la_excepcion()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        Dim lobjEx As New ErrorTasaMoraGuardadaException(5, #2026-09-01#, "la tasa anterior no quedó cerrada", Nothing)
        lobjT.ErrorAlRegistrar = lobjEx
        Dim lstrMens As String = String.Empty
        Assert.IsFalse(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 1.3,
                False, MdtmHoy, lstrMens))
        Assert.AreEqual(lobjEx.Message, lstrMens)
    End Sub

    <TestMethod>
    Public Sub Validacion_previa_al_guardar_devuelve_false_con_mensaje_y_nada_escrito()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        lobjT.ErrorAlRegistrar = FexcInesperadoPanL("La fecha desde de la Tasa de Mora no es válida")
        Dim lstrMens As String = String.Empty
        Assert.IsFalse(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 1.3,
                False, MdtmHoy, lstrMens))
        StringAssert.Contains(lstrMens, "La fecha desde de la Tasa de Mora no es válida")
        StringAssert.Contains(lstrMens, "No se registró")
        Assert.AreEqual(0, lobjT.Registros.Count)
    End Sub

    <TestMethod>
    Public Sub Otras_excepciones_al_registrar_no_se_ocultan()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        lobjT.ErrorAlRegistrar = New InvalidOperationException("falla de BD")
        Dim lstrMens As String = String.Empty
        Dim lobjS = FobjSinc(lobjP, lobjA, lobjT)
        Assert.ThrowsException(Of InvalidOperationException)(
            Sub() lobjS.FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 1.3, False, MdtmHoy, lstrMens))
    End Sub
End Class
