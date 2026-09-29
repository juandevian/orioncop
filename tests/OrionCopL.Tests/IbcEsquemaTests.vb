Imports System.IO
Imports System.Linq
Imports System.Xml.Linq
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class IbcEsquemaTests

    Private Shared Function FxdocEsquema() As XDocument
        Dim ldirActual As DirectoryInfo = New DirectoryInfo(AppContext.BaseDirectory)
        While ldirActual IsNot Nothing
            Dim lstrComunes As String = Path.Combine(ldirActual.FullName, "Comunes")
            If Directory.Exists(lstrComunes) Then
                Dim lstrXml As String = Path.Combine(lstrComunes, "PanDat", "XmlBd", "OrionCop_Net.xml")
                If File.Exists(lstrXml) Then Return XDocument.Load(lstrXml)
            End If
            ldirActual = ldirActual.Parent
        End While
        Return Nothing
    End Function

    Private Shared Function FxelTabla(axdocEsquema As XDocument) As XElement
        Return axdocEsquema.Descendants("Tabla").FirstOrDefault(
                Function(x) CStr(x.Attribute("Nombre")) = "OriIbcCertificados")
    End Function

    <TestMethod>
    Public Sub VersionBd_Es268()
        Dim lxdoc = FxdocEsquema()
        If lxdoc Is Nothing Then Assert.Inconclusive("No se encontro la carpeta Comunes.")
        Assert.AreEqual("268", CStr(lxdoc.Descendants("BD").First().Attribute("Version")))
    End Sub

    <TestMethod>
    Public Sub TablaOriIbcCertificados_TieneLas8Columnas()
        Dim lxdoc = FxdocEsquema()
        If lxdoc Is Nothing Then Assert.Inconclusive("No se encontro la carpeta Comunes.")
        Dim lxelTabla = FxelTabla(lxdoc)
        Assert.IsNotNull(lxelTabla, "Falta la tabla OriIbcCertificados")
        Dim lstrEsperadas As String() = {"FechaCertificado:DATE", "FechaDescarga:DATE", "FechaDesde:DATE",
                "FechaHasta:DATE", "IdCarpeta:SHORT", "IdCentroUtil:SHORT", "Idfile:STRING:20", "Ibc:DOUBLE"}
        Dim lstrReales As String() = lxelTabla.Element("Columnas").Elements("Columna").Select(
                Function(c) CStr(c.Attribute("Nombre")) & ":" & CStr(c.Attribute("TipoDato")) &
                        If(c.Attribute("Longitud") Is Nothing, "", ":" & CStr(c.Attribute("Longitud")))).ToArray()
        CollectionAssert.AreEquivalent(lstrEsperadas, lstrReales)
    End Sub

    <TestMethod>
    Public Sub TablaOriIbcCertificados_IndicesReferencianColumnasExistentes()
        Dim lxdoc = FxdocEsquema()
        If lxdoc Is Nothing Then Assert.Inconclusive("No se encontro la carpeta Comunes.")
        Dim lxelTabla = FxelTabla(lxdoc)
        Assert.IsNotNull(lxelTabla, "Falta la tabla OriIbcCertificados")
        Dim lstrColumnas As New HashSet(Of String)(
                lxelTabla.Element("Columnas").Elements("Columna").Select(Function(c) CStr(c.Attribute("Nombre"))))
        For Each lxelIndice As XElement In lxelTabla.Element("Indices").Elements("Indice")
            For Each lxelCol As XElement In lxelIndice.Descendants("ColumnaIndice")
                Assert.IsTrue(lstrColumnas.Contains(CStr(lxelCol.Attribute("Nombre"))),
                        "Indice " & CStr(lxelIndice.Attribute("Nombre")) & " referencia columna inexistente " &
                        CStr(lxelCol.Attribute("Nombre")))
            Next
        Next
    End Sub

    <TestMethod>
    Public Sub TablaOriIbcCertificados_PkEsPrincipal()
        Dim lxdoc = FxdocEsquema()
        If lxdoc Is Nothing Then Assert.Inconclusive("No se encontro la carpeta Comunes.")
        Dim lxelTabla = FxelTabla(lxdoc)
        Assert.IsNotNull(lxelTabla, "Falta la tabla OriIbcCertificados")
        Dim lxelPk = lxelTabla.Element("Indices").Elements("Indice").FirstOrDefault(
                Function(i) CStr(i.Attribute("Nombre")) = "PK_OriIbcCertificados")
        Assert.IsNotNull(lxelPk, "Falta el indice PK_OriIbcCertificados")
        Assert.AreEqual("S", CStr(lxelPk.Attribute("Principal")))
    End Sub
End Class
