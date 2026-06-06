Public Class WinEstadoDocEFac
#Region "Definiciones"
    Private MstrCUDE As String = String.Empty
    Private MstrCUFE As String = String.Empty
    Private MstrEstado As String = String.Empty
    Private MstrEstadoApp As String = String.Empty
    Private MstrFechaEmision As String = String.Empty
    Private MstrNroDocEstado As String = String.Empty
    Private MstrIdCliente As String = String.Empty
    Private MstrNombreCliente As String = String.Empty
    Private MstrNroFacBase As String = String.Empty
    Private ReadOnly MblnVer2 As Boolean = False
    Private MdtbErroresDian As DataTable = Nothing
#End Region
#Region "Propiedades"
    Friend Property ObjEstadoDoc As ClsEstadoDoc = Nothing
    Friend Property EnuTipoDoc As EnuTipoDocOri = EnuTipoDocOri.None
    Friend Property StrIdDoc As String = String.Empty
    Friend Property EnuEstadoEDocEnApp As EnuEstadoEDoc
#End Region
    Sub New()
        InitializeComponent()
        MblnVer2 = True
    End Sub
    Sub New(astrEstado As String)
        InitializeComponent()
        MstrEstado = astrEstado
        MblnVer2 = False
    End Sub
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        If MblnVer2 Then
            SProceseRespuesta()
        Else
            SProceseEstado()
        End If
        SMuestreDatos()
    End Sub
    Private Sub SMuestreDatos()
        cnvFact.Visibility = Visibility.Collapsed
        cnvNota.Visibility = Visibility.Collapsed
        If EnuTipoDoc = EnuTipoDocOri.EnuFactura Then
            cnvFact.Visibility = Visibility.Visible
            SMuestreFactura()
        ElseIf EnuTipoDoc <> EnuTipoDocOri.None Then
            cnvNota.Visibility = Visibility.Visible
            SMuestreNota()
        End If
    End Sub
    Private Sub SMuestreFactura()
        lblTipoDoc.Content = lblTipoDoc.Content & " " & " FACTURA"
        txtCUFE.Content = MstrCUFE
        txtCUFE.ToolTip = MstrCUFE
        txtIdDocEFac.Content = StrIdDoc
        txtFechaEst.Content = MstrFechaEmision
        txtNroFac.Content = MstrNroDocEstado
        txtIdCliente.Content = MstrIdCliente
        txtNombreCliente.Content = MstrNombreCliente
        txtNombreCliente.ToolTip = MstrNombreCliente
        txtEstadoEnAPI.Content = MstrEstado
        txtEstadoEnAPI.ToolTip = MstrEstado
        txtEstadoEnApp.Content = MstrEstadoApp
        txtEstadoEnApp.ToolTip = MstrEstadoApp
    End Sub
    Private Sub SMuestreNota()
        lblTipoDoc.Content = lblTipoDoc.Content & " " & " NOTA"
        txtCUDE.Content = MstrCUDE
        txtCUDE.ToolTip = MstrCUDE
        txtCUFEFacBase.Content = MstrCUFE
        txtCUFEFacBase.ToolTip = MstrCUFE
        txtIdDocENot.Content = StrIdDoc
        txtFechaEstNot.Content = MstrFechaEmision
        txtNroNot.Content = MstrNroDocEstado
        txtIdClienteNot.Content = MstrIdCliente
        txtNombClienteNot.Content = MstrNombreCliente
        txtNombClienteNot.ToolTip = MstrNombreCliente
        txtIdFacturaBase.Content = MstrNroFacBase
        txtEstadoEnAPINot.Content = MstrEstado
        txtEstadoEnAPINot.ToolTip = MstrEstado
        txtEstadoEnAppNot.Content = MstrEstadoApp
        txtEstadoEnAppNot.ToolTip = MstrEstadoApp
    End Sub
    Private Sub SProceseRespuesta()
        Dim lentEstadoAPI As Integer
        If ObjEstadoDoc IsNot Nothing Then
            ' Estado Documento
            lentEstadoAPI = ObjEstadoDoc.DocumentStatus
            MstrEstado = FstrEstado(lentEstadoAPI)
            MstrEstadoApp = FstrNombreEstadoEFacEnApp(EnuEstadoEDocEnApp)
            ' Fecha Estado
            MstrFechaEmision = ObjEstadoDoc.StatusDate
            ' Número documento
            MstrNroDocEstado = ObjEstadoDoc.DocumentNumber
            ' Cliente
            MstrIdCliente = ObjEstadoDoc.CustomerPartyID
            MstrNombreCliente = ObjEstadoDoc.CustomerParty
            MstrCUFE = ObjEstadoDoc.CUFE
            MstrCUDE = ObjEstadoDoc.CUDE
            MstrNroFacBase = ObjEstadoDoc.InvoiceNumber
            SGenereDataTableErrores()
        Else
            lentEstadoAPI = 80
            MstrEstado = FstrEstado(lentEstadoAPI)
            MstrEstadoApp = FstrNombreEstadoEFacEnApp(EnuEstadoEDocEnApp)
            ' Fecha Estado
            MstrFechaEmision = "01-01-1900"
            ' Número documento
            MstrNroDocEstado = "Sin"
            ' Cliente
            MstrIdCliente = "Sin"
            MstrNombreCliente = "Sin"
            MstrCUFE = "Sin"
            MstrCUDE = "Sin"
            MstrNroFacBase = "Sin"
        End If
    End Sub
    Private Sub SProceseEstado()
        Dim lstrPartes As String() = MstrEstado.Split(",")
        Dim lentEstado As Integer = CType(lstrPartes(0).Split(":")(1).Trim, Integer)
        Dim lentPosIniFecha = lstrPartes(1).IndexOf(":") + 1
        MstrFechaEmision = lstrPartes(1).Substring(lentPosIniFecha).Trim
        MstrFechaEmision = MstrFechaEmision.Replace(Chr(34), "")
        MstrCUDE = String.Empty
        If EnuTipoDoc = EnuTipoDocOri.EnuFactura Then
            MstrCUFE = (lstrPartes(2).Split(":")(1).Trim).Replace(Chr(34), "")
            MstrNroDocEstado = (lstrPartes(3).Split(":")(1).Trim).Replace(Chr(34), "")
            MstrIdCliente = (lstrPartes(5).Split(":")(1).Trim).Replace(Chr(34), "")
            MstrNombreCliente = (lstrPartes(4).Split(":")(1).Trim).Replace(Chr(34), "")
        Else
            MstrCUFE = String.Empty
            MstrNroDocEstado = (lstrPartes(2).Split(":")(1).Trim).Replace(Chr(34), "")
            MstrIdCliente = (lstrPartes(4).Split(":")(1).Trim).Replace(Chr(34), "")
            MstrNombreCliente = (lstrPartes(3).Split(":")(1).Trim).Replace(Chr(34), "")
        End If
        MstrEstado = FstrDetalleEstado(lentEstado)
    End Sub
    Private Sub SGenereDataTableErrores()
        Dim ldrwNuevoError As DataRow, i = 0
        Dim ldclOrdinal As New DataColumn("Id", System.Type.GetType("System.Int16"))
        Dim ldclCodErr As New DataColumn("CodError", System.Type.GetType("System.String"))
        Dim ldclDescripcion As New DataColumn("Descripcion", System.Type.GetType("System.String"))
        MdtbErroresDian = New DataTable
        MdtbErroresDian.Columns.Add(ldclOrdinal)
        MdtbErroresDian.Columns.Add(ldclCodErr)
        MdtbErroresDian.Columns.Add(ldclDescripcion)
        If ObjEstadoDoc.DIANErrors IsNot Nothing Then
            For Each lobjDianErr As ClsErrorDian In ObjEstadoDoc.DIANErrors
                i += 1
                ldrwNuevoError = MdtbErroresDian.NewRow
                ldrwNuevoError("Id") = i
                ldrwNuevoError("CodError") = lobjDianErr.Code
                ldrwNuevoError("Descripcion") = lobjDianErr.Description
                MdtbErroresDian.Rows.Add(ldrwNuevoError)
            Next
        End If
        dgrErroresDian.DataContext = MdtbErroresDian
    End Sub
    Private Shared Function FstrDetalleEstado(aentEstado As Integer)
        Dim lenuProvEFac As EnuProveedorEFac = GobjParametros.ObjIdProvEFacByt.ObjValorPro
        Dim lstrDetEstado As String
        Select Case lenuProvEFac
            Case EnuProveedorEFac.EnuProtecdataMisFac
                If aentEstado <= 10 Then
                    lstrDetEstado = FstrDetEsta_1_10MF(aentEstado)
                Else
                    lstrDetEstado = FstrDetEsta_11MF(aentEstado)
                End If
            Case Else
                Throw New ErrorInesperadoPanLException("Proveedor eFactura no esperado")
        End Select
        Return lstrDetEstado
    End Function
    Private Shared Function FstrDetEsta_1_10MF(aentEstado As Integer)
        Dim lstrDetEstado As String
        Select Case aentEstado
            Case 1
                lstrDetEstado = "Creado"
            Case 2
                lstrDetEstado = "Emitido"
            Case 3
                lstrDetEstado = "Acusado"
            Case 4
                lstrDetEstado = "Aceptado"
            Case 5
                lstrDetEstado = "Rechazado"
            Case 6
                lstrDetEstado = "Anulado"
            Case 7
                lstrDetEstado = "Modificado"
            Case 8
                lstrDetEstado = "Eliminado"
            Case 9
                lstrDetEstado = "En proceso de generación XML"
            Case 10
                lstrDetEstado = "En proceso de firmado"
            Case 11
                lstrDetEstado = "En proceso de envío a la DIAN"
            Case 12
                lstrDetEstado = "12 - Sin Descripción"
            Case 13
                lstrDetEstado = "13 - Sin Descripción"
            Case 14
                lstrDetEstado = "14 - Sin Descripción"
            Case 15
                lstrDetEstado = "15 - Sin Descripción"
            Case Else
                lstrDetEstado = aentEstado.ToString & " - Sin Descripción"
        End Select
        Return lstrDetEstado
    End Function
    Private Shared Function FstrDetEsta_11MF(aentEstado As Integer)
        Dim lstrDetEstado As String
        Select Case aentEstado
            Case 11
                lstrDetEstado = "En proceso de envío a la DIAN"
            Case 12
                lstrDetEstado = "Para Firma"
            Case 15
                lstrDetEstado = "Fallo"
            Case 18
                lstrDetEstado = "Emitido"
            Case 19
                lstrDetEstado = "Reasignado"
            Case 20
                lstrDetEstado = "Recibido"
            Case 21
                lstrDetEstado = "Pendiente de Pago"
            Case 22
                lstrDetEstado = "Fallo Pago"
            Case 23
                lstrDetEstado = "Pago exitoso"
            Case Else
                lstrDetEstado = aentEstado.ToString & " - Sin Descripción"
        End Select
        Return lstrDetEstado
    End Function
    Private Sub BttCerrar_Click(sender As Object, e As RoutedEventArgs) Handles bttCerrar.Click
        Me.Close()
    End Sub
End Class
