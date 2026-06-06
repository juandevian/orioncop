Imports System.IO
Module mdefOriInt
#Region "Definiciones"
    Private MSwErrorEFac As StreamWriter = Nothing
    Friend Const CMSTRURL_V1 = "https://www.misfacturas.com.co/IntegrationAPI/api/"
    Friend GblnEstadoRechazado As Boolean = False
#Region "Enumeradores"

#End Region
#End Region

#Region "Procedimientos"
    Friend Function FstrTipoPersonaDian(aEnuTipoTercero As EnuTipoTerceroDef) As String
        Dim lenuTipoPerDian As EnuTipoPersonaDian = EnuTipoPersonaDian.None
        Select Case aEnuTipoTercero
            Case EnuTipoTerceroDef.enuPersonaJuridica
                lenuTipoPerDian = EnuTipoPersonaDian.EnuJuridica
            Case EnuTipoTerceroDef.enuPersonaNatural
                lenuTipoPerDian = EnuTipoPersonaDian.EnuNatural
        End Select
        Dim lstrTipoPerDian As String = CInt(lenuTipoPerDian).ToString
        Return lstrTipoPerDian
    End Function

    Friend Function FstrMedioPagoDian(aenuTipoMedPago As EnuTipoMedioPagoDef) As EnuTipoMedioPagoDian
        Dim lenuTipoMedPagDian As EnuTipoMedioPagoDian = EnuTipoMedioPagoDian.None
        Select Case aenuTipoMedPago
            Case EnuTipoMedioPagoDef.EnuEfectivo
                lenuTipoMedPagDian = EnuTipoMedioPagoDian.EnuEfectivo
            Case EnuTipoMedioPagoDef.EnuCheque
                lenuTipoMedPagDian = EnuTipoMedioPagoDian.EnuCheque
            Case EnuTipoMedioPagoDef.EnuTarjetaCR
                lenuTipoMedPagDian = EnuTipoMedioPagoDian.EnuTarjetaCR
            Case EnuTipoMedioPagoDef.EnuTarjetaDB
                lenuTipoMedPagDian = EnuTipoMedioPagoDian.EnuTarjetaDB
            Case EnuTipoMedioPagoDef.EnuConsignacion
                lenuTipoMedPagDian = EnuTipoMedioPagoDian.EnuConsignacion
            Case EnuTipoMedioPagoDef.EnuTransferencia
                lenuTipoMedPagDian = EnuTipoMedioPagoDian.EnuTransferencia
            Case Else
        End Select
        Dim lstrMedPagDian As String = CType(lenuTipoMedPagDian, Integer).ToString
        Return lstrMedPagDian
    End Function

    Friend Function FstrIdImptoDian(ablnIva As Boolean, aenuTipoDescuento As EnuTipoDescuento) As String
        Dim lstrIdImpto = String.Empty
        If ablnIva Then
            lstrIdImpto = "01"
        Else
            Select Case aenuTipoDescuento
                Case EnuTipoDescuento.EnuReteIva
                    lstrIdImpto = "05"
                Case EnuTipoDescuento.EnuReteFuente
                    lstrIdImpto = "06"
                Case EnuTipoDescuento.EnuReteIca
                    lstrIdImpto = "07"
            End Select
        End If
        Return FstrComillas(lstrIdImpto)
    End Function

    Friend Function FstrDireccionCliente(aobjCliente As ClsCliente) As String
        Dim lstrDir = aobjCliente.ObjDireccionUnoStr.ToString
        Return FstrComillas(lstrDir)
    End Function

    ''' <summary>
    ''' Devuelve la fecha como texto según formato de Protecdata
    ''' </summary>
    ''' <param name="adtmfecha">Fecha aser formateada</param>
    ''' <returns>Fecha formateada según Protecdata</returns>
    Friend Function FstrFechaPro(adtmfecha As DateTime) As String
        Dim lstrFecha = adtmfecha.Year.ToString & "-" &
                Format(adtmfecha.Month, "00") & "-" & Format(adtmfecha.Day, "00") & "T" &
                Format(adtmfecha.Hour, "00") & ":" & Format(adtmfecha.Minute, "00") &
                ":" & Format(adtmfecha.Second, "00")
        Return lstrFecha
    End Function
    ' Aqui se traduce el estado recibido de mis facturas al estado en la aplicación
    Friend Function FenuEstadoEDoc(aobjEstadoDoc As ClsEstadoDoc) As EnuEstadoEDoc
        Dim lenuEstadoEDoc As EnuEstadoEDoc
        Dim lentDocStatus As Integer
        If aobjEstadoDoc Is Nothing OrElse (aobjEstadoDoc.DocumentStatus = 0) Then
            lentDocStatus = 0
        Else
            lentDocStatus = aobjEstadoDoc.DocumentStatus
        End If
        Select Case lentDocStatus
            Case Is = 0
                lenuEstadoEDoc = EnuEstadoEDoc.EnuErrorFtp
            Case Is < 70
                lenuEstadoEDoc = EnuEstadoEDoc.EnuEnProceso
            Case Is = 70
                lenuEstadoEDoc = EnuEstadoEDoc.EnuInvalida
            Case Is = 72
                lenuEstadoEDoc = EnuEstadoEDoc.EnuRegi
            Case Is = 73
                lenuEstadoEDoc = EnuEstadoEDoc.EnuRegi
            Case Is = 74
                lenuEstadoEDoc = EnuEstadoEDoc.EnuEnviada
            Case Is = 90
                lenuEstadoEDoc = EnuEstadoEDoc.EnuAceptada
            Case Is = 94
                lenuEstadoEDoc = EnuEstadoEDoc.EnuRechazada
            Case Else
                lenuEstadoEDoc = EnuEstadoEDoc.EnuOtro
        End Select
        Return lenuEstadoEDoc
    End Function
#End Region

#Region "Escribe archivo Error Post Efactura"
    Friend Sub SregistreError(aobjEstadoDoc As ClsEstadoDoc,
            aenuTipoDocOrigen As EnuTipoDocOri, astrNroDoc As String)
        Dim lstrDoc = String.Empty
        SInicialiceRepErrEfac()
        Select Case aenuTipoDocOrigen
            Case EnuTipoDocOri.EnuFactura
                lstrDoc = "Fac"
            Case EnuTipoDocOri.EnuNotaDb
                lstrDoc = "NDb"
            Case EnuTipoDocOri.EnuNotaCr
                lstrDoc = "NCr"
            Case EnuTipoDocOri.EnuReciboCaja
                lstrDoc = "RCa"
            Case EnuTipoDocOri.EnuNotaRevCr
                lstrDoc = "RCr"
            Case EnuTipoDocOri.EnuNotaCon
                lstrDoc = "NConAjuste"
        End Select
        Dim lstrFecha As String, lstrNroDoc As String, lstrEstado As String
        If aobjEstadoDoc Is Nothing Then
            lstrNroDoc = astrNroDoc
            lstrFecha = ClsPanoramaDat.FstrFechaHoraNormalizada(Date.Now)
            lstrEstado = "El Servidor reportó Error!"
        Else
            lstrFecha = aobjEstadoDoc.StatusDate
            lstrNroDoc = aobjEstadoDoc.DocumentNumber
            lstrEstado = FstrEstado(aobjEstadoDoc.DocumentStatus)
        End If
        SEscribaLinea(lstrFecha, lstrDoc, lstrNroDoc, lstrEstado, "", "")
        If aobjEstadoDoc IsNot Nothing Then
            If aobjEstadoDoc IsNot Nothing Then
                For Each lobjErrDian As ClsErrorDian In aobjEstadoDoc.DIANErrors
                    SEscribaLinea(lstrFecha, lstrDoc, lstrNroDoc, lstrEstado, lobjErrDian.Code,
                            lobjErrDian.Description)
                Next
            End If
        End If
        SCierreSW()
    End Sub
    Friend Sub SRegistreError(astrFecha As String, astrDoc As String, astrNroDoc As String,
                astrEstado As String, astrCodErr As String, astrDesc As String)
        SInicialiceRepErrEfac()
        Dim lstrLinea = astrFecha & "," & astrDoc & "," & astrNroDoc & "," & astrEstado & "," &
                astrCodErr & "," & astrDesc
        SEscribaErrEFac(lstrLinea)
        SCierreSW()
    End Sub
    Private Sub SEscribaLinea(astrFecha As String, astrDoc As String, astrNroDoc As String,
                astrEstado As String, astrCodErr As String, astrDesc As String)
        Dim lstrLinea = astrFecha & "," & astrDoc & "," & astrNroDoc & "," & astrEstado & "," &
                astrCodErr & "," & astrDesc
        SEscribaErrEFac(lstrLinea)
    End Sub
    Private Sub SInicialiceRepErrEfac()
        Dim lstrTray As String = GstrTrayEFac & "\" & ClsPanorama.FstrFechayyyymmdd(Date.Today)
        Dim lstrArchErrEFac = lstrTray & "_ErroresEFact.csv"
        If My.Computer.FileSystem.FileExists(lstrArchErrEFac) Then
            MSwErrorEFac = New StreamWriter(lstrArchErrEFac, True)
        Else
            MSwErrorEFac = New StreamWriter(lstrArchErrEFac)
            SEscribaErrEFac("")
        End If
    End Sub
    Private Sub SEscribaErrEFac(astrErrEFac As String)
        Dim lstrLinea As String
        If String.IsNullOrEmpty(astrErrEFac) Then
            lstrLinea = "Fecha, Doc, IdDoc, Estado, Cod. Err, Descripción"
        Else
            lstrLinea = astrErrEFac
        End If
        MSwErrorEFac.WriteLine(lstrLinea)
    End Sub
    Private Sub SCierreSW()
        If Not IsNothing(MSwErrorEFac) Then
            MSwErrorEFac.Close()
            MSwErrorEFac.Dispose()
        End If
    End Sub
#End Region

#Region "DIAN"
    Friend Function FstrTipoDocIdDian(aenuTipoDocIdOrion As EnuTipoDocIdDef) As String
        Dim lenuTipoDocIdDian As EnuTipoDocIdDian
        Select Case aenuTipoDocIdOrion
            Case EnuTipoDocIdDef.enuCedulaCiudadania
                lenuTipoDocIdDian = EnuTipoDocIdDian.enuCedulaCiudadania
            Case EnuTipoDocIdDef.enuTarjetaIdentidad
                lenuTipoDocIdDian = EnuTipoDocIdDian.enuTarjetaIdentidad
            Case EnuTipoDocIdDef.enuCedulaExtranjeria
                lenuTipoDocIdDian = EnuTipoDocIdDian.enuCedulaExtranjeria
            Case EnuTipoDocIdDef.enuNit
                lenuTipoDocIdDian = EnuTipoDocIdDian.enuNIT
            Case EnuTipoDocIdDef.enuNuip
                lenuTipoDocIdDian = EnuTipoDocIdDian.enuNUIP
            Case EnuTipoDocIdDef.enuRegistroCivil
                lenuTipoDocIdDian = EnuTipoDocIdDian.enuRegistroCivil
            Case EnuTipoDocIdDef.enuPasaporte
                lenuTipoDocIdDian = EnuTipoDocIdDian.enuPasaporte
            Case EnuTipoDocIdDef.enuDocIdExtranjero
                lenuTipoDocIdDian = EnuTipoDocIdDian.enuDocIdentidadExtranjero
            Case EnuTipoDocIdDef.enuTarjetaExtranjeria
                lenuTipoDocIdDian = EnuTipoDocIdDian.enuTarjetaExtranjeria
            Case EnuTipoDocIdDef.enuNitOtroPais
                lenuTipoDocIdDian = EnuTipoDocIdDian.enuNITotroPais
            Case EnuTipoDocIdDef.enuPEP
                lenuTipoDocIdDian = EnuTipoDocIdDian.enuPEP
            Case Else
                Throw New ErrorInesperadoPanLException("Tipo identidad sin definir!")
        End Select
        Dim lstrTipoDocDian As String = CType(lenuTipoDocIdDian, Integer).ToString
        Return lstrTipoDocDian
    End Function
    Friend Function FstrEstado(aentIdEstado As Integer) As String
        Dim lstrEstado = String.Empty
        Select Case aentIdEstado
            Case 40
                lstrEstado = "Creado"
            Case 41
                lstrEstado = "Anulado"
            Case 42
                lstrEstado = "Actualizado"
            Case 46
                lstrEstado = "Paso 1 de 4"
            Case 47
                lstrEstado = "Paso 2 de 4"
            Case 48
                lstrEstado = "Paso 3 de 4"
            Case 49
                lstrEstado = "Paso 4 de 4"
            Case 70
                lstrEstado = "Invalido"
            Case 72
                lstrEstado = "Valido"
            Case 73
                lstrEstado = "AD Valido"
            Case 74
                lstrEstado = "AD Enviado"
            Case 75
                lstrEstado = "Esperando RG"
            Case 80
                lstrEstado = "Fallido"
            Case 90
                lstrEstado = "Aceptación Expresa"
            Case 91
                lstrEstado = "Recibo Bien o Servicio"
            Case 92
                lstrEstado = "Acusado"
            Case 93
                lstrEstado = "Aceptación Tácita"
            Case 94
                lstrEstado = "Reclamado"
            Case 96
                lstrEstado = "Contingencia"
            Case 97
                lstrEstado = "Pendiente por envío de correo"
            Case 98
                lstrEstado = "Pendiente por validación DIAN"
        End Select
        lstrEstado = aentIdEstado.ToString & "-" & lstrEstado
        Return lstrEstado
    End Function
#End Region
End Module