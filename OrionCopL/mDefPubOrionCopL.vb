Imports System.Drawing
Imports System.IO
Imports System.Net.NetworkInformation
Module MDefPubOrionCopL
#Region "Constantes globales Solucion"
    Friend Const GCSHRANOMINIMO As Short = 2000
    Friend Const GCSHRANOMAXIMO As Short = 2199
    Friend Const GCSHRIDMORA As Short = 999
    Friend Const GCSTRPREDIONULO As String = "*"
    Friend Const GCSTRPREFPREFACTURA As String = "***"
    Friend Const GCSTRCUENTADSCTOCAP As String = "******"
    Friend Const GCSTRCUENTADSCTOINT As String = "%%%%%%"
    Friend Const GCSTRSINPA As String = "Sin Predio Agr."
#End Region
#Region "Variables globales de la solucion"
    Friend GobjParametros As ClsCentroUtilOriCop = Nothing
    ' Limites Edad Cartera
    Friend GentLimite1 As Integer = 30
    Friend GentLimite2 As Integer = 60
    Friend GentLimite3 As Integer = 120
    Friend GentLimite4 As Integer = 210
    '
    Friend GarlServiciosImportados As ArrayList
    Friend GblnCausandoFM As Boolean = False
#End Region
#Region "Funciones y procedimientos Globales"
    Friend Function FblnStrArrayIguales(astrArrayUno As String(), astrArrayDos As String()) As Boolean
        If astrArrayUno Is Nothing OrElse astrArrayDos Is Nothing Then
            Return False
        End If
        Dim lblnIguales = (astrArrayUno.Length = astrArrayDos.Length)
        If lblnIguales Then
            For Each lstrInte As String In astrArrayUno
                If Not astrArrayDos.Contains(lstrInte) Then
                    lblnIguales = False
                    Exit For
                End If
            Next
        End If
        Return lblnIguales
    End Function
    Public Function FstrTrayecCalculadoraExe() As String
        Dim dir As String
        dir = Environment.SystemDirectory
        If dir & "\CALC.EXE" <> "" Then
            Return dir & "\CALC.EXE"
        Else
            Return ""
        End If
    End Function
    Friend Function FstrTrayecBlockNotasExe() As String
        Dim dir As String
        dir = Environment.SystemDirectory
        If dir & "\NOTEPAD.EXE" <> "" Then
            Return dir & "\NOTEPAD.EXE"
        Else
            Return ""
        End If
    End Function
    Friend Function FstrDiaFecha(adtmFecha As Date)
        Dim lstrDia = String.Empty
        Select Case adtmFecha.DayOfWeek
            Case DayOfWeek.Monday
                lstrDia = "Lunes"
            Case DayOfWeek.Tuesday
                lstrDia = "Martes"
            Case DayOfWeek.Wednesday
                lstrDia = "Miercoles"
            Case DayOfWeek.Thursday
                lstrDia = "Jueves"
            Case DayOfWeek.Friday
                lstrDia = "Viernes"
            Case DayOfWeek.Saturday
                lstrDia = "Sabado"
            Case DayOfWeek.Sunday
                lstrDia = "Domingo"
        End Select
        Return lstrDia
    End Function
    Friend Function FentDayOfWeek(astrNombreDia As String) As DayOfWeek
        Dim lentDiaDeLaSemana As DayOfWeek = DayOfWeek.Sunday
        Select Case astrNombreDia
            Case "Lunes"
                lentDiaDeLaSemana = DayOfWeek.Monday
            Case "Martes"
                lentDiaDeLaSemana = DayOfWeek.Tuesday
            Case "Miercoles"
                lentDiaDeLaSemana = DayOfWeek.Wednesday
            Case "Jueves"
                lentDiaDeLaSemana = DayOfWeek.Thursday
            Case "Viernes"
                lentDiaDeLaSemana = DayOfWeek.Friday
            Case "Sabado"
                lentDiaDeLaSemana = DayOfWeek.Saturday
            Case "Domingo"
                lentDiaDeLaSemana = DayOfWeek.Sunday
        End Select
        Return lentDiaDeLaSemana
    End Function
    ''' <summary>
    ''' Devuelve la semana a la cual pertenece una fecha
    ''' </summary>
    ''' <param name="adtmFecha">Fecha de la cual se va a devolver la semana a la cual pertenece</param>
    ''' <returns></returns>
    ''' <remarks>Semana 1: 1-7; Semana 2: 8-14; Semana 3: 15-24 Semana 4: >=25 </remarks>
    Friend Function FentSemanaFecha(adtmFecha As Date) As Integer
        Dim lIntDiaFecha = adtmFecha.Day
        Dim lintSemana = 0
        If lIntDiaFecha >= 1 AndAlso lIntDiaFecha <= 7 Then
            lintSemana = 1
        End If
        If lIntDiaFecha >= 8 AndAlso lIntDiaFecha <= 14 Then
            lintSemana = 2
        End If
        If lIntDiaFecha >= 15 AndAlso lIntDiaFecha <= 21 Then
            lintSemana = 3
        End If
        If lIntDiaFecha > 21 Then
            lintSemana = 4
        End If
        Return lintSemana
    End Function
    Friend Function FstrMesFecha(adtmFecha As Date)
        Dim lstrMes = String.Empty
        Select Case adtmFecha.Month
            Case 1
                lstrMes = "Enero"
            Case 2
                lstrMes = "Febrero"
            Case 3
                lstrMes = "Marzo"
            Case 4
                lstrMes = "Abril"
            Case 5
                lstrMes = "Mayo"
            Case 6
                lstrMes = "Junio"
            Case 7
                lstrMes = "Julio"
            Case 8
                lstrMes = "Agosto"
            Case 9
                lstrMes = "Septiembre"
            Case 10
                lstrMes = "Octubre"
            Case 11
                lstrMes = "Noviembre"
            Case 12
                lstrMes = "Diciembre"
        End Select
        Return lstrMes
    End Function
    Friend Function FstrMesNroMes(aentMes As Integer)
        Dim lstrMes = String.Empty
        Select Case aentMes
            Case 1
                lstrMes = "Enero"
            Case 2
                lstrMes = "Febrero"
            Case 3
                lstrMes = "Marzo"
            Case 4
                lstrMes = "Abril"
            Case 5
                lstrMes = "Mayo"
            Case 6
                lstrMes = "Junio"
            Case 7
                lstrMes = "Julio"
            Case 8
                lstrMes = "Agosto"
            Case 9
                lstrMes = "Septiembre"
            Case 10
                lstrMes = "Octubre"
            Case 11
                lstrMes = "Noviembre"
            Case 12
                lstrMes = "Diciembre"
        End Select
        Return lstrMes
    End Function
    Friend Function FstrCadenaRecortada(astrCadenaOriginal As String, aentLongitudMax As Integer)
        Dim lstrCadenaRec As String
        If astrCadenaOriginal.Length > aentLongitudMax Then
            lstrCadenaRec = astrCadenaOriginal.Substring(0, aentLongitudMax)
        Else
            lstrCadenaRec = astrCadenaOriginal
        End If
        Return lstrCadenaRec
    End Function
    Friend Sub SLimpieCarpeta(astrNombreCar As String, astrTipoArch As String)
        Try
            Dim lstrArchivos = My.Computer.FileSystem.GetFiles(astrNombreCar,
                          FileIO.SearchOption.SearchAllSubDirectories, astrTipoArch)
            For Each lstrArchivo In lstrArchivos
                My.Computer.FileSystem.DeleteFile(lstrArchivo)
            Next
        Catch ex As ArgumentNullException
            Throw
        Catch ex As ArgumentException
            Throw
        Catch ex As PathTooLongException
            Throw
        Catch ex As NotSupportedException
            Throw
        Catch ex As FileNotFoundException
            Throw
        Catch ex As IOException
            Throw
        Catch ex As UnauthorizedAccessException
            Throw
        End Try
    End Sub
    Friend Function FenuTipoDsctoOReten(aenuTipoItemTec As EnuTipoItemRecCajaDef) As EnuTipoDescuento
        Dim lenuTipoDscto As EnuTipoDescuento = EnuTipoDescuento.None
        Select Case aenuTipoItemTec
            Case EnuTipoItemRecCajaDef.EnuDsctoCapital
                lenuTipoDscto = EnuTipoDescuento.EnuDsctoCapital
            Case EnuTipoItemRecCajaDef.EnuDsctoIntMora
                lenuTipoDscto = EnuTipoDescuento.EnuDsctoIntMora
            Case EnuTipoItemRecCajaDef.EnuDsctoPP
                lenuTipoDscto = EnuTipoDescuento.EnuDsctoPP
            Case EnuTipoItemRecCajaDef.EnuReteFuente
                lenuTipoDscto = EnuTipoDescuento.EnuReteFuente
            Case EnuTipoItemRecCajaDef.EnuReteIca
                lenuTipoDscto = EnuTipoDescuento.EnuReteIca
            Case EnuTipoItemRecCajaDef.EnuReteIva
                lenuTipoDscto = EnuTipoDescuento.EnuReteIva
        End Select
        Return lenuTipoDscto
    End Function
    Friend Function FenuTipoNov(aenuTipoDscto As EnuTipoDescuento) As EnuTipoNov
        Dim lenuTipoNov As EnuTipoNov = EnuTipoNov.None
        Select Case aenuTipoDscto
            Case EnuTipoDescuento.EnuDsctoCapital, EnuTipoDescuento.EnuDsctoPP
                lenuTipoNov = EnuTipoNov.EnuCrPagoCap
            Case EnuTipoDescuento.EnuDsctoIntMora
                lenuTipoNov = EnuTipoNov.EnuCrPagoInt
            Case EnuTipoDescuento.EnuReteFuente
                lenuTipoNov = EnuTipoNov.EnuCrRetFte
            Case EnuTipoDescuento.EnuReteIca
                lenuTipoNov = EnuTipoNov.EnuCrRetIca
            Case EnuTipoDescuento.EnuReteIva
                lenuTipoNov = EnuTipoNov.EnuCrRetIva
        End Select
        Return lenuTipoNov
    End Function
    ''' <summary>
    ''' Devuelve la letra de la columna de Excel que corresponde al indice indicado por el argumento "aentIndice"
    ''' </summary>
    ''' <param name="aentIndice">Indica la posición de la columna en la Hoja de Excel</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FstrLetraDeIndice(aentIndice As Integer) As String
        Dim lstrLetra As String = String.Empty
        If aentIndice > 60 Then
            Throw New ErrorInesperadoPanLException("No se soportan mas de sesenta columnas")
        End If
        If aentIndice > 10 AndAlso aentIndice <= 20 Then
            lstrLetra = FstrLetraDeIndice_10(aentIndice)
        ElseIf aentIndice > 20 AndAlso aentIndice <= 40 Then
            lstrLetra = FstrLetraDeIndice_20(aentIndice)
        ElseIf aentIndice > 40 AndAlso aentIndice <= 60 Then
            lstrLetra = FstrLetraDeIndice_40(aentIndice)
        Else
            Select Case aentIndice
                Case Is = 1
                    lstrLetra = "A"
                Case Is = 2
                    lstrLetra = "B"
                Case Is = 3
                    lstrLetra = "C"
                Case Is = 4
                    lstrLetra = "D"
                Case Is = 5
                    lstrLetra = "E"
                Case Is = 6
                    lstrLetra = "F"
                Case Is = 7
                    lstrLetra = "G"
                Case Is = 8
                    lstrLetra = "H"
                Case Is = 9
                    lstrLetra = "I"
                Case Is = 10
                    lstrLetra = "J"
            End Select
        End If
        Return lstrLetra
    End Function
    Private Function FstrLetraDeIndice_10(aentIndice As Integer) As String
        Dim lstrLetra = String.Empty
        Select Case aentIndice
            Case Is = 11
                lstrLetra = "K"
            Case Is = 12
                lstrLetra = "L"
            Case Is = 13
                lstrLetra = "M"
            Case Is = 14
                lstrLetra = "M"
            Case Is = 15
                lstrLetra = "O"
            Case Is = 16
                lstrLetra = "P"
            Case Is = 17
                lstrLetra = "Q"
            Case Is = 18
                lstrLetra = "R"
            Case Is = 19
                lstrLetra = "S"
            Case Is = 20
                lstrLetra = "T"
        End Select
        Return lstrLetra
    End Function
    Private Function FstrLetraDeIndice_20(aentIndice As Integer) As String
        Dim lstrLetra = String.Empty
        Select Case aentIndice
            Case Is = 21
                lstrLetra = "U"
            Case Is = 22
                lstrLetra = "V"
            Case Is = 23
                lstrLetra = "W"
            Case Is = 24
                lstrLetra = "X"
            Case Is = 25
                lstrLetra = "Y"
            Case Is = 26
                lstrLetra = "Z"
            Case Is = 27
                lstrLetra = "AA"
            Case Is = 28
                lstrLetra = "AB"
            Case Is = 29
                lstrLetra = "AC"
            Case Is = 30
                lstrLetra = "AD"
            Case Is = 31
                lstrLetra = "AE"
            Case Is = 32
                lstrLetra = "AF"
            Case Is = 33
                lstrLetra = "AG"
            Case Is = 34
                lstrLetra = "AH"
            Case Is = 35
                lstrLetra = "AI"
            Case Is = 36
                lstrLetra = "AJ"
            Case Is = 37
                lstrLetra = "AK"
            Case Is = 38
                lstrLetra = "AL"
            Case Is = 39
                lstrLetra = "AM"
            Case Is = 40
                lstrLetra = "AN"
        End Select
        Return lstrLetra
    End Function
    Private Function FstrLetraDeIndice_40(aentIndice As Integer) As String
        Dim lstrLetra = String.Empty
        Select Case aentIndice
            Case Is = 41
                lstrLetra = "AO"
            Case Is = 42
                lstrLetra = "AP"
            Case Is = 43
                lstrLetra = "AQ"
            Case Is = 44
                lstrLetra = "AR"
            Case Is = 45
                lstrLetra = "AS"
            Case Is = 46
                lstrLetra = "AT"
            Case Is = 47
                lstrLetra = "AU"
            Case Is = 48
                lstrLetra = "AV"
            Case Is = 49
                lstrLetra = "AW"
            Case Is = 50
                lstrLetra = "AX"
            Case Is = 51
                lstrLetra = "AY"
            Case Is = 52
                lstrLetra = "AZ"
            Case Is = 53
                lstrLetra = "BA"
            Case Is = 54
                lstrLetra = "BB"
            Case Is = 55
                lstrLetra = "BC"
            Case Is = 56
                lstrLetra = "BD"
            Case Is = 57
                lstrLetra = "BE"
            Case Is = 58
                lstrLetra = "BF"
            Case Is = 59
                lstrLetra = "BG"
            Case Is = 60
                lstrLetra = "BH"
        End Select
        Return lstrLetra
    End Function
    Friend Function FstrComillas(astrCadena As String) As String
        Dim lstrCadena = Chr(34) & astrCadena.Trim & Chr(34)
        Return lstrCadena
    End Function
    Friend Sub SEspere(aentMins As Integer, aentSegs As Integer, aentMilSeg As Integer)
        Dim ldtmInicio As DateTime
        Dim ldtmFin = Now.AddMinutes(aentMins).AddSeconds(aentSegs).AddMilliseconds(aentMilSeg)
        Do While ldtmInicio <= ldtmFin
            ldtmInicio = Now
        Loop
    End Sub
    Friend Function FblnHayInternet()
        Try
            Dim ping As New Ping()
            Dim reply As PingReply = ping.Send("8.8.8.8", 1000) ' Google DNS
            Return reply.Status = IPStatus.Success
        Catch ex As Exception
            Return False
        End Try
    End Function
    Friend Function FblnEstaConectado(astrUrl As String, ByRef astrMens As String) As Boolean
        Const LCMF = "https://www.misfacturas.com.co/"
        Const LCG = "https://accounts.google.com/"
        Const LCM = "https://www.msn.com/es-co"
        Dim lblnEstaCon As Boolean = My.Computer.Network.IsAvailable
        If lblnEstaCon Then
            If String.IsNullOrEmpty(astrUrl) Then
                lblnEstaCon = FblnEstaCon(LCMF) AndAlso
                        FblnEstaCon(LCG) AndAlso
                        FblnEstaCon(LCM)
                If Not lblnEstaCon Then
                    astrMens = "No está conectado a Internet en este momento!"
                End If
            Else
                Dim lstrUrlOrigen As String = GobjParametros.ObjURLStr.ToString
                Dim i = 0, lstrUrl = String.Empty
                For Each lchrUrl As Char In lstrUrlOrigen
                    If lchrUrl = "/" Then
                        i += 1
                    End If
                    lstrUrl &= lchrUrl
                    If i = 3 Then Exit For
                Next
                lblnEstaCon = FblnEstaCon(lstrUrl)
                If Not lblnEstaCon Then
                    astrMens = "El Servidor parece no estar disponible en este momento!"
                End If
            End If
        Else
            astrMens = "No está conectado a Internet en este momento!"
        End If
        Return lblnEstaCon
    End Function
    Friend Function FblnEstaConectado() As Boolean
        Try
            Using lwecCliente = New System.Net.WebClient()
                Using stream = lwecCliente.OpenRead("http://www.google.com")
                    Return True
                End Using
            End Using
        Catch
            Return False
        End Try
    End Function
    Private Function FblnEstaCon(astrUrl As String) As Boolean
        Dim lobjUrl As New Uri(astrUrl)
        Dim lobjWebReq As System.Net.WebRequest
        Dim lobjResp As System.Net.WebResponse
        Try
            lobjWebReq = System.Net.WebRequest.Create(lobjUrl)
            lobjResp = lobjWebReq.GetResponse
            lobjResp.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Friend Function FstrNombreTipoDocId(aenuTipoDocId As EnuTipoDocIdDef) As String
        Dim lstrNom As String
        Select Case aenuTipoDocId
            Case EnuTipoDocIdDef.enuCedulaCiudadania
                lstrNom = "Cédula de Ciudadanía"
            Case EnuTipoDocIdDef.enuCedulaExtranjeria
                lstrNom = "Cédula de Extranjería"
            Case EnuTipoDocIdDef.enuDocIdExtranjero
                lstrNom = "Documento de Identidad extranjero"
            Case EnuTipoDocIdDef.enuNit
                lstrNom = "NIT"
            Case EnuTipoDocIdDef.enuPasaporte
                lstrNom = "Pasaporte"
            Case EnuTipoDocIdDef.enuPEP
                lstrNom = "Permiso especial de Permanencia"
            Case EnuTipoDocIdDef.enuTarjetaExtranjeria
                lstrNom = "Tarjeta de Extranjería"
            Case EnuTipoDocIdDef.enuRegistroCivil
                lstrNom = "Registro Civil"
            Case EnuTipoDocIdDef.enuTarjetaIdentidad
                lstrNom = "Tarjeta de Identidad"
            Case Else
                lstrNom = String.Empty
        End Select
        Return lstrNom
    End Function
    Friend Function FBytQR(abmQR As Bitmap) As Byte()
        Dim lbytQR As Byte() = Array.Empty(Of Byte)()
        If abmQR IsNot Nothing Then
            Using lmstImagenBinaria As New MemoryStream
                abmQR.Save(lmstImagenBinaria, Imaging.ImageFormat.Jpeg)
                lbytQR = lmstImagenBinaria.GetBuffer
            End Using
        End If
        Return lbytQR
    End Function
    Friend Function FblnFechaDocEsPeriodoActual(adtmFechaDoc As Date) As Boolean
        Dim ldtmFechaFinPeract = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
        Dim lblnEsPeriodoActual = ldtmFechaFinPeract.Year = adtmFechaDoc.Year AndAlso
                ldtmFechaFinPeract.Month = adtmFechaDoc.Month
        Return lblnEsPeriodoActual
    End Function
    Friend Function FdtbClientesConCorreo(adblIdClienteAPArtirDe As Double) As DataTable
        Dim lstrTablaPri = ClsPropietario.SstrNombreTabla
        Dim lstrTablaSec = ClsCliente.SstrNombreTabla
        Dim lstrCamSelPri As String() = {"DISTINCT " & ClsIdCliente_PropDbl.SstrNombreCampoBd}
        Dim lstrCamSelSec As String() = {}
        Dim lstrCamRelPri As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsIdCliente_PropDbl.SstrNombreCampoBd}
        Dim lstrCamRelSec As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsRecibeDocsPorEmailBln.SstrNombreCampoBd & " = TRUE AND " &
                ClsPorcentajePartiDbl.SstrNombreCampoBd & " > 0" & " AND " &
                ClsIdCliente_PropDbl.SstrNombreCampoBd & " > " & adblIdClienteAPArtirDe
        Dim lstrOrden As String(,) = {{ClsIdCliente_PropDbl.SstrNombreCampoBd, "ASC"}}
        Dim ldtbResu = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamSelPri, lstrTablaSec,
                lstrCamSelSec, lstrCamRelPri, lstrCamRelSec, lstrOrden, False, lstrFiltro, {})
        Return ldtbResu
    End Function
    Friend Function FblnExisteExcel() As Boolean
        Dim excelType As Type = Type.GetTypeFromProgID("Excel.Application")
        Dim lblnExiste As Boolean = excelType IsNot Nothing
        Return lblnExiste
    End Function
#End Region
End Module