Imports System.IO
Imports Microsoft.Win32
Friend Module mOrionCopIU
#Region "Definiciones"
#Region "Constantes"
    Friend Const GCOBJREGISTRO As Object = "A0b1f9*hjBó^23ö~"
    ' Constantes globales
    Friend Const GCSTRNINGUNO As String = "0 - Ninguno"
#End Region
#Region "Variables"
    Friend GstrVersionApp As String = My.Application.Info.Version.ToString
    Friend GstrVerAntApp As String = String.Empty
    Friend Property GblnCorreoOn As Boolean = False
    Friend BlnFactAuto As Boolean = False
    Friend BlnFactConti As Boolean = False
    Friend EnuEstadoAyuda As EnuEstadoAyudaDef = EnuEstadoAyudaDef.EnuOff
    ' Variables de Modulo
    Private WithEvents MwinCorreo As WinCorreoE = Nothing
    Private MwinProcesoEFac As WinProcesaEFac = Nothing
    Private MblnActualizoApl As Boolean = False
#End Region
#End Region

#Region "Inicio Aplicacion"
    Friend Sub SInicialiceApp()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GenuIdAplicacion = EnuListaAplicaciones.EnuOrionCop
            Dim lobjOrionCop As New ClsOrionCop(GCOBJREGISTRO, True)
            GobjAdministrador.SInicieApp()
            SInicieSesion()
            SSetAddRemoveProgramsIcon()
            lblnNoHayError = True
        Catch ex As ModuloNoRegistradoPanException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ErrorInesperadoPanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                Dim lblnConfiguRegionalOk = FblnConfNumeroOk()
                If lblnConfiguRegionalOk Then
                    lblnConfiguRegionalOk = FblnConfMonedaOk()
                    If Not lblnConfiguRegionalOk Then
                        lstrMens = "La configuración regional de Windows de Moneda no es adecauda." & vbCrLf &
                                        "Orión Plus no se puede correr!"
                    End If
                Else
                    lstrMens = "La configuración regional de Windows de Numero no es adecauda." & vbCrLf &
                                        "Orión Plus no se puede correr!"
                End If
                If Not lblnConfiguRegionalOk Then
                    SMuestreMensajeInicio(lstrMens)
                    End
                End If
                Select Case GenuTipoInstanciamiento
                    Case EnuTipoInstanciamiento.enuInstalacion
                        lstrMens = "Orión Plus se instaló exitosamente!"
                    Case EnuTipoInstanciamiento.enuActualizacion
                        lstrMens = "Orión Plus se actualizó exitosamente!"
                    Case EnuTipoInstanciamiento.enuNormal
                        lstrMens = String.Empty
                End Select
                If String.IsNullOrEmpty(lstrMens) AndAlso MblnActualizoApl Then
                    lstrMens = "Orión Plus se actualizó exitosamente!"
                End If
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SMuestreMensajeInicio(lstrMens)
                End If
            Else
                Select Case GenuTipoInstanciamiento
                    Case EnuTipoInstanciamiento.enuInstalacion
                        lstrMens &= vbCrLf & "La instalación de Orión Plus no se llevo a cabo!"
                    Case EnuTipoInstanciamiento.enuActualizacion
                        GobjPanDat.SAborteTransaccion()
                        lstrMens &= vbCrLf & "La actualización de Orión Plus no se llevo a cabo!"
                    Case EnuTipoInstanciamiento.enuNormal
                        lstrMens &= vbCrLf & "Orión Plus no se puede abrir!"
                End Select
                SMuestreMensajeInicio(lstrMens)
                If Not String.IsNullOrEmpty(lstrMensEx) Then
                    ClsPanorama.SEscribaArchivoError(lstrMensEx)
                End If
                End
            End If
        End Try
    End Sub
    Private Sub SInicieSesion()
        If GenuTipoInstanciamiento = EnuTipoInstanciamiento.enuInstalacion Then
            GobjAdministrador.SRegistreApp(EnuListaAplicaciones.EnuOrionCop, 1, GstrVersionApp)
        Else
            GobjPanorama.ObjAppActual = ClsAdministrador.FobjAppActual
            GstrVerAntApp = GobjPanorama.ObjAppActual.ObjVersionStr.ObjValorPro
        End If
        Dim lstrVerAnt = GstrVerAntApp.Replace(".", "")
        Dim lstrVerApp = GstrVersionApp.Replace(".", "")
        If lstrVerAnt.Length > 0 AndAlso lstrVerAnt < lstrVerApp Then
            ClsActualizacionApl.SActualiceVer(GstrVerAntApp, MblnActualizoApl)
            ClsAdministrador.SRegistreVersion(GstrVersionApp)
        ElseIf lstrVerAnt > lstrVerApp Then
#If DES = 0 Then
            Throw New ErrorInesperadoPanLException("Versión anterior mayor a vesión actual")
#End If
        End If
    End Sub
#End Region

#Region "Funciones y procedimientos"
    Friend Function FdblTasa(astrTasa As String) As Double
        Dim ldblTasa As Double = 0
        If astrTasa.EndsWith("%") Then
            astrTasa = astrTasa.Substring(0, astrTasa.Length - 1)
        End If
        If Not String.IsNullOrEmpty(astrTasa) Then
            If IsNumeric(astrTasa) Then
                ldblTasa = CType(astrTasa, Double) / 100
            End If
        End If
        Return ldblTasa
    End Function

    Friend Sub SPuebleBarraEstado(acolLabels As Collection)
        acolLabels(1).Content = "Carpeta: " & ClsOrionCop.StrNombreCarpetaActual
        acolLabels(1).ToolTip = acolLabels(1).Content
        acolLabels(2).Content = "Copropiedad: " & ClsOrionCop.StrNombreCentroUtilActual
        acolLabels(2).ToolTip = acolLabels(2).Content
        acolLabels(3).Content = "Usuario Actual: " & GstrIdUsuario
        acolLabels(3).ToolTip = acolLabels(3).Content
        acolLabels(4).Content = "Periodo Actual: "
        If Not IsNothing(GobjParametros.ObjAnoActual) Then
            acolLabels(4).Content &= GobjParametros.ObjAnoActual.StrNombrePeriodoActual
        End If
        acolLabels(4).ToolTip = acolLabels(4).Content
    End Sub

    Friend Function FblnYaEstaCorriendo() As Boolean
        Dim lblnEstaCorr = Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1
        Return lblnEstaCorr
    End Function

    Private Function FblnConfMonedaOk() As Boolean
        Dim lblnOk = False
        Dim lstrValorFormateado = Format(1000.1, "c")
        lstrValorFormateado = Replace(lstrValorFormateado, " ", String.Empty)
        If lstrValorFormateado.Length = 9 Then
            lblnOk = Not (lstrValorFormateado.Substring(2, 1) <> "," OrElse
                    lstrValorFormateado.Substring(6, 1) <> "." OrElse
                    (lstrValorFormateado.Substring(0, 1) <> "$" AndAlso
                    lstrValorFormateado.Substring(0, 1) <> ChrW(8353)))
        End If
        Return lblnOk
    End Function

    Private Function FblnConfNumeroOk() As Boolean
        Dim lblnOk = False
        Dim lstrValorFormateado = FormatNumber(1000.1)
        If lstrValorFormateado.Length = 8 Then
            lstrValorFormateado = Replace(lstrValorFormateado, " ", String.Empty)
            lblnOk = Not (lstrValorFormateado.Substring(1, 1) <> "," OrElse
                    lstrValorFormateado.Substring(5, 1) <> ".")
        End If
        Return lblnOk
    End Function

    Private Sub SMuestreMensajeInicio(astrMensaje As String)
        Dim lentRes = 0
        Do While lentRes <> 1
            lentRes = MsgBox(astrMensaje, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Información")
        Loop
    End Sub

    Friend Sub SImprimaNotasDb(astrIdNotasDb As String, ablnPregunte As Boolean)
        Dim lblnImprima = Not String.IsNullOrEmpty(astrIdNotasDb)
        If ablnPregunte Then
            If Not String.IsNullOrEmpty(astrIdNotasDb) Then
                lblnImprima = MsgBox("Desea imprimir las Notas de Intereses", MsgBoxStyle.YesNo,
                        "Imprimir Notas ?") = vbYes
            End If
        End If
        If lblnImprima Then
            Try
                Mouse.OverrideCursor = Cursors.Wait
                If Not String.IsNullOrEmpty(astrIdNotasDb) Then
                    Dim lobjRep As New ClsRepOrionCop(GCOBJREGISTRO)
                    Dim lstrPrefNotas = astrIdNotasDb.Split(";")(0)
                    Dim lentIdNota_1 = CInt(astrIdNotasDb.Split(";")(1))
                    Dim lentIdNota_N = CInt(astrIdNotasDb.Split(";")(2))
                    Dim lobjParaNotaDb As New ClsParametrosReportesDocs(lstrPrefNotas, lentIdNota_1, lentIdNota_N)
                    lobjRep.ObjParRepDocs = lobjParaNotaDb
                    lobjRep.EnuReporte = EnuReporteDef.enuNotasDb
                    If ablnPregunte Then
                        lobjRep.SGenereReporteDialog()
                    Else
                        lobjRep.SGenereReporte()
                    End If
                End If
            Catch ex As PanLException
                Throw
            Catch ex As PanDatException
                Throw
            Catch ex As ArgumentNullException
                Throw
            Catch ex As Exception
                Throw
            Finally
                Mouse.OverrideCursor = Cursors.Arrow
            End Try
        End If
    End Sub

#End Region

#Region "EFac"
    Friend Sub SMuestreEstadoEFac(astrCudoc As String,
            aenuTipoDoc As EnuTipoDocOri, astrNroDoc As String,
            ablnV1 As Boolean, aenuEstadoEDocEnApp As EnuEstadoEDoc, ByRef astrMens As String)
        If GobjParametros.ObjIdProvEFacByt.ObjValorPro > 0 Then
            If FblnEstaConectado(GobjParametros.ObjURLStr.ObjValorPro, astrMens) Then
                If ablnV1 Then
                    SMuestreEstadoEFac_V1(astrCudoc, aenuTipoDoc, aenuEstadoEDocEnApp)
                Else
                    SMuestreEstadoEFac_V2(astrCudoc, aenuTipoDoc, astrNroDoc,
                            aenuEstadoEDocEnApp)
                End If
            Else
                If String.IsNullOrEmpty(astrMens) Then
                    astrMens = "Parece no estar conectado a Internet!"
                End If
            End If
        End If
    End Sub

    Private Async Sub SMuestreEstadoEFac_V1(astrCudoc As String,
            aenuTipoDoc As EnuTipoDocOri, aenuEstadoEDoc As EnuEstadoEDoc)
        Dim lstrEstado = "V1"
        Using lIntMisFact As New ClsInterfazMisFacturas(GCOBJREGISTRO)
            Select Case aenuTipoDoc
                Case EnuTipoDocOri.EnuFactura
                    lstrEstado = Await lIntMisFact.FstrObtengaEstado_V1(astrCudoc, aenuTipoDoc)
                Case EnuTipoDocOri.EnuNotaCr
                    lstrEstado = Await lIntMisFact.FstrObtengaEstado_V1(astrCudoc, aenuTipoDoc)
                Case EnuTipoDocOri.EnuNotaDb
                    lstrEstado = Await lIntMisFact.FstrObtengaEstado_V1(astrCudoc, aenuTipoDoc)
                Case EnuTipoDocOri.EnuReciboCaja
                    lstrEstado = Await lIntMisFact.FstrObtengaEstado_V1(astrCudoc, aenuTipoDoc)
                Case EnuTipoDocOri.EnuNotaRevCr
                    lstrEstado = Await lIntMisFact.FstrObtengaEstado_V1(astrCudoc, aenuTipoDoc)
            End Select
        End Using
        If String.IsNullOrEmpty(lstrEstado) Then
            lstrEstado = "Con Problema"
        End If
        Dim lwinEstado As New WinEstadoDocEFac(lstrEstado) With {
                            .EnuTipoDoc = aenuTipoDoc,
                            .StrIdDoc = astrCudoc,
                            .EnuEstadoEDocEnApp = aenuEstadoEDoc
                        }
        lwinEstado.ShowDialog()
    End Sub

    Private Async Sub SMuestreEstadoEFac_V2(astrCudoc As String,
            aenuTipoDoc As EnuTipoDocOri, astrNroDoc As String,
                    aenuEstadoEDoc As EnuEstadoEDoc)
        Dim lobjEstadoDoc As ClsEstadoDoc = Nothing
        Using lIntMisFact As New ClsInterfazMisFacturas(GCOBJREGISTRO)
            lobjEstadoDoc = Await lIntMisFact.FobjObtengaEstado(astrCudoc,
                            aenuTipoDoc, astrNroDoc)
        End Using
        Dim lwinEstado As New WinEstadoDocEFac() With {
                    .ObjEstadoDoc = lobjEstadoDoc,
                    .EnuTipoDoc = aenuTipoDoc,
                    .StrIdDoc = astrCudoc,
                    .EnuEstadoEDocEnApp = aenuEstadoEDoc
                }
        lwinEstado.ShowDialog()
    End Sub

    Friend Function FstrNombreEstadoEFacEnApp(aenuEstadoEDoc As EnuEstadoEDoc) As String
        Dim lstrNomEst As String
        Select Case aenuEstadoEDoc
            Case EnuEstadoEDoc.EnuNoReg
                lstrNomEst = "No Registrado"
            Case EnuEstadoEDoc.EnuInvalida
                lstrNomEst = "Invalido (Rechazado por la DIAN)"
            Case EnuEstadoEDoc.EnuEnProceso
                lstrNomEst = "En Proceso"
            Case EnuEstadoEDoc.EnuRegi
                lstrNomEst = "Esperando RG (Rep. Gráfica)"
            Case EnuEstadoEDoc.EnuEnviada
                lstrNomEst = "Enviado"
            Case EnuEstadoEDoc.EnuAceptada
                lstrNomEst = "Aceptación Expresa"
            Case EnuEstadoEDoc.EnuRechazada
                lstrNomEst = "Rechazada por el Cliente"
            Case EnuEstadoEDoc.EnuOtro
                lstrNomEst = "Otro"
            Case Else
                lstrNomEst = String.Empty
        End Select
        Return lstrNomEst
    End Function

    Friend Sub SInicieProcEFac(awinMW As MWOrionCop)
#If PRFEL = 0 Then
        If GobjParametros.BlnEFacAutorizado Then
            If MwinProcesoEFac Is Nothing Then
                MwinProcesoEFac = New WinProcesaEFac(awinMW)
                MwinProcesoEFac.Show()
                MwinProcesoEFac.Visibility = Visibility.Hidden
            End If
        End If
#End If
    End Sub

    Friend Sub SProceseEFac(ByRef astrMens As String)
#If PRFEL = 0 Then
        If FblnEstaConectado(GobjParametros.ObjURLStr.ObjValorPro, astrMens) Then
            If GobjParametros.EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuDocPorProEFac OrElse
                    ClsOrionCop.FblnDocPorProcesarEFac OrElse Not GblnEstadoRechazado Then
                MwinProcesoEFac?.SProceseEDocs()
            End If
        End If
#End If
    End Sub

    Friend Sub SCambieVisibilidad()
        If MwinProcesoEFac IsNot Nothing Then
            Dim lblnOcultar = MwinProcesoEFac.Visibility = Visibility.Visible
            If lblnOcultar Then
                MwinProcesoEFac.Visibility = Visibility.Hidden
            Else
                MwinProcesoEFac.Visibility = Visibility.Visible
                MwinProcesoEFac.Activate()
            End If
        End If
    End Sub

    Friend Sub SOculteVenProcesoEFac()
        If MwinProcesoEFac IsNot Nothing Then
            If MwinProcesoEFac.Visibility = Visibility.Visible Then
                MwinProcesoEFac.Visibility = Visibility.Hidden
            End If
        End If
    End Sub
#End Region

#Region "Procedimientos"
    Friend Sub SLeaArchivoIni()
        Dim lsrArchivoIni As StreamReader
        Dim lstrLinea As String
        Dim lstrArg() As String
        Dim lstrArchivoIni As String = GstrTrayDatPrg & "OrionCop.ini"
        lsrArchivoIni = ClsPanorama.FsrStreamReader(lstrArchivoIni)
        If IsNothing(lsrArchivoIni) Then Exit Sub
        lstrLinea = lsrArchivoIni.ReadLine
        SLeaArchivoAdminIni(False)
        Do While Not IsNothing(lstrLinea)
            If lstrLinea.Contains("=") Then
                lstrArg = lstrLinea.Split("=")
                If Not FblnLeyo(lstrArg(0), lstrArg(1)) Then
                    Select Case lstrArg(0)
                        Case "Carpeta"
                            SCarpeta(lstrArg(1))
                        Case "CentroUtil"
                            SCentroUtil(lstrArg(1))
                        Case "Limite1"
                            GentLimite1 = CType(lstrArg(1), Integer)
                        Case "Limite2"
                            GentLimite2 = CType(lstrArg(1), Integer)
                        Case "Limite3"
                            GentLimite3 = CType(lstrArg(1), Integer)
                        Case "Limite4"
                            GentLimite4 = CType(lstrArg(1), Integer)
                        Case "MostrarAyuda"
                            EnuEstadoAyuda = CType(lstrArg(1), Integer)
                    End Select
                End If
            End If
            lstrLinea = lsrArchivoIni.ReadLine
        Loop
        lsrArchivoIni.Close()
    End Sub

    Friend Sub SEscribaArchivoIni()
        Dim lstrLinea As String = Nothing
        Dim lstrArchivoIni As String = GstrTrayDatPrg & "OrionCop.ini"
        Dim lswArchivoIni = ClsPanorama.FswStreamWriter(lstrArchivoIni)
        Select Case GenuTamanoIcono
            Case EnuTamanoIconos.EnuGrande
                lstrLinea = "TamanoIcono=L"
            Case EnuTamanoIconos.EnuMediano
                lstrLinea = "TamanoIcono=M"
            Case EnuTamanoIconos.EnuPequeño
                lstrLinea = "TamanoIcono=S"
        End Select
        lswArchivoIni.WriteLine(lstrLinea)
        If GblnMostrandoTitulos Then
            lstrLinea = "TitulosMenu=S"
        Else
            lstrLinea = "TitulosMenu=N"
        End If
        lswArchivoIni.WriteLine(lstrLinea)
        If GobjPanorama.ObjUsuarioActual.BlnExiste OrElse GstrIdUsuario = GCSTRUSUARIOU Then
            lstrLinea = "Usuario=" & GstrIdUsuario
        Else
            lstrLinea = "Usuario="
        End If
        lswArchivoIni.WriteLine(lstrLinea)
        lstrLinea = "Carpeta=" & GshrIdCarpeta.ToString
        lswArchivoIni.WriteLine(lstrLinea)
        lstrLinea = "CentroUtil=" & GshrIdCentroUtil.ToString
        lswArchivoIni.WriteLine(lstrLinea)
        lstrLinea = "Limite1=" & GentLimite1.ToString
        lswArchivoIni.WriteLine(lstrLinea)
        lstrLinea = "Limite2=" & GentLimite2.ToString
        lswArchivoIni.WriteLine(lstrLinea)
        lstrLinea = "Limite3=" & GentLimite3.ToString
        lswArchivoIni.WriteLine(lstrLinea)
        lstrLinea = "Limite4=" & GentLimite4.ToString
        lswArchivoIni.WriteLine(lstrLinea)
        lstrLinea = "MostrarAyuda=" & EnuEstadoAyuda
        lswArchivoIni.WriteLine(lstrLinea)
        lswArchivoIni.Close()
    End Sub

    Private Function FblnLeyo(astrArg0 As String, astrArg1 As String) As Boolean
        Dim lblnLeyo = True
        Select Case astrArg0
            Case "TamanoIcono"
                STamanoIcono(astrArg1)
            Case "TitulosMenu"
                STitulos(astrArg1)
            Case "Usuario"
                SUsuario(astrArg1)
            Case Else
                lblnLeyo = False
        End Select
        Return lblnLeyo
    End Function

    Private Sub STamanoIcono(astrTamaño As String)
        Select Case astrTamaño
            Case "L"
                GenuTamanoIcono = EnuTamanoIconos.EnuGrande
            Case "M"
                GenuTamanoIcono = EnuTamanoIconos.EnuMediano
            Case "S"
                GenuTamanoIcono = EnuTamanoIconos.EnuPequeño
        End Select
    End Sub

    Private Sub STitulos(astrArg As String)
        If astrArg = "S" Then
            GblnMostrandoTitulos = True
        Else
            GblnMostrandoTitulos = False
        End If
    End Sub

    Private Sub SUsuario(astrArg As String)
        If Not IsNothing(astrArg) Then
            GstrIdUsuario = astrArg
        Else
            GstrIdUsuario = String.Empty
        End If
    End Sub

    Private Sub SCarpeta(astrArg As String)
        If Not IsNothing(astrArg) Then
            GshrIdCarpeta = CType(astrArg, Short)
        Else
            GshrIdCarpeta = 0
        End If
    End Sub

    Private Sub SCentroUtil(astrArg As String)
        If Not IsNothing(astrArg) Then
            GshrIdCentroUtil = CType(astrArg, Short)
        Else
            GshrIdCentroUtil = 0
        End If
    End Sub

    ''' <summary>
    ''' Establece el icono de la aplicación en la ventana de desinstalar o cambiar un programa de windows
    ''' </summary>
    Private Sub SSetAddRemoveProgramsIcon()
        Try
            Dim lstrTrayIcono As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "OrionPlus.ico")
            If Not File.Exists(lstrTrayIcono) Then
                Return
            End If
            Dim lrkMyUninstallKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Uninstall")
            Dim lstrMySubKeyNames As String() = lrkMyUninstallKey.GetSubKeyNames()
            For i = 0 To lstrMySubKeyNames.Length - 1
                Dim lrkMyKey As RegistryKey = lrkMyUninstallKey.OpenSubKey(lstrMySubKeyNames(i), True)
                Dim lobjMyValue As Object = lrkMyKey.GetValue("DisplayName")
                If Not IsNothing(lobjMyValue) AndAlso (lobjMyValue.ToString() = "Orión Plus x32" OrElse
                        lobjMyValue.ToString() = "Orión Plus x64") Then
                    lrkMyKey.SetValue("DisplayIcon", lstrTrayIcono)
                    Exit For
                End If
            Next
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Imprime las facturas al terminar el proceso de facturación automática
    ''' </summary>
    Friend Sub SImprimaFactAut(ablnContingencia As Boolean)
        Dim lobjRep As New ClsRepOrionCop(GCOBJREGISTRO)
        Dim lstrMens As String, lstrMensTitulo As String
        If GobjParametros.BlnEFacAutorizado Then
            lstrMens = "Desea imprimir las Facturas generadas?"
            lstrMensTitulo = "Imprimir Facturas Generadas?"
        Else
            lstrMens = "Desea imprimir las Cuentas de Cobro generadas?"
            lstrMensTitulo = "Imprimir Cuentas de Cobro Generadas?"
        End If
        If MsgBox(lstrMens, vbYesNo, lstrMensTitulo) = vbYes Then
            SImprimaFacturas(ablnContingencia)
        End If
        If Not ablnContingencia Then
            Dim lstrUltimasCtaCob = ClsOrionCop.FstrIdUltimasCtasCobro
            If Not String.IsNullOrEmpty(lstrUltimasCtaCob) Then
                lstrMens = "Desea imprimir las Cuentas de Cobro Adicionales generadas?"
                If MsgBox(lstrMens, vbYesNo, "Imprimir Cuentas de Cobro Adicionales") = vbYes Then
                    SImprimaCtasCobro(lobjRep)
                End If
            End If
            SImprimaNotasCon(lobjRep)
        End If
    End Sub

    Private Sub SImprimaFacturas(ablnContingencia As Boolean)
        Mouse.OverrideCursor = Cursors.Wait
        Try
            Dim lobjRep As New ClsRepOrionCop(GCOBJREGISTRO)
            Dim lstrIdUltimFac = ClsOrionCop.FstrIdUltimasFras(ablnContingencia)
            If Not String.IsNullOrEmpty(lstrIdUltimFac) Then
                Dim lstrFacts = lstrIdUltimFac.Split(";")
                Dim lstrPrefFact = lstrFacts(0)
                Dim lentIdFacPrimera = CType(lstrFacts(1), Integer)
                Dim lentIdFacUltima = CType(lstrFacts(2), Integer)
                Dim lobjParaFact As New ClsParametrosReportesDocs(lstrPrefFact,
                    lentIdFacPrimera, lentIdFacUltima)
                If GobjParametros.ObjAnoActual.ObjTipoIncentivoByt.ObjValorPro =
                        EnuTipoIncentivo.EnuDescuentoPP Then
                    lobjRep.EnuReporte = EnuReporteDef.enuFacturaDscto
                ElseIf GobjParametros.BlnEFacAutorizado Then
                    lobjRep.EnuReporte = EnuReporteDef.enuFacturaEFac
                Else
                    lobjRep.EnuReporte = EnuReporteDef.enuFactura
                End If
                lobjParaFact.BlnExcluirFacEnvEmail = False
                lobjRep.ObjParRepDocs = lobjParaFact
                lobjRep.SGenereReporteDialog()
            End If
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub

    Friend Sub SImprimaFactura(aobjFactura As ClsFactura, ByRef astrMens As String)
        Dim lblnPuede = aobjFactura.BlnEstaRegEFac OrElse Not aobjFactura.BlnEsFacEle
        If lblnPuede Then
            Dim lstrPrefFact = aobjFactura.ObjPrefijo_FactStr.ObjValorPro
            Dim lentIdFacPrimera = aobjFactura.ObjIdFacturaEnt.ObjValorPro
            Dim lentIdFacUltima = aobjFactura.ObjIdFacturaEnt.ObjValorPro
            Dim lobjParaFact As New ClsParametrosReportesDocs(lstrPrefFact, lentIdFacPrimera,
                    lentIdFacUltima)
            Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
                    .ObjParRepDocs = lobjParaFact
                }
            If aobjFactura.BlnEsFacEle Then
                lobjRep.EnuReporte = EnuReporteDef.enuFacturaEFac
            Else
                If aobjFactura.ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuImportada Then
                    lobjRep.EnuReporte = EnuReporteDef.enuFactImportada
                Else
                    If GobjParametros.ObjAnoActual.ObjTipoIncentivoByt.ObjValorPro =
                            EnuTipoIncentivo.EnuDescuentoPP Then
                        lobjRep.EnuReporte = EnuReporteDef.enuFacturaDscto
                    Else
                        lobjRep.EnuReporte = EnuReporteDef.enuFactura
                    End If
                End If
            End If
            lobjRep.SGenereReporte()
        Else
            astrMens = "La Factura no se puede imprimir porque aún no está registrada en API de EFac!"
        End If
    End Sub

    Friend Sub SImprimaCtasCobro(aobjRepo As ClsRepOrionCop)
        Mouse.OverrideCursor = Cursors.Wait
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError As Boolean
        Try
            Dim lstrUltimasCtaCob = ClsOrionCop.FstrIdUltimasCtasCobro
            Dim lstrPref = String.Empty
            If Not String.IsNullOrEmpty(lstrUltimasCtaCob) Then
                Dim lentIdCtaCob_1 As Integer = CType(lstrUltimasCtaCob.Split(";")(0), Integer)
                Dim lentIdCtaCob_N As Integer = CType(lstrUltimasCtaCob.Split(";")(1), Integer)
                Dim lobjParaNotaCon As New ClsParametrosReportesDocs("", lentIdCtaCob_1,
                        lentIdCtaCob_N)
                aobjRepo.ObjParRepDocs = lobjParaNotaCon
                aobjRepo.EnuReporte = EnuReporteDef.enuCtaCobroDet
                aobjRepo.SGenereReporteDialog()
            End If
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString()
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString()
        Catch ex As ArgumentNullException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString()
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString()
        Finally
            If Not lblnNoHayError Then
                SMuestreMensajeInicio(lstrMens)
                If Not String.IsNullOrEmpty(lstrMensEx) Then
                    ClsPanorama.SEscribaArchivoError(lstrMensEx)
                End If
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub

    Friend Sub SImprimaNotasCon(aobjRepo As ClsRepOrionCop)
        Dim lstrIdNotasCon = ClsOrionCop.FstrIdUltimasNotasAplAnt()
        Mouse.OverrideCursor = Cursors.Wait
        If Not String.IsNullOrEmpty(lstrIdNotasCon) Then
            Dim lstrMens = "Imprimir las Notas de Contabilidad correspondientes a los" &
                " Anticipos Aplicados?"
            If MsgBox(lstrMens, vbYesNo, "Imprimir Notas de Contabilidad") = vbYes Then
                Dim lstrPrefNotas = lstrIdNotasCon.Split(";")(0)
                Dim lentIdNota_1 = lstrIdNotasCon.Split(";")(1)
                Dim lentIdNota_N = lstrIdNotasCon.Split(";")(2)
                If lentIdNota_1 > 0 Then
                    Dim lobjParaNotaCon As New ClsParametrosReportesDocs(lstrPrefNotas,
                        lentIdNota_1, lentIdNota_N)
                    aobjRepo.ObjParRepDocs = lobjParaNotaCon
                    aobjRepo.EnuReporte = EnuReporteDef.enuNotasCon
                    aobjRepo.SGenereReporteDialog()
                End If
            End If
        End If
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
#End Region

#Region "Correo electrónico"
    Friend Sub SEscribaReporteEmails(ablnInicio As Boolean, astrNumeroDoc As String,
                                     astrMensaje As String)
        Dim lstrArchivo = GstrTrayEmails & "\DocumentosNoEnviados.txt"
        Dim lstrMens = String.Empty
        If Not String.IsNullOrEmpty(astrNumeroDoc) Then
            lstrMens = "El Documento número " & astrNumeroDoc & " no pudo ser publicado. " &
                    astrMensaje
        End If
        Try
            If ablnInicio OrElse Not My.Computer.FileSystem.FileExists(lstrArchivo) Then
                Using lswRepEmail = File.CreateText(lstrArchivo)
                    lswRepEmail.WriteLine("DOCUMENTOS NO ENVIADOS " & Date.Now.ToString)
                    lswRepEmail.WriteLine(lstrMens)
                    lswRepEmail.WriteLine(String.Empty)
                    lswRepEmail.Flush()
                End Using
            Else
                Using lswRepEmail = File.AppendText(lstrArchivo)
                    lswRepEmail.WriteLine(lstrMens)
                    lswRepEmail.WriteLine(String.Empty)
                    lswRepEmail.Flush()
                End Using
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Friend Sub SMuestreCorreosNoEnviados()
        Dim lstrArchivo = GstrTrayEmails & "\DocumentosNoEnviados.txt"
        If My.Computer.FileSystem.FileExists(lstrArchivo) Then
            Process.Start("notepad.exe", lstrArchivo)
        End If
    End Sub

    Friend Sub SEnvieCorreo(aenuTipoCorreo As EnuTipoCorreoE,
            adblIdCliente As Double, astrPredioAgr As String, astrNroDoc As String,
            ByRef astrMens As String)
        If Not GblnCorreoOn Then
            If FblnPuedeEnviarEmail(astrMens) Then
                Dim lobjCorreoE As New ClsCorreoE, lstrMens = String.Empty
                If aenuTipoCorreo >= EnuTipoCorreoE.EnuFactAuto Then
                    With lobjCorreoE
                        .EnuTipoCorreo = aenuTipoCorreo
                        .DblIdCliente = adblIdCliente
                        .FblnEsValidoIdPreAgr(astrPredioAgr)
                        .FblnEsValidoNroDoc(astrNroDoc, lstrMens)
                    End With
                End If
                MwinCorreo = New WinCorreoE With {
                    .ObjCorreoE = lobjCorreoE
                }
                MwinCorreo.Show()
                GblnCorreoOn = True
            End If
        Else
            astrMens = "La ventana de correo ya está abierta!"
        End If
    End Sub

    Private Function FblnPuedeEnviarEmail(ByRef astrMens As String) As Boolean
        Dim lblnPuede As Boolean
        lblnPuede = ClsPanorama.FblnEmailsHabilitado
        If Not lblnPuede Then
            astrMens = "No tiene habilitado el Módulo de Correo Electrónico!"
        End If
        If lblnPuede Then
            GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.FblnExisteCtaCorreoValida()
            If Not lblnPuede Then
                astrMens = "La Cuenta de Correo no esta debidamente parametrizada!"
            End If
        End If
        If lblnPuede Then
            lblnPuede = Not GblnPosteando
            If Not lblnPuede Then
                astrMens = "No es posible enviar Emails mientras haya un proceso " &
                            "de facturación electrónica activo!"
            End If
        End If
        Return lblnPuede
    End Function
#End Region

#Region "Monto escrito"
    Friend Function FstrMontoEscrito(adecValor As Decimal) As String
        Dim lstrValor As String
        Dim lstrEnteros As String
        Dim lstrDeci As String
        Dim lstrV1 As String
        Dim lstrVStr As String
        ' normalizo que el cero (0) sea igual a 0.00
        lstrValor = Format(adecValor, "#########0.00")
        ' descompongo la cadena valor en dos cadenas. una que contiene los enteros y la otra los decimales
        lstrEnteros = Left(lstrValor, InStr(lstrValor, ".") - 1)
        lstrDeci = Right(lstrValor, Len(lstrValor) - InStr(lstrValor, "."))
        If lstrEnteros = "0" And lstrDeci = "00" Then
            lstrV1 = String.Empty
        Else
            lstrV1 = String.Empty
            If lstrEnteros = "0" Then
                lstrEnteros = String.Empty
            End If
            While lstrEnteros.Length > 0
                Select Case Len(lstrEnteros)
                    Case 10
                        If lstrEnteros = "1000000000" Then
                            lstrV1 = "UN MIL MILLONES DE PESOS"
                            lstrEnteros = String.Empty
                        Else
                            lstrVStr = FstrGrupo1(Left(lstrEnteros, 1))
                            lstrV1 = lstrV1 & lstrVStr & "MIL "
                            lstrEnteros = Right(lstrEnteros, 9)
                        End If
                    Case 7 To 9
                        lstrEnteros = Left(Right("  " & lstrEnteros, 9), 9)
                        lstrVStr = FstrGrupo3(Left(lstrEnteros, 3))
                        If lstrEnteros = "  1000000" Or lstrEnteros = "001000000" Then
                            If lstrV1.Length > 0 Then
                                lstrV1 &= "UN MILLONES DE PESOS "
                            Else
                                lstrV1 = "UN MILLON DE PESOS "
                            End If
                            lstrEnteros = String.Empty
                        ElseIf lstrEnteros = " 10000000" Or lstrEnteros = "010000000" Then
                            lstrV1 &= "DIEZ MILLONES DE PESOS "
                            lstrEnteros = String.Empty
                        ElseIf lstrEnteros = "100000000" Or lstrEnteros = "100000000" Then
                            lstrV1 &= "CIEN MILLONES DE PESOS "
                            lstrEnteros = String.Empty
                        ElseIf lstrVStr = "UN " Then
                            If lstrV1.Length > 0 Then
                                lstrV1 &= "UN MILLONES "
                            Else
                                lstrV1 = "UN MILLON "
                            End If
                        Else
                            lstrV1 = lstrV1 & lstrVStr & "MILLONES "
                        End If
                        lstrEnteros = Right(lstrEnteros, 6)
                    Case 4 To 6
                        lstrVStr = FstrGrupo3(Left(Right("  " & lstrEnteros, 6), 3))
                        If lstrVStr.Length > 0 Then
                            lstrV1 = lstrV1 & lstrVStr & "MIL "
                        End If
                        lstrEnteros = Right(lstrEnteros, 3)
                    Case 1 To 3
                        If Right("  " & lstrEnteros, 3) = "  1" Then
                            lstrV1 = "UN PESO "
                        Else
                            lstrVStr = FstrGrupo3(Right("  " & lstrEnteros, 3))
                            lstrV1 = lstrV1 & lstrVStr & "PESOS "
                        End If
                        lstrEnteros = String.Empty
                    Case Else
                        lstrV1 = String.Empty
                End Select
            End While
            If CInt(lstrDeci) > 0 Then
                If lstrV1.Length > 0 Then
                    lstrV1 = lstrV1 & "CON " & FstrGrupo2(Right(lstrDeci, 2)) & "CENTAVOS "
                Else
                    lstrV1 = FstrGrupo2(Right(lstrDeci, 2)) & "CENTAVOS "
                End If
            End If
            lstrV1 &= "MDA/CTE"
        End If
        Return lstrV1
    End Function

    Private Function FstrGrupo1(astrVlr As String) As String
        If astrVlr <> " " And astrVlr <> "0" Then
            Select Case CInt(astrVlr)
                Case 1
                    Return "UN "
                Case 2
                    Return "DOS "
                Case 3
                    Return "TRES "
                Case 4
                    Return "CUATRO "
                Case 5
                    Return "CINCO "
                Case 6
                    Return "SEIS "
                Case 7
                    Return "SIETE "
                Case 8
                    Return "OCHO "
                Case 9
                    Return "NUEVE "
                Case Else
                    Return ""
            End Select
        Else
            Return ""
        End If
    End Function

    Private Function FstrGrupo2(astrVlr As String) As String
        Dim lstrGrupo2 = String.Empty
        If astrVlr <> "  " And astrVlr <> "00" Then
            lstrGrupo2 = FstrGrupo2_1(astrVlr)
            If String.IsNullOrEmpty(lstrGrupo2) Then
                lstrGrupo2 = FstrGrupo2_2(astrVlr)
                If String.IsNullOrEmpty(lstrGrupo2) Then
                    Select Case CInt(astrVlr)
                        Case 51 To 59
                            Return "CINCUENTA Y " & FstrGrupo1(Right(astrVlr, 1))
                        Case 60
                            Return "SESENTA "
                        Case 61 To 69
                            Return "SESENTA Y " & FstrGrupo1(Right(astrVlr, 1))
                        Case 70
                            Return "SETENTA "
                        Case 71 To 79
                            Return "SETENTA Y " & FstrGrupo1(Right(astrVlr, 1))
                        Case 80
                            Return "OCHENTA "
                        Case 81 To 89
                            Return "OCHENTA Y " & FstrGrupo1(Right(astrVlr, 1))
                        Case 90
                            Return "NOVENTA "
                        Case 91 To 99
                            Return "NOVENTA Y " & FstrGrupo1(Right(astrVlr, 1))
                    End Select
                End If
            End If
        End If
        Return lstrGrupo2
    End Function

    Private Function FstrGrupo2_1(astrVlr As String) As String
        Dim lstrGrupo = String.Empty
        Select Case CInt(astrVlr)
            Case 1 To 9
                lstrGrupo = FstrGrupo1(Right(astrVlr, 1))
            Case 10
                lstrGrupo = "DIEZ "
            Case 11
                lstrGrupo = "ONCE "
            Case 12
                lstrGrupo = "DOCE "
            Case 13
                lstrGrupo = "TRECE "
            Case 14
                lstrGrupo = "CATORCE "
            Case 15
                lstrGrupo = "QUINCE "
            Case 16 To 19
                lstrGrupo = "DIEZ Y " & FstrGrupo1(Right(astrVlr, 1))
            Case 20
                lstrGrupo = "VEINTE "
        End Select
        Return lstrGrupo
    End Function

    Private Function FstrGrupo2_2(astrVlr As String) As String
        Dim lstrGrupo = String.Empty
        Select Case CInt(astrVlr)
            Case 21 To 29
                lstrGrupo = "VEINTI" & FstrGrupo1(Right(astrVlr, 1))
            Case 30
                lstrGrupo = "TREINTA "
            Case 31 To 39
                lstrGrupo = "TREINTA Y " & FstrGrupo1(Right(astrVlr, 1))
            Case 40
                lstrGrupo = "CUARENTA "
            Case 41 To 49
                lstrGrupo = "CUARENTA Y " & FstrGrupo1(Right(astrVlr, 1))
            Case 50
                lstrGrupo = "CINCUENTA "
        End Select
        Return lstrGrupo
    End Function

    Private Function FstrGrupo3(astrVlr As String) As String
        If astrVlr <> "   " And astrVlr <> "000" Then
            Select Case CInt(astrVlr)
                Case 0 To 99
                    Return FstrGrupo2(Right(astrVlr, 2))
                Case 100
                    Return "CIEN "
                Case 101 To 199
                    Return "CIENTO " & FstrGrupo2(Right(astrVlr, 2))
                Case 200 To 299
                    Return "DOSCIENTOS " & FstrGrupo2(Right(astrVlr, 2))
                Case 300 To 399
                    Return "TRECIENTOS " & FstrGrupo2(Right(astrVlr, 2))
                Case 400 To 499
                    Return "CUATROCIENTOS " & FstrGrupo2(Right(astrVlr, 2))
                Case 500 To 599
                    Return "QUINIENTOS " & FstrGrupo2(Right(astrVlr, 2))
                Case 600 To 699
                    Return "SEISCIENTOS " & FstrGrupo2(Right(astrVlr, 2))
                Case 700 To 799
                    Return "SETECIENTOS " & FstrGrupo2(Right(astrVlr, 2))
                Case 800 To 899
                    Return "OCHOCIENTOS " & FstrGrupo2(Right(astrVlr, 2))
                Case 900 To 999
                    Return "NOVECIENTOS " & FstrGrupo2(Right(astrVlr, 2))
                Case Else
                    Return ""
            End Select
        Else
            Return ""
        End If
    End Function
#End Region

#Region "Abre Ventanas"
    Friend Sub SAbraWinPredio(lstrIdPredio As String)
        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
        lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrIdPredio})
        Dim lwinVentana As New WinPredios() With {
            .BlnVentanaAux = True,
            .ObjObjetoWin = lobjPredio
        }
        lwinVentana.ShowDialog()
    End Sub

    Friend Sub SAbraWinCliente(lstrIdCliente As String)
        Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
        lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrIdCliente})
        Dim lwinVentana As New WinClientes() With {
            .BlnVentanaAux = True,
            .ObjObjetoWin = lobjCliente
        }
        lwinVentana.ShowDialog()
    End Sub

    Friend Sub SAbraFactura(lstrPrefDoc As String, aentIdDoc As Integer)
        Dim lobjDoc As New ClsFactura()
        lobjDoc.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDoc, aentIdDoc})
        Dim lwinVentana As New WinFacturas() With {
            .BlnVentanaAux = True,
            .ObjObjetoWin = lobjDoc
        }
        lwinVentana.ShowDialog()
    End Sub

    Friend Sub SAbraNotaAju(lstrPrefDoc As String, aentIdDoc As Integer)
        Dim lobjDoc As New ClsNotaAjusteCuotaAdmin()
        lobjDoc.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDoc, aentIdDoc})
        Dim lwinVentana As New WinNotasAjuste() With {
            .BlnVentanaAux = True,
            .ObjObjetoWin = lobjDoc
        }
        lwinVentana.ShowDialog()
    End Sub

    Friend Sub SAbraNotaCon(lstrPrefDoc As String, aentIdDoc As Integer)
        Dim lobjDoc As New ClsNotaCon()
        lobjDoc.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDoc, aentIdDoc})
        Dim lwinVentana As New WinNotasAplicaAnt With {
            .BlnVentanaAux = True,
            .ObjObjetoWin = lobjDoc
        }
        lwinVentana.ShowDialog()
    End Sub

    Friend Sub SAbraNotaCr(lstrPrefDoc As String, aentIdDoc As Integer)
        Dim lobjDoc As New ClsNotaCr()
        lobjDoc.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDoc, aentIdDoc})
        Dim lwinVentana As New WinNotasCr With {
            .BlnVentanaAux = True,
            .ObjObjetoWin = lobjDoc
        }
        lwinVentana.ShowDialog()
    End Sub

    Friend Sub SAbraNotaDb(lstrPrefDoc As String, aentIdDoc As Integer)
        Dim lobjDoc As New ClsNotaDb()
        lobjDoc.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDoc, aentIdDoc})
        Dim lwinVentana As New WinNotasIntMora With {
            .BlnVentanaAux = True,
            .ObjObjetoWin = lobjDoc
        }
        lwinVentana.ShowDialog()
    End Sub

    Friend Sub SAbraNotaRCr(lstrPrefDoc As String, aentIdDoc As Integer)
        Dim lobjDoc As New ClsNotaReversionCr()
        lobjDoc.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDoc, aentIdDoc})
        Dim lwinVentana As New WinNotasReversionCr With {
            .BlnVentanaAux = True,
            .ObjObjetoWin = lobjDoc
        }
        lwinVentana.ShowDialog()
    End Sub

    Friend Sub SAbraRecibo(lstrPrefDoc As String, aentIdDoc As Integer)
        Dim lobjDoc As New ClsReciboCaja()
        lobjDoc.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDoc, aentIdDoc})
        Dim lwinVentana As New WinRecibosCaja With {
            .BlnVentanaAux = True,
            .ObjObjetoWin = lobjDoc
        }
        lwinVentana.ShowDialog()
    End Sub

    Friend Sub SAbraNotaDevAnt(lstrPrefDoc As String, aentIdDoc As Integer)
        Dim lobjDoc As New ClsNotaDevAnt()
        lobjDoc.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDoc, aentIdDoc})
        Dim lwinVentana As New WinNotasDevAnt With {
            .BlnVentanaAux = True,
            .ObjObjetoWin = lobjDoc
        }
        lwinVentana.ShowDialog()
    End Sub

    Friend Sub SAbraAnticipo(aentIdAnt As Integer)
        Dim lobjAnt As New ClsAnticipo(EnuModoInstanciaObjDef.EnuUnico)
        Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, aentIdAnt}
        lobjAnt.SAbra(lobjValorLlave)
        Dim lwinAnt As New WinAnticipos With {
            .BlnVentanaAux = True,
            .ObjObjetoWin = lobjAnt
        }
        Dim unused = lwinAnt.ShowDialog()
    End Sub

    Friend Sub SGenereBK()
        Dim lwinCopiaSeg As New WinCopiaSeg With {
            .BlnAutom = True
        }
        Dim unused = lwinCopiaSeg.ShowDialog()
    End Sub
#End Region

#Region "Mensajes ayuda Ventana principal"
    Friend Function FstrMensajeAyuda(ByRef astrMens As String) As String
        Dim lstrMens = String.Empty, lblnInstalacionOk = False
        Select Case GobjParametros.EnuEstadoInstalacion
            Case EnuEstadoInstalacion.Todos
                lblnInstalacionOk = True
            Case EnuEstadoInstalacion.None
                lstrMens = FstrProcParaCuentasCont(astrMens)
            Case EnuEstadoInstalacion.CuentasCont
                lstrMens = FstrProcParaParametros(astrMens)
            Case 3
                lstrMens = FstrProcCtasBancos(astrMens)
            Case 7
                lstrMens = FstrProcSectores(astrMens)
            Case 15
                lstrMens = FstrProcModulos(astrMens)
            Case 31
                lstrMens = FstrProcSectoresModulo(astrMens)
            Case 63
                lstrMens = FstrProcAnoInicial(astrMens)
            Case 127
                lstrMens = FstrProcServicioAFacturar(astrMens)
            Case 255
                lstrMens = FstrProcTasaMora(astrMens)
            Case 511
                lstrMens = FstrProcTerceros(astrMens)
            Case 1023
                lstrMens = FstrProcClientes(astrMens)
            Case 2047
                lstrMens = FstrProcPredios(astrMens)
            Case 8191
                lstrMens = FstrProcPropietarios(astrMens)
            Case 16383
                lstrMens = FstrProcDocContables(astrMens)
            Case 32767
                lstrMens = FstrProcSerIdentificacion(astrMens)
        End Select
        If lblnInstalacionOk Then
            Select Case GobjParametros.EnuEstadoAplicacion
                Case EnuEstadoAplicacionDef.EnuListoImportar
                    lstrMens = FstrProcSaldosIniciales(astrMens)
                Case EnuEstadoAplicacionDef.EnuParaCierreMes
                    lstrMens = FstrProcParaCierre(astrMens)
                Case EnuEstadoAplicacionDef.EnuServicioNotOk
                    lstrMens = FstrProcServiciosPerNoOk(astrMens)
                Case EnuEstadoAplicacionDef.EnuListoImpNDb
                    lstrMens = FstrProcInteresesDebidos(astrMens)
                Case EnuEstadoAplicacionDef.EnuSinModulos
                    lstrMens = FstrProcModulo(astrMens)
                Case EnuEstadoAplicacionDef.EnuSinPresupuesto
                    lstrMens = FstrProcPresupuesto(astrMens)
                Case EnuEstadoAplicacionDef.EnuSinCalcAdmin
                    lstrMens = FstrProcCuotasAdmin(astrMens)
                Case EnuEstadoAplicacionDef.EnuServicioPorCal
                    lstrMens = FstrProcCalcularSer(astrMens)
                Case EnuEstadoAplicacionDef.EnuHayItemsProgFactPorProcesar
                    lstrMens = FstrProcPrefacturar(astrMens)
                Case EnuEstadoAplicacionDef.EnuHayPrefacturas
                    lstrMens = FstrProcFacturar(astrMens)
                Case EnuEstadoAplicacionDef.EnuDebeAjustarCuotasAdmin
                    lstrMens = FstrProcRetroactivo(astrMens)
                Case EnuEstadoAplicacionDef.EnuCrearServicioAno
                    lstrMens = FstrProcServicioAFacturar(astrMens)
                Case EnuEstadoAplicacionDef.EnuCrearAno
                    lstrMens = FstrProcCrearAno(astrMens)
                Case EnuEstadoAplicacionDef.EnuCausarInt
                    lstrMens = FstrProcCausaInt(astrMens)
                Case EnuEstadoAplicacionDef.EnuDocPorProEFac
                    lstrMens = FstrProcesarEDocs(astrMens)
                Case EnuEstadoAplicacionDef.EnuFactFueraDeFecha
                    lstrMens = FstrProcFactFueraFecha(astrMens)
                Case Else
                    lstrMens = String.Empty
            End Select
        End If
        Return lstrMens
    End Function

    Private Function FstrProcParaCuentasCont(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para crear las cuentas contables"
        Dim lstrMens = "1. De clic en las opción " & Chr(34) & "Acciones->Parametrizar" &
                Chr(34) & " en el menú de la ventana principal." & vbCrLf & "2. En la ventana " &
                Chr(34) & "Parametrización" & Chr(34) & " abierta, oprima simultaneamente las " &
                "teclas " & Chr(34) & "Ctrl" & Chr(34) & " y " & Chr(34) & "E" & Chr(34) &
                " (Ctrl+E), para expandir el arbol de la aplicaciòn ubicado en el panel " &
                "izquierdo de la ventana." & vbCrLf & "3. De clic sobre el nodo " & Chr(34) &
                "Cuentas Contables" & Chr(34) & " para seleccionarlo." & vbCrLf & "4. De clic " &
                "izquierdo (contrario) sobre el nodo seleccionado y luego de clic sobre la opción " &
                Chr(34) & "Abrir Cuentas Contabilidad" & Chr(34) & " del menú contextual " &
                "desplegado." & vbCrLf & "5. Si va a ingresar cuenta por cuenta siga en el punto 6, " &
                "de lo contrario, si las va a importar, siga en el punto 9." & vbCrLf & "6.  En la " &
                "ventana abierta, de clic sobre la opcion " & Chr(34) & "Acciones->Crear Nuevo" &
                Chr(34) & " del menú de la ventana, con lo cual el estado de la ventana cambia a " &
                Chr(34) & "Creando" & Chr(34) & ", como se muestra en la barra de herramientas." &
                vbCrLf & "7. Ingrese los datos requeridos." & vbCrLf & "8. De clic en el bóton " &
                Chr(34) & "Aceptar" & Chr(34) & " o en la opción del menú de la ventana " & Chr(34) &
                "Acciones->Guardar" & Chr(34) & ". Repita este proceso para ingresar cada cuenta " &
                "contable." & vbCrLf & "9. Para  importar las cuentas desde un archivo de Excel " &
                "de clic sobre la opción " & Chr(34) & "Acciones->Importar cuentas contables" &
                Chr(34) & " del menú de de la ventana y siga las instrucciones siguientes:" &
                FstrprocImportación("PlantillaCuentasCont_OrionPlus.xlsx")
        Return lstrMens
    End Function

    Private Function FstrProcParaParametros(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Opciones de la copropiedad"
        Dim lstrMens = "1. En la ventana " & Chr(34) & "Parametrización" & Chr(34) &
                " de clic en el nodo " & Chr(34) & "Copropiedad: <Nombre de la Copropiedad>" &
                Chr(34) & " para seleccionarlo." & vbCrLf & "2. De clic izquierdo (contrario) " &
                "sobre el nodo seleccionado y en el menú contextual displegado, de clic en " &
                "la opción " & Chr(34) & "Abrir Opciones de la Copropiedad" & Chr(34) & "." &
                vbCrLf & "3. En la ventana desplegada diligencie todos los campos requeridos, " &
                "que son aquellos cuyo nombre aparece en letras rojas; y los no requeridos " &
                "diligéncielos según las necesidades y las preferencias de la copropiedad." &
                vbCrLf & "4. Despues de satisfechos todos los campos requeridos se habilita el " &
                "botón " & Chr(34) & "Aceptar" & Chr(34) & ". De clic sobre éste botón para " &
                "grabar la información en la base de datos." & vbCrLf & "5. Cierre la ventana " &
                " dando clic sobre el botón " & Chr(34) & "Aceptar" & Chr(34) & " o sobre el botón " &
                Chr(34) & "Cerrar" & Chr(34) & "." & vbCrLf & vbCrLf & "NOTA: si se cancela " &
                "el proceso, es necesario cerrar la ventana y volverla a abrir para repetirlo."
        Return lstrMens
    End Function

    Private Function FstrProcCtasBancos(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Crear Cuentas Bancarias"
        Dim lstrMens = "OBSERVACION: Antes de crear una cuenta bancaria, verifique que el código " &
                "de la cuenta contable para el banco este creada!" & vbCrLf & "1. En la ventana " &
                Chr(34) & "Parametrización" & Chr(34) & ", expanda el arbol de la aplicaciòn " &
                "ubicado en el panel izquierdo, dando clic en el menú de la ventana, en la opción" &
                Chr(34) & "Acciones->" & "Expandir Todo" & Chr(34) & "." & vbCrLf & "2. Seleccione " &
                "el nodo " & Chr(34) & "Cuentas Bancarias" & Chr(34) & " dando clic sobre el." &
                vbCrLf & "3. De clic derecho (contrario) sobre el nodo seleccionado y de clic " &
                "sobre el menú contextual desplegado, enla opción " & Chr(34) & "Abrir Cuentas " &
                "Bancarias" & Chr(34) & vbCrLf & "4. En la ventana " & Chr(34) &
                "Cuentas Bancarias de la Copropiedad" & Chr(34) & " de clic sobre la opcion " &
                Chr(34) & "Acciones->Crear Nuevo" & Chr(34) & " del menú de la ventana con " &
                "lo cual el estado de la ventana cambia a " & Chr(34) & "Creando" & Chr(34) &
                " como se muestra en la barra de herramientas." & vbCrLf & "5. Diligencie todos " &
                "los campos requeridos (nombre en rojo). Para encontrar el " & Chr(34) & " Código " &
                "Cuenta Contanilidad" & Chr(34) & ", de clic sobre el boton con el signo " &
                Chr(34) & "?" & Chr(34) & " y seleccionela en lista de las cuentas contables " &
                "ya creadas." & vbCrLf & "6. De clic en el bóton " & Chr(34) & "Aceptar" & Chr(34) &
                " o en la opción del menú de la ventana, " & Chr(34) & "Acciones->Guardar" &
                Chr(34) & "." & vbCrLf & "7. Repita el proceso para ingresar cada cuenta bancaria " &
                "y por último cierre la ventana."
        Return lstrMens
    End Function

    Private Function FstrProcSectores(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Crear los Sectores que componen la copropiedad"
        Dim lstrMens = "OBSERVACION: Un sector está compuesto por todos aquellos predios o unidades " &
                "de área privada, que por sus caracterísitcas, contribuyen de la misma manera al " &
                "presupuesto de ingresos de la copropiedad!" & vbCrLf & vbCrLf &
                "1. Abra la ventana " & Chr(34) & "Parametrización" & Chr(34) &
                " (Ctrl+Z) y de clic en la opción del menú, " & Chr(34) & "Tablas Auxiliares->" &
                "Sectores" & Chr(34) & "." & vbCrLf & "2. Establezca el estado de la ventana " &
                "abierta en modo " & Chr(34) & "Creando" & Chr(34) & " oprimiendo simultaneamente " &
                "las  teclas " & Chr(34) & "Ctrl" & Chr(34) & " y " & Chr(34) & "N" & Chr(34) &
                " (Ctrl+N)." & vbCrLf & "3. Escriba el nombre del sector en el campo " &
                "correspondiente y luego oprima la " & "tecla " & Chr(34) & "Tab" & Chr(34) &
                ", lo cual ubica el cursor en el " & "siguiente  campo." & vbCrLf & "4. De clic " &
                "sobre el botón " & Chr(34) & "Aceptar" & Chr(34) & "para grabar la informaación. " &
                "La ventana permanece en modo " & Chr(34) & "Creando" & Chr(34) & " con el fin de " &
                "siguir creando todos los sectores." & vbCrLf &
                "5. Una vez se haya terminado de crear los sectores de clic en el botón" & Chr(34) &
                "Cancelar" & Chr(34) & " con lo cual le ventana queda en modo " & Chr(34) &
                "Consultando" & Chr(34) & ", después de lo cual puede cerrar la ventana." & vbCrLf &
                vbCrLf & "NOTAS: " & vbCrLf & "1.En el caso de que esten establecidos descuentos por pronto pago, " &
                "es requerido ingresar en esta ventana el valor del descuento para cada sector." &
                vbCrLf & "2. Los datos correpondientes a áreas, que aparecen en la parte superior " &
                "derecha de la ventana, seran calculados posteriormente por la aplicación!"
        Return lstrMens
    End Function

    Private Function FstrProcModulos(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Crear los Módulos que contribuyen con la Cuota de Administración"
        Dim lstrMens = "NOTA: es muy probable que solo exista un solo módulo al cual contribuyan " &
                "todos los sectores de la copropiedad." & vbCrLf & vbCrLf & "1. Abra la ventana " & Chr(34) & "Parametrización" & Chr(34) & " (Ctrl + Z)" &
                " y de clic en la opción del menú " & Chr(34) & "Tablas Auxiliares->Módulos de " &
                "Contribución" & Chr(34) & "." & vbCrLf & "2. Establezca el estado de la " &
                "ventana abierta en modo " & Chr(34) & "Creando" & Chr(34) & " dando clic en el " &
                "menú de la ventana sobre la opción " & Chr(34) & "Acciones->Crear Nuevo" & Chr(34) &
                "." & vbCrLf & "3. Diligencie los campos de la ventana." & vbCrLf &
                "4. Grabe los datos dando clic sobre la opción del menu " & Chr(34) & "Acciones" &
                "->Guardar" & Chr(34) & "." & vbCrLf & "5. Ingrese todos los Móduos de " &
                "Contribución y cuando haya terminado cierre la ventana oprimiendo la tecla " &
                Chr(34) & "Esc" & Chr(34) & "."
        Return lstrMens
    End Function

    Private Function FstrProcSectoresModulo(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Asignar Sectores que contribuyen a los Módulos"
        Dim lstrMens = "1. Vaya a la ventana " & Chr(34) & "Parametrización" & Chr(34) &
                " y abra la ventana " & Chr(34) & "Módulos de Contribución" & Chr(34) &
                ", dando clic en la opción del menu " & Chr(34) & "Tablas Auxiliares->Módulos de " &
                "Contribución" & vbCrLf & "2. En la ventana, de clic en la opción del menú " &
                Chr(34) & "Acciones->Abrir Sectores Contribuyentes" & Chr(34) & " o en el botón " &
                Chr(34) & "Abrir Sectores Contribuyentes" & Chr(34) & " ubicado en la parte " &
                "inferior izquierda de la ventana. De esta manera se abre la ventana " & Chr(34) &
                "Sectores Contribuyentes al Módulo" & Chr(34) & "." & vbCrLf & "3. Si el caso es " &
                "que todos los sectores contribuyen con una Tasa de Contribución del 100%, de clic " &
                "en la opción del menú " & Chr(34) & "Acciones->Agregar todos los sectores" & Chr(34) &
                ", de lo contrario siga los siguientes pasos." & vbCrLf & "4. De clic en la " &
                "opción del menú " & Chr(34) & "Acciones->Agregar Sector Contribuyente" & Chr(34) &
                "." & vbCrLf & "5. Con la ventana en estado " & Chr(34) & "Creando" & Chr(34) &
                ", seleccione de la lista descolgante " & Chr(34) & "Sector Contribuye al Módulo" &
                Chr(34) & ", el sector que desea agregar." & vbCrLf & "6. Ingrese la tasa de " &
                "contribución en el campo respectivo." & vbCrLf & "7. Si los datos ingresados son " &
                "validos, se activa el botón " & Chr(34) & "Aceptar" & Chr(34) & ". De clic " &
                "sobre este botón para guardar los datos" & vbCrLf & "8. Repita este proceso " &
                "para cada Sector de cada Módulo de Contribución."
        Return lstrMens
    End Function

    Private Function FstrProcAnoInicial(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para crear el Año Inicial"
        Dim lstrDocs As String
        If GobjParametros.BlnEFacAutorizado Then
            lstrDocs = "facturas"
        Else
            lstrDocs = "cuentas de cobro"
        End If
        Dim lstrMens = "OBSERVACION: Toda la infromación relacionada con las Cuotas de " &
                "Administración, como son su valor y todos los documentos generados (Facturas o " &
                "Cuentas de Cobor, Recibos de Caja y Notas) son guardados en la base de datos " &
                "relacionados con el Año respectivo" & vbCrLf & vbCrLf & "1. Abra la ventana " &
                Chr(34) & "Parametrización" & Chr(34) & ", expanda el arbol (Ctrl+E) y " &
                "seleccione el nodo " & Chr(34) & "Copropiedad: <Nombre Copropiedad>" & Chr(34) &
                vbCrLf & "2. De clic derecho (contrario) sobre el nodo seleccionado, y en el " &
                "menú contextual desplegado, de clic sobre la opción " & Chr(34) & "Crear Año" &
                Chr(34) & "." & vbCrLf & "3. En la ventana desplegada, indique si las primeras " &
                lstrDocs & " de las cuotas de administración que generará la aplicación, " &
                "corresponde al mes calendario actual o al anterior." & vbCrLf & "4. Luego de " &
                "clic sobre el botón " & Chr(34) & "Aceptar" & Chr(34) & ". En este momento el " &
                "programa crea el año correspondiente, y se debe proceder con su parametrización." &
                vbCrLf & "PARAMETRIZACION AÑO INICIAL: " & vbCrLf & "1. Abra la ventana del año " &
                "seleccionando el nodo respectivo y luego de clic derecho (contrario) sobre el." &
                vbCrLf & "2. En el menú contextual desplegado, seleccione la opcion " & Chr(34) &
                "Abrir Año" & Chr(34) & "." & vbCrLf & "3. Establezca el estado de la ventana del " &
                "año en " & Chr(34) & "Modificando" & Chr(34) & ", dando clic en el icono " &
                "correspondiente de la " &
                "barra de tareas ubicada en la parte superior de la ventana." & vbCrLf & "4. Haga los " &
                "cambios necesarios segun las necesidades y las condiciones de la copropiedad!" &
                vbCrLf & "5. Guarde los cambios hechos dando clic en el botón " & Chr(34) &
                "Aceptar" & Chr(34) & " y cierre la ventana."
        Return lstrMens
    End Function

    Private Function FstrProcServicioAFacturar(ByRef astrMensTitulo As String) As String
        Dim lstrProce As String
        If GobjParametros.BlnEFacAutorizado Then
            lstrProce = "facturar"
        Else
            lstrProce = "cobrar"
        End If

        astrMensTitulo = "Proceso para crear los Servicios a " & lstrProce
        Dim lstrMens = "OBSERVACION: Hay dos clases de servicios a " & lstrProce &
                ": a) Los servicios anuales, los cuales tienen vigencia solo durante el año " &
                "correspondiente, y b) los servicios permanentes, que como su nombre lo dice, " &
                "son para ser usados en el momento que se necesiten." & vbCrLf & vbCrLf &
                "1. Abra la ventana " & Chr(34) & "Parametrización" & Chr(34) & " (Ctrl+Z), " &
                "expanda el arbol de la aplicación (Ctrl+E) y seleccione el nodo " &
                "correspondiente al año requerido " & Chr(34) & "aaaa" & Chr(34) &
                " para el servicio anual, y para el servicio permanente seleccione el nodo " &
                Chr(34) & "Servicios Permanentes" & Chr(34) & "." & vbCrLf & "2. De clic " &
                "derecho (contrario) sobre el nodo seleccionado, y en el menú contextual " &
                "desplegado, de clic sobre la opción " & Chr(34) & "Abrir Servicio Anual" &
                Chr(34) & " o " & Chr(34) & "Abrir Servicio Permanente" & Chr(34) & " según " &
                "sea el caso." & vbCrLf & "3. Establezca el estado de la " & "ventana en modo " &
                Chr(34) & "Creando" & Chr(34) & "." & vbCrLf & "4. Diligencie la información " &
                "requerida, de acuerdo a las condiciones de la copropiedad." & vbCrLf &
                "5. Cuando todos los campos requeridos esten satisfechos, se habiltará el " &
                "botón " & Chr(34) & "Aceptar" & Chr(34) & ". De clic sobre él para guardar " &
                "el Servicio en la base de datos."
        Return lstrMens
    End Function

    Private Function FstrProcTasaMora(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para establecer la tasa de Mora"
        Dim lstrMens = "1. Abra la ventana " & Chr(34) & "Parametrización" & Chr(34) &
                " (Ctrl+Z), expanda el arbol (Ctrl+E) y " & "seleccione el nodo " & Chr(34) &
                "Tasas Interés de Mora" & Chr(34) & vbCrLf & "2. De clic derecho (contrario) " &
                "sobre el nodo seleccionado, y en el menú contextual desplegado, de clic sobre " &
                "la opción " & Chr(34) & "Abrir Tasas de Mora" & Chr(34) & "." & vbCrLf & "3. " &
                "Establezca la ventana en estado " & Chr(34) & "Creando" & Chr(34) & " (Ctrl+N)." &
                vbCrLf & "4. Diligencie la información requerida." & vbCrLf & "5. Cuando todos " &
                "los campos requeridos esten satisfechos, de clic sobre el botón " &
                Chr(34) & "Aceptar" & Chr(34) & " para guardar la información." & vbCrLf &
                vbCrLf & "NOTA: Siempre la última tasa ingresada regirá desde la fecha ingresada " &
                "hasta el día de Hoy."
        Return lstrMens
    End Function

    Private Function FstrProcTerceros(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para crear Terceros"
        Dim lstrMens = "1. Estando en la ventana principal, abra ventana " & Chr(34) &
                "Parametrización" & Chr(34) & " (Ctrl+Z)." & vbCrLf & "2. De clic sobre la opción " &
                Chr(34) & "Tablas Auxiliares->" & "Terceros" & Chr(34) & " del menú de la " &
                "ventana, para abrir la ventana de terceros." & vbCrLf & "3. Si va crear un " &
                "tercero siga con el punto 4; ahora bien si va a importar los terceros vaya " &
                "al punto 6." & vbCrLf & "4. Establezca la ventana en estado " & Chr(34) &
                "Creando" & Chr(34) & " (Ctrl+N)." & vbCrLf & "5. Diligencie la " & "información " &
                "correspondiente, y cuando todos los campos requeridos esten satisfechos, de clic " &
                "sobre el botón " & Chr(34) & "Aceptar" & Chr(34) & " para guardar la " &
                "información." & vbCrLf & "6. Para importar los tercero de clic en la opción " &
                Chr(34) & "Acciones->Importar Terceros" & Chr(34) & " del menú de la ventana " &
                "y siga las instrucciones siguientes:" &
                FstrprocImportación("PlantillaTerceros_OrionPlus.xlsx")
        Return lstrMens
    End Function

    Private Function FstrProcClientes(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para crear los Clientes"
        Dim lstrMens = "1. Estando en la ventana principal, de clic sobre la opción " &
                Chr(34) & "Clientes" & Chr(34) & " del menú principal para abrir la ventana " &
                Chr(34) & "Ficha Clientes" & Chr(34) & "." & vbCrLf & "2. Si va crear un " &
                "cliente siga con el punto 3; si va a importar los clientes vaya " &
                "al punto 5." & vbCrLf & "3. Establezca la ventana en estado " & Chr(34) &
                "Creando" & Chr(34) & " (Ctrl+N)." & vbCrLf & "4. Diligencie la " & "información " &
                "correspondiente, y cuando todos los campos requeridos esten satisfechos, de clic " &
                "sobre el botón " & Chr(34) & "Aceptar" & Chr(34) & " para guardar la " &
                "información." & vbCrLf & "5. Para importar los clientes de clic en la opción " &
                Chr(34) & "Acciones->Importar Clientes" & Chr(34) & " del menú de la ventana " &
                "y siga las instrucciones siguientes:" &
                FstrprocImportación("PlantillaClientes_OrionPLus.xlsx")
        Return lstrMens
    End Function

    Private Function FstrProcPredios(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para crear los Predios"
        Dim lstrMens = "NOTA : Al finalizar el proceso de importación de los predios, y cada vez " &
                "que se crea un nuevo predio o se modifica en el programa el área de un " &
                "predio, la aplicación calcula o recalcula los coeficiente de propiedad de todos " &
                "los predios, a partir de las área y los factores de ponderación establecidos en " &
                "el " & Chr(34) & "Reglamento de la Copropiedad" & Chr(34) & vbCrLf & vbCrLf &
                "1. Estando en la ventana principal, de clic sobre la opción " & Chr(34) &
                "Predios" & Chr(34) & " del menú principal para abrir la ventana " & Chr(34) &
                "Ficha Predios" & Chr(34) & "." & vbCrLf & "2. Si va crear un predio siga con " &
                "el punto 3; si va a importar los predios vaya al punto 5." & vbCrLf &
                "3. Establezca la ventana en estado " & Chr(34) & "Creando" & Chr(34) &
                " (Ctrl+N)." & vbCrLf & "4. Diligencie la información correspondiente, y cuando " &
                "todos los campos requeridos esten satisfechos, de clic " & "sobre el botón " &
                Chr(34) & "Aceptar" & Chr(34) & " para guardar la " & "información." & vbCrLf &
                "5. Para importar los predios de clic en la opción " & Chr(34) & "Acciones->" &
                "Importar Predios" & Chr(34) & " del menú de la ventana y siga las instrucciones " &
                "siguientes:" & FstrprocImportación("PlantillaPredios_OrionPlus.xlsx")
        Return lstrMens
    End Function

    Private Function FstrProcPropietarios(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para vincular Propietarios a Predios"
        Dim lstrMens = "NOTA 1: Un predio puede tener varios propietarios; todos ellos deben " &
                "estar creados como clientes. La suma de los porcentajes de propiedad de los propietarios del predio debe ser " &
                "exactamente el ciento por ciento. El valor a cobrar a cada propietario sera " &
                "calculado teniendo en cuenta el porcentaje de propiedad de cada uno de ellos!" &
                vbCrLf & "NOTA 2: Cuando se crea un predio desde la ventana " & Chr(34) &
                "Ficha Predios" & Chr(34) & ", se requiere ingresar los propietarios del predio; " &
                "pero si hay predios que fueron importados se puede utilizar ese metodo o " &
                "importar los propietarios siguiendo el siguiente procedimiento:" & vbCrLf &
                "1. Estando en la ventana principal, de clic sobre la opción " & Chr(34) &
                "Predios" & Chr(34) & " del menú principal para abrir la ventana " & Chr(34) &
                "Ficha Predios" & Chr(34) & "." & vbCrLf & "2. De clic en la opción " & Chr(34) &
                "Acciones->Importar Propietarios" & Chr(34) & " del menú de la ventana y siga " &
                "las instrucciones siguientes:" &
                FstrprocImportación("PlantillaPropietarios_OrionPlus.xlsx")
        Return lstrMens
    End Function

    Private Function FstrProcDocContables(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para definir los Documentos Contables"
        Dim lstrMens = "NOTA 1: Cuando se ha definido que el programa debe exportar el movimiento " &
                "contable generado en un período determinado, para ser importado por la aplicación " &
                "contable, es necesario definir los documentos contables de acuerdo a las " &
                "caracterísitcas de la aplicación contable. Para tal fin siga el siguiente " &
                "procedimiento: " & vbCrLf & "1. Estando en la ventana principal, abra la ventana " &
                Chr(34) & "Parametrización" & Chr(34) & " oprimiendo las teclas (Ctrl+Z), y " &
                "expanda el arbol oprimiendo las teclas (Ctrl+E)." & vbCrLf & "2. Seleccione el " &
                "nodo (rama) " & Chr(34) & "Documentos" & Chr(34) & " dando clic sobre él." &
                vbCrLf & "3. De clic izquierdo sobre él nodo seleccionado, y luego clic sobre " &
                "la opción " & Chr(34) & "Abrir Documentos" & Chr(34) & " del menú contextual " &
                "desplegado." & vbCrLf & "4. En la tabla ubicada en el parte inferior de la " &
                "ventana mostrada, seleccione la fila correspondiente al documento que va a " &
                "modificar." & vbCrLf & "5. Establezca la ventana en modo " & Chr(34) &
                "Modificando" & Chr(34) & " oprimiendo las teclas " & Chr(34) & "Ctrl+M" & Chr(34) &
                "." & vbCrLf & "6. Diligencie los campos necesarios y luego guarde los cambios " &
                "oprimirndo las teclas " & Chr(34) & "Ctrl+G" & "." & vbCrLf & "Repita estos " &
                "pasos para todos y cada uno de los documentos listados en la tabla." & vbCrLf &
                vbCrLf & "NOTA: Cuando la interfaz contables esta definida para hacerse por " &
                "documento es necesario definir los ocho primeros documentos; y cuando esta " &
                "definida para hacerse por Comprobante solo se necesita definir el " & Chr(34) &
                "Comprobante Interfaz Contable" & Chr(34) & ". Esta definición se hace en los " &
                "parámetros de la copropiedad."
        Return lstrMens
    End Function

    Private Function FstrProcSerIdentificacion(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para crear el Servicio de Indentificación"
        Dim lstrMens = "NOTA 1: El servicio de indentificación sirve para conocer el predio al " &
                "cual corresponde una determinada consignación. Esto funciona muy bien siempre y " &
                "cuando el valor de la consignación coincida exactamente con el valor cobrado en " &
                "la factura o cuenta de cobro." & vbCrLf & "NOTA 2: Se recomienda que en el " &
                "caso de ser utilizado este servicio, los valores a cobrar por él, se establezcan " &
                "en la plantilla de importación de predios, de lo contrario habría que entrar a " &
                "modificar cada predio e ingresar el valor uno por uno. Es de notar que este valor " &
                "solo lo requieren los predios agrupadores!" & vbCrLf & vbCrLf & "1. Estando en la " &
                "ventana principal, abra la ventana " & Chr(34) & "Parametrización" & Chr(34) &
                " oprimiendo las teclas (Ctrl+Z), y expanda el arbol oprimiendo las teclas " &
                "(Ctrl+E)." & vbCrLf & "2. Seleccione el nodo (rama) " & Chr(34) & "Servicios " &
                "Permanentes" & Chr(34) & " dando clic sobre él." & vbCrLf & "3. De clic izquierdo " &
                "sobre él nodo seleccionado, y luego clic sobre la opción " & Chr(34) & "Abrir " &
                "Servicio Permanente" & Chr(34) & " del menú contextual desplegado." & vbCrLf &
                "4. Establezca la ventana en modo " & Chr(34) & "Creando" & Chr(34) & " oprimiendo " &
                "las teclas " & Chr(34) & "Ctrl+N" & Chr(34) & " o dando clic sobre el botón " &
                "correspondiente de la barra de herramientas." & vbCrLf & "5. Diligencie los " &
                "campos requeridos teniendo en cuenta que debe seleccionar el campo " & Chr(34) &
                "Es Servicio de Identificación" & Chr(34) & vbCrLf & "6. Después de que todos los " &
                "campos requeridos esten satisfechos, de clic " & "sobre el botón " & Chr(34) &
                "Aceptar" & Chr(34) & " para guardar la información."
        Return lstrMens
    End Function

    Private Function FstrProcSaldosIniciales(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para registrar los Saldos Iniciales"
        Dim lstrMens = "NOTA: Normalmente, en una nueva instalación, es necesario registrar " &
                "todos los saldos de las cuentas por cobrar al último día del mes " &
                "inmediatamente anterior al mes de la primera generaciòn de los documentos " &
                "de cobro por parte del programa. Si no hay necesidad de importar los saldos " &
                "iniciales, siga en el punto 5." & vbCrLf & vbCrLf & "1. Cree los servicios " &
                "permanentes necesarios para los saldos debidos a la fecha. Para este fin, " &
                "estando en la ventana principal, abra la ventana " & Chr(34) & "Parametrización" &
                Chr(34) & " (Ctrl+Z), expanda el arbol (Ctrl+E), seleccione el nodo " & Chr(34) &
                "Servicios Permanentes" & Chr(34) & ", de clic izquierdo sobre él y luego " &
                "sobre la opción " & Chr(34) & "Abrir Servicio Permanente" & Chr(34) &
                ". Establezca la ventana en modo " & Chr(34) & "Creando" & Chr(34) &
                " (Ctrl+N), diligencie los campos requeridos y guarde la información (" &
                "Ctrl+G)." & vbCrLf & "2. Abra el archivo de Excel " & Chr(34) &
                "PlantillaFacturas_OrionPLus.xlsx" & Chr(34) & " y diligencie la información " &
                "alli establecida." & vbCrLf & "3. Copie el archivo de Excel en la carpeta " &
                Chr(34) & "C:\Panorama.Net\DatPrg" & Chr(34) & vbCrLf & "4. De clic en la " &
                "opción " & Chr(34) & "Acciones -> Importar Datos Iniciales -> Importar " &
                "Facturas Iniciales" & Chr(34) & " en el menú de la ventana principal. " &
                "En este momento se lleva a cabo el proceso, mostrando el avance en la barra " &
                "de progreso ubicada en la parte inferior de la ventana." & vbCrLf & "5. Si no " &
                "es necesario llevar a cabo la importación, de clic en la opción " & Chr(34) &
                "Acciones -> Importar Datos Iniciales -> No Importar Facturas Iniciales" &
                Chr(34) & " en el menú de la ventana principal."
        Return lstrMens
    End Function

    Private Function FstrProcInteresesDebidos(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para registrar Intereses de Mora debidos"
        Dim lstrMens = "NOTA: Con el fin de que la información registrada en la contabilidad y los " & "datos en Orión Plus coincidan exactamente, es necesario importar los intereses " &
                "de mora registrados en la contabilidad. Ahora bien, los intereses del último mes " &
                "pueden estar ya calculados e incluidos en la contabilidad, o pueden ser calculados" &
                " por el programa al cierre del mes anterior al mes de inicio de operaciones del " &
                "programa. Para importar los intereses siga con el numeral 1, de los contrario siga" &
                " con el numeral 5." & vbCrLf & vbCrLf & "1. Abra el archivo de Excel " & Chr(34) &
                "PlantillaNotasDb_OrionPLus.xlsx" & Chr(34) & " y diligencie la información alli " &
                "establecida. Si los intereses de mora a importar contienen los del mes actual, " &
                "en la columna " & Chr(34) & "CierreMes" & Chr(34) & " esctriba " & Chr(34) &
                "VERDADERO" & Chr(34) & ", de lo contrario excriba " & Chr(34) & "FALSO" & Chr(34) &
                ". Si es " & Chr(34) & "VERDADERO" & Chr(34) & " el proceso hará el cierre del mes." &
                vbCrLf & "2. Copie el " & "archivo de Excel en la carpeta " & Chr(34) &
                "C:\Panorama.Net\DatPrg" & Chr(34) & vbCrLf & "3. Abra la ventana " & Chr(34) &
                "Ficha Nota Intereses Mora" & Chr(34) & " dando clic en la opción " & Chr(34) &
                " Notas -> Notas Intereses por Mora" & Chr(34) & " en el menú de la ventana  " &
                "principal." & vbCrLf & "4. En la ventana  abierta, de clic en la opción " & Chr(34) &
                "Acciones -> Importar Notas Intereses " & "Mora" & Chr(34) & "; en este momento se " &
                "valida la información contenida en el archivo de Excel; si se encuentra algún " &
                "problema, será informado en el área de notificaciones; en este caso abrá el archivo, " &
                "corrijalo y repita el proceso; de lo contrario, se importarán las notas y se " &
                "informará el fin del proceso en el área de notificaciones." & vbCrLf & "5. Si no es " &
                "necesario llevar a cabo la importación, cierre el mes dando clic en la opción " &
                Chr(34) & "Acciones -> Cerrar Mes" & Chr(34) & " del menú de la ventana principal."
        Return lstrMens
    End Function

    Private Function FstrProcModulo(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para vincular Módulos de Contribución al Servicio"
        Dim lstrMens = "NOTA: A los servicios anuales, cuyo año está parametrizado para que no a " &
                "cada servicio le contribuya solo un módulo, y a los servicios permanentes cuyo " &
                "valor a cobrar es calculado por Orión Plus, se les debe vincular uno o más " &
                "módulos con el correspondiente valor de contribución. Este paso no es necesario " &
                "si los servicios van a ser importados, caso en el cual el valor de los modulos " &
                "será actualizado a cero pesos por el programa, cuando el servicio es anual; si es " &
                "un servicio permanente, el valor exacto será actualizado por Orión Plus." & vbCrLf &
                vbCrLf & "1. " &
                "Estando en la ventana principal, abra la ventana " & Chr(34) & "Parametrización" &
                Chr(34) & " oprimiendo las teclas" & Chr(34) & "Ctrl+Z" & Chr(34) & ", expanda el" &
                " arbol oprimiendo las teclas " & Chr(34) & "Ctrl+E" & Chr(34) & ", y seleccione " &
                "el nodo del año concerniente para que se expandan los servicios." & vbCrLf & "2. " &
                "Seleccione el servicio a tratar, de clic derecho sobre él y luego clic sobre la " &
                "opción " & Chr(34) & "Abrir Servicio Anual" & Chr(34) & " con el fin de abrir " &
                "la ventana del servicio." & vbCrLf & "3. Establezca la ventana del servicio en " &
                "modo " & Chr(34) & "Modificando" & Chr(34) & " dando clic sobre el botón " &
                "correpondiente en la barra de herramientas, u oprimirndo las teclas " & Chr(34) &
                "Ctrl+M" & Chr(34) & vbCrLf & "4. De clic sobre la pestaña " & Chr(34) & "Valores" &
                 Chr(34) & " y agregue, modifique o elimine un módulo de contribución con su " &
                 "respectivo valor de contribución, utilizando los botones ubicados debajo de la " &
                 "tabla " & Chr(34) & "Módulos que contribuyen al Servicio" & Chr(34) & "." &
                 vbCrLf & "5. Cuando termine de vincular los  módulos, guarde la información " &
                 "dando clic en el botón " & Chr(34) & "Aceptar" & Chr(34) & "." & vbCrLf &
                 "6. Siempre en estos casos, el proceso para el calculo de los valores a cobrar " &
                 "se genera desde la ventana del servicio, dando clic en la opción " & Chr(34) &
                 "Acciones -> Calcular valores a cobrar" & Chr(34) & " del menú de la ventana ." &
                 vbCrLf & "7. Por último cierre la ventana oprimirndo la tecla " & Chr(34) & "Esc" &
                 Chr(34) & ". Repita el proceso para todos los servicios  del año."
        Return lstrMens
    End Function

    Private Function FstrProcPresupuesto(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para ingresar el Presupuesto Anual"
        Dim lstrMens = "NOTA: El Presupuesto Anual de Ingresos por Cuotas de Administración se " &
                "requiere, para que conjuntamente con los Coeficientes de Propiedad, Orión Plus " &
                "calcule las cuotas de administración. Si por algún motivo esto no es posible, el " &
                "presupuesto será calculado a partir de las cuotas de administración importadas, " &
                "proceso este que se explicará en el siguiente paso. Sin embargo, es necesario " &
                "ingresar un valor para poder continuar, y en el caso de que se importen las " &
                "cuotas de administración, este valor será sobrescrito por el valor calculado " &
                "en el proceso de importación." & vbCrLf & vbCrLf & "1. Estando en la ventana " &
                "principal, abra la ventana " & Chr(34) & "Parametrización" & Chr(34) &
                " oprimiendo las teclas" & Chr(34) & "Ctrl+Z" & Chr(34) & "; ahora expanda el" &
                " arbol oprimiendo las teclas " & Chr(34) & "Ctrl+E" & Chr(34) & "." & vbCrLf &
                "2. Seleccione el nodo (rama) correspondinte al año (aaaa) y de clic derecho sobre " &
                "él; luego clic sobre la opción " & Chr(34) & "Abrir Año" & Chr(34) & " del menú " &
                "contextual desplegado." & vbCrLf & "3. Establezca la ventana del año en modo " &
                Chr(34) & "Modificando" & Chr(34) & " dando clic sobre el botón correpondiente en " &
                "la barra de herramientas, u oprimirndo las teclas " & Chr(34) & "Ctrl+M" & Chr(34) &
                ", e ingrese el valor del presupuesto anual en el campo " & Chr(34) &
                "Presupuesto Ingresos por Cuota de Administracón" & Chr(34) & vbCrLf & "4. Guarde " &
                "los cambios oprimiendo las teclas " & Chr(34) & "Ctrl+G" & Chr(34) &
                ", o dando clic sobre el botón correspondiente de la barra de herramientas." &
                vbCrLf & "5. Cierre la ventana. Oprimiendo la tecla " & Chr(34) & "Esc" & Chr(34) &
                " puede cerrar cualquier ventana de Orión Plus."
        Return lstrMens
    End Function

    Private Function FstrProcCuotasAdmin(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para Cálculo de las Cuotas de Administración"
        Dim lstrOpcionMen As String
        lstrOpcionMen = "Cobranza -> Cobranza Automática -> Programación de Cobros"
        Dim lstrMens = "NOTA: Las cuotas de administración son calculadas a partir del Presupuesto " &
                "Anual de Ingresos por Cuotas de Administración, y de los Coeficientes de " &
                "Propiedad; si por algún motivo esto no es posible, se deben importar las cuotas " &
                "de administración, y el programa calcula el Presupuesto Anual a partir de ellas. " &
                "Para información sobre la importación vaya a la sección " & Chr(34) & "PROCESO " &
                "IMPORTACION CUOTAS ADMINISTRACION" & Chr(34) & " de esta ayuda." & vbCrLf & vbCrLf &
                "1. Estando en la ventana principal, abra la ventana " & Chr(34) & "Parametrización" &
                Chr(34) & " oprimiendo las teclas" & Chr(34) & "Ctrl+Z" & Chr(34) & ", y expanda el" &
                " arbol oprimiendo las teclas " & Chr(34) & "Ctrl+E" & Chr(34) & "." & vbCrLf &
                "2. Seleccione el nodo (rama) correspondinte al año (aaaa) y de clic derecho sobre " &
                "él; luego clic sobre la opción " & Chr(34) & "Abrir Año" & Chr(34) & " del menú " &
                "contextual desplegado." & vbCrLf & "3. De clic en la opción " & Chr(34) &
                "Acciones -> Calcular Cuotas Administración" & Chr(34) & " del menú de la ventana " &
                "del año y el proceso será ejecutado." & vbCrLf & "4.  Con el fin  de verificar " &
                "los valores calculados, cierre la ventana; de clic sobre el nodo  del año " &
                "correspondiente para expandir los nodos de los servicios del año; de clic sobre " &
                "el nodo del servicio para seleccionarlo; luego clic derecho sobre él; y por " &
                "último, clic sobre la opción " & Chr(34) & "Abrir Servicio Anual" & Chr(34) &
                " del menú desplegado." & vbCrLf & "5. De clic sobre la opcion " & Chr(34) &
                "Reportes -> Valores a Cobrar" & Chr(34) & " del menú de la ventana; se abrirá " &
                "el reporte de los valores a cobrar por el servicio, el cual puede ser impreso " &
                "y/o exportado a archivos de varios formatos (pdf , Excel, Word, etc.)" & vbCrLf &
                vbCrLf & "PROCESO IMPORTACION CUOTAS ADMINISTRACION" & vbCrLf & "1. Diligencie " &
                "adecuadamente la plantila contenida en el Archivo " & Chr(34) &
                "PlantillaItemsProgramaFact.xlsx" & Chr(34) & ". A este archivo le puede cambiar " &
                "de nombre para diferenciarlo de otros archivos que contegan datos de diferentes " &
                "servicios para ser importados." & vbCrLf & "2. IMPORTANTE: Cuando el servicio " &
                "a importar es cuota de administración, la columna de la plantilla " & Chr(34) &
                "IdTerceroCliente" & Chr(34) & " se debe dejar en blanco; la columna " & Chr(34) &
                "PeriodoInicioFact" & Chr(34) & " debe contener un texto formado por los cuatros " &
                "digitos del respectivo año seguido de " & Chr(34) & "01" & Chr(34) & " (aaaa01); " &
                "y la " & "columna " & Chr(34) & "CantidadPeriodos" & Chr(34) & " debe ser " &
                "siempre el número " & Chr(34) & "12" & Chr(34) & "." & vbCrLf & "3. Estando en " &
                "la ventana principal, de clic sobre la opción " & Chr(34) & lstrOpcionMen &
                Chr(34) & "." & vbCrLf & "4. De clic sobre la opción " & Chr(34) &
                "Acciones -> Importar Programación de Cobros" & Chr(34) & " del menú de la ventana y siga las instrucciones " &
                "siguientes: " & FstrprocImportación("PlantillaServicioAImportar.xlsx")
        Return lstrMens
    End Function

    Private Function FstrProcCalcularSer(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para cálcular los valores a cobrar de Servicio Permanente"
        Dim lstrMens = "1. Estando en la ventana" & Chr(34) & "Parametrización" & Chr(34) &
            " expanda el arbol de la aplicación oprimiendo " & Chr(34) & "Ctrl+E" & Chr(34) &
            ", seleccione el nodo " & Chr(34) & "Servicios Permanentes" & Chr(34) & " dando clic " &
            "sobre él y luego el nodo del servicio requerido." & vbCrLf & "2. De clic derecho " &
            "sobre le nodo del servicio seleccionado y luego clic sobre la opciòn " & Chr(34) &
            "Abrir Servicio Permanente" & Chr(34) & " con lo cual se abre la ventana del servicio " &
            "requerido." & vbCrLf & "3. De clic sobre la opción " & Chr(34) & "Acciones -> " &
            "Calcular valores a cobrar" & Chr(34) & " del menú de la ventana, con lo cual la " &
            "aplicación procede a hacer el calculo correspondiente, y cuando finaliza, lo notifica " &
            "en el área de notificaciones." & vbCrLf & "4. Verifique el resultado dando clic en la " &
            "opción " & Chr(34) & "Reportes -> Valores a cobrar" & Chr(34) & " del menú de " &
            "la ventana. Si encuentra algún error o inconsistencia corrijalo y repita el procedimiento."
        Return lstrMens
    End Function

    Private Function FstrProcRetroactivo(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para generar Retroactivo"
        Dim lstrMens = "NOTA: Siempre que se introduce un cambio en el valor de los ingresos por " &
                "Cuotas de Administración, el programa solicita generar el servicio " & Chr(34) &
                "Retroactivo Cuota Administración" & Chr(34) & "; sin embargo se da la posiblidad de " &
                "no generarlo. Este proceso calcula el valor debido a cobrar hasta la fecha según " &
                "el nuevo valor de las cuotas de administración, y el valor cobrado desde el mes de " &
                "enero del año actual. Si la diferencia entre los dos valores es positivo, será " &
                "cobrado en el mes o los meses siguientes; y si es negativo, se genera un anticipo " &
                "por la diferencia para ser aplicado en la factura del mes siguiente." & vbCrLf &
                vbCrLf & "1. Estando en la ventana principal, abra la ventana " & Chr(34) &
                "Parametrización" & Chr(34) & " oprimiendo las teclas" & Chr(34) & "Ctrl+Z" &
                Chr(34) & ", y expanda el arbol oprimiendo las teclas " & Chr(34) & "Ctrl+E" &
                Chr(34) & "." & vbCrLf & "2. Seleccione el nodo (rama) correspondinte al año " &
                "actual y de clic derecho sobre él; luego clic sobre la opción " & Chr(34) &
                "Abrir Año" & Chr(34) & " del menú contextual desplegado." & vbCrLf & "3. En la " &
                "ventana del año abierta, de clic en la opción " & Chr(34) & "Acciones -> Generar " &
                "Retroactivo Cuotas " & "Administración" & Chr(34) & ", con lo cual se abre la " &
                "ventana respectiva." & vbCrLf & "4. Digite la cantidad de cuotas en que se " &
                "cobrará el retroactivo, si es que se va a cobrar; de lo contarrio de clic " &
                "sobre el control de chequeo " & Chr(34) & "No generar Retroactivo" & Chr(34) & "." &
                vbCrLf & "5. De clic en el botón " & Chr(34) & "Aceptar" & Chr(34) & " con lo que " &
                "se calculará el valor del retroactivo que debe pagar cada predio, si es que se " &
                "va a cobrar, de lo contrario se marcará que no se cobrará retroactivo." & vbCrLf &
                "6. Este proceso creará un nuevo servicio del año llamado " & Chr(34) &
                "Retroactivo Cuota de Administración" & Chr(34) & " adicionandolo al año actual " &
                "y en donde se puede consultar el valor total del retoactivo y los valores a " &
                "cobrar a cada predio."
        Return lstrMens
    End Function

    Private Function FstrProcParaCierre(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para el cierre de mes"
        Dim lstrMens = "OBSERVACION: Antes de cerrar el último período del año (diciembre) se debe " &
                "crear el nuevo año!" & vbCrLf & vbCrLf & "1. De clic en las opciones del menú " &
                Chr(34) & "Acciones->Cerrar mes" & Chr(34) & "." & vbCrLf & "2. Verifique el período " &
                "y fecha de cierre." & vbCrLf & "3. De clic en el botón " & Chr(34) & "Aceptar" &
                Chr(34) & " para proceder con el cierre del mes, o de clic en el botón " & Chr(34) &
                "Cancelar" & Chr(34) & " para no ejecutar la acción." & vbCrLf &
                vbCrLf & "NOTA: En este proceso el grograma generalmente ejecuta " &
                "las siguientes acciones:" & vbCrLf & "1. Genera una copia de seguridad." &
                vbCrLf & "3. Genera un reporte de la edad de la cartera al momento del cierre." &
                vbCrLf & "4. Genera un reporte de los anticipos por aplicar al momento del cierre." &
                vbCrLf & "5. Causa los intereses de mora a todas las facturas vencidas," & vbCrLf &
                "6. Genera la interfaz contable conteniendo el movimiento contable del período " &
                "cerrado." & vbCrLf & "7. Da la opción de imprimir las notas débito de " &
                "intereses generadas."
        Return lstrMens
    End Function

    Private Function FstrProcCrearAno(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para la creación de un nuevo año"
        Dim lstrMens = "1. Estando en la ventana principal, abra la ventana " & Chr(34) &
                "Parametrización" & Chr(34) & " oprimiendo las teclas" & Chr(34) & "Ctrl+Z" &
                Chr(34) & ", y luego, expanda el arbol de ésta ventana oprimiendo las teclas " &
                Chr(34) & "Ctrl+E" & Chr(34) & "." & vbCrLf & "2. Seleccione el nodo (rama) " &
                "correspondinte a la " & Chr(34) & "Copropiedad" & Chr(34) &
                " dando clic sobre el" & vbCrLf & "3. " & "Expanda el menú contextual dando " &
                "clic contrario (derecho) sobre el dicho nodo y de clic en la opción " & Chr(34) &
                "Crear Año" & vbCrLf & "4. Diligencie la información solicitada en la ventana " &
                Chr(34) & "Incremento Presupuesto." & Chr(34) & vbCrLf & "5. Por último de " &
                "clic en el botón " & Chr(34) & "Aceptar" & Chr(34) & " con l cual el programa " &
                "procede a crear el nuevo año y a calcular las nuevas cuotas de administración " &
                "de acuerdo a la información ingresada."
        Return lstrMens
    End Function

    Private Function FstrProcCausaInt(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para Causar Intereses"
        Dim lstrMens = "1. De clic en las opciones del menú de la ventana pricipal en " &
                Chr(34) & "Acciones->Causar Intereses de Mora" & Chr(34) &
                ", con lo cual se abre la ventana " & Chr(34) & "Causa Intereses de Mora" &
                Chr(34) & "." & vbCrLf & "2. Verifique la información de la ventana." & vbCrLf &
                "3. De clic en el botón " & Chr(34) & "Aceptar" & Chr(34) &
                " para proceder con la causación de intereses, o de clic en el botón " &
                Chr(34) & "Cancelar" & Chr(34) & " para no ejecutar la acción."
        Return lstrMens
    End Function

    Private Function FstrProcServiciosPerNoOk(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para revisar servicios permanentes"
        Dim lstrMens = "1. Vaya a la ventana " & Chr(34) & "Parametrización" & Chr(34) & " y " &
                "expanda el arbol de Oriòn (Ctrl+E)." & vbCrLf & "2. De clic sobre el nodo " &
                Chr(34) & "Servicios Permanentes" & Chr(34) & " para seleccionarlo y luego de clic " &
                "derecho sobre él para mostrar el menú contextual." & vbCrLf & "3. De clic sobre " &
                "el ítem " & Chr(34) & "Abrir Servicio Permanente" & Chr(34) & "." & vbCrLf &
                "4. Navegue por todos los servicios, utilizando los botones correspondientes " &
                "de la barra de herramientas, y revise uno por uno hasta que encuentre el, o los " &
                "servicios que tengan campos requeridos por satisfacer." & vbCrLf & "5. Para " &
                "corregir, establezca la ventana en modo " & Chr(34) & "Modificando" & Chr(34) &
                " (Ctrl+M), e ingrese la información requerida." & vbCrLf & "6. De clic sobre el" &
                "botón " & Chr(34) & "Aceptar" & Chr(34) & " para gravar el servicio." & vbCrLf &
                "7. Por último cierre la ventana oprimiendo la tecla " & Chr(34) & "ESC" & Chr(34) &
                "."
        Return lstrMens
    End Function

    Private Function FstrProcPrefacturar(ByRef astrMensTitulo As String) As String
        Dim lstrNombreDoc As String, lstrOpcion As String, lstrReversarDoc As String,
                lstrOpcionAuto As String
        If GobjParametros.BlnEFacAutorizado Then
            lstrNombreDoc = "Pre-Facturas"
            lstrOpcion = "Facturación -> Facturas"
            lstrOpcionAuto = "Facturación -> Facturación Automática -> Generar Pre-Facturas"
            lstrReversarDoc = "6. REVERSAR PRE-FACTURAS: "
        Else
            lstrNombreDoc = "Pre-Cuentas de Cobro"
            lstrOpcion = "Cobranza -> Cuentas de Cobro"
            lstrOpcionAuto = "Cobranza -> Cobranza Automática  -> Generar Pre-Cuentas de Cobro "
            lstrReversarDoc = "6. REVERSAR PRE-CUENTAS DE COBRO: "
        End If
        astrMensTitulo = "Proceso para generar " & lstrNombreDoc
        Dim lstrMens = "OBSERVACION: Los pre-documentos conservan el consecutivo de los documentos, " &
                "pero se les asigna un prefijo compuesto por tres asteriscos (" & Chr(34) & "***" &
                Chr(34) & ")." & vbCrLf & vbCrLf & "1. Estando en la ventana principal, de clic sobre la " &
                "opción " & Chr(34) & lstrOpcionAuto & Chr(34) & " del menú de la ventana." & vbCrLf & "2. En la " &
                "ventana " & Chr(34) & "Generación de " & lstrNombreDoc & Chr(34) & " abierta, verifique, y si " &
                "es necesario modifique las fechas de vencimiento y limite de pago sin mora." &
                vbCrLf & "3. De clic en el botón " & Chr(34) & "Aceptar" & Chr(34) & " con lo que " &
                "el proceso se ejecuta, mostrando el avence en la barra de progreso. Al terminar, " &
                "se muestra el mensaje correspondiente en el área de notificaciones." & vbCrLf &
                "4. Cierre la ventana, y para revisar las " & lstrNombreDoc & "generadas, de clic en la " &
                "opción " & Chr(34) & lstrOpcion & Chr(34) & " del menú de la ventana principal." &
                vbCrLf & "5. Navegue por lod documents en la ficha abierta, y si " &
                "encuentra que debe hacer algún ajuste o correción, reverse los documentos como " &
                "se indica a continuación." & vbCrLf & lstrReversarDoc & "De clic en la " &
                "opción " & Chr(34) & "Facturación -> Facturación Automática -> " &
                "Reversar " & lstrNombreDoc & Chr(34) & " del menú de la ventana principal." & vbCrLf &
                "7. De clic en el botón " & Chr(34) & "Aceptar" & Chr(34) & " de la ventana, y los " &
                "documentos son reversadas mostrando el avence del proceso en la barra de progreso." &
                vbCrLf & "8. Terminado el proceso, cierre la ventana y si se reversaron las " &
                lstrNombreDoc & ", haga las modificaciones pertinentes y repita el proceso."
        Return lstrMens
    End Function

    Private Function FstrProcFacturar(ByRef astrMensTitulo As String) As String
        If GobjParametros.BlnEFacAutorizado Then
            astrMensTitulo = "Proceso para generar Facturas definitivas"
        Else
            astrMensTitulo = "Proceso para generar Cuenta de Cobro definitivos"
        End If
        Dim lstrMens = "OBSERVACION: Después de generadas las facturas definitivas, no hay " &
                "posibilidad de modificarlas. La opción para corregir una factura, es anularla " &
                "como se indica al final de esta ayuda y corregir lo que sea necesario. Después de " &
                "anulada una factura generada de forma automática, el programa solicita volverla " &
                "a generar con el fin de guardar la integridad entre el valor del presupuesto del " &
                "servicio y los valores realmente cobrados." & vbCrLf & vbCrLf & "1. Estando en la " &
                "ventana principal, de clic sobre la opción " & Chr(34) & "Facturación -> " &
                "Facturación Automática -> Generar Facturas Definitivas" & Chr(34) & " del menú de " &
                "la ventana." & vbCrLf & "2. En la ventana " & Chr(34) & "Genera Facturas " &
                "Definitivas" & Chr(34) & " abierta, de clic en el botón " & Chr(34) & "Aceptar" &
                Chr(34) & ", con lo que a continuación se llevan a cabo los siguientes procesos: " &
                vbCrLf & Chr(9) & "a. Generación estados de cuenta: Por cada cuenta por cobrar al " &
                "momento de la facturación, se genera un documento " & Chr(34) & "Estado de Cuenta" &
                 Chr(34) & ", el cual contiene la información de las facturas debidas por cada " &
                 "cliente (Copropietario, Arrendatario o Tercero). Este estado de cuenta se vincula " &
                 "a la respectiva factura generada y es mostrado en la cuenta de cobro o factura." &
                 vbCrLf & Chr(9) & "b. Genera Facturas Definitivas: Convierte las prefacturas en " &
                 "facturas definitivas y actualiza el prefijo de la factura." & vbCrLf & Chr(9) &
                 "c. Aplica Anticipos: Si existen anticipos por aplicar, son aplicados a la " &
                 "factura generando una " & Chr(34) & "Nota Aplicación de Anticipos" & Chr(34) &
                 "; se incorpora el anticipo aplicado al estado de la cuenta, y por último se " &
                 "actualizan los anticipos por aplicar." & vbCrLf & vbCrLf & "Anulación de " &
                 "factura: El proceso es el siguiente:" & vbCrLf & "1. Estando en la ventana " &
                 "principal, de clic en la  opción " & Chr(34) & "Facturación -> Facturas " &
                 Chr(34) & " del menú de la ventana, con lo cual se abre  la ficha de las facturas, " &
                 "en donde puede consultar todas las facturas registradas en Orión." & vbCrLf &
                 "2. Ubique la facturas deseada utilizando las opciones que se muestran en la " &
                 "opción " & Chr(34) & "Navegar" & Chr(34) & " del menú de la ventana o utilizando " &
                 "los botones respectivos en la barra de herramientas." & vbCrLf & "3. De clic en " &
                 "la opción " & Chr(34) & "Acciones Anular" & Chr(34) & " o de clic en el botón " &
                 " correspondiente de la barra de herramientas." & vbCrLf & "4. Confirme o rechaze " &
                 "la anulación en el cuadro emergente que se presenta." & vbCrLf & vbCrLf &
                 "NOTA: al anular una factura se genera automáticamente un documento " & Chr(34) &
                 "Nota Crédito" & Chr(34) & " el cual se refleja en la pestaña " & Chr(34) &
                 "Movimiento Contable" & Chr(34) & "de la misma ventana. Dando clic derecho sobre " &
                 "la fila de la nota crédito en el movimiento contable, se abre la ficha de la " &
                 "nota crédito correspondiente."
        Return lstrMens
    End Function

    Private Function FstrprocImportación(astrArchivoPlantilla As String) As String
        Dim lstrMens = vbCrLf & vbCrLf & "IMPORTACION: " & "Despúes de haber abierto la " &
                "ventana " & Chr(34) & "Importación" & Chr(34) & " siga los siguientes pasos: " &
                vbCrLf & "1. De clic en el botón " & Chr(34) & "Examinar ..." & Chr(34) &
                " y en el explorador de archivos busque el archivo de Excel " & Chr(34) &
                astrArchivoPlantilla & Chr(34) & ", seleccionelo y luego de clic en el " &
                "botón abrir, con lo cual se mostrará en el campo " & Chr(34) &
                "Ruta y Nombre Archivo de  Origen" & Chr(34) & "." & vbCrLf &
                "2. De clic en el botón " & Chr(34) & "Cargar Tablas" & Chr(34) &
                " y en la lista descolgante " & Chr(34) & "Tablas Cargadas" & Chr(34) &
                ", seleccione la hoja del archivo de Excel que contiene la plantilla con los " &
                "datos a ser importados." & vbCrLf & "3. De clic en el botón " & Chr(34) & "Validar" &
                Chr(34) & " con el fin de validar la información contenida en la tabla." & vbCrLf &
                "4. Si se encuantra algún error, abra el archivo de Excel y corríjalo. Cuando " &
                "todos los datos esten correctos, se habilitará el botón " & Chr(34) & "Aceptar" &
                Chr(34) & "." & vbCrLf & "5. De clic en el botón " & Chr(34) & "Aceptar" & Chr(34) &
                ". En este momento se inicia el proceso de importación, el cual se muestra " &
                " en una barra de progreso en la parte inferior de la ventana. Cuando termina la " &
                "importación se informa en el área de notificación y se abre un archivo de texto " &
                "donde se puede ver la cantidad de elementos importados y el tiempo que duró " &
                "el proceso." & vbCrLf & "6. Cierre la ventana dando clic en el botón " & Chr(34) &
                "Aceptar" & Chr(34) & " o en el botón " & Chr(34) & "Cancelar" & Chr(34)
        Return lstrMens
    End Function

    Private Function FstrProcesarEDocs(ByRef astrMensTitulo As String)
        astrMensTitulo = "Acciones para procesar Documentos Electrónicos"
        Dim lstrMens = "1. Estando en la Ventana Principal, de clic sobre el nodo " & Chr(34) &
                "EFactura" & Chr(34) & " del menú de la ventana." & vbCrLf & "2. En el menú " &
                "desplegado, de clic sobre el nodo " & Chr(34) & "Procesar Documentos " &
                "Electrónicos" & Chr(34) & ". Con esto, Orión tratará de terminar de procesar " &
                "los documentos que aún no han llegado al estado " & Chr(34) & "Enviado" &
                Chr(34) & "." & vbCrLf & vbCrLf & "NOTA: Para ver los documentos que aún no " &
                "han sido enviados siga los siguientes pasos:" & vbCrLf & "1. En la Ventana " &
                "Principal de clic sobre el nodo " & Chr(34) & "EFactura" & Chr(34) &
                " del menú de la ventana." & vbCrLf & "2. En el menú desplegado de clic " &
                "sobre el nodo " & Chr(34) & "Reporte Documentos no registrados" & Chr(34) &
                "; asi se genera el reporte donde se listan los documentos pendientes de " &
                "procesar y su estado actual." & vbCrLf & vbCrLf & "NOTA: Si se quiere revisar " &
                "el estado de cualquier documento electrónico, ejecute los siguientes pasos: " &
                vbCrLf & "1. Abra el documento requerido." & vbCrLf & "2. De clic sobre el nodo " &
                Chr(34) & "EDocumento" & Chr(34) & " del menú de la ventana del documento." &
                vbCrLf & "3. De clic sobre el nodo " & Chr(34) & "Estado en Proveedor eFactura" &
                Chr(34) & "." & vbCrLf & "Con esta acción se abre una ventana que contiene toda " &
                "la información relacionada con el registro del documento en el Proveedor de " &
                "facturación electrónica y en la DIAN."
        Return lstrMens
    End Function

    Private Function FstrProcFactFueraFecha(ByRef astrMensTitulo As String) As String
        astrMensTitulo = "Proceso para las facturas fuera de fecha"
        Dim lstrMens = "1. Estando en la Ventana Principal, de clic sobre el nodo " & Chr(34) &
                "EFactura" & Chr(34) & " del menú de la ventana." & vbCrLf & "2. En el menú " &
                "desplegado, de clic sobre el nodo " & Chr(34) & "Procesar Facturas fuera " &
                "de Fecha" & Chr(34) & ". Asi la aplicación procede a marcar las facturas " &
                "fuera de fecha, o sea, las facturas con fecha anterior al día de hoy que " &
                "no fueron procesadas adecuadamente, como facturas no electrónicas, y son " &
                "anuladas." & vbCrLf & "NOTA: Es necesario que no este pendiente por hacer " &
                "cierre de mes!"

        Return lstrMens
    End Function
#End Region
End Module