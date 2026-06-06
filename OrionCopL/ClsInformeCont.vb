Friend Class ClsInformeCont
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriInformeCont"
    ' Variables de propiedad
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Panorama.
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Indica si se instancia como un objeto navegable o como un Objeto único.</param>
    ''' <remarks>Si se instancia como un objeto navegable, se crea un datatable que contiene las columnas de
    ''' la llave con las llaves de todos los objetos y queda a la espera de que se indique que objeto abrir.
    ''' Si se instancia como un objeto único, queda a la espera de recibir el valor de los campos de la llave 
    ''' para abrir dicho objeto. </remarks>
    Public Sub New(aenuModoInstanciaObj As EnuModoInstanciaObjDef)
        If aenuModoInstanciaObj <> EnuModoInstanciaObjDef.enuDeColeccion Then
            HobjPadre = Nothing
            Dim lstrCamposSelect As String()
            If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
                HblnEsAnulable = False
                HblnEsSuprimible = False
                HcolFiltros.Add(ClsOrionCop.StrFiltroUbicacion)
                lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                    ClsIdInformeContEnt.SstrNombreCampoBd}
            Else
                HblnEsCreable = False
                HblnEsModificable = False
                HblnEsSuprimible = False
                HblnEsAnulable = False
                HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
                lstrCamposSelect = {"*"}
            End If
            HcolTablas.Add(MCSTRNOMBRETABLA)
            HcolCamposSelect.Add(lstrCamposSelect)
        Else
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
    End Sub
#End Region
#Region "Propiedades"
#Region "Propiedades indentificadoras"
    Protected Overrides ReadOnly Property HstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Protected Overrides ReadOnly Property HenuIdClase As EnuIdClasesPanDef
        Get
            Return EnuIdClasesPanDef.enuInformeCont
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Infrome Contingencia"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & "Nro. " & ObjIdInformeContEnt.ToString & Chr(34)
        End Get
    End Property
    Friend Const StrNombreInfContingencia As String = "Informe de Contingencia"
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjComentario_InfContStr As New ClsComentario_InfContStr(Me)
    Friend ReadOnly Property ObjFechaFinContDtm As New ClsFechaFinContDtm(Me)
    Friend ReadOnly Property ObjFechaInicioContDtm As New ClsFechaInicioContDtm(Me)
    Friend ReadOnly Property ObjFechaRadicoDtm As New ClsFechaRadicoDtm(Me)
    Friend ReadOnly Property ObjIdCarpetaInfConShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtilInfConShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdInformeContEnt As New ClsIdInformeContEnt(Me)
    Friend ReadOnly Property ObjIdFactContFinEnt As New ClsIdFactContFinEnt(Me)
    Friend ReadOnly Property ObjIdFactContIniEnt As New ClsIdFactContIniEnt(Me)
    Friend ReadOnly Property ObjIdUsuario_InfStr As New ClsIdUsuarioStr(Me)
    Friend ReadOnly Property ObjPrefFactContStr As New ClsPrefFactContStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjOrigenInstanciaStr)
                HcolPropiedades.Add(ObjComentario_InfContStr)
                HcolPropiedades.Add(ObjFechaInicioContDtm)
                HcolPropiedades.Add(ObjFechaFinContDtm)
                HcolPropiedades.Add(ObjFechaRadicoDtm)
                HcolPropiedades.Add(ObjIdCarpetaInfConShr)
                HcolPropiedades.Add(ObjIdCentroUtilInfConShr)
                HcolPropiedades.Add(ObjIdInformeContEnt)
                HcolPropiedades.Add(ObjIdFactContIniEnt)
                HcolPropiedades.Add(ObjIdFactContFinEnt)
                HcolPropiedades.Add(ObjIdUsuario_InfStr)
                HcolPropiedades.Add(ObjPrefFactContStr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SInicialiceObj()
        ObjIdCarpetaInfConShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtilInfConShr.ObjValorPro = GshrIdCentroUtil
        ObjOrigenInstanciaStr.ObjValorPro = GstrOrigenActual
        ObjIdUsuario_InfStr.ObjValorPro = GstrIdUsuario
        ObjAnuladoBln.ObjValorPro = False
        ObjIdUsuarioAnuloStr.ObjValorPro = String.Empty
        ObjOrigenInstanciaAnuloStr.ObjValorPro = String.Empty
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        GobjPanDat.SControleProcesoObj(True)
        Try
            GobjPanDat.SInicialiceTransaccion()
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                SNumereObj()
                ObjFechaCreacionDtm.ObjValorPro = Date.Now
                MyBase.SActualice(ablnExigeRequeridos)
            Else
            End If
            MyBase.SActualice(ablnExigeRequeridos)
            lblnNoHayError = True
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lentIdInfCon As Integer
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion
            lentIdInfCon = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                    ClsIdInformeContEnt.SstrNombreCampoBd, ObjIdInformeContEnt.EnuTipoValor,
                    lstrFiltro)
            lentIdInfCon += 1
            ObjIdInformeContEnt.ObjValorPro = lentIdInfCon
        End If
    End Sub
#End Region
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsComentario_InfContStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Comentario"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Comentario"
        HshrLongitud = 500
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return HobjValorPro
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsFechaAnulacion_InfContDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaAnulacion"
    Private ReadOnly MobjPadre As ClsInformeCont = Nothing
    Friend Sub New(aobjPadre As ClsInformeCont)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha Anulacion Informe cont."
        HenuTipoValor = EnuTipoValor.enuDateTime
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = True
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            Dim ldtmFechaMin = GCDTMFECHANULA
            Dim ldtmFechaMax = GCDTMFECHANULA
            HblnEsRequerido = MobjPadre.ObjAnuladoBln.ObjValorPro
            If HblnEsRequerido Then
                ldtmFechaMin = Now.AddHours(-Now.Hour)
                ldtmFechaMax = Now
            End If
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                                BlnEsRequerido)
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, GCSTRFMTFECHASIMPLE)
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class
Friend Class ClsFechaFinContDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaFinCont"
    Private ReadOnly MobjPadre As ClsInformeCont = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha Fin Contingencia"
        HenuTipoValor = EnuTipoValor.enuDateTime
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = Date.Now
        HobjValorPro = HobjValorNew
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = True
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            Dim ldtmFechaMin As DateTime = MobjPadre.ObjFechaInicioContDtm.ObjValorPro
            Dim ldtmFechaMax As DateTime = Date.Now
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                BlnEsRequerido)
            If Not HblnEsValido Then
                HstrMens = "La Fecha del fin de la Contingencia no es valida!"
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Friend Property EntHoraFecFin As Integer
        Set(value As Integer)
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                Dim ldtmFec As DateTime = HobjValorNew
                Dim lentHora = ldtmFec.Hour
                Dim lentHorasSuman As Integer = value - lentHora
                ldtmFec = ldtmFec.AddHours(lentHorasSuman)
                ObjValorPro = ldtmFec
            End If
        End Set
        Get
            If HblnEsValido Then
                Dim ldtmFecha As Date = HobjValorNew
                Return ldtmFecha.Hour
            Else
                Return 0
            End If
        End Get
    End Property
    Friend Property EntMinFecFin As Integer
        Set(value As Integer)
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                Dim ldtmFec As DateTime = HobjValorNew
                Dim lentMin = ldtmFec.Minute
                Dim lentMinSuman As Integer = value - lentMin
                ldtmFec = ldtmFec.AddMinutes(lentMinSuman)
                ObjValorPro = ldtmFec
            End If
        End Set
        Get
            If HblnEsValido Then
                Dim ldtmFecha As Date = HobjValorNew
                Return ldtmFecha.Minute
            Else
                Return 0
            End If
        End Get
    End Property
    Private Sub EPosSetValor() Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando AndAlso BlnEsValido Then
            If HobjValorPro <> GCDTMFECHANULA Then
                MobjPadre.ObjFechaInicioContDtm.SValide()
                MobjPadre.ObjFechaRadicoDtm.SValide()
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, GCSTRFMTFECHASIMPLE)
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class
Friend Class ClsFechaInicioContDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaInicioCont"
    Private ReadOnly MobjPadre As ClsInformeCont = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha Inicio Cont"
        HenuTipoValor = EnuTipoValor.enuDateTime
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = Date.Now
        HobjValorPro = HobjValorNew
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = True
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            Dim ldtmFechaMin As DateTime = FdtmFechaFinConAnt()
            Dim ldtmFechaMax As DateTime = Date.Now
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                BlnEsRequerido)
            If Not HblnEsValido Then
                HstrMens = "La Fecha de inicio de la Contingencia no es valida!"
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Friend Property EntHoraFecIni As Integer
        Set(value As Integer)
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                Dim ldtmFec As DateTime = HobjValorNew
                Dim lentHora = ldtmFec.Hour
                Dim lentHorasSuman As Integer = value - lentHora
                ldtmFec = ldtmFec.AddHours(lentHorasSuman)
                ObjValorPro = ldtmFec
            End If
        End Set
        Get
            If HblnEsValido Then
                Dim ldtmFecha As Date = HobjValorNew
                Return ldtmFecha.Hour
            Else
                Return 0
            End If
        End Get
    End Property
    Friend Property EntMinFecIni As Integer
        Set(value As Integer)
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                Dim ldtmFec As DateTime = HobjValorNew
                Dim lentMin = ldtmFec.Minute
                Dim lentMinSuman As Integer = value - lentMin
                ldtmFec = ldtmFec.AddMinutes(lentMinSuman)
                ObjValorPro = ldtmFec
            End If
        End Set
        Get
            If HblnEsValido Then
                Dim ldtmFecha As Date = HobjValorNew
                Return ldtmFecha.Minute
            Else
                Return 0
            End If
        End Get
    End Property
    Private Sub EPosSetValor() Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando AndAlso BlnEsValido Then
            If HobjValorPro <> GCDTMFECHANULA Then
                MobjPadre.ObjFechaFinContDtm.SValide()
                MobjPadre.ObjFechaRadicoDtm.SValide()
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Private Function FdtmFechaFinConAnt() As DateTime
        Dim lstrCamSel As String() = {"MAX(" & ClsFechaFinContDtm.SstrNombreCampoBd & ")"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion
        Dim ldtb = ClsPanorama.FdtbDataTable(ClsInformeCont.SstrNombreTabla, lstrCamSel, Nothing,
                lstrFiltro)
        Dim ldtmFecFin As DateTime = ClsPanorama.FobjValorCampo(ldtb.Rows(0)(0), EnuTipoValor.enuDateTime)
        Return ldtmFecFin
    End Function
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, GCSTRFMTFECHASIMPLE)
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class
Friend Class ClsFechaRadicoDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaRadicacion"
    Private ReadOnly MobjPadre As ClsInformeCont = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha Radicacion"
        HenuTipoValor = EnuTipoValor.enuDateTime
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = Date.Now
        HobjValorPro = HobjValorNew
    End Sub
    Friend Property EntHoraFecRad As Integer
        Set(value As Integer)
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                Dim ldtmFec As DateTime = HobjValorNew
                Dim lentHora = ldtmFec.Hour
                Dim lentHorasSuman As Integer = value - lentHora
                ldtmFec = ldtmFec.AddHours(lentHorasSuman)
                ObjValorPro = ldtmFec
            End If
        End Set
        Get
            If HblnEsValido Then
                Dim ldtmFecha As Date = HobjValorNew
                Return ldtmFecha.Hour
            Else
                Return 0
            End If
        End Get
    End Property
    Friend Property EntMinFecRad As Integer
        Set(value As Integer)
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                Dim ldtmFec As DateTime = HobjValorNew
                Dim lentMin = ldtmFec.Minute
                Dim lentMinSuman As Integer = value - lentMin
                ldtmFec = ldtmFec.AddMinutes(lentMinSuman)
                ObjValorPro = ldtmFec
            End If
        End Set
        Get
            If HblnEsValido Then
                Dim ldtmFecha As Date = HobjValorNew
                Return ldtmFecha.Minute
            Else
                Return 0
            End If
        End Get
    End Property
    Public Overrides Sub SValide()
        HblnEsValido = True
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            Dim ldtmFechaMin As DateTime = MobjPadre.ObjFechaFinContDtm.ObjValorPro
            Dim ldtmFechaMax As DateTime = Date.Now
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                BlnEsRequerido)
            If Not HblnEsValido Then
                HstrMens = "La Fecha de radicación debe estar entre la Fecha " &
                        " de fin de Contingencia y este Momento!"
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, GCSTRFMTFECHASIMPLE)
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class
Friend Class ClsIdInformeContEnt
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsInformeCont = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdInformeCon"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id Documento Contingencia"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HstrOrdenIndice = "ASC"
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido,
                EnuTipoValor)
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            If Not BlnLeyendoOrigen Then
                HstrMens = String.Empty
                If HblnEsValido Then
                    Dim lobjLlavePrincipal() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                    If Not MobjPadre.FblnExisteLlave(lobjLlavePrincipal) Then
                        HblnEsValido = False
                        HstrMens = "La Id. del Documento ingresada no existe!"
                    Else
                        MobjPadre.SAbra(lobjLlavePrincipal)
                    End If
                Else
                    If Not (ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso
                            (HobjValorNew = String.Empty OrElse IsNothing(HobjValorNew))) Then
                        HstrMens = "La Id. del Documento ingresada no es válida!"
                    End If
                End If
                If Not String.IsNullOrEmpty(HstrMens) Then
                    SNotifiqueDatInv()
                End If
            End If
        Else
            HblnEsValido = True
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdFactContFinEnt
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsInformeCont = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdFacturaFin"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id Factura Cont. inicial"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng As Object = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido,
                       EnuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                Dim lentIdFacMin = GobjParametros.ObjRangoFraConIniEnt.ObjValorPro
                Dim lentIdFacMax = GobjParametros.ObjRangoFraConFinEnt.ObjValorPro
                HblnEsValido = (HobjValorNew >= lentIdFacMin AndAlso HobjValorNew <= lentIdFacMax)
                If Not HblnEsValido Then
                    HstrMens = "La Id. de la Factura ingresada está por fuera del Rango " &
                            "de Nmueración autorizado en la Resolución!"
                End If
                If HblnEsValido Then
                    HblnEsValido = HobjValorNew > MobjPadre.ObjIdFactContIniEnt.ObjValorPro
                    If Not HblnEsValido Then
                        HstrMens = "La Id. de la Factura ingresada debe ser mayor " &
                            "al Número de la Factura inicial!"
                    End If
                End If
            End If
        Else
            HstrMens = "La Id. de la Factura ingresada, '" & lobjValorIng.ToString &
                    ",  no es válida!"
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Private Sub EPosSetValor() Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando AndAlso BlnEsValido Then
            MobjPadre.ObjIdFactContIniEnt.SValide()
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdFactContIniEnt
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsInformeCont = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdFacturaIni"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id Factura Cont. final"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng As Object = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido,
                       EnuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                Dim lentIdFacMin = GobjParametros.ObjRangoFraConIniEnt.ObjValorPro
                Dim lentIdFacMax = GobjParametros.ObjRangoFraConFinEnt.ObjValorPro
                HblnEsValido = (HobjValorNew >= lentIdFacMin AndAlso HobjValorNew <= lentIdFacMax)
                If Not HblnEsValido Then
                    HstrMens = "La Id. de la Factura ingresada está por fuera del Rango " &
                            "de Nmueración autorizado en la Resolución!"
                End If
                If HblnEsValido Then
                    HblnEsValido = HobjValorNew < MobjPadre.ObjIdFactContFinEnt.ObjValorPro
                    If Not HblnEsValido Then
                        HstrMens = "La Id. de la Factura ingresada debe ser menor " &
                            "al Número de la Factura final!"
                    End If
                End If
            End If
        Else
            HstrMens = "La Id. de la Factura ingresada, '" & lobjValorIng.ToString &
                    ",  no es válida!"
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsPrefFactContStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoFacCon"
    Private ReadOnly MobjPadre As ClsInformeCont = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Prefijo Factura Cont."
        HshrLongitud = 5
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then
            HobjValorNew = String.Empty
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                Dim lstrPrefResCon As String = GobjParametros.ObjPrefijoFactContStr.ObjValorPro
                HblnEsValido = (lstrPrefResCon = HobjValorNew)
                If Not HblnEsValido Then
                    HstrMens = "El Prefijo ingresado es diferente al de la Resolución!"
                End If
            End If
        Else
            HstrMens = "Las Facturas deben tener un Prefijo!"
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
#End Region
