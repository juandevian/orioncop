Friend Class ClsNovedadAnticipo
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriNovedadesAnt"
    ' Variables de modulo
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwNovedadAnt">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As clsCBObjetoPan, adrwNovedadAnt As DataRow)
        HobjPadre = aobjPadre
        henuTipoObjeto = enuModoInstanciaObjDef.enuDeColeccion
        HblnEsSuprimible = False
        hblnEsModificable = False
        '
        drwRegistroActual = adrwNovedadAnt
        DtbTablaColeccion = drwRegistroActual.Table
        henuTipoPermiso = EnuPermisosDef.enuHeredado
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
            Return EnuIdClasesPanDef.enuNovedadAnt
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Novedad Anticipo"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjAliasCont_NovAntStr As New ClsAliasCont_NovAntStr(Me)
    Friend ReadOnly Property ObjFechaNovedadAntDtm As New ClsFechaNovedadAntDtm(Me)
    Friend ReadOnly Property ObjIdAnticipo_NovEnt As New ClsIdAnticipo_NovEnt(Me)
    Friend ReadOnly Property ObjIdCarpeta_NovAntShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_NovAntShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdCuentaCr_NovAntStr As New ClsIdCuentaCr_NovAntStr(Me)
    Friend ReadOnly Property ObjIdCuentaDb_NovAntStr As New ClsIdCuentaDb_NovAntStr(Me)
    Friend ReadOnly Property ObjIdDocOrigen_NovAntEnt As New ClsIdDocOrigen_NovAntEnt(Me)
    Friend ReadOnly Property ObjIdNovedadAntShr As New ClsIdNovedadAntShr(Me)
    Friend ReadOnly Property ObjIdTercero_NovAntDbl As New ClsIdTercero_NovAntDbl(Me)
    Friend ReadOnly Property ObjIdTipoDocOrigen_NovAntByt As New ClsIdTipoDocOrigen_NovAntByt(Me)
    Friend ReadOnly Property ObjIdTipoNovedad_NovAntByt As New ClsIdTipoNovedad_NovAntByt(Me)
    Friend ReadOnly Property ObjPrefijoDocOrigen_NovAntStr As New ClsPrefijoDocOrigen_NovAntStr(Me)
    Friend ReadOnly Property ObjValor_NovAntDec As New ClsValor_NovAntDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAnuladoBln)
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjFechaNovedadAntDtm)
                HcolPropiedades.Add(ObjAliasCont_NovAntStr)
                HcolPropiedades.Add(ObjIdAnticipo_NovEnt)
                HcolPropiedades.Add(ObjIdCarpeta_NovAntShr)
                HcolPropiedades.Add(ObjIdCentroUtil_NovAntShr)
                HcolPropiedades.Add(ObjIdCuentaCr_NovAntStr)
                HcolPropiedades.Add(ObjIdCuentaDb_NovAntStr)
                HcolPropiedades.Add(ObjIdDocOrigen_NovAntEnt)
                HcolPropiedades.Add(ObjIdNovedadAntShr)
                HcolPropiedades.Add(ObjIdTercero_NovAntDbl)
                HcolPropiedades.Add(ObjIdTipoDocOrigen_NovAntByt)
                HcolPropiedades.Add(ObjIdTipoNovedad_NovAntByt)
                HcolPropiedades.Add(ObjPrefijoDocOrigen_NovAntStr)
                HcolPropiedades.Add(ObjValor_NovAntDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Public Overrides Function FblnEsAnulable() As Boolean
        Return hblnEsAnulable
    End Function
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        SNumereObj()
        ObjFechaCreacionDtm.ObjValorPro = Date.Now
        ObjValor_NovAntDec.SValide()
        MyBase.SActualice(ablnExigeRequeridos)
    End Sub
    Protected Overrides Function SAnuleEnObj() As Boolean
        Dim lblnAnulado = FblnEsAnulable()
        If lblnAnulado Then
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                ObjAnuladoBln.ObjValorPro = True
                ObjValor_NovAntDec.ObjValorPro = 0
            Else
                Throw New ErrorInesperadoPanLException("Anulando. Estado inesperado del objeto Novedad")
            End If
        End If
        Return lblnAnulado
    End Function
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdNovedadAntShr.ToString
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If enuEstadoActualizacion = enuEstadoObjetoDef.enuCreando Then
            Dim lshrIdNovedadAnt As Short
            Dim lstrFiltro = clsOrionCop.strFiltroUbicacion &
                    " AND " & clsIdAnticipo_NovEnt.sstrNombreCampoBd & " = " &
                    objIdAnticipo_NovEnt.ObjValorPro
            lshrIdNovedadAnt = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                    ClsIdNovedadAntShr.SstrNombreCampoBd, ObjIdNovedadAntShr.EnuTipoValor,
                    lstrFiltro) + 1
            ObjIdNovedadAntShr.ObjValorPro = lshrIdNovedadAnt
        End If
    End Sub
    Friend Function FenuTipoNovReversaRC() As EnuTipoNov
        Dim lenutipoNovRevRC As EnuTipoNov = EnuTipoNov.None
        Select Case ObjIdTipoNovedad_NovAntByt.ObjValorPro
            Case EnuTipoNov.EnuCrAntRec
                lenutipoNovRevRC = EnuTipoNov.EnuRCrAntRec
            Case EnuTipoNov.EnuDbAntDev
                lenutipoNovRevRC = EnuTipoNov.EnuRDbAntDev
            Case EnuTipoNov.EnuDbAntApl
                lenutipoNovRevRC = EnuTipoNov.EnuRDbAntApl
        End Select
        Return lenutipoNovRevRC
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsAliasCont_NovAntStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "AliasCont"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Alias Contable"
        HshrLongitud = 50
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1,
                ShrLongitud, BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsFechaNovedadAntDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaNovedad"
    Private ReadOnly MobjPadre As ClsNovedadAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaNovedadAnticipo"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin = DateSerial(1990, 1, 1)
        Dim ldtmFechaMax = Now
        If ClsOrionCop.BlnProcesoEspecial Then
            ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If Not GblnActualizandoApp Then
                    ldtmFechaMin = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
                    ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                End If
                HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin,
                        ldtmFechaMax, BlnEsRequerido)
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
Friend Class ClsIdAnticipo_NovEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAnticipo"
    Private ReadOnly MobjPadre As ClsNovedadAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdAnticipo_Nov"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue,
                        BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            Dim lobjAnticipo As ClsAnticipo = MobjPadre.ObjPadre
            Dim lentIdNovedad As Integer = lobjAnticipo.ObjIdAnticipoEnt.ObjValorPro
            HblnEsValido = (HobjValorNew = lentIdNovedad)
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
Friend Class ClsIdCuentaCr_NovAntStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCuentaCr"
    Private ReadOnly MobjPadre As ClsNovedadAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdCuentaCr_NovAnt"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        Else
            If MobjPadre.ObjIdTipoNovedad_NovAntByt.ObjValorPro = EnuTipoNov.EnuDbAntApl Then
                HblnEsValido = (HobjValorNew = "*")
            End If
        End If
    End Sub
    Friend ReadOnly Property StrNombreCuenta As String
        Get
            Dim lstrNombreCuenta = String.Empty
            If HblnEsValido Then
                lstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            End If
            Return lstrNombreCuenta
        End Get
    End Property
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
Friend Class ClsIdCuentaDb_NovAntStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCuentaDb"
    Private ReadOnly MobjPadre As ClsNovedadAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdCuentaDb_NovAnt"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        Else
            If MobjPadre.ObjIdTipoNovedad_NovAntByt.ObjValorPro = EnuTipoNov.EnuRDbAntApl Then
                HblnEsValido = (HobjValorNew = "*")
            End If
        End If
    End Sub
    Friend ReadOnly Property StrNombreCuenta As String
        Get
            Dim lstrNombreCuenta = String.Empty
            If HblnEsValido Then
                lstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorPro.ToString)
            End If
            Return lstrNombreCuenta
        End Get
    End Property
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
Friend Class ClsIdDocOrigen_NovAntEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdDocOrigen"
    Private ReadOnly MobjPadre As ClsNovedadAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdDocumentoOrigenAnticipo"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido,
                EnuTipoValor.enuInteger)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                If Not GblnActualizandoApp Then
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
                End If
            End If
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
Friend Class ClsIdNovedadAntShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdNovedad"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdNovedadAnticipo"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Byte.MaxValue,
                BlnEsRequerido, EnuTipoValor)
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
Friend Class ClsIdTercero_NovAntDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCliente"
    Private ReadOnly MobjPadre As ClsNovedadAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTerceroCliente"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC,
                GCDBLMAXTERC, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                Dim lobjLlaveCliente() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
                lobjCliente.SAbra(lobjLlaveCliente)
                HblnEsValido = lobjCliente.BlnExiste
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
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
Friend Class ClsIdTipoDocOrigen_NovAntByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoDocOrigen"
    Private ReadOnly MobjPadre As ClsNovedadAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTipoDocOrigen"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoDocOri.EnuReciboCaja,
                EnuTipoDocOri.EnuNotaAjuste, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                If Not GblnActualizandoApp Then
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
                End If
            End If
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
Friend Class ClsIdTipoNovedad_NovAntByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoNovedad"
    Private ReadOnly MobjPadre As ClsNovedadAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTipoNovedad_Anticipo"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoNov.EnuCrAntRec,
                EnuTipoNov.EnuDbAntApl, BlnEsRequerido)
        If Not HblnEsValido Then
            HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoNov.EnuRCrAntRec,
                    EnuTipoNov.EnuRDbAntApl, BlnEsRequerido)
        End If
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        MobjPadre.ObjIdCuentaCr_NovAntStr.SValide()
        MobjPadre.ObjIdCuentaDb_NovAntStr.SValide()
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
Friend Class ClsPrefijoDocOrigen_NovAntStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoDocOrigen"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PrefijoDocOrigen_Anticipo"
        HshrLongitud = 5
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then
            HobjValorNew = String.Empty
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
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
Friend Class ClsValor_NovAntDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"
    Private ReadOnly MobjPadre As ClsNovedadAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Valor_NovedadAnticipo"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0.00, Decimal.MaxValue,
                BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                HblnEsValido = False
                If HobjValorNew = 0 Then
                    HblnEsValido = (MobjPadre.ObjAnuladoBln.ObjValorPro = True)
                End If
            ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
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
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
#End Region