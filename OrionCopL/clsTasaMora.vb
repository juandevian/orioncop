Friend Class ClsTasaMora
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriTasasMora"
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
    Public Sub New(aenuModoInstanciaObj As enuModoInstanciaObjDef)
        If aenuModoInstanciaObj = enuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = Nothing
        If aenuModoInstanciaObj = enuModoInstanciaObjDef.enuNavegable Then
            HcolFiltros.Add(ClsOrionCop.StrFiltroUbicacion)
            lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                    clsOrdinalTasaMoraEnt.sstrNombreCampoBd}
        Else
            hblnEsCreable = False
            hblnEsModificable = False
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        HblnEsAnulable = False
        HblnEsSuprimible = False
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsCentroUtilOriCop, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        HblnEsSuprimible = False
        '
        DrwRegistroActual = adrwObjeto
        DtbTablaColeccion = DrwRegistroActual.Table
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
            Return EnuIdClasesPanDef.enuInteresMora
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Tasa de Mora"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjFechaDesdeTasaMoraDtm As New ClsFechaDesdeTasaMoraDtm(Me)
    Friend ReadOnly Property ObjFechaHastaTasaMoraDtm As New ClsFechaHastaTasaMoraDtm(Me)
    Friend ReadOnly Property ObjIdCarpetaTasaMoraShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtilTasaMoraShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjOrdinalTasaMoraEnt As New ClsOrdinalTasaMoraEnt(Me)
    Friend ReadOnly Property ObjTasaMoraDbl As New ClsTasaMoraDbl(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjFechaDesdeTasaMoraDtm)
                HcolPropiedades.Add(ObjFechaHastaTasaMoraDtm)
                HcolPropiedades.Add(ObjIdCarpetaTasaMoraShr)
                HcolPropiedades.Add(ObjIdCentroUtilTasaMoraShr)
                HcolPropiedades.Add(ObjOrdinalTasaMoraEnt)
                HcolPropiedades.Add(ObjTasaMoraDbl)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        If enuEstadoActualizacion = enuEstadoObjetoDef.enuCreando Then
            gobjPanDat.sControleProcesoObj(True)
            Try
                sNumereObj()
                ObjIdCarpetaTasaMoraShr.ObjValorPro = GshrIdCarpeta
                ObjIdCentroUtilTasaMoraShr.ObjValorPro = GshrIdCentroUtil
                MyBase.sActualice(ablnExigeRequeridos)
                sActualiceFechaHasta()
            Catch ex As PanLException
                Throw
            Catch ex As PanDatException
                Throw
            Catch ex As ArgumentNullException
                Throw
            Catch ex As Exception
                Throw
            Finally
                gobjPanDat.sControleProcesoObj(False)
            End Try
        Else
            Dim lblnCambiarFechaHasta = ObjFechaDesdeTasaMoraDtm.BlnCambio
            MyBase.sActualice(ablnExigeRequeridos)
            If lblnCambiarFechaHasta Then
                sActualiceFechaHasta()
            End If
        End If
    End Sub
    Protected Overrides Sub SCreeObj(aobjValorLlave() As Object)
        MyBase.SCreeObj(aobjValorLlave)
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjOrdinalTasaMoraEnt.ToString
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If enuEstadoActualizacion = enuEstadoObjetoDef.enuCreando Then
            Dim lentOrdinal As Integer
            Dim lstrFiltro = objIdCarpetaTasaMoraShr.strNombreCampoBD & " = " & gshrIdCarpeta &
                    " AND " & objIdCentroUtilTasaMoraShr.strNombreCampoBD & " = " & gshrIdCentroUtil
            lentOrdinal = clsPanorama.fobjUltimaIdNumericaObjeto(sstrNombreTabla,
                    objOrdinalTasaMoraEnt.strNombreCampoBD, objOrdinalTasaMoraEnt.enuTipoValor, lstrFiltro) + 1
            objOrdinalTasaMoraEnt.objValorPro = lentOrdinal
        End If
    End Sub
    Private Sub SActualiceFechaHasta()
        If DtbTablaNavegacion.Rows.Count > 1 Then
            SVayaAlAnterior()
            SModifique()
            ObjFechaHastaTasaMoraDtm.ObjValorPro = FdtmFechaDesdeUltima.AddDays(-1)
            SActualice(True)
            SVayaAlSiguiente()
        End If
    End Sub
    Friend Function FdtmFechaDesdeUltima() As Date
        Dim ldtmFechaDesdeUltima As Date
        Dim ldtbTasaMora = GobjParametros.FdtbTasasMora()
        If ldtbTasaMora.Rows.Count > 0 Then
            Dim ldrwUltima As DataRow = ldtbTasaMora.Rows(ldtbTasaMora.Rows.Count - 1)
            ldtmFechaDesdeUltima = ClsPanorama.FobjValorCampo(ldrwUltima(ObjFechaDesdeTasaMoraDtm.StrNombreCampoBD),
                        EnuTipoValor.enuDate)
        Else
            ldtmFechaDesdeUltima = GCDTMFECHANULA
        End If
        Return ldtmFechaDesdeUltima
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsFechaDesdeTasaMoraDtm
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaDesde"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FechaDesde"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsTasaMora = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, GCDTMFECHANULA, Today, HblnEsRequerido)
        If HblnEsValido Then
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                Dim ldtmFechaMinima As Date = lobjPadre.FdtmFechaDesdeUltima.AddDays(1)
                HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMinima, Today, HblnEsRequerido)
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsTasaMora = ObjPadre
        If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso HobjValorNew = GCDTMFECHANULA Then
            HobjValorPro = Today
        End If
        lobjPadre.ObjFechaHastaTasaMoraDtm.SValide()
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return GCDTMFECHANULA
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsFechaHastaTasaMoraDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaHasta"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FechaHasta"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        HobjValorNew = Date.Today
        HobjValorPro = HobjValorNew
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsTasaMora = ObjPadre
        If Not lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HblnEsRequerido = Not lobjPadre.FblnEstaVacioOrigenDatos OrElse
                    Not lobjPadre.FblnEsElUltimoRegistro
        Else
            HblnEsRequerido = False
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, GCDTMFECHANULA, Today, HblnEsRequerido)
        If HblnEsValido Then
            Dim ldtmFechaMinima As Date = lobjPadre.ObjFechaDesdeTasaMoraDtm.ObjValorPro
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMinima, Today, HblnEsRequerido)
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return GCDTMFECHANULA
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsOrdinalTasaMoraEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Ordinal"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Ordinal"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
            HstrMens = String.Empty
            If HblnEsValido Then
                If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    If Not BlnLeyendoOrigen Then
                        If HblnEsValido Then
                            Dim lobjLlavePrincipal() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                            If Not ObjPadre.FblnExisteLlave(lobjLlavePrincipal) Then
                                HstrMens = "El Ordinal de la Tasa de Mora ingresado no existe!"
                                HblnEsValido = False
                            End If
                        End If
                    End If
                ElseIf ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                    HblnEsValido = (HobjValorOriginal = HobjValorNew)
                    If Not HblnEsValido Then
                        HstrMens = "No es permitido cambiar la identidad a objeto alguno!"
                    End If
                End If
            Else
                HstrMens = "El Ordinal de ls tasa de Mora ingresada no es válido!"
            End If
            If Not String.IsNullOrEmpty(HstrMens) Then
                SNotifiqueDatInv()
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
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsTasaMoraDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "TasaMora"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TasaMora"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
        HblnEsRequerido = True
    End Sub
    Private Sub ClsTasaMoraDbl_evnPreSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPreSetValor
        Dim lobjPadre As ClsTasaMora = ObjPadre
        Dim lstrValor = String.Empty
        If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If String.IsNullOrEmpty(HobjValorNew.ToString) Then
                lstrValor = "0mv"
            Else
                If IsNumeric(HobjValorNew) Then
                    lstrValor = HobjValorNew.ToString & "mv"
                End If
            End If
            HobjValorNew = ClsOrionCop.FdblTraduceATasaEfectivaAnual(lstrValor,
                            EnuTipoInteres.EnuInteresSimple)
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, 1, HblnEsRequerido,
                EnuTipoValor.enuDouble)
        If HblnEsValido Then
            HobjValorNew = Math.Round(HobjValorNew, 6)
        End If
    End Sub
    Friend ReadOnly Property DblTasaMensual As Single
        Get
            Dim ldblTasaMensual As Double = Math.Round(HobjValorPro / 12, 6)
            Return ldblTasaMensual
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