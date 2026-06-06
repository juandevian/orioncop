Friend Class ClsSector
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriSectores"
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Panorama
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Indica si se instancia como un objeto navegable o como un Objeto único.</param>
    ''' <remarks>Si se instancia como un objeto navegable, se crea un datatable que contiene las columnas de
    ''' la llave con las llaves de todos los objetos y queda a la espera de que se indique que objeto abrir.
    ''' Si se instancia como un objeto único, queda a la espera de recibir el valor de los campos de la llave 
    ''' para abrir dicho objeto. </remarks>
    Public Sub New(aenuModoInstanciaObj As EnuModoInstanciaObjDef)
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = Nothing
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
            hblnEsAnulable = False
            HblnEsSuprimible = False
            lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                                ClsIdSectorShr.sstrNombreCampoBd}
            hcolFiltros.Add(ClsOrionCop.strFiltroUbicacion)
        Else
            hblnEsCreable = False
            hblnEsModificable = False
            HblnEsSuprimible = False
            hblnEsAnulable = False
            henuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        hcolTablas.Add(MCSTRNOMBRETABLA)
        hcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As ClsCentroUtilOriCop, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        henuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        hblnEsAnulable = False
        HblnEsSuprimible = False
        '
        drwRegistroActual = adrwObjeto
        DtbTablaColeccion = drwRegistroActual.Table
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
            Return EnuIdClasesPanDef.enuSector
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Sector"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & ObjNombreSectorStr.ObjValorPro & Chr(34)
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    ''' <summary>
    ''' Devuelve la base de participación del sector en la cuota de administración de acuerdo a la base
    ''' de participación definida el calculo del coeficiente de propiedad definida en los 
    ''' parametros de la Copropiedad
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property ObjDctoProntoPago_SecDbl As New ClsDctoProntoPago_SecDbl(Me)
    Friend ReadOnly Property ObjIdCarpetaSectorShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtilSectorShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdSectorShr As New ClsIdSectorShr(Me)
    Friend ReadOnly Property ObjNombreSectorStr As New ClsNombreSectorStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjDctoProntoPago_SecDbl)
                HcolPropiedades.Add(ObjIdCarpetaSectorShr)
                HcolPropiedades.Add(ObjIdCentroUtilSectorShr)
                HcolPropiedades.Add(ObjIdSectorShr)
                HcolPropiedades.Add(ObjNombreSectorStr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend ReadOnly Property DblBaseParticipacionSector(aenuBasePartic As EnuTipoBaseCalculo) As Double
        Get
            Dim ldblBaseParSec As Double
            If aenuBasePartic = EnuTipoBaseCalculo.EnuUnidad Then
                ldblBaseParSec = FentCantidadPrediosSector()
            Else
                ldblBaseParSec = FdblTotalAreaPrediosSector()
            End If
            Return Math.Round(ldblBaseParSec, 3)
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            gobjPanDat.sControleProcesoObj(True)
            Try
                sNumereObj()
                ObjIdCarpetaSectorShr.ObjValorPro = GshrIdCarpeta
                ObjIdCentroUtilSectorShr.ObjValorPro = GshrIdCentroUtil
                MyBase.sActualice(ablnExigeRequeridos)
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
            MyBase.sActualice(ablnExigeRequeridos)
        End If
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdSectorShr.ToString
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lshrIdSector As Short
            lshrIdSector = ClsPanorama.FobjUltimaIdNumericaObjeto(sstrNombreTabla,
                    ClsIdSectorShr.sstrNombreCampoBd, ObjIdSectorShr.EnuTipoValor,
                    ClsOrionCop.strFiltroUbicacion) + 1
            ObjIdSectorShr.ObjValorPro = lshrIdSector
        End If
    End Sub
    Friend Function FentCantidadPrediosSector() As Integer
        GobjPanDat.SControleProcesoObj(True)
        Dim lshrIdSector As Short = 0
        Dim lentCanPreSec As Integer
        If Not IsNothing(ObjIdSectorShr.ObjValorPro) Then
            lshrIdSector = ObjIdSectorShr.ObjValorPro
        End If
        Dim lstrCamposSelect = {"COUNT(" & ClsIdPredioStr.SstrNombreCampoBd & ")"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdSector_PredioShr.SstrNombreCampoBd &
                " = " & lshrIdSector
        Dim ldtbResu = ClsPanorama.FdtbDataTable(ClsPredio.SstrNombreTabla, lstrCamposSelect, {{"", ""}},
                                                lstrFiltro)
        Dim ldrwResu = ldtbResu.Rows(0)
        lentCanPreSec = ClsPanorama.FobjValorCampo(ldrwResu(0), EnuTipoValor.enuInteger)
        GobjPanDat.SControleProcesoObj(False)
        Return lentCanPreSec
    End Function
    ''' <summary>
    ''' Devuelve el total de las areas del sector teniendo en cuenta el factor de ponderación
    ''' de los predios
    ''' </summary>
    ''' <returns></returns>
    Friend Function FdblTotalAreaPrediosSector() As Double
        GobjPanDat.SControleProcesoObj(True)
        Dim lshrIdSector As Short = 0
        Dim ldblTotalAreaSec As Double
        If Not IsNothing(ObjIdSectorShr.ObjValorPro) Then
            lshrIdSector = ObjIdSectorShr.ObjValorPro
        End If
        Dim lstrCamposSelect
        lstrCamposSelect = {"SUM(" & ClsAreaPredioDec.SstrNombreCampoBd & " * " &
                ClsFactorPonderaCPDbl.SstrNombreCampoBd & ") AS Area"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdSector_PredioShr.SstrNombreCampoBd & " = " & lshrIdSector
        Dim ldtbResu = ClsPanorama.FdtbDataTable(ClsPredio.SstrNombreTabla, lstrCamposSelect, {{"", ""}},
                                                lstrFiltro)
        Dim ldrwResu = ldtbResu.Rows(0)
        ldblTotalAreaSec = ClsPanorama.FobjValorCampo(ldrwResu(0), EnuTipoValor.enuDouble)
        GobjPanDat.SControleProcesoObj(False)
        Return ldblTotalAreaSec
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsDctoProntoPago_SecDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DctoProntoPago"
    Private lenuTipoDsctoPP As EnuTipoDsctoPP = EnuTipoDsctoPP.None
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DctoProntoPago"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        Dim ldblValorMin = 0.0, ldblValorMax = 0.5
        If Not IsNothing(GobjParametros.ObjAnoActual) Then
            HblnEsRequerido = GobjParametros.ObjAnoActual.ObjAplicaDsctoPPBln.ObjValorPro
            If HblnEsRequerido Then
                lenuTipoDsctoPP = GobjParametros.ObjAnoActual.ObjTipoDsctoPPByt.ObjValorPro
                If lenuTipoDsctoPP = EnuTipoDsctoPP.EnuValorFijo Then
                    HenuTipoValor = EnuTipoValor.enuDecimal
                    ldblValorMin = GobjParametros.ObjBaseRedondeoGeneralDbl.ObjValorPro
                    ldblValorMax = Decimal.MaxValue
                ElseIf lenuTipoDsctoPP = EnuTipoDsctoPP.EnuProcentaje Then
                    ldblValorMin = 0.01
                End If
            End If
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, ldblValorMin, ldblValorMax, HblnEsRequerido,
                    HenuTipoValor)
        If Not HblnEsValido Then
            If HblnEsRequerido Then
                HstrMens = "EL Valor ingresado debe estar comprendido entre 1 y 50!"
                SNotifiqueDatInv()
            Else
                HstrMens = "No está parametrizado para Descuento por Pronto Pago!"
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
        Dim lstrValor = String.Empty
        If Not IsNothing(HobjValorPro) AndAlso HblnEsValido Then
            If lenuTipoDsctoPP = EnuTipoDsctoPP.EnuValorFijo Then
                lstrValor = Format(HobjValorPro, "c")
            ElseIf lenuTipoDsctoPP = EnuTipoDsctoPP.EnuProcentaje Then
                lstrValor = Format(HobjValorPro, "p")
            End If
        End If
        Return lstrValor
    End Function
End Class
Friend Class ClsIdSectorShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdSector"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdSector"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HstrOrdenIndice = "ASC"
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Byte.MaxValue, BlnEsRequerido, EnuTipoValor)
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
Friend Class ClsNombreSectorStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Nombre"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "NombreSector"
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HshrLongitud = 40
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, HshrLongitud, HblnEsRequerido)
        If HblnEsValido Then
            HobjValorNew = HobjValorNew.ToString.ToUpper
        Else
            If Not String.IsNullOrEmpty(HobjValorNew) Then
                HstrMens = "EL Nombre del Sector debe tener entre 3 y 40 Caracteres!"
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
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
#End Region
