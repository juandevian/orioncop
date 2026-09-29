''' <summary>
''' Certificado de Interés Bancario Corriente (IBC) descargado de la API y guardado localmente por centro
''' de utilidad. Tabla OriIbcCertificados.
''' </summary>
Friend Class ClsIbcCertificado
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriIbcCertificados"
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Panorama.
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Indica si se instancia como un objeto navegable o como un Objeto único.</param>
    Public Sub New(aenuModoInstanciaObj As enuModoInstanciaObjDef)
        If aenuModoInstanciaObj = enuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = Nothing
        If aenuModoInstanciaObj = enuModoInstanciaObjDef.enuNavegable Then
            HcolFiltros.Add(ClsOrionCop.StrFiltroUbicacion)
            lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil, ClsIdfileIbcStr.SstrNombreCampoBd}
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
            Return EnuIdClasesPanDef.EnuIbcCertificado
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Certificado IBC"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjFechaCertificadoIbcDtm As New ClsFechaCertificadoIbcDtm(Me)
    Friend ReadOnly Property ObjFechaDescargaIbcDtm As New ClsFechaDescargaIbcDtm(Me)
    Friend ReadOnly Property ObjFechaDesdeIbcDtm As New ClsFechaDesdeIbcDtm(Me)
    Friend ReadOnly Property ObjFechaHastaIbcDtm As New ClsFechaHastaIbcDtm(Me)
    Friend ReadOnly Property ObjIbcDbl As New ClsIbcDbl(Me)
    Friend ReadOnly Property ObjIdCarpetaIbcShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtilIbcShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdfileIbcStr As New ClsIdfileIbcStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjFechaCertificadoIbcDtm)
                HcolPropiedades.Add(ObjFechaDescargaIbcDtm)
                HcolPropiedades.Add(ObjFechaDesdeIbcDtm)
                HcolPropiedades.Add(ObjFechaHastaIbcDtm)
                HcolPropiedades.Add(ObjIbcDbl)
                HcolPropiedades.Add(ObjIdCarpetaIbcShr)
                HcolPropiedades.Add(ObjIdCentroUtilIbcShr)
                HcolPropiedades.Add(ObjIdfileIbcStr)
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
                ObjIdCarpetaIbcShr.ObjValorPro = GshrIdCarpeta
                ObjIdCentroUtilIbcShr.ObjValorPro = GshrIdCentroUtil
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
    Protected Overrides Sub SCreeObj(aobjValorLlave() As Object)
        MyBase.SCreeObj(aobjValorLlave)
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdfileIbcStr.ToString
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    ''' <summary>Crea y guarda un certificado nuevo en la BD (el objeto debe ser navegable y estar consultando).</summary>
    Friend Sub SGuardeCertificado(astcCertificado As StcIbcCertificado)
        SCreeObj(Nothing)
        ' SCreeObj no hace nada (sin error) si el usuario no puede crear: no dejar el certificado sin guardar en silencio
        If EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuCreando Then
            Throw New ErrorInesperadoPanLException("No se pudo guardar el certificado IBC " & astcCertificado.StrIdfile &
                    ": el usuario no tiene permiso para crear certificados IBC.")
        End If
        ObjIdfileIbcStr.ObjValorPro = astcCertificado.StrIdfile
        ObjFechaCertificadoIbcDtm.ObjValorPro = astcCertificado.DtmFechaCertificado
        ObjFechaDesdeIbcDtm.ObjValorPro = astcCertificado.DtmFechaDesde
        ObjFechaHastaIbcDtm.ObjValorPro = astcCertificado.DtmFechaHasta
        ObjIbcDbl.ObjValorPro = astcCertificado.DblIbc
        ObjFechaDescargaIbcDtm.ObjValorPro = Date.Today
        SActualice(True)
        SNormaliceEstado(True)
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsIdfileIbcStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Idfile"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Idfile"
        HshrLongitud = 20
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
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
Friend Class ClsFechaCertificadoIbcDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaCertificado"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FechaCertificado"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, DateSerial(2000, 1, 1),
                DateSerial(2100, 12, 31), HblnEsRequerido)
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
Friend Class ClsFechaDescargaIbcDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaDescarga"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FechaDescarga"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, DateSerial(2000, 1, 1),
                DateSerial(2100, 12, 31), HblnEsRequerido)
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
Friend Class ClsFechaDesdeIbcDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaDesde"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FechaDesde"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, DateSerial(2000, 1, 1),
                DateSerial(2100, 12, 31), HblnEsRequerido)
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
Friend Class ClsFechaHastaIbcDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaHasta"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FechaHasta"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, DateSerial(2000, 1, 1),
                DateSerial(2100, 12, 31), HblnEsRequerido)
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
Friend Class ClsIbcDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Ibc"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Ibc"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
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
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
#End Region
#Region "Almacén de certificados"
''' <summary>Almacén local (BD del centro de utilidad) de los certificados IBC.</summary>
Friend Class ClsIbcAlmacenBd
    Implements IIbcAlmacen

    Friend Function FlstCertificados() As List(Of StcIbcCertificado) Implements IIbcAlmacen.FlstCertificados
        Dim llstResultado As New List(Of StcIbcCertificado)
        Dim lstrCampos As String() = {ClsIdfileIbcStr.SstrNombreCampoBd, ClsFechaCertificadoIbcDtm.SstrNombreCampoBd,
                ClsFechaDesdeIbcDtm.SstrNombreCampoBd, ClsFechaHastaIbcDtm.SstrNombreCampoBd,
                ClsIbcDbl.SstrNombreCampoBd}
        Dim lstrIndice(,) As String = {{ClsFechaDesdeIbcDtm.SstrNombreCampoBd, "ASC"}}
        Dim ldtbCert = ClsPanorama.FdtbDataTable(ClsIbcCertificado.SstrNombreTabla, lstrCampos, lstrIndice,
                ClsOrionCop.StrFiltroUbicacion)
        For Each ldrwCert As DataRow In ldtbCert.Rows
            Dim lstcCert As New StcIbcCertificado
            lstcCert.StrIdfile = ClsPanorama.FobjValorCampo(ldrwCert(ClsIdfileIbcStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString).ToString
            lstcCert.DtmFechaCertificado = ClsPanorama.FobjValorCampo(
                    ldrwCert(ClsFechaCertificadoIbcDtm.SstrNombreCampoBd), EnuTipoValor.enuDate)
            lstcCert.DtmFechaDesde = ClsPanorama.FobjValorCampo(
                    ldrwCert(ClsFechaDesdeIbcDtm.SstrNombreCampoBd), EnuTipoValor.enuDate)
            lstcCert.DtmFechaHasta = ClsPanorama.FobjValorCampo(
                    ldrwCert(ClsFechaHastaIbcDtm.SstrNombreCampoBd), EnuTipoValor.enuDate)
            lstcCert.DblIbc = ClsPanorama.FobjValorCampo(ldrwCert(ClsIbcDbl.SstrNombreCampoBd),
                    EnuTipoValor.enuDouble)
            llstResultado.Add(lstcCert)
        Next
        Return llstResultado
    End Function

    ''' <summary>Guarda solo los certificados que aún no existen localmente (por Idfile).</summary>
    Friend Sub SGuarde(alstCertificados As IEnumerable(Of StcIbcCertificado)) Implements IIbcAlmacen.SGuarde
        Dim lcolExistentes As New HashSet(Of String)(FlstCertificados().Select(Function(x) x.StrIdfile))
        Dim lobjCert As New ClsIbcCertificado(EnuModoInstanciaObjDef.enuNavegable)
        For Each lstcCert As StcIbcCertificado In alstCertificados
            If Not lcolExistentes.Contains(lstcCert.StrIdfile) Then
                lobjCert.SGuardeCertificado(lstcCert)
                lcolExistentes.Add(lstcCert.StrIdfile)
            End If
        Next
    End Sub
End Class
#End Region
