Public Class ClsServicioEstadoCta
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "ServiciosEstadoCta"
    ' Otras variables de modulo
    Private ReadOnly MobjPadre As ClsEstadoCuenta = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte 
    ''' de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta 
    ''' instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades 
    ''' del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsEstadoCuenta, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        MobjPadre = aobjPadre
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
            Return EnuIdClasesPanDef.enuServicioEstadoCta
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Servicio Estado Cuenta"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdCarpeta_SerEstShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_SerEstShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdEstadoCta_SerEnt As New ClsIdEstadoCta_SerEstEnt(Me)
    Friend ReadOnly Property ObjIdServicioEstado_SerEstShr As New ClsIdServicioEstado_SerEstShr(Me)
    Friend ReadOnly Property ObjDeudaCapital_SerEstDec As New ClsDeudaCapital_SerEstDec(Me)
    Friend ReadOnly Property ObjDeudaIntMora_SerEstDec As New ClsDeudaIntMora_SerEstDec(Me)
    Friend ReadOnly Property ObjNombre_SerEstStr As New ClsNombre_SerEstStr(Me)
    Friend ReadOnly Property ObjServicio_EstCtaStr As New ClsServicio_EstCtaStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjIdCarpeta_SerEstShr)
                HcolPropiedades.Add(ObjIdCentroUtil_SerEstShr)
                HcolPropiedades.Add(ObjIdEstadoCta_SerEnt)
                HcolPropiedades.Add(ObjIdServicioEstado_SerEstShr)
                HcolPropiedades.Add(ObjDeudaCapital_SerEstDec)
                HcolPropiedades.Add(ObjDeudaIntMora_SerEstDec)
                HcolPropiedades.Add(ObjNombre_SerEstStr)
                HcolPropiedades.Add(ObjServicio_EstCtaStr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    '
#End Region
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsDeudaCapital_SerEstDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DeudaCapital"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DeudaCapital"
        HenuTipoValor = EnuTipoValorDef.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
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
Friend Class ClsDeudaIntMora_SerEstDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DeudaIntMora"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Deuda Int. Mora"
        HenuTipoValor = EnuTipoValorDef.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
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
Friend Class ClsIdEstadoCta_SerEstEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdEstadoCuenta"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdEstadoCuenta"
        HenuTipoValor = EnuTipoValorDef.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1,
                Short.MaxValue, BlnEsRequerido)
        If Not BlnLeyendoOrigen Then
            If HblnEsValido Then
                Dim lobjPadre As ClsServicioEstadoCta = ObjPadre
                Dim lobjAbuelo As ClsEstadoCuenta = lobjPadre.ObjPadre
                HblnEsValido = HobjValorNew = lobjAbuelo.ObjIdEstadoCuentaEnt.ObjValorPro
            End If
            If Not HblnEsValido Then
                HstrMens = "La Id. del Estado ingresado no es valida!"
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
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdServicioEstado_SerEstShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdServicioEstado"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Id Servicio Estado"
        HenuTipoValor = EnuTipoValorDef.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsSectorModulo = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                    HblnEsRequerido, EnuTipoValorDef.enuShort)
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
Friend Class ClsNombre_SerEstStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "NombreServicio"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Nombre Servicio"
        HshrLongitud = 50
        HenuTipoValor = EnuTipoValorDef.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        If HobjValorNew.GetType.Name = "String" Then
            If HobjValorNew.ToString.Length > 50 Then
                HobjValorNew = HobjValorNew.ToString.Substring(0, 50)
            End If
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud, BlnEsRequerido)
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
Friend Class ClsServicio_EstCtaStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Servicio"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Acrónimo Servicio"
        HshrLongitud = 3
        HenuTipoValor = EnuTipoValorDef.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud,
                BlnEsRequerido)
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