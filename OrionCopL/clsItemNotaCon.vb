Friend Class ClsItemNotaCon
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriItemsNotaCon"
    ' Variables de modulo
    Private ReadOnly MobjPadre As ClsNotaCon = Nothing
    Private MobjFactura As ClsFactura = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwItemNotaCon">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As clsNotaCon, adrwItemNotaCon As DataRow)
        HobjPadre = aobjPadre
        MobjPadre = aobjPadre
        henuTipoObjeto = enuModoInstanciaObjDef.enuDeColeccion
        HblnEsSuprimible = False
        '
        drwRegistroActual = adrwItemNotaCon
        DtbTablaColeccion = DrwRegistroActual.Table
        HenuTipoPermiso = EnuPermisosDef.enuHeredado
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
            Return EnuIdClasesPanDef.enuItemNotaCon
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Item Nota Contable"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdCarpeta_ItemNotaConShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_ItemNotaConShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdFactura_ItemNotaConEnt As New ClsIdFactura_ItemNotaConEnt(Me)
    Friend ReadOnly Property ObjIdItemFac_ItemNotaConShr As New ClsIdItemFac_ItemNotaConShr(Me)
    Friend ReadOnly Property ObjIdItemNotaConShr As New ClsIdItemNotaConShr(Me)
    Friend ReadOnly Property ObjIdNotaCon_ItemNotaConEnt As New ClsIdNotaCon_ItemNotaConEnt(Me)
    Friend ReadOnly Property ObjIdTipoItemNotaConByt As New ClsIdTipoItemNotaConByt(Me)
    Friend ReadOnly Property ObjPrefijoFact_ItemNotaConStr As New ClsPrefijoFact_ItemNotaConStr(Me)
    Friend ReadOnly Property ObjValor_ItemNotaConDec As New ClsValor_ItemNotaConDec(Me)
    Friend ReadOnly Property ObjPrefijoNotaCon_ItemNotaConStr As New ClsPrefijo_NotaConStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAnuladoBln)
                HcolPropiedades.Add(ObjIdCarpeta_ItemNotaConShr)
                HcolPropiedades.Add(ObjIdCentroUtil_ItemNotaConShr)
                HcolPropiedades.Add(ObjIdFactura_ItemNotaConEnt)
                HcolPropiedades.Add(ObjIdItemFac_ItemNotaConShr)
                HcolPropiedades.Add(ObjIdItemNotaConShr)
                HcolPropiedades.Add(ObjIdNotaCon_ItemNotaConEnt)
                HcolPropiedades.Add(ObjIdTipoItemNotaConByt)
                HcolPropiedades.Add(ObjPrefijoFact_ItemNotaConStr)
                HcolPropiedades.Add(ObjPrefijoNotaCon_ItemNotaConStr)
                HcolPropiedades.Add(ObjValor_ItemNotaConDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend ReadOnly Property ObjFactura As ClsFactura
        Get
            If ObjPrefijoFact_ItemNotaConStr.BlnEsValido AndAlso ObjIdFactura_ItemNotaConEnt.BlnEsValido Then
                If IsNothing(MobjFactura) Then
                    Dim lstrPrefFac As String = ObjPrefijoFact_ItemNotaConStr.ObjValorPro
                    Dim lentIdFac As Integer = ObjIdFactura_ItemNotaConEnt.ObjValorPro
                    Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil,
                            lstrPrefFac, lentIdFac}
                    MobjFactura = New ClsFactura()
                    MobjFactura.SAbra(lobjValorLlave)
                End If
            Else
                MobjFactura = Nothing
            End If
            Return MobjFactura
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        MobjFactura = Nothing
    End Sub
    Public Overrides Function FblnEsAnulable() As Boolean
        Return MobjPadre.fblnEsAnulable()
    End Function
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdItemNotaConShr.ToString
        End Get
    End Property
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsIdFactura_ItemNotaConEnt
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdFactura"
    Private ReadOnly MobjPadre As clsItemNotaCon = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdFactura_ItemNotaCon"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
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
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdItemFac_ItemNotaConShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdItemFactura"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Id Item Factura Item Nota Con"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            Dim lobjPadre As ClsCBObjetoPan = ObjPadre
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
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
Friend Class ClsIdItemNotaConShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdItemNotaCon"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdItemNotaCon"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 4
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HobjValorPro = True
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
Friend Class ClsIdNotaCon_ItemNotaConEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdNotaCon"
    Private ReadOnly MobjPadre As ClsItemNotaCon = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdNotaCon_ItemNotaCon"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        Dim lobjNotaCon As ClsNotaCon = MobjPadre.ObjPadre
        If lobjNotaCon.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HblnEsValido = True
        Else
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido)
            If HblnEsValido Then
                HblnEsValido = (HobjValorNew = lobjNotaCon.ObjIdNotaConEnt.ObjValorPro)
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
Friend Class ClsPrefijoFact_ItemNotaConStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoFactura"
    Private ReadOnly MobjPadre As ClsItemNotaCon = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "PrefijoFactura_NotaCon"
        HshrLongitud = 5
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then
            HobjValorNew = String.Empty
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
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
Friend Class ClsIdTipoItemNotaConByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoItemNotaCon"
    Private ReadOnly MobjPadre As ClsItemNotaCon = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTipoItemNotaCon"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuTipoItemNotaConDef.EnuAplicaAntCap,
                EnuTipoItemNotaConDef.EnuReteIva, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
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
Friend Class ClsValor_ItemNotaConDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"
    Private ReadOnly MobjPadre As ClsItemNotaCon = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "ValorItem_NotaCon"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0.01, Decimal.MaxValue,
                BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                If MobjPadre.ObjAnuladoBln.ObjValorPro Then
                    HblnEsValido = (HobjValorNew = 0)
                Else
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
                End If
            End If
        Else
            HblnEsValido = (HobjValorNew = 0 AndAlso MobjPadre.ObjAnuladoBln.ObjValorPro)
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
