Friend Class ClsItemRecCaja
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriItemsRecCaja"
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwRecCaja">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As clsReciboCaja, adrwRecCaja As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        '
        drwRegistroActual = adrwRecCaja
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
            Return EnuIdClasesPanDef.enuItemRecCaja
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Item Recibo Caja"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjBaseDsctoDec As New ClsBaseDsctoDec(Me)
    Friend ReadOnly Property ObjIdCarpeta_ItemRecShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_ItemRecShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdFactura_ItemRecEnt As New ClsIdFactura_ItemRecEnt(Me)
    Friend ReadOnly Property ObjIdCuentaDb_ItemRecStr As New ClsIdCuentaDb_ItemRecStr(Me)
    Friend ReadOnly Property ObjIdItemFac_ItemRecShr As New ClsIdItemFac_ItemRecShr(Me)
    Friend ReadOnly Property ObjIdItemRecCajaShr As New ClsIdItemRecCajaShr(Me)
    Friend ReadOnly Property ObjIdRecCaja_ItemRecEnt As New ClsIdRecCaja_ItemRecEnt(Me)
    Friend ReadOnly Property ObjIdTipoItemRecByt As New ClsIdTipoItemRecByt(Me)
    Friend ReadOnly Property ObjPrefijoFact_ItemRecStr As New ClsPrefijoFact_ItemRecStr(Me)
    Friend ReadOnly Property ObjPrefijoRec_ItemRecStr As New ClsPrefijo_RecStr(Me)
    Friend ReadOnly Property ObjTasaDsctoDbl As New ClsTasaDsctoDbl(Me)
    Friend ReadOnly Property ObjValor_ItemRecDec As New ClsValor_ItemRecDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjIdCarpeta_ItemRecShr)
                HcolPropiedades.Add(ObjIdCentroUtil_ItemRecShr)
                HcolPropiedades.Add(ObjIdFactura_ItemRecEnt)
                HcolPropiedades.Add(ObjIdCuentaDb_ItemRecStr)
                HcolPropiedades.Add(ObjIdItemFac_ItemRecShr)
                HcolPropiedades.Add(ObjIdItemRecCajaShr)
                HcolPropiedades.Add(ObjIdRecCaja_ItemRecEnt)
                HcolPropiedades.Add(ObjIdTipoItemRecByt)
                HcolPropiedades.Add(ObjBaseDsctoDec)
                HcolPropiedades.Add(ObjPrefijoFact_ItemRecStr)
                HcolPropiedades.Add(ObjPrefijoRec_ItemRecStr)
                HcolPropiedades.Add(ObjTasaDsctoDbl)
                HcolPropiedades.Add(ObjValor_ItemRecDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend ReadOnly Property StrNumeroFactura As String
        Get
            Dim lstrPrefijo As String = ObjPrefijoFact_ItemRecStr.ObjValorPro
            Dim lentIdFactura As Integer = ObjIdFactura_ItemRecEnt.ObjValorPro
            Return clsPanorama.fstrNumeroDcto(lstrPrefijo, lentIdFactura)
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.sVacie()
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Dim lstrPrefijo As String = ObjPrefijoFact_ItemRecStr.ObjValorPro
            Dim lstrIdObjeto = String.Empty
            If lstrPrefijo <> "" Then
                lstrIdObjeto = lstrPrefijo & "-"
            End If
            lstrIdObjeto &= ObjIdItemRecCajaShr.ToString
            Return lstrIdObjeto
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Friend Function FenuTipoNovedad() As EnuTipoNov
        Dim lenuTipoItem As EnuTipoItemRecCajaDef = ObjIdTipoItemRecByt.ObjValorPro
        Dim lenuTipoNov As EnuTipoNov = EnuTipoNov.None
        Select Case lenuTipoItem
            Case EnuTipoItemRecCajaDef.enuAbonoCapital
                lenuTipoNov = EnuTipoNov.enuCrPagoCap
            Case EnuTipoItemRecCajaDef.enuAbonoIntMora
                lenuTipoNov = EnuTipoNov.enuCrPagoInt
            Case EnuTipoItemRecCajaDef.enuDsctoCapital, EnuTipoItemRecCajaDef.enuDsctoPP
                lenuTipoNov = EnuTipoNov.enuCrDctoCap
            Case EnuTipoItemRecCajaDef.enuDsctoIntMora
                lenuTipoNov = EnuTipoNov.enuCrDctoInt
            Case EnuTipoItemRecCajaDef.enuReteFuente
                lenuTipoNov = EnuTipoNov.enuCrRetFte
            Case EnuTipoItemRecCajaDef.enuReteIca
                lenuTipoNov = EnuTipoNov.enuCrRetIca
            Case EnuTipoItemRecCajaDef.enuReteIva
                lenuTipoNov = EnuTipoNov.enuCrRetIva
            Case EnuTipoItemRecCajaDef.enuReteCree
                lenuTipoNov = EnuTipoNov.enuCrRetCre
        End Select
        Return lenuTipoNov
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsBaseDsctoDec
    Inherits clsCBPropiedad
    Private ReadOnly MobjPadre As clsItemRecCaja
    Private Const MCSTRNOMBRECAMPOBD As String = "BaseDscto"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Base"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Decimal.MaxValue,
                BlnEsRequerido, EnuTipoValor.enuDecimal)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If MobjPadre.ObjIdTipoItemRecByt.ObjValorPro >= EnuTipoItemRecCajaDef.enuDsctoIntMora Then
                    HblnEsValido = (HobjValorNew > 0)
                End If
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
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdFactura_ItemRecEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdFactura"
    Private ReadOnly MobjPadre As ClsItemRecCaja = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdFactura_ItemRec"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lentIdFraMin = 1
        If MobjPadre.ObjIdTipoItemRecByt.ObjValorPro = EnuTipoItemRecCajaDef.EnuAnticipo Then
            lentIdFraMin = 0
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, lentIdFraMin, Integer.MaxValue, BlnEsRequerido)
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
Friend Class ClsIdCuentaDb_ItemRecStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCuentaDb"
    Private MstrNombreCuenta As String = String.Empty
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CodigoCuentaDb"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If lblnEsValido Then
            lblnEsValido = (HobjValorNew = GCSTRCUENTADSCTOCAP) OrElse
                    (HobjValorNew = GCSTRCUENTADSCTOINT)
            If lblnEsValido Then
                MstrNombreCuenta = "Cuenta de Descuento según Servicio"
            Else
                lblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
                If lblnEsValido Then
                    MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
                End If
            End If
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Friend ReadOnly Property StrNombreCuenta As String
        Get
            Return MstrNombreCuenta
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
Friend Class ClsIdItemFac_ItemRecShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdItemFactura"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Id Item Factura ItemRec"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Short.MaxValue,
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
Friend Class ClsIdItemRecCajaShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdItemRecCaja"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdItemRecCaja"
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
Friend Class ClsIdRecCaja_ItemRecEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdReciboCaja"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdRecCaja_ItemRec"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsItemRecCaja = ObjPadre
        Dim lobjRecCaja As ClsReciboCaja = lobjPadre.ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido)
        If HblnEsValido Then
            HblnEsValido = (HobjValorNew = lobjRecCaja.ObjIdRecCajaEnt.ObjValorPro)
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
Friend Class ClsIdTipoItemRecByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoItemRec"
    Private ReadOnly MobjPadre As ClsItemRecCaja = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTipoItemRec"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoItemRecCajaDef.EnuAbonoCapital,
                EnuTipoItemRecCajaDef.EnuDsctoPP, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        With MobjPadre
            .ObjIdFactura_ItemRecEnt.SValide()
            .ObjBaseDsctoDec.SValide()
            .ObjValor_ItemRecDec.SValide()
        End With
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
Friend Class ClsPrefijoFact_ItemRecStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoFactura"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PrefijoFactura_ItemRec"
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
            Dim lobjPadre As ClsItemRecCaja = ObjPadre
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
Friend Class ClsTasaDsctoDbl
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsItemRecCaja = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "TasaDscto"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Tasa de descuento"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, 1.0, BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If MobjPadre.ObjIdTipoItemRecByt.ObjValorPro >= EnuTipoItemRecCajaDef.EnuDsctoIntMora Then
                    HblnEsValido = (HobjValorNew > 0)
                End If
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
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsValor_ItemRecDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"
    Private ReadOnly MobjPadre As ClsItemRecCaja = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Valor ItemRec"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0.01, Decimal.MaxValue,
                BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            With MobjPadre
                If .EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
                End If
            End With
        Else
            If HobjValorNew = 0 Then
                Dim lobjRecCaja As ClsReciboCaja = MobjPadre.ObjPadre
                HblnEsValido = lobjRecCaja.ObjAnuladoBln.ObjValorPro
            End If
        End If
    End Sub
    Private Sub ClsValor_ItemRecDec_evnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            MobjPadre.ObjTasaDsctoDbl.SValide()
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