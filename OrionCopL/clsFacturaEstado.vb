Friend Class ClsFacturaEstado
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriFacturasEstado"
    ' Variables de modulo
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwFacturaEstado">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As ClsEstadoCuenta, adrwFacturaEstado As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        HblnEsSuprimible = False
        HblnEsModificable = False
        '
        DrwRegistroActual = adrwFacturaEstado
        DtbTablaColeccion = DrwRegistroActual.Table
        EnuPermisosObj = HobjPadre.EnuPermisosObj
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
            Return EnuIdClasesPanDef.enuFacturaEstado
        End Get
    End Property

    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Factura de Estado"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjCreditos_ItFacEstadoDec As New ClsCreditos_ItFacEstadoDec(Me)
    Friend ReadOnly Property ObjDebitos_ItFacEstadoDec As New ClsDebitos_ItFacEstadoDec(Me)
    Friend ReadOnly Property ObjDetalleItemFac_EstadoStr As New ClsDetalleItemFac_EstadoStr(Me)
    Friend ReadOnly Property ObjDeudaCap_ItFacEstDec As New ClsDeudaCap_ItFacEstDec(Me)
    Friend ReadOnly Property ObjDeudaIntMes_ItFacEstDec As New ClsDeudaIntMes_ItFacEstDec(Me)
    Friend ReadOnly Property ObjDeudaIntMora_ItFacEstDec As New ClsDeudaIntMora_ItFacEstDec(Me)
    Friend ReadOnly Property ObjFecha_FacEstadoDtm As New ClsFecha_FacEstadoDtm(Me)
    Friend ReadOnly Property ObjIdAno_ItemFactEstadoShr As New ClsIdAno_ItemFactEstadoShr(Me)
    Friend ReadOnly Property ObjIdCarpeta_FacEstadoShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_FacEstadoShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdCliente_FacEstadoDbl As New ClsIdCliente_FacEstadoDbl(Me)
    Friend ReadOnly Property ObjIdEstadoCta_FacEstadoShr As New ClsIdEstadoCuentaEnt(Me)
    Friend ReadOnly Property ObjIdFacturaVivaEnt As New ClsIdFacturaVivaEnt(Me)
    Friend ReadOnly Property ObjIdPredioAgr_FacEstadoStr As New ClsIdPredioAgr_EstadoStr(Me)
    Friend ReadOnly Property ObjIdServicioItemFac_EstadoShr As New ClsIdServicioItemFac_EstadoShr(Me)
    Friend ReadOnly Property ObjOrdinal_FacEstadoShr As New ClsOrdinal_FacEstadoShr(Me)
    Friend ReadOnly Property ObjPrefijoFacturaVivaStr As New ClsPrefijoFacturaVivaStr(Me)
    Friend ReadOnly Property ObjVlrItemFac_EstadoDec As New ClsVlrItemFac_EstadoDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjCreditos_ItFacEstadoDec)
                HcolPropiedades.Add(ObjDeudaCap_ItFacEstDec)
                HcolPropiedades.Add(ObjDeudaIntMes_ItFacEstDec)
                HcolPropiedades.Add(ObjDeudaIntMora_ItFacEstDec)
                HcolPropiedades.Add(ObjFecha_FacEstadoDtm)
                HcolPropiedades.Add(ObjDebitos_ItFacEstadoDec)
                HcolPropiedades.Add(ObjDetalleItemFac_EstadoStr)
                HcolPropiedades.Add(ObjIdAno_ItemFactEstadoShr)
                HcolPropiedades.Add(ObjIdCarpeta_FacEstadoShr)
                HcolPropiedades.Add(ObjIdCentroUtil_FacEstadoShr)
                HcolPropiedades.Add(ObjIdEstadoCta_FacEstadoShr)
                HcolPropiedades.Add(ObjIdFacturaVivaEnt)
                HcolPropiedades.Add(ObjIdPredioAgr_FacEstadoStr)
                HcolPropiedades.Add(ObjIdCliente_FacEstadoDbl)
                HcolPropiedades.Add(ObjIdServicioItemFac_EstadoShr)
                HcolPropiedades.Add(ObjOrdinal_FacEstadoShr)
                HcolPropiedades.Add(ObjPrefijoFacturaVivaStr)
                HcolPropiedades.Add(ObjVlrItemFac_EstadoDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    ' 
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        GobjPanDat.SControleProcesoObj(True)
        Try
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                SNumereObj()
            End If
            MyBase.SActualice(ablnExigeRequeridos)
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            GobjPanDat.SControleProcesoObj(False)
        End Try
    End Sub
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim ldblIdCliente As Double = ObjIdCliente_FacEstadoDbl.ObjValorPro
            Dim lstrIdpredioAgr As String = ObjIdPredioAgr_FacEstadoStr.ObjValorPro
            Dim lentIdEstadoCuenta As Integer = ObjIdEstadoCta_FacEstadoShr.ObjValorPro
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdCliente_EstadoDbl.SstrNombreCampoBd & " = " & ldblIdCliente &
                    " AND " & ClsIdPredioAgr_EstadoStr.SstrNombreCampoBd & " = '" & lstrIdpredioAgr & "'" &
                    " AND " & ClsIdEstadoCuentaEnt.SstrNombreCampoBd & " = " & lentIdEstadoCuenta
            Dim lshrOrdinal As Short = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                    ClsOrdinal_FacEstadoShr.SstrNombreCampoBd, ObjOrdinal_FacEstadoShr.EnuTipoValor,
                    lstrFiltro) + 1
            ObjOrdinal_FacEstadoShr.ObjValorPro = lshrOrdinal
        End If
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsCreditos_ItFacEstadoDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Creditos"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Credito_FacEstado"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Decimal.MaxValue,
                    BlnEsRequerido, EnuTipoValor.enuDecimal)
        If HblnEsValido Then
            Dim lobjPadre As ClsFacturaEstado = ObjPadre
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
Friend Class ClsDebitos_ItFacEstadoDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Debitos"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Debitos_FacEstado"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Decimal.MaxValue,
                    BlnEsRequerido, EnuTipoValor.enuDecimal)
        If HblnEsValido Then
            Dim lobjPadre As ClsFacturaEstado = ObjPadre
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
Friend Class ClsDetalleItemFac_EstadoStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Detalle"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DetalleItesmFact"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, HshrLongitud, HblnEsRequerido)
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
Friend Class ClsDeudaCap_ItFacEstDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DeudaCapital"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Deuda Capital"
        HenuTipoValor = EnuTipoValor.enuDecimal
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
Friend Class ClsDeudaIntMes_ItFacEstDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DeudaIntMes"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Deuda Intereses Mes"
        HenuTipoValor = EnuTipoValor.enuDecimal
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
Friend Class ClsDeudaIntMora_ItFacEstDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DeudaIntMora"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Deuda Intereses Mora"
        HenuTipoValor = EnuTipoValor.enuDecimal
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
Friend Class ClsFecha_FacEstadoDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaFactura"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Fecha Factura"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = Date.Today
        HobjValorPro = Date.Today
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin = GCDTMFECHANULA
        Dim ldtmFechaMax = Now
        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
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
Friend Class ClsIdAno_ItemFactEstadoShr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsFacturaEstado = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAno"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id Año Item Factura Estado"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Year(Date.MaxValue),
                BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If HobjValorNew > 0 Then
                    HblnEsValido = GobjParametros.ColAnos.Contains(HobjValorNew.ToString)
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
Friend Class ClsIdCliente_FacEstadoDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCliente"
    Private MobjCliente As ClsCliente = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdTerceroCliente_FacturaEstado"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng = HobjValorNew
        Dim lobjPadre As ClsFacturaEstado = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC, BlnEsRequerido)
        If Not HblnEsValido Then
            HblnEsValido = Not String.IsNullOrEmpty(
                    lobjPadre.ObjIdPredioAgr_FacEstadoStr.ToString())
            If Not HblnEsValido Then
                Throw New ErrorInesperadoPanLException("La Id. del Cliente ingresada, '" &
                        lobjValorIng.ToString & "',  no es válida!")
            End If
        End If
    End Sub
    Friend ReadOnly Property ObjCliente As ClsCliente
        Get
            Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
            MobjCliente = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
            MobjCliente.SAbra(lobjValorLlave)
            Return MobjCliente
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
Friend Class ClsIdEstado_FactEstadoEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdEstadoCuenta"
    Private ReadOnly MobjPadre As ClsFacturaEstado = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdEstado_FacturaEstado"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando OrElse
                ClsOrionCop.BlnFacturando Then
            Dim lobjEstadoCuenta As ClsEstadoCuenta = MobjPadre.ObjPadre
            HblnEsValido = (HobjValorNew = lobjEstadoCuenta.ObjIdEstadoCuentaEnt.ObjValorPro)
        Else
            HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
Friend Class ClsIdFacturaVivaEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdFacturaViva"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdFacturaViva"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido)
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
Friend Class ClsIdServicioItemFac_EstadoShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdServicio"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Id Servicio Item Factura Estado"
        HenuTipoValor = EnuTipoValor.enuUShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                    BlnEsRequerido, EnuTipoValor)
        If Not HblnEsValido AndAlso GblnActualizandoApp Then
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
Friend Class ClsOrdinal_FacEstadoShr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsFacturaEstado = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "Ordinal"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Ordinal_FacEstado"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsLlave = True
        HbytPosicionLlave = 3
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, Short.MinValue,
                Short.MaxValue, BlnEsRequerido, EnuTipoValor)
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
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
Friend Class ClsPrefijoFacturaVivaStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoFactViva"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PrefijoFacturaViva"
        HshrLongitud = 5
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then
            HobjValorNew = String.Empty
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            Dim lobjPadre As ClsFacturaEstado = ObjPadre
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
Friend Class ClsVlrItemFac_EstadoDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ValorItem"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Valor"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
        Dim lobjPadre As ClsFacturaEstado = ObjPadre
        If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
            HblnEsValido = (HobjValorNew = HobjValorOriginal)
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