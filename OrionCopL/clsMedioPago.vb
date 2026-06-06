Friend Class ClsMedioPago
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriMediosPago"
    '
    Private ReadOnly MobjPadre As clsReciboCaja = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwMedPago">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As clsReciboCaja, adrwMedPago As DataRow)
        HobjPadre = aobjPadre
        MobjPadre = aobjPadre
        henuTipoObjeto = enuModoInstanciaObjDef.enuDeColeccion
        hblnEsAnulable = False
        '
        drwRegistroActual = adrwMedPago
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
            Return EnuIdClasesPanDef.enuMediosPago
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Medio de Pago"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdCarpeta_MedPagoShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_MedPagoShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdCtaContabIngresoStr As New ClsIdCtaContabIngresoStr(Me)
    Friend ReadOnly Property ObjIdRecCaja_MedPagoEnt As New ClsIdRecCaja_MedPagoEnt(Me)
    Friend ReadOnly Property ObjIdTipoMedPago_MedPagoByt As New ClsIdTipoMedPago_MedPagoByt(Me)
    Friend ReadOnly Property ObjNumeroMedPagoStr As New ClsNumeroMedPagoStr(Me)
    Friend ReadOnly Property ObjOrdinal_MedPagoShr As New ClsOrdinal_MedPagoShr(Me)
    Friend ReadOnly Property ObjPrefijo_MedPagoStr As New ClsPrefijo_RecStr(Me)
    Friend ReadOnly Property ObjValor_MedPagoDec As New ClsValor_MedPagoDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjIdCarpeta_MedPagoShr)
                HcolPropiedades.Add(ObjIdCentroUtil_MedPagoShr)
                HcolPropiedades.Add(ObjIdCtaContabIngresoStr)
                HcolPropiedades.Add(ObjIdRecCaja_MedPagoEnt)
                HcolPropiedades.Add(ObjIdTipoMedPago_MedPagoByt)
                HcolPropiedades.Add(ObjNumeroMedPagoStr)
                HcolPropiedades.Add(ObjOrdinal_MedPagoShr)
                HcolPropiedades.Add(ObjPrefijo_MedPagoStr)
                HcolPropiedades.Add(ObjValor_MedPagoDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend ReadOnly Property StrCuentaIngreso As String
        Get
            Dim lstrCuentaIngreso = String.Empty
            If Not IsNothing(ObjIdCtaContabIngresoStr.ObjValorPro) Then
                lstrCuentaIngreso = ClsOrionCop.FstrCuentaBanco(ObjIdCtaContabIngresoStr.ObjValorPro)
            End If
            Return lstrCuentaIngreso
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SInicialiceObj()
        ObjIdCarpeta_MedPagoShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtil_MedPagoShr.ObjValorPro = GshrIdCentroUtil
        ObjPrefijo_MedPagoStr.ObjValorPro = GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuReciboCaja)
        ObjNumeroMedPagoStr.ObjValorPro = String.Empty
        ObjIdCtaContabIngresoStr.ObjValorPro = String.Empty
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjOrdinal_MedPagoShr.ObjValorPro.ToString
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    ''' <summary>
    ''' Devuelve un factor que indica en que porcentaje del pago total participa este medio de pago 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FdblTasaParticipaEnPago() As Double
        Dim ldblTasaPart As Double
        ldblTasaPart = ObjValor_MedPagoDec.objValorPro / MobjPadre.objValor_RecDec.objValorPro
        Return ldblTasaPart
    End Function
    Friend Function FblnMedioPagoUnico(astrIdCntaContIng As String, astrNroMediPago As String,
                                       aenuTipoMediPago As EnuTipoMedioPagoDef) As Boolean
        Dim lobjPadre As ClsReciboCaja = HobjPadre
        Dim lblnEsUnico = True
        If ObjIdTipoMedPago_MedPagoByt.BlnEsValido AndAlso ObjNumeroMedPagoStr.BlnEsValido AndAlso
                ObjIdCtaContabIngresoStr.BlnEsValido Then
            lblnEsUnico = lobjPadre.FblnEsUnicoMediPago(astrIdCntaContIng, astrNroMediPago,
                aenuTipoMediPago)
        End If
        Return lblnEsUnico
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Class ClsIdCtaContabIngresoStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsMedioPago = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCtaContIngreso"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "CuentaContIngreso"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            If HblnEsValido Then
                HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
                If HblnEsValido Then
                    If MobjPadre.ObjIdTipoMedPago_MedPagoByt.ObjValorPro <= 2 Then
                        HblnEsValido = (HobjValorNew = GobjParametros.ObjIdCtaCajaStr.ObjValorPro)
                    Else
                        HblnEsValido = (HobjValorNew <> GobjParametros.ObjIdCtaCajaStr.ObjValorPro)
                    End If
                    If Not HblnEsValido Then
                        HstrMens = "La Cuenta donde ingresó el Pago no tiene relación con el Tipo " &
                                "del Medio de Pago!"
                        SNotifiqueDatInv()
                    End If
                End If
                If HblnEsValido Then
                    HblnEsValido = MobjPadre.FblnMedioPagoUnico(HobjValorNew,
                            MobjPadre.ObjNumeroMedPagoStr.ObjValorPro,
                            MobjPadre.ObjIdTipoMedPago_MedPagoByt.ObjValorPro)
                    If Not HblnEsValido Then
                        HstrMens = "Este Medio de Pago ya se ingreso. El Medio de Pago debe " &
                            "ser único!"
                        SNotifiqueDatInv()
                    End If
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
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            If HblnEsValido Then
                MobjPadre.ObjIdTipoMedPago_MedPagoByt.SValide()
                MobjPadre.ObjNumeroMedPagoStr.SValide()
            End If
        End If
    End Sub
End Class
Friend Class ClsIdRecCaja_MedPagoEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdReciboCaja"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdReciboCaja_MedPago"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsMedioPago = ObjPadre
        HblnEsValido = True
        If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido)
            If HblnEsValido Then
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
Friend Class ClsIdTipoMedPago_MedPagoByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoMediosPago"
    Private ReadOnly MobjPadre As ClsMedioPago = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTipoMedioPago"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoMedioPagoDef.enuEfectivo,
                EnuTipoMedioPagoDef.enuTransferencia, BlnEsRequerido)
        If HblnEsValido Then
            If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If GobjParametros.ObjExigeFechaHoyCajaBln.ObjValorPro Then
                    If HobjValorNew = EnuTipoMedioPagoDef.enuCheque OrElse
                            HobjValorNew = EnuTipoMedioPagoDef.enuEfectivo Then
                        Dim lobjRecCaja As ClsReciboCaja = ObjPadre.ObjPadre
                        If Not (lobjRecCaja.ObjFechaRecDtm.ObjValorPro = Date.Today) Then
                            HblnEsValido = False
                            HstrMens = "La Fecha del Recibo de Caja para este Medio de Pago " &
                                    "debe ser la fecha de Hoy!"
                            SNotifiqueDatInv()
                        End If
                    End If
                End If
                If HblnEsValido Then
                    HblnEsValido = MobjPadre.FblnMedioPagoUnico(
                            MobjPadre.ObjIdCtaContabIngresoStr.ObjValorPro,
                            MobjPadre.ObjNumeroMedPagoStr.ObjValorPro, HobjValorNew)
                    If Not HblnEsValido Then
                        HstrMens = "Este Medio de Pago ya se ingreso. El Medio de Pago debe " &
                            "ser único!"
                        SNotifiqueDatInv()
                    End If
                End If
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            If HblnEsValido Then
                MobjPadre.ObjIdCtaContabIngresoStr.SValide()
                MobjPadre.ObjNumeroMedPagoStr.SValide()
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
            Return ClsOrionCop.FstrNombreDatoConstanteOri(EnuGrupoConstantesOriDef.enuMediosPago,
                    HobjValorPro)
        End If
    End Function
End Class
Friend Class ClsNumeroMedPagoStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "NumeroMedioPago"
    Private ReadOnly MobjPadre As ClsMedioPago = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Numero_MedPago"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        If String.IsNullOrEmpty(HobjValorNew) Then HobjValorNew = Nothing
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            If HblnEsValido AndAlso HobjValorNew <> "" AndAlso
                    MobjPadre.ObjIdTipoMedPago_MedPagoByt.BlnEsValido Then
                Dim lstrIdRecCaja = String.Empty
                If FblnExisteMedioPago(lstrIdRecCaja) Then
                    HstrMens = "Un Medio de Pago con este número ya existe " &
                            "en el Recibo de Caja Nro." & lstrIdRecCaja
                    SLevanteEveNot("", 0, EnuSeveridadNot.EnuAdvertencia)
                End If
            ElseIf Not IsNothing(HobjValorNew) AndAlso HobjValorNew.ToString.Length > 30 Then
                HblnEsValido = False
                HstrMens = "El Número del Medio de Pago excede 30 caracteres"
                SNotifiqueDatInv()
            End If
            If HblnEsValido Then
                HblnEsValido = MobjPadre.FblnMedioPagoUnico(
                            MobjPadre.ObjIdCtaContabIngresoStr.ObjValorPro, HobjValorNew,
                            MobjPadre.ObjIdTipoMedPago_MedPagoByt.ObjValorPro)
                If Not HblnEsValido Then
                    HstrMens = "Este Medio de Pago ya se ingreso. El Medio de Pago debe " &
                            "ser único!"
                    SNotifiqueDatInv()
                End If
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            If HblnEsValido Then
                MobjPadre.ObjIdTipoMedPago_MedPagoByt.SValide()
                MobjPadre.ObjIdCtaContabIngresoStr.SValide()
            End If
        End If
    End Sub
    Private Function FblnExisteMedioPago(ByRef astrNroRecCaja As String) As Boolean
        Dim lblnExiste = False
        Dim lstrCamposSelect = {ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd, ClsNumeroMedPagoStr.SstrNombreCampoBd,
                ClsPrefijo_RecStr.SstrNombreCampoBd, ClsIdRecCaja_MedPagoEnt.SstrNombreCampoBd}
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd &
                " = " & MobjPadre.ObjIdTipoMedPago_MedPagoByt.ObjValorPro & " AND " &
                ClsNumeroMedPagoStr.SstrNombreCampoBd & " = '" & HobjValorNew & "'"
        Dim ldtbMedPago = ClsPanorama.FdtbDataTable(ClsMedioPago.SstrNombreTabla, lstrCamposSelect,
                {{"", ""}}, lstrFiltro)
        If ldtbMedPago.Rows.Count > 0 Then
            Dim lstrPrefRecCaj As String = ClsPanorama.FobjValorCampo(
                    ldtbMedPago.Rows(0)(ClsPrefijo_RecStr.SstrNombreCampoBd), EnuTipoValor.enuString)
            Dim lentIdRecCaj As Integer = ClsPanorama.FobjValorCampo(
                    ldtbMedPago.Rows(0)(ClsIdRecCaja_MedPagoEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
            astrNroRecCaja = ClsPanorama.FstrNumeroDcto(lstrPrefRecCaj, lentIdRecCaj)
            lblnExiste = True
        End If
        Return lblnExiste
    End Function
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
Friend Class ClsOrdinal_MedPagoShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Ordinal"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "OrdinalMedPago"
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
Friend Class ClsValor_MedPagoDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"
    Private ReadOnly MobjPadre As ClsMedioPago = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Valor Medio de Pago"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Decimal.MaxValue,
                BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            Else
                Dim lobjReciboCaja As ClsReciboCaja = MobjPadre.ObjPadre
                If lobjReciboCaja.ObjValor_RecDec.ObjValorPro < HobjValorNew Then
                    HstrMens = "El Valor total de los Medios de Pago ingresados es mayor al " &
                            "Valor Recibido!"
                    SNotifiqueDatInv()
                    HblnEsValido = False
                End If
                If HblnEsValido Then
                    HblnEsValido = HobjValorNew - Int(HobjValorNew) = 0
                    If Not HblnEsValido Then
                        HstrMens = "El Valor ingresado debe ser sin Centavos!"
                        SNotifiqueDatInv()
                    End If
                End If
            End If
        Else
            If HobjValorNew = 0 Then
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                    Dim lobjRecCaja As ClsReciboCaja = MobjPadre.ObjPadre
                    HblnEsValido = (lobjRecCaja.ObjAnuladoBln.ObjValorPro)
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
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
#End Region