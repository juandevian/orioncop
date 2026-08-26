Imports System.Runtime.CompilerServices
#Region "Definiciones"
<Assembly: CLSCompliant(True)>
<Assembly: InternalsVisibleTo("OrionCopIU")>
<Assembly: InternalsVisibleTo("AdminOrionIU")>
<Assembly: InternalsVisibleTo("WinCom")>
<Assembly: InternalsVisibleTo("RepOriCop")>
<Assembly: InternalsVisibleTo("OriIntCon")>
#End Region
#Region "Estructuras"
Friend Structure StcIntMoraFactura
    Implements IEquatable(Of StcIntMoraFactura)
    Friend Property ShrIdItemFactura As Short
    Friend Property DecVlrMora As Decimal
    Friend Property DecBaseIntereses As Decimal
    Friend Property EntDiasMora As Integer
    Friend Property DtmFechaCauso As Date
    Friend Property DblTarifaIva As Double
    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then Return False
        Dim lblnEsIgual As Boolean = obj.GetType.Name = "stcIntMoraFactura"
        If lblnEsIgual Then
            lblnEsIgual = Equals(obj)
        End If
        Return lblnEsIgual
    End Function
    Public Overloads Function Equals(other As StcIntMoraFactura) As Boolean Implements IEquatable(Of StcIntMoraFactura).Equals
        Dim lblnEsIgual As Boolean = (ShrIdItemFactura = other.ShrIdItemFactura)
        If lblnEsIgual Then
            lblnEsIgual = (DecVlrMora = other.DecVlrMora) AndAlso
                    (DecBaseIntereses = other.DecBaseIntereses) AndAlso (
                    DblTarifaIva = other.DblTarifaIva)
        End If
        Return lblnEsIgual
    End Function
    Public Shared Operator =(astcIntMoraFac1 As StcIntMoraFactura, astcIntMoraFac2 As StcIntMoraFactura) As Boolean
        Return astcIntMoraFac1.Equals(astcIntMoraFac2)
    End Operator
    Public Shared Operator <>(astcIntMoraFac1 As StcIntMoraFactura, astcIntMoraFac2 As StcIntMoraFactura) As Boolean
        Return Not astcIntMoraFac1.Equals(astcIntMoraFac2)
    End Operator
    Public Overrides Function GetHashCode() As Integer
        Return 0
    End Function
End Structure
#End Region
#Region "Enumeradores"
Friend Enum EnuAppConta As Byte
    None = 0
    EnuApoloAP
    EnuApoloBD
    EnuContaPyme
    EnuSIIGO
    EnuSIIGON
    EnuColon
    EnuMekano
    EnuPodium
End Enum

Friend Enum EnuDestinatarioFacturaDef As Byte
    None = 0
    EnuPropietario
    EnuArrendatario
End Enum

Friend Enum EnuDestiItemProgramaFact As Byte
    None = 0
    EnuPredio
    EnuCliente
    EnuTodos
End Enum

<Flags()>
Friend Enum EnuDocsIntegridad As Short
    None = 0
    EnuFac = 1
    EnuRC = 2
    EnuNcr = 4
    EnuNdb = 8
    EnuNco = 16
    EnuAnt = 32
    EnuNrrc = 64
    EnuEstadoCta = 128
    EnuTodos = 255
End Enum
Friend Enum EnuDocReversado As Byte
    None = 0
    EnuReciboC
    EnuNotaCr
End Enum

Friend Enum EnuEstadoAplicacionDef As Byte
    None = 0
    EnuCrearServicioAno
    EnuParaSerAdminNotOK
    EnuServicioNotOk
    EnuServicioPorCal
    EnuDocPorDefinir
    EnuListoImportar
    EnuListoImpNDb
    EnuSinPresupuesto
    EnuSinModulos
    EnuHayItemsProgFactPorProcesar
    EnuDebeAjustarCuotasAdmin
    EnuDebeImportarAjuste
    EnuHayPrefacturas
    EnuParaCierreMes
    EnuCausarInt
    EnuCrearAno
    EnuSinCalcAdmin
    EnuDocPorProEFac
    EnuFactFueraDeFecha
    EnuEmailNoOk
    EnuSectoresSinDsctoPP
    EnuNormal
End Enum

Friend Enum EnuEstadoDeudaDef As Byte
    None = 0
    EnuNormal
    EnuPersuasivo
    EnuPrejuridico
    EnuJuridico
    EnuPerdida
    EnuSuspendida
End Enum

Friend Enum EnuEstadoEDoc As Short
    None = -1
    EnuErrorFtp = 0
    EnuNoReg
    EnuInvalida
    EnuEnProceso
    EnuRegi
    EnuEnviada
    EnuAceptada
    EnuRechazada
    EnuOtro
    EnuNoEDoc = 10
End Enum

Friend Enum EnuEstadoFacturaDef As Byte
    None = 0
    EnuNormal
    EnuCancelada
    EnuPeriodoGracia
    EnuVencida
    EnuAnulada
End Enum

Friend Enum EnuEstadoInstalacion As Integer
    None = 0
    CuentasCont = 1
    OpcionesCentroUtil = 2
    CuentasBancos = 4
    Sectores = 8
    Modulos = 16
    SectoresModulo = 32
    Anos = 64
    Servicios = 128
    TasasMora = 256
    Terceros = 512
    Clientes = 1024
    Predios = 2048
    CPCalculados = 4096
    Propietarios = 8192
    Docum = 16384
    SerId = 32768
    Todos = 65535
End Enum

Friend Enum EnuEstadoItemTareaDef As Byte
    None = 0
    EnuNormal
    EnuCumplida
    EnuVencida
    EnuCancelada
    EnuTodos
End Enum

Friend Enum EnuEstadoResDian
    EnuOk
    EnuSinResVigente
    EnuResPorVencer
    EnuVenceHoy
    EnuVencida
    EnuNumPorAgotarse
    EnuNumAgotada
End Enum

Friend Enum EnuEstadoTareaDef As Byte
    None = 0
    EnuVigente
    EnuTerminada
    EnuCancelada
End Enum

Friend Enum EnuFormaPago As Byte
    None = 0
    EnuContado
    EnuCredito
End Enum

Friend Enum EnuGrupoConstantesOriDef As Byte
    None = 0
    EnuDestinatarioFactura
    EnuMediosPago
    EnuRegimenIva
    EnuTipoBaseCalculo
    EnuTipoServicio
    EnuOrigenItemProgramaFact
    EnuEstadoDeuda
    EnuTipoDescuento
    EnuAppContable
    EnuTipoInterfaz
    EnuEstadoTarea
    EnuEstadoItemTarea
    EnuModoCausaMora
    EnuTipoTercero
    EnuTipoTerCtaCrSer
    EnuProvEFac
    EnuDocReversado
    EnuFormaPago
    EnuModoNotaCr
    EnuTipoIncentivo
    EnuTipoDsctoPP
End Enum

Friend Enum EnuIdDocumentoDef As Byte
    None = 0
    EnuFacturaVenta
    EnuReciboCaja
    EnuNotaAplicacionAnt
    EnuNotaIntMora
    EnuNotaCr
    EnuNotaReintegroAnt
    EnuNotaReversaCr
    EnuNotaAjuste
    EnuComprobanteInterfaz
End Enum

<Flags()>
Friend Enum EnuIntervaloDiaDef As Integer
    None = 0
    EnuLunes = 1
    EnuMartes = 2
    EnuMiercoles = 4
    EnuJueves = 8
    EnuViernes = 16
    EnuSabado = 32
    EnuDomingo = 64
    EnuTodos = 127
End Enum

<Flags()>
Friend Enum EnuIntervaloMesDef As Integer
    None = 0
    EnuEne = 1
    EnuFeb = 2
    EnuMar = 4
    EnuAbr = 8
    EnuMay = 16
    EnuJun = 32
    EnuJul = 64
    EnuAgo = 128
    EnuSep = 256
    EnuOct = 512
    EnuNov = 1024
    EnuDic = 2048
    EnuTodos = 4095
End Enum

Friend Enum EnuModoCausaMora
    None
    EnuNoCausa
    EnuEnFecha
    EnuFinMes
    EnuUltimoDia
    EnuAlReciboCaja
End Enum

Friend Enum EnuModoFacturacionDef As Byte
    None = 0
    EnuManual
    EnuSistema
    EnuImportada
    EnuContingencia
    EnuMulta
End Enum

Friend Enum EnuModoNotaCr As Byte
    None
    EnuPorFactura
    EnuPorValor
End Enum

Friend Enum EnuModoPagoIntereses As Byte
    None = 0
    EnuVencido
    EnuAnticipado
End Enum

Friend Enum EnuOrigenItemProgramaFactDef
    None = 0
    EnuAplicacion
    EnuUsuario
    EnuImportado
End Enum

Friend Enum EnuOrigenNotaDb
    None = 0
    EnuAplicacion
    EnuImportado
End Enum

Friend Enum EnuPeriodicidadDePagoDef
    None = 0
    EnuDiaria = 1
    EnuMensual = 30
    EnuBimestral = 60
    EnuTrimestral = 90
    EnuSemestral = 180
    EnuAnual = 360
End Enum

Friend Enum EnuProveedorEFac
    None
    EnuProtecdataMisFac
End Enum

Friend Enum EnuRegimenVentasDef As Byte
    None = 0
    EnuNoResponsable
    EnuResponsable
End Enum

Friend Enum EnuTipoBaseCalculo As Byte
    None = 0
    EnuCoeficientePro
    EnuUnidad
    EnuCuotaAnterior
    EnuImportadas
End Enum

Friend Enum EnuTipoIncentivo
    None = 0
    EnuDescuentoPP
    EnuPenalización
End Enum

Friend Enum EnuTipoCorreoE As Byte
    None = 0
    EnuSoloMens
    EnuArchExt
    EnuFactAuto
    EnuFac
    EnuRC
    EnuNCR
    EnuNDB
    EnuNAA
    EnuRecibos
    EnuCobroPers
End Enum

Friend Enum EnuTipoDescuento As Byte
    None = 0
    EnuDsctoCapital
    EnuDsctoIntMora
    EnuReteFuente
    EnuReteIca
    EnuReteIva
    EnuReteCree
    EnuDsctoPP
    EnuCancelaIva
End Enum

Friend Enum EnuTipoDeudorDef As Byte
    None = 0
    EnuPredio
    EnuCliente
End Enum

Friend Enum EnuTipoDocOri As Byte
    None = 0
    EnuFactura
    EnuReciboCaja
    EnuNotaCon
    EnuNotaDb
    EnuNotaCr
    EnuNotaDevAnt
    EnuNotaRevCr
    EnuNotaAjuste
End Enum

Friend Enum EnuTipoDsctoPP As Byte
    None = 0
    EnuProcentaje
    EnuValorFijo
End Enum

Friend Enum EnuTipoInteres As Byte
    None = 0
    EnuInteresSimple
    EnuInteresCompuesto
End Enum

Friend Enum EnuTipoInterfazDef As Byte
    None = 0
    EnuPorComprobante
    EnuPorDocumento
End Enum

Friend Enum EnuTipoItemNotaConDef As Byte
    None
    EnuAplicaAntCap
    EnuAplicaAntInt
    EnuDsctoPP
    EnuReteFuente
    EnuReteIca
    EnuReteIva
End Enum

Friend Enum EnuTipoItemRecCajaDef As Byte
    None = 0
    EnuAbonoCapital
    EnuAbonoIntMora
    EnuAnticipo
    EnuDsctoIntMora
    EnuDsctoCapital
    EnuReteFuente
    EnuReteIca
    EnuReteIva
    EnuReteCree
    EnuDsctoPP
End Enum

Friend Enum EnuTipoMedioPagoDef As Byte
    None = 0
    EnuEfectivo
    EnuCheque
    EnuTarjetaCR
    EnuTarjetaDB
    EnuConsignacion
    EnuTransferencia
End Enum

Friend Enum EnuTipoNotaCrDef As Byte
    None = 0
    EnuDescuento
    EnuAnulaFac
    EnuRetenciones
End Enum

Friend Enum EnuTipoNov As Byte
    None = 0
    EnuDbCap        '1  Db a CxC Servicio; Cr a Ingresos Servicio o equivalente (Cuentas definidas en el Servicio facturado) 
    EnuDbIva        '2  Db a CxC Servicio; Cr a Iva Generado (Cuentas definidas en el Servicio facturado)
    EnuDbInt        '3  Db a CxC Intereses; Cr Ingresos Inte  (Cuentas definidas en Parametros Centro Utilidades)
    EnuCrPagoCap    '4  Db a Caja o Banco; Cr a CxC Servicio 
    EnuCrPagoInt    '5  Db a Caja o Banco; Cr a CxC Intereses
    EnuCrAnApCap    '6  Db a Anticipos recibidos; Cr a CxC Servicio
    EnuCrAnApInt    '7  Db a Anticipos recibidos; Cr a CxC Intereses
    EnuCrDctoCap    '8  Db a Devoluciones Cap; Cr a CxC Servicio
    EnuCrDctoInt    '9  Db a Devoluciones Int; Cr a CxC Intereses
    EnuCrRetFte     '10 Db a Cta Retefuente; Cr a CxC Servicio
    EnuCrRetIva     '11 Db a Cta ReteIva; Cr a CxC Servicio
    EnuCrRetIca     '12 Db a Cta ReteIca; Cr a CxC Servicio
    EnuCrRetCre     '13 Db a Cta ReteCre; Cr a CxC Servicio
    ' Anticipos
    EnuCrAntRec     '14 Db a Activo Corriente (Según Medio de Pago); Cr a Anticipos recibidos 
    EnuDbAntDev     '15 Db a Anticipos recibidos; Cr a Caja (Devolución de Anticipos)
    EnuDbAntApl     '16 Db a Anticipos recibidos; Cr a CxC Servicio o intereses de mora 
    ' Rversar Movimiento
    EnuRDbCap        '17 Cr a CxC Servicio; Db a Ingresos Servicio o equivalente (Cuentas definidas en el Servicio facturado) 
    EnuRDbIva        '18 Cr a CxC Servicio; Db a Iva Generado (Cuentas definidas en el Servicio facturado)
    EnuRDbInt        '19 Cr a CxC Intereses; Db Ingresos Inte  (Cuentas definidas en Parametros Centro Utilidades)
    EnuRCrPagoCap    '20 Cr a Caja o Banco; Db a CxC Servicio 
    EnuRCrPagoInt    '21 Cr a Caja o Banco; Db a CxC Intereses
    EnuRCrAnApCap    '22 Cr a Anticipos recibidos; Db a CxC Servicio
    EnuRCrAnApInt    '23 Cr a Anticipos recibidos; Db a CxC Intereses
    EnuRCrDctoCap    '24 Cr a Devoluciones Cap; Db a CxC Servicio
    EnuRCrDctoInt    '25 Cr a Devoluciones Int; Db a CxC Intereses
    EnuRCrRetFte     '26 Cr a Cta Retefuente; Db a CxC Servicio
    EnuRCrRetIva     '27 Cr a Cta ReteIva; Db a CxC Servicio
    EnuRCrRetIca     '28 Cr a Cta ReteIca; Db a CxC Servicio
    EnuRCrRetCre     '29 Cr a Cta ReteCre; Db a CxC Servicio
    ' Reversar Anticipos
    EnuRCrAntRec     '30 Cr a Activo Corriente (Según Medio de Pago); Db a Anticipos recibidos  
    EnuRDbAntDev     '31 Cr a Anticipos recibidos; Db a Caja (Devolución de Anticipos)
    EnuRDbAntApl     '32 Cr a Anticipos recibidos; Db a CxC Servicio o intereses de mora
    ' Cancelar deuda del IVA
    EnuCrIvaGas      '33 Cr a CxC Servicio; Db al Gasto (Impuestos Asumidos)
    EnuRCrIvaGas     '34 Cr al Gasto (Impuestos Asumidos); Db a CxC Servicio
    ' IVA intereses de mora
    EnuDbIvaInt      '35 Db a CxC Intereses; Cr a Iva Generado (Cuentas definidas en el Servicio facturado)
    EnuRDbIvaInt     '36 Cr a CxC Intereses; Db a IvaGenerado
End Enum

Friend Enum EnuTipoServicio As Byte
    None = 0
    EnuAnual
    EnuPermanente
End Enum

Friend Enum EnuTipoTerceroCajaDef As Byte
    None = 0
    EnuSinTercero
    EnuCopropiedad
    EnuCliente
End Enum

Friend Enum EnuTipoTerCtaCrServicio As Byte
    None = 0
    EnuProveedor
    EnuCliente
End Enum

Friend Enum EnuVerEFac As Byte
    EnuNinguna = 0
    EnuV1
    EnuV2
End Enum
#End Region
#Region "Enumeradores Dian"
Friend Enum EnuTipoRegVentasPD As SByte
    None = -1
    EnuSimple = 0
    EnuResponsable = 2
End Enum
Friend Enum EnuTipoPersonaDian As Byte
    None = 0
    EnuJuridica
    EnuNatural
End Enum
Friend Enum EnuConceptoNotaCrDian As Byte
    None = 0
    EnuDevolucionParcial
    EnuAnulacion    ' Anulación factura
    EnuRebajaDscto  ' Rebaja o descuento parcial ototal
    EnuAjustePrecio ' Ajuste de precio
    EnuOtros
End Enum
Friend Enum EnuConceptoNotaDbDian As Byte
    None = 0
    Intereses           ' Intereses de mora
    EnuGastosXCobrar    ' Gastos por cobrar
    EnuCambioValor      ' Cambio Valor
    EnuOtros
End Enum
#End Region
#Region "Clases de propiedad comunes (Ubicación, Origen y Usuario)"
Friend Class ClsIdCarpetaShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCarpeta"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdCarpeta"
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HenuTipoValor = EnuTipoValor.enuShort
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 0
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        Dim lobjValorIng As Object = HobjValorNew
        If Not BlnLeyendoOrigen Then
            If HblnEsValido Then
                HblnEsValido = (HobjValorNew = GshrIdCarpeta)
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
Friend Class ClsIdCentroUtilShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCentroUtil"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdCentroUtil"
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HenuTipoValor = EnuTipoValor.enuShort
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 1
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
    BlnEsRequerido, EnuTipoValor)
        If Not BlnLeyendoOrigen Then
            If HblnEsValido Then
                HblnEsValido = (HobjValorNew = GshrIdCentroUtil)
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
Friend Class ClsIdUsuarioStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdUsuario"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdUsuario"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud,
BlnEsRequerido)
        If HblnEsValido Then
            Dim lobjPadre As ClsCBObjetoPan = ObjPadre
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = GstrIdUsuario)
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
Friend Class ClsCUDocStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "CUDoc"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Código único documento"
        HshrLongitud = 300
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
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsCUDEStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "CUDE"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Código único documento electrónico"
        HshrLongitud = 300
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
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdEstadoEDocEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdEstadoEDoc"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Id estado Documento electrónico"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = GobjParametros.BlnEFacAutorizado
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuEstadoEDoc.EnuErrorFtp,
                EnuEstadoEDoc.EnuOtro, BlnEsRequerido)
        If Not HblnEsValido Then
            HblnEsValido = HobjValorNew = EnuEstadoEDoc.EnuNoEDoc
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
            Return HobjValorPro.ToString()
        End If
    End Function
End Class
Friend Class ClsVerEFacEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "VerEFac"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Versión Fact. Elec."
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = (GobjParametros.BlnEFacAutorizado)
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuVerEFac.EnuNinguna,
EnuVerEFac.EnuV2, BlnEsRequerido)
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
            Return CType(HobjValorPro, Integer).ToString
        End If
    End Function
End Class
#End Region
Friend Class ClsFechasServicio
    Friend DtmFechaFac As Date = GCDTMFECHANULA
    Friend StrKeysSer As String() = {}
    Friend Sub New(adtmFechaFac As Date)
        DtmFechaFac = adtmFechaFac
    End Sub
    Friend Function FblnEsDistinta(adtmFechaFac As Date) As Boolean
        Dim lblnEsDistinta = DtmFechaFac <> adtmFechaFac
        Return lblnEsDistinta
    End Function
    Friend Sub SAdicioneKeySer(astrKeySer As String)
        If Not StrKeysSer.Contains(astrKeySer) Then
            Dim lentLargo = StrKeysSer.Length
            ReDim Preserve StrKeysSer(lentLargo)
            StrKeysSer(lentLargo) = astrKeySer
        End If
    End Sub
End Class