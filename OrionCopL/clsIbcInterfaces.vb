Friend Interface IIbcProveedor
    ''' <summary>Consulta los últimos N certificados en la API (IBC como fracción).</summary>
    Function FlstConsulteCertificados(aentCantidad As Integer) As List(Of StcIbcCertificado)
End Interface

Friend Interface IIbcAlmacen
    Function FlstCertificados() As List(Of StcIbcCertificado)
    Sub SGuarde(alstCertificados As IEnumerable(Of StcIbcCertificado))
End Interface

Friend Interface IIbcTasasMora
    ''' <summary>Tasa anual vigente a la fecha en OriTasasMora (0 si no hay).</summary>
    Function FdblTasaVigente(adtmFecha As Date) As Double
    ''' <summary>FechaDesde de la última fila de OriTasasMora (GCDTMFECHANULA si no hay).</summary>
    Function FdtmFechaDesdeUltima() As Date
    ''' <summary>Agrega una fila de tasa anual (fracción) desde la fecha indicada.</summary>
    Sub SRegistreTasa(adtmFechaDesde As Date, adblTasaAnual As Double)
End Interface

Friend Class ErrorIbcApiException
    Inherits Exception
    Private ReadOnly MentCodigoHttp As Integer
    Friend Sub New(astrMensaje As String, aentCodigoHttp As Integer)
        MyBase.New(astrMensaje)
        MentCodigoHttp = aentCodigoHttp
    End Sub
    Friend ReadOnly Property EntCodigoHttp As Integer
        Get
            Return MentCodigoHttp
        End Get
    End Property
End Class

''' <summary>
''' La fila nueva de OriTasasMora YA quedó guardada, pero algo posterior falló (cerrar la fila anterior o
''' la verificación de lo guardado). No se debe reintentar: requiere corrección manual en Tasas de Mora.
''' </summary>
Friend Class ErrorTasaMoraGuardadaException
    Inherits Exception
    Private ReadOnly MentOrdinal As Integer
    Private ReadOnly MdtmFechaDesde As Date
    Friend Sub New(aentOrdinal As Integer, adtmFechaDesde As Date, astrDetalle As String,
            aobjInterna As Exception)
        MyBase.New("La nueva Tasa de Mora (Ordinal " & aentOrdinal & ", desde " &
                adtmFechaDesde.ToString("dd/MM/yyyy", Globalization.CultureInfo.InvariantCulture) &
                ") SÍ se guardó, pero quedó inconsistente: " & astrDetalle &
                ". Corríjala manualmente en Tasas de Mora (cierre la fecha hasta de la tasa anterior o " &
                "corrija la tasa) antes de causar intereses. No reintente la sincronización.", aobjInterna)
        MentOrdinal = aentOrdinal
        MdtmFechaDesde = adtmFechaDesde
    End Sub
    Friend ReadOnly Property EntOrdinal As Integer
        Get
            Return MentOrdinal
        End Get
    End Property
    Friend ReadOnly Property DtmFechaDesde As Date
        Get
            Return MdtmFechaDesde
        End Get
    End Property
End Class
