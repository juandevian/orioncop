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
