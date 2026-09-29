#Region "Enumeradores"
Friend Enum EnuModoInteres As Byte
    None = 0
    EnuFijo
    EnuVariable
End Enum
#End Region

#Region "Estructuras"
Friend Structure StcIbcCertificado
    Friend Property StrIdfile As String
    Friend Property DtmFechaCertificado As Date
    Friend Property DtmFechaDesde As Date
    Friend Property DtmFechaHasta As Date
    ''' <summary>IBC como fracción (0.1938).</summary>
    Friend Property DblIbc As Double
End Structure
#End Region

Friend Class ClsIbcCalculo
#Region "Definiciones"
    ' Ley 675 de 2001, art. 30: el interés de mora no puede exceder 1.5 veces el IBC certificado
    Friend Const CDBLFACTORMAXIMO As Double = 1.5
#End Region
#Region "Reglas de cálculo"
    ''' <summary>
    ''' Busca el certificado cuya vigencia cubre la fecha. Ignora IBC no positivos. Si hay varios,
    ''' toma el de emisión más reciente (desempate por Idfile descendente).
    ''' </summary>
    Friend Shared Function FblnCertificadoVigente(alstCertificados As IEnumerable(Of StcIbcCertificado),
            adtmFecha As Date, ByRef astcVigente As StcIbcCertificado) As Boolean
        Dim lblnHay = False
        astcVigente = Nothing
        If Not IsNothing(alstCertificados) Then
            For Each lstcCert As StcIbcCertificado In alstCertificados
                If lstcCert.DblIbc > 0 AndAlso lstcCert.DtmFechaDesde <= adtmFecha AndAlso
                        adtmFecha <= lstcCert.DtmFechaHasta Then
                    If Not lblnHay OrElse lstcCert.DtmFechaCertificado > astcVigente.DtmFechaCertificado OrElse
                            (lstcCert.DtmFechaCertificado = astcVigente.DtmFechaCertificado AndAlso
                            String.CompareOrdinal(lstcCert.StrIdfile, astcVigente.StrIdfile) > 0) Then
                        astcVigente = lstcCert
                        lblnHay = True
                    End If
                End If
            Next
        End If
        Return lblnHay
    End Function

    ''' <summary>
    ''' Convierte el IBC de la API (porcentaje, ej. 19.38) a fracción (0.1938).
    ''' </summary>
    Friend Shared Function FdblIbcComoFraccion(adblPorcentaje As Double) As Double
        Return Math.Round(adblPorcentaje / 100, 6)
    End Function

    Friend Shared Function FdblTopeMaximo(adblIbc As Double) As Double
        Return Math.Round(adblIbc * CDBLFACTORMAXIMO, 6)
    End Function

    ''' <summary>
    ''' Tasa fija que realmente se cobra: la deseada, limitada por el tope vigente.
    ''' </summary>
    Friend Shared Function FdblTasaFijaEfectiva(adblDeseada As Double, adblIbc As Double) As Double
        Return Math.Min(adblDeseada, FdblTopeMaximo(adblIbc))
    End Function

    Friend Shared Function FblnFactorValido(adblFactor As Double) As Boolean
        Return adblFactor > 0 AndAlso adblFactor <= CDBLFACTORMAXIMO
    End Function

    Friend Shared Function FdblTasaVariable(adblIbc As Double, adblFactor As Double) As Double
        Return Math.Round(adblIbc * adblFactor, 6)
    End Function

    ''' <summary>
    ''' Equivalente mensual que se muestra al usuario (misma convención de FdtbTasasMora: anual / 12).
    ''' </summary>
    Friend Shared Function FdblMensualParaMostrar(adblTasaAnual As Double) As Double
        Return Math.Round(adblTasaAnual / 12, 6)
    End Function
#End Region
End Class
