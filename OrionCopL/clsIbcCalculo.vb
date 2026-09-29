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

    ''' <summary>
    ''' Tasa fija que se acepta al digitar: nunca negativa ni mayor al tope. Si se ajusta, avisa.
    ''' </summary>
    Friend Shared Function FdblTasaFijaPermitida(adblDigitada As Double, adblIbc As Double,
            ByRef ablnSeAjusto As Boolean) As Double
        Dim ldblTope = FdblTopeMaximo(adblIbc)
        Dim ldblPermitida = Math.Max(0, Math.Min(adblDigitada, ldblTope))
        ablnSeAjusto = ldblPermitida <> adblDigitada
        Return ldblPermitida
    End Function

    ''' <summary>
    ''' Tasa fija que se guarda como "deseada". Si lo digitado es igual a lo ya guardado se conserva tal cual
    ''' (el tope vigente solo limita lo que se COBRA, nunca reemplaza lo deseado: 24 % que baja a 21 % y
    ''' vuelve a 24 %). Solo un valor NUEVO se limita al tope y avisa.
    ''' </summary>
    Friend Shared Function FdblTasaFijaAGuardar(adblDigitada As Double, adblGuardada As Double, adblIbc As Double,
            ablnHayIbc As Boolean, ByRef ablnSeAjusto As Boolean) As Double
        ablnSeAjusto = False
        If FblnTasaCoincide(adblDigitada, adblGuardada) Then Return adblGuardada
        If Not ablnHayIbc Then Return adblDigitada
        Return FdblTasaFijaPermitida(adblDigitada, adblIbc, ablnSeAjusto)
    End Function

    ''' <summary>
    ''' Factor que se acepta al digitar: nunca mayor al factor maximo. Si se ajusta, avisa.
    ''' </summary>
    Friend Shared Function FdblFactorPermitido(adblDigitado As Double, ByRef ablnSeAjusto As Boolean) As Double
        Dim ldblPermitido = Math.Min(adblDigitado, CDBLFACTORMAXIMO)
        ablnSeAjusto = ldblPermitido <> adblDigitado
        Return ldblPermitido
    End Function

    ''' <summary>
    ''' Interpreta el porcentaje que digita el usuario ("24", "24,5", "24.5 %") como fracción (0.245).
    ''' Acepta coma o punto decimal sin depender de la cultura del hilo.
    ''' </summary>
    Friend Shared Function FblnTryParsePorcentaje(astrTexto As String, ByRef adblFraccion As Double) As Boolean
        Dim ldblPorcentaje As Double
        adblFraccion = 0
        Dim lblnOk = FblnTryParseNumero(astrTexto, ldblPorcentaje)
        If lblnOk Then
            adblFraccion = Math.Round(ldblPorcentaje / 100, 6)
        End If
        Return lblnOk
    End Function

    ''' <summary>
    ''' Interpreta el factor que digita el usuario ("1,3", "1.5") sin depender de la cultura del hilo.
    ''' </summary>
    Friend Shared Function FblnTryParseFactor(astrTexto As String, ByRef adblFactor As Double) As Boolean
        Dim ldblFactor As Double
        adblFactor = 0
        Dim lblnOk = FblnTryParseNumero(astrTexto, ldblFactor)
        If lblnOk Then
            adblFactor = Math.Round(ldblFactor, 4)
        End If
        Return lblnOk
    End Function

    Private Shared Function FblnTryParseNumero(astrTexto As String, ByRef adblNumero As Double) As Boolean
        adblNumero = 0
        If String.IsNullOrWhiteSpace(astrTexto) Then Return False
        Dim lstrTexto = astrTexto.Replace("%", "").Replace(" ", "").Replace(",", ".")
        Dim ldblNumero As Double
        If Double.TryParse(lstrTexto, Globalization.NumberStyles.Float,
                Globalization.CultureInfo.InvariantCulture, ldblNumero) AndAlso
                Not Double.IsInfinity(ldblNumero) AndAlso Not Double.IsNaN(ldblNumero) Then
            adblNumero = ldblNumero
            Return True
        End If
        Return False
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
#Region "Escritura en OriTasasMora"
    ''' <summary>
    ''' Valor que se asigna a ClsTasaMoraDbl.ObjValorPro para guardar la tasa anual indicada: la propiedad
    ''' interpreta lo asignado como tasa MENSUAL vencida y la multiplica por 12 (interés simple).
    ''' </summary>
    Friend Shared Function FdblTasaMensualParaAsignar(adblTasaAnual As Double) As Double
        Return Math.Round(adblTasaAnual / 12, 8)
    End Function

    ''' <summary>Texto de la tasa mensual con punto decimal, sin importar la cultura del hilo.</summary>
    Friend Shared Function FstrTasaMensualInvariante(adblTasaAnual As Double) As String
        Return FdblTasaMensualParaAsignar(adblTasaAnual).ToString("0.########",
                Globalization.CultureInfo.InvariantCulture)
    End Function

    ''' <summary>Indica si la tasa leída de la BD coincide con la esperada (tolerancia 0.000002).</summary>
    Friend Shared Function FblnTasaCoincide(adblEsperada As Double, adblLeida As Double) As Boolean
        Return Math.Abs(adblEsperada - adblLeida) <= 0.000002
    End Function
#End Region
End Class
