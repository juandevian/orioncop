Imports System.Runtime.CompilerServices
#Region "Definiciones"
<Assembly: CLSCompliant(True)>
<Assembly: InternalsVisibleTo("OrionCopIU")>
#End Region
#Region "Enumeradores"
Friend Enum EnuTipoDocIdDian As Integer
    None = 0
    enuRegistroCivil = 11
    enuTarjetaIdentidad
    enuCedulaCiudadania
    enuTarjetaExtranjeria = 21
    enuCedulaExtranjeria
    enuNIT = 31
    enuPasaporte = 41
    enuDocIdentidadExtranjero
    enuPEP = 47
    enuNITotroPais = 50
    enuNUIP = 91
End Enum
Friend Enum EnuTipoPersonaDian As Integer
    None = 0
    EnuJuridica
    EnuNatural
End Enum
Friend Enum EnuTipoMedioPagoDian As Integer
    None = 0
    EnuEfectivo = 10
    EnuCheque = 20
    EnuTransferencia = 31
    EnuConsignacion = 42
    EnuTarjetaCR = 48
    EnuTarjetaDB = 49
End Enum
#End Region