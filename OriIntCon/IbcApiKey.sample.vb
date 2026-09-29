' Copie este archivo como IbcApiKey.vb (ignorado por Git) y complete las claves OFUSCADAS con:
'   powershell -File scripts\build\Ofusca-IbcKey.ps1 -Key "<API key en texto plano>"
' Debe ser la API key PLANA (la que se envia en el header X-Api-Key), NO su hash SHA-256.
'
' Produccion: se usa siempre en compilaciones Release.
' API local (Docker): solo en compilaciones Debug y solo si MBLNUSARAPILOCAL = True.
Friend Module MdefIbcApiKey
    Private Const MSTRURLPRODUCCION As String = "https://api-ibc-certificados-sif.onrender.com"
    Private Const MSTRCLAVEPRODOFUSCADA As String = ""
    Private Const MBLNUSARAPILOCAL As Boolean = False
    Private Const MSTRURLLOCAL As String = "http://localhost:5080"
    Private Const MSTRCLAVELOCALOFUSCADA As String = ""

    Private Function FblnUsaApiLocal() As Boolean
#If DEBUG Then
        Return MBLNUSARAPILOCAL
#Else
        Return False
#End If
    End Function

    Friend Function FstrUrlApi() As String
        Return If(FblnUsaApiLocal(), MSTRURLLOCAL, MSTRURLPRODUCCION)
    End Function

    Friend Function FstrApiKey() As String
        Return FstrDesofusque(If(FblnUsaApiLocal(), MSTRCLAVELOCALOFUSCADA, MSTRCLAVEPRODOFUSCADA))
    End Function

    Private Function FstrDesofusque(astrOfuscada As String) As String
        If astrOfuscada.Length = 0 Then Return String.Empty
        Dim lbytDatos = Convert.FromBase64String(astrOfuscada)
        Dim lbytMascara = System.Text.Encoding.UTF8.GetBytes("OrionPlus-IBC")
        For i As Integer = 0 To lbytDatos.Length - 1
            lbytDatos(i) = lbytDatos(i) Xor lbytMascara(i Mod lbytMascara.Length)
        Next
        Return System.Text.Encoding.UTF8.GetString(lbytDatos)
    End Function
End Module
