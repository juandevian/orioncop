' Copia este archivo a IbcApiKey.vb (ignorado por Git) y reemplaza MSTRCLAVEOFUSCADA
' con el resultado de: powershell -File scripts\build\Ofusca-IbcKey.ps1 -Key "<tu key>"
Friend Module MdefIbcApiKey
    Private Const MSTRCLAVEOFUSCADA As String = ""
    Friend Function FstrApiKey() As String
        If MSTRCLAVEOFUSCADA.Length = 0 Then Return String.Empty
        Dim lbytDatos = Convert.FromBase64String(MSTRCLAVEOFUSCADA)
        Dim lbytMascara = System.Text.Encoding.UTF8.GetBytes("OrionPlus-IBC")
        For i As Integer = 0 To lbytDatos.Length - 1
            lbytDatos(i) = lbytDatos(i) Xor lbytMascara(i Mod lbytMascara.Length)
        Next
        Return System.Text.Encoding.UTF8.GetString(lbytDatos)
    End Function
End Module
