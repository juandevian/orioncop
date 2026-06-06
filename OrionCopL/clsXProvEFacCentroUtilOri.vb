Partial Friend Class ClsCentroUtilOriCop
    Friend Property ObjContrasenaAPIEFacStr As ClsContrasenaAPIEFacStr = ObjProveedorEFac.ObjContrasenaAPIEFacStr
    Friend Property ObjIdProveedorEFacEnt As ClsIdProveedorEFacEnt = ObjProveedorEFac.ObjIdProveedorEFacEnt
    Friend Property ObjIdUsuarioProvEFacStr As ClsIdUsuarioProvEFacStr = ObjProveedorEFac.ObjIdUsuarioProvEFacStr
    Friend Property ObjIdCarpetaEFacShr As ClsIdCarpetaShr = ObjProveedorEFac.ObjIdCarpeta
    Friend Property ObjIdCentroUtilEFacShr As ClsIdCentroUtilShr = ObjProveedorEFac.ObjIdCentroUtil
    Friend Property ObjURLStr As ClsURLStr = ObjProveedorEFac.ObjURLStr
    Friend Property ObjSubirFacBln As ClsSubirFacBln = ObjProveedorEFac.ObjSubirFacBln

    Friend Function FblnEsValidoContAPIEFac() As Boolean
        Dim lblnEsValido = True
        ObjProveedorEFac.ObjContrasenaAPIEFacStr.SValide
        If EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If ObjIdProvEFacByt.ObjValorPro <> EnuProveedorEFac.None Then
                lblnEsValido = ObjProveedorEFac.ObjContrasenaAPIEFacStr.BlnEsValido
            End If
        End If
        Return lblnEsValido
    End Function

    Friend Function FblnEsValidoIdProvEFac() As Boolean
        Dim lblnEsValido = True
        ObjProveedorEFac.ObjIdProveedorEFacEnt.SValide()

        If EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If ObjIdProvEFacByt.ObjValorPro <> EnuProveedorEFac.None Then
                lblnEsValido = ObjProveedorEFac.ObjIdProveedorEFacEnt.BlnEsValido
            End If
        End If
        Return lblnEsValido
    End Function

    Friend Function FblnEsValidoIdUsuarioProvEFac() As Boolean
        Dim lblnEsValido = True
        ObjProveedorEFac.ObjIdUsuarioProvEFacStr.SValide()
        If EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If ObjIdProvEFacByt.ObjValorPro <> EnuProveedorEFac.None Then
                lblnEsValido = ObjProveedorEFac.ObjIdUsuarioProvEFacStr.BlnEsValido
            End If
        End If
        Return lblnEsValido
    End Function

    Friend Function FblnEsValidoUrl() As Boolean
        Dim lblnEsValido = True
        ObjProveedorEFac.ObjURLStr.SValide()
        If EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If ObjIdProveedorEFacEnt.ObjValorPro <> EnuProveedorEFac.None Then
                lblnEsValido = ObjProveedorEFac.ObjURLStr.BlnEsValido
            End If
        End If
        Return lblnEsValido
    End Function

    Friend Function FblnEsValidoSubirFac() As Boolean
        Dim lblnEsValido = True
        ObjProveedorEFac.ObjSubirFacBln.SValide()
        If EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If ObjIdProveedorEFacEnt.ObjValorPro <> EnuProveedorEFac.None Then
                lblnEsValido = ObjProveedorEFac.ObjSubirFacBln.BlnEsValido
            End If
        End If
        Return lblnEsValido
    End Function
End Class
