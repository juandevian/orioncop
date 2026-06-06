Friend Class ClsCalculosPredios
    ''' <summary>
    ''' Calcula el total de la base de calculo para el coeficiente de propiedad y calcula el
    ''' coeficiente de propiedad de cada uno de los predios. Ademas calcula la Base de Participación 
    ''' de los Sectores que conforman la Copropiedad. Actualiza los datos en la BD.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Shared Sub SActualiceBasesParticipacion()
        gobjPanDat.sControleProcesoObj(True)
        Dim ldecTotalBaseParticipacion = 0D
        Try
            Dim lstrCamposSelect() As String
            If gobjCentroUtilOriCop.objIdTipoBaseCalculoCPByt.objValorPro = enuTipoBaseCalculoDef.enuArea Then
                lstrCamposSelect = {"SUM(" & clsBaseParticipacionDec.sstrNombreCampoBd & ")"}
            Else
                lstrCamposSelect = {"COUNT(" & clsIdPredioStr.sstrNombreCampoBd & ")"}
            End If
            Dim ldtbCalculo = clsPanorama.fdtbDataTable(clsPredio.sstrNombreTabla, lstrCamposSelect, Nothing,
                                       clsOrionCop.strFiltroUbicacion)
            Dim ldrwCalculos As DataRow() = ldtbCalculo.Select
            If ldrwCalculos.Length > 0 Then
                Dim ldrwCalculo As DataRow = ldrwCalculos(0)
                ldecTotalBaseParticipacion = Math.Round(ClsPanorama.FobjValorCampo(ldrwCalculo(0), EnuTipoValorDef.enuDecimal), 4)
            End If
            If ldecTotalBaseParticipacion <> gobjCentroUtilOriCop.objTotalBaseCPDec.objValorPro Then
                GobjCentroUtilOriCop.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                GobjCentroUtilOriCop.objTotalBaseCPDec.objValorPro = ldecTotalBaseParticipacion
                gobjCentroUtilOriCop.sActualice(False)
            End If
            sActualiceCoeficientesProp()
            sActualiceBaseParticipacionSectores()
        Catch ex As ProveedorBdPanException
            clsPanorama.sProceseExcepcion(ex, ex.NumeroErr)
        Catch ex As PanDatException
            clsPanorama.sProceseExcepcion(ex, ex.NumeroErr)
        Catch ex As ArgumentNullException
            clsPanorama.sProceseExcepcion(ex)
        Catch ex As Exception
            clsPanorama.sProceseExcepcion(ex)
        Finally
            gobjPanDat.sControleProcesoObj(False)
        End Try
    End Sub

    Friend Shared Sub SActualiceCoeficientesProp()
        Dim lobjOrionCop = New ClsOrionCop(GCOBJREGISTRO, False)
        Dim lobjPredio As ClsPredio = ClsOrionCop.FobjNuevoPredio(EnuModoInstanciaObjDef.enuNavegable)
        lobjPredio.SVayaAlPrimero()
        Dim ldecCoeficientePro = 0D
        Do While lobjPredio.BlnExiste
            With lobjPredio
                ldecCoeficientePro = .ObjBaseParticipacionDec.ObjValorPro /
                        GobjCentroUtilOriCop.ObjTotalBaseCPDec.ObjValorPro
                Select Case GobjCentroUtilOriCop.ObjBaseRedondeoCPByt.ObjValorPro
                    Case 1
                        .ObjCoeficientePropiedadDec.ObjValorPro = ldecCoeficientePro
                    Case Else
                        ldecCoeficientePro = Math.Round(ldecCoeficientePro,
                                GobjCentroUtilOriCop.ObjBaseRedondeoCPByt.ObjValorPro + 2)
                        .ObjCoeficientePropiedadDec.ObjValorPro = ldecCoeficientePro
                End Select
                .SModifique()
                .SActualice(False)
            End With
            lobjPredio.SVayaAlSiguiente()
        Loop
    End Sub

    Friend Shared Sub SActualiceBaseParticipacionSectores()
        Dim lcolSectores = gobjCentroUtilOriCop.colSectores
        If lcolSectores.Count > 0 Then
            Dim lstrSql = String.Empty
            If gobjCentroUtilOriCop.objIdTipoBaseCalculoCPByt.objValorPro = enuTipoBaseCalculoDef.enuArea Then
                lstrSql = "SELECT " & clsIdSector_PredioShr.sstrNombreCampoBd &
                        ", SUM(" & clsBaseParticipacionDec.sstrNombreCampoBd & ") AS Base FROM " &
                        clsPredio.sstrNombreTabla & " WHERE " & clsOrionCop.strFiltroUbicacion &
                        " GROUP BY " & clsIdSector_PredioShr.sstrNombreCampoBd
            Else
                lstrSql = "SELECT " & clsIdSector_PredioShr.sstrNombreCampoBd &
                        ", COUNT(" & clsIdPredioStr.sstrNombreCampoBd & ") AS Base FROM " &
                        clsPredio.sstrNombreTabla & " WHERE " & clsOrionCop.strFiltroUbicacion &
                        " GROUP BY " & clsIdSector_PredioShr.sstrNombreCampoBd
            End If
            Using ldstCalculo As New DataSet
                gobjPanDat.sdsDataSet(ldstCalculo, lstrSql)
                Dim ldtbCalculo = ldstCalculo.Tables(0)
                Dim ldrwSectores() As DataRow = Nothing
                Dim lstrFiltro = String.Empty
                Dim ldecBase = 0D
                For Each lobjSector As clsSector In lcolSectores
                    lstrFiltro = clsIdSector_PredioShr.sstrNombreCampoBd & " = " &
                            lobjSector.objIdSectorShr.ToString
                    ldrwSectores = ldtbCalculo.Select(lstrFiltro)
                    If ldrwSectores.Count > 0 Then
                        ldecBase = clsPanorama.fobjValorCampo(ldrwSectores(0)("Base"),
                                enuTipoValorDef.enuDecimal)
                    Else
                        ldecBase = 0D
                    End If
                    With lobjSector
                        If .enuEstadoActualizacion = enuEstadoObjetoDef.enuConsultando Then
                            .enuEstadoActualizacion = enuEstadoObjetoDef.enuModificando
                        End If
                        .objBaseParticipacionSectorDec.objValorPro = ldecBase
                        .sActualice(False)
                        .enuTipoPermisoObj -= enuTipoPermisoDef.enuModificar
                    End With
                Next
            End Using
        End If
    End Sub
End Class
