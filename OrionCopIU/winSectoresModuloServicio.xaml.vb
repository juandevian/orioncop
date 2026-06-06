Public Class WinSectoresModuloServicio
    Private ReadOnly MobjModuloServicio As ClsModuloServicio = Nothing
    Friend Sub New(aobjModuloServicio As ClsModuloServicio)
        InitializeComponent()
        MobjModuloServicio = aobjModuloServicio
    End Sub
    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        Dim lobjCarAct = GobjPanorama.ObjCarpetaActual
        Dim lobjCenUtil = lobjCarAct.ObjCentroUtilidadActual
        txtCarpeta.Content = lobjCarAct.ObjIdCarpetaShr.ToString & " - " &
                lobjCarAct.ObjNombreStr.ObjValorPro
        txtCentroUtilidad.Content = lobjCenUtil.ObjIdCarpetaCenUtilShr.ToString & " - " &
                lobjCenUtil.ObjNombreCentroUtilStr.ObjValorPro
        txtAno.Content = MobjModuloServicio.ObjIdAno_ModuloServicioShr.ToString
        Dim lobjPadreModSer As ClsServicio = MobjModuloServicio.ObjPadre
        txtServicio.Content = lobjPadreModSer.ObjIdServicioShr.ToString & " - " &
                lobjPadreModSer.ObjNombreServicioStr.ObjValorPro
        Dim lobjModuloCont As ClsModuloContribucion =
                GobjParametros.ColModulos(MobjModuloServicio.ObjIdModulo_ModuloServicioShr.ToString)
        txtModulo.Content = lobjModuloCont.ObjIdModuloShr.ToString & " - " &
                lobjModuloCont.ObjNombreModuloStr.ObjValorPro
        Dim ldtbSecModSer = MobjModuloServicio.FdtbSectoresModuloServicio
        Dim ldecTotVlrIni = 0D
        For Each ldrwSectorModSer As DataRow In ldtbSecModSer.Rows
            ldecTotVlrIni += ClsPanorama.FobjValorCampo(ldrwSectorModSer(
                    ClsValor_SectorModuloServicioDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
        Next
        txtValorIni.Content = Format(ldecTotVlrIni, "c")
        dgrSectoresModulo.DataContext = ldtbSecModSer
        SVisibiliceCtlsServicio()
    End Sub
    ''' <summary>
    ''' Visibiliza u oculta algunos controles dependiendo del tipo de servicio (Anual o Permanente)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SVisibiliceCtlsServicio()
        Dim lshrIdAno As Short = MobjModuloServicio.ObjIdAno_ModuloServicioShr.ObjValorPro
        If lshrIdAno = 0 AndAlso dgrSectoresModulo.Items.Count > 0 Then
            lblValorDef.Visibility = Visibility.Collapsed
            txtValorDef.Visibility = Visibility.Collapsed
            Dim ldgcColumna As DataGridColumn = dgrSectoresModulo.Columns(3)
            ldgcColumna.Visibility = Visibility.Collapsed
            Canvas.SetLeft(dgrSectoresModulo, 75)
            Canvas.SetLeft(lblValorIni, 320)
            Canvas.SetLeft(txtValorIni, 420)
        End If
    End Sub
End Class
