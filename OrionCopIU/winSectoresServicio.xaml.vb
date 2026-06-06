Public Class WinSectoresServicio
    Private ReadOnly MobjServicio As ClsServicio = Nothing
    Friend Sub New(aobjServicio As ClsServicio)
        InitializeComponent()
        MobjServicio = aobjServicio
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
        txtAno.Content = MobjServicio.ObjIdAno_ServicioShr.ToString
        txtServicio.Content = MobjServicio.ObjIdServicioShr.ToString & " - " &
                MobjServicio.ObjNombreServicioStr.ObjValorPro
        Dim ldtbSecServicio = MobjServicio.FdtbSectoresServicio()
        Dim ldecTotVlrIni = 0D
        For Each ldrwSectorServicio As DataRow In ldtbSecServicio.Rows
            ldecTotVlrIni += ClsPanorama.FobjValorCampo(ldrwSectorServicio(
                    ClsValor_SectorModuloServicioDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
        Next
        txtValorIni.Content = Format(ldecTotVlrIni, "c")
        dgrSectoresServicio.DataContext = ldtbSecServicio
        SVisibiliceCtlsServicio()
    End Sub
    ''' <summary>
    ''' Visibiliza u oculta algunos controles dependiendo del tipo de servicio (Anual o Permanente)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SVisibiliceCtlsServicio()
        If Not MobjServicio.BlnEsCuotaAdministracion Then
            If dgrSectoresServicio.Items.Count > 0 Then
                Dim ldgcColumna As DataGridColumn = dgrSectoresServicio.Columns(3)
                ldgcColumna.Visibility = Visibility.Collapsed
                Canvas.SetLeft(dgrSectoresServicio, 80)
                Canvas.SetLeft(lblValorIni, 290)
                Canvas.SetLeft(txtValorIni, 425)
            End If
        End If
    End Sub
End Class
