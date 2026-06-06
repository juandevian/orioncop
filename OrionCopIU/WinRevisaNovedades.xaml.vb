Public Class WinRevisarNovedades
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
    ' Variables
    Private ReadOnly MobjObjetoWin As ClsCentroUtilOriCop = GobjParametros
    Private ReadOnly MstrNombreVentana As String = "Revisar novedades"
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuRevisaNovs
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SCargueForma(EnuElementosAdicionalesDef.None, 2,
                Nothing, Nothing, True)
        dtpFechaFinNovs.Focus()
        EnuOperacionEnWin = EnuOperacionEnVentana.cenuModificando
        dtpFechaFinNovs.IsEnabled = True
    End Sub
    Protected Overrides ReadOnly Property StrNombreVentana As String
        Get
            Return MstrNombreVentana
        End Get
    End Property
    Protected Overrides ReadOnly Property EnuIdVentana As EnuIdVentanaDef
        Get
            Return HenuIdVentana
        End Get
    End Property
    Protected Overrides Sub SInicialiceObjeto()
        If IsNothing(ObjObjetoWin) Then
            ObjObjetoWin = MobjObjetoWin
        End If
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        dtpFechaFinNovs.SelectedDate = DateSerial(1900, 1, 1)
        '
        HbttAceptar.TabIndex = 9
        HbttCancelar.TabIndex = 10
    End Sub
    Protected Overrides Sub SMuestreDatos()
        '
    End Sub
    Protected Overrides Sub SValide()
        '
    End Sub
    Protected Overrides Sub SRegistre()
        '
    End Sub
    Protected Overrides Sub SConfigureMenuesPropios()
        '
    End Sub
#End Region
#Region "Procedimientos sobrescritos"
    Protected Overrides Sub SCancele()
        SCerrarClic()
    End Sub
    Protected Overrides Sub SGuarde()
        Dim lblnNoHayError = False
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty
        Try
            lstrMens = "Revisión en proceso!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            Mouse.OverrideCursor = Cursors.Wait
            ClsActualizacionApl.SCorrijaProblemasNovedades(dtpFechaFinNovs.SelectedDate)
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                lstrMens = "Termino el proceso de revisión!"
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
        SFinaliceOperacion()
    End Sub
#End Region
End Class
