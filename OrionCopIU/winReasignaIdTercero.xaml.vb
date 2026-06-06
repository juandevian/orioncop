Imports System.Windows.Controls
Public Class WinReasignaIdTercero
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuIdNuevo
    End Enum
#End Region

    ' Variables
    Private ReadOnly MobjTer As ClsTercero = Nothing
    Private ReadOnly MstrNombreVentana As String = "REASIGNACION ID TERCERO"
    Private MblnDejoUltimoControl As Boolean = False
    Private MdblIdNuevo As Double = 0
    Private ReadOnly MdblIdTerceroActual As Double = 0
    Private ReadOnly MstrNombreTerceroActual As String = String.Empty
#End Region
#Region "Constructor"
    Friend Sub New(aobjTercero As ClsTercero)
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuReasignaId
        MdblIdTerceroActual = aobjTercero.ObjIdTerceroDbl.ObjValorPro
        MstrNombreTerceroActual = aobjTercero.FstrNombreCompleto
        MobjTer = aobjTercero
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SCargueForma(EnuElementosAdicionalesDef.None, 1,
                Nothing, txtIdTerceroNuevo, True)
        SCree()
        txtIdTerceroNuevo.Focus()
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
        ObjObjetoWin = GobjPanorama.ObjUsuarioActual
        EnuTipoPermisoObjWin = GobjPanorama.ObjUsuarioActual.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuIdNuevo) = lblIdTerceroNuevo

        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub
    Protected Overrides Sub SMuestreDatos()
        txtIdTerceroAct.Content = MdblIdTerceroActual.ToString()
        txtNombreTerAct.Content = MstrNombreTerceroActual
        If Not HblnCargandoForma Then
            SValide()
        End If
    End Sub
    Protected Overrides Sub SValide()
        Dim lstrMens = String.Empty
        StcValidValido(EnuValidEntradaDef.enuIdNuevo) = FblnEsValidoIdTerceroNuevo(lstrMens)
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
        '
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        SValide()
    End Sub
    Protected Overrides Sub SConfigureMenuesPropios()
        '
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    Protected Overrides Sub SCree()
        SFrmAdicione()
        EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuCreando
    End Sub
    Protected Overrides Sub SGuarde()
        If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
            Close()
            Exit Sub
        End If
        SRegistre()
        SValide()
        If FblnEstanTodosBien() Then
            Mouse.OverrideCursor = Cursors.Wait
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            If FblnEsValidoIdTerceroNuevo(lstrMens) Then
                Dim lblnOk = MsgBox("Realmente desea asignar esta identificiación al Tercero?", MsgBoxStyle.YesNo,
                                    "Confirmación cambio Id.")
                If lblnOk Then
                    lstrMens = "Reasignado Id. del Tercero!"
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    Try
                        MdblIdNuevo = CType(txtIdTerceroNuevo.Text, Double)
                        ClsPanorama.SReasigneIdTer(MdblIdTerceroActual, MdblIdNuevo)
                        ClsOrionCop.SReasigneAliasCont(MdblIdTerceroActual, MdblIdNuevo)
                        lblnNoHayError = True
                    Catch ex As ProveedorBdPanException
                        lstrMens = ex.Message
                        lstrMensEx = ex.ToString
                    Catch ex As ArgumentOutOfRangeException
                        lstrMens = ex.Message
                        lstrMensEx = ex.ToString
                    Catch ex As PanDatException
                        lstrMens = ex.Message
                        lstrMensEx = ex.ToString
                    Catch ex As PanLException
                        lstrMens = ex.Message
                        lstrMensEx = ex.ToString
                    Catch ex As Exception
                        lstrMens = ex.Message
                        lstrMensEx = ex.ToString
                    Finally
                        If lblnNoHayError Then
                            lstrMens = "Proceso terminado exitosamente"
                            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                            EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando
                        Else
                            SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                            MdblIdNuevo = 0
                            Mouse.OverrideCursor = Cursors.Arrow
                        End If
                    End Try
                Else
                    SRefresqueWin()
                End If
            Else
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                txtIdTerceroNuevo.Text = String.Empty
            End If
        End If
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Protected Overrides Sub SCancele()
        SCerrarClic()
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            ltxtTextBox.SelectAll()
        ElseIf lelmElemento.Equals(HbttCancelar) Then
            If MblnDejoUltimoControl Then
                HbttAceptar.Focus()
            End If
        End If
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        If Not HblnSeEstaCerrando Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is TextBox Then
                MblnDejoUltimoControl = False
                If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
                    If lelmElemento.Name = "txtIdTerceroNuevo" Then
                        SValide()
                        MblnDejoUltimoControl = True
                    End If
                End If
            End If
        End If
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Function FblnEsValidoIdTerceroNuevo(ByRef astrMens As String) As Boolean
        Dim ldblIdTerNue As Object = txtIdTerceroNuevo.Text
        Dim lblnEsValido = ClsPanorama.FblnEsValidoNumero(ldblIdTerNue, GCDBLMINTERC,
                    GCDBLMAXTERC, True)
        If lblnEsValido Then
            MobjTer.SAbra({ldblIdTerNue})
            lblnEsValido = Not MobjTer.BlnExiste
            If Not lblnEsValido Then
                astrMens = "Un Tercero con éste número de Identificación ya existe!"
            End If
        Else
            astrMens = "El Dato ingresado no es valido!"
        End If
        Return lblnEsValido
    End Function
#End Region
End Class
