Imports System.Windows.Controls
Public Class WinOpcionesCopiasSeg
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuTray = 0
        enuHoraMin
    End Enum
#End Region
    ' Variables
    Private MblnDejoUltimoControl As Boolean = False
    Private MobjObjetoWin As ClsAplicacion = Nothing
    Private ReadOnly MstrNombreVentana As String = "CONFIGURAR COPIAS SEGURIDAD"

    Private MentHoras As Integer = 0
    Private MentMinutos As Integer = 0
    Private MblnPoblandoCbo As Boolean = False
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuOpcionesBK
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SCargueForma(EnuElementosAdicionalesDef.None, 2,
                Nothing, txtTrayCopiaSeg, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
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
            ObjObjetoWin = ClsAdministrador.FobjAppActual
        End If
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuTray) = lblTrayCopiaSeg
        StcValidaControl(EnuValidEntradaDef.enuHoraMin) = lblHoraBK
        SPuebleCboDias()
        SActiveBkProg()
        '
        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub
    Protected Overrides Sub SMuestreDatos()
        With MobjObjetoWin
            If String.IsNullOrEmpty(.ObjTrayCopiaSeguridadStr.ObjValorPro) Then
                txtTrayCopiaSeg.Text = .StrTrayCopiaSeguridad
            Else
                txtTrayCopiaSeg.Text = .ObjTrayCopiaSeguridadStr.ObjValorPro
            End If
            chkActivaProgBK.IsChecked = .ObjActivaProgramaBKBln.ObjValorPro
            cboDiaBK.SelectedIndex = .ObjDiaCopiaSeguridadEnt.ObjValorPro
            MentHoras = .ObjHoraCopiaSeguridadEnt.ObjValorPro
            MentMinutos = .ObjMinutosCopiaSeguridadEnt.ObjValorPro
            txtHoraBk.Text = Format(MentHoras, "00") & ":" & Format(MentMinutos, "00")
        End With
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        With MobjObjetoWin
            StcValidValido(EnuValidEntradaDef.enuTray) = .ObjTrayCopiaSeguridadStr.BlnEsValido
            StcValidValido(EnuValidEntradaDef.enuHoraMin) = .ObjHoraCopiaSeguridadEnt.BlnEsValido AndAlso
                    .ObjMinutosCopiaSeguridadEnt.BlnEsValido
        End With
        '
        SHabiliteBotonesTlb()
        SHabiliteBotonTlb(True, HbttModificar)
        SHabiliteMenuItem(True, HmnuModificar)
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjActivaProgramaBKBln.ObjValorPro = chkActivaProgBK.IsChecked
            .ObjTrayCopiaSeguridadStr.ObjValorPro = txtTrayCopiaSeg.Text
            .ObjHoraCopiaSeguridadEnt.ObjValorPro = MentHoras
            .ObjMinutosCopiaSeguridadEnt.ObjValorPro = MentMinutos
        End With
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        '
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    Protected Overrides Sub SHabiliteMenues()
        Dim lblnHabilitarModif As Boolean = FblnHabilitarMenuPan(1)
        If Not lblnHabilitarModif Then
            HmnuModificar.IsEnabled = False
            HbttModificar.IsEnabled = False
        End If
    End Sub
    Protected Overrides Sub SModifique()
        MyBase.SModifique()
        SActiveBkProg()
        txtTrayCopiaSeg.Focus()
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            If lelmElemento.Name = "bttExaminar" Then
                SSeleccioneTray()
            End If
        ElseIf TypeOf lelmElemento Is CheckBox Then
            If lelmElemento.Name = "chkActivaProgBK" Then
                MobjObjetoWin.ObjActivaProgramaBKBln.ObjValorPro = chkActivaProgBK.IsChecked
                SActiveBkProg()
            End If
        End If
    End Sub

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
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    Select Case lelmElemento.Name
                        Case "txtTrayCopiaSeg"
                            MobjObjetoWin.ObjTrayCopiaSeguridadStr.ObjValorPro = txtTrayCopiaSeg.Text
                        Case "txtHoraBk"
                            SRegistreHoraMin()
                    End Select
                    SMuestreDatos()
                End If
            End If
        End If
    End Sub
    Private Sub OnCambioSeleccion(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuModificando Then
            GobjPanDat.SControleProcesoObj(True)
            If Not MblnPoblandoCbo Then
                If TypeOf lelmElemento Is ComboBox Then
                    If lelmElemento.Name = "cboDiaBK" Then
                        MobjObjetoWin.ObjDiaCopiaSeguridadEnt.ObjValorPro = cboDiaBK.SelectedIndex
                    End If
                End If
            End If
            GobjPanDat.SControleProcesoObj(False)
        End If
    End Sub
#End Region
#Region "Registra entradas"
    Private Sub SRegistreHoraMin()
        Dim lblnEsValido = True
        Dim lstrHora = txtHoraBk.Text
        Dim lstrCompo() As String = Array.Empty(Of String)
        If lstrHora.Contains(":") Then
            lstrCompo = lstrHora.Split(":")
            lblnEsValido = (IsNumeric(lstrCompo(0)) AndAlso
                    Val(lstrCompo(0)) >= 0 AndAlso Val(lstrCompo(0)) < 24)
            If lblnEsValido Then
                lblnEsValido = (IsNumeric(lstrCompo(1)) AndAlso
                    Val(lstrCompo(1)) >= 0 AndAlso Val(lstrCompo(1)) < 59)
            End If
        Else
            lblnEsValido = False
        End If
        If lblnEsValido Then
            MentHoras = CType(lstrCompo(0), Integer)
            MentMinutos = CType(lstrCompo(1), Integer)
        End If
        With MobjObjetoWin
            .ObjHoraCopiaSeguridadEnt.ObjValorPro = MentHoras
            .ObjMinutosCopiaSeguridadEnt.ObjValorPro = MentMinutos
        End With
        SMuestreDatos()
        StcValidValido(EnuValidEntradaDef.enuHoraMin) = lblnEsValido
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SSeleccioneTray()
        Dim lfbdOrigenDatos As New System.Windows.Forms.FolderBrowserDialog With {
            .Description = "Seleccione la Carpeta dondea serán guardadas las Copias de Seguridad.",
            .RootFolder = Environment.SpecialFolder.MyComputer
        }
        Dim lblnOk As Boolean = lfbdOrigenDatos.ShowDialog
        If lblnOk Then
            txtTrayCopiaSeg.Text = lfbdOrigenDatos.SelectedPath
            MobjObjetoWin.ObjTrayCopiaSeguridadStr.ObjValorPro = txtTrayCopiaSeg.Text
            SMuestreDatos()
        End If
    End Sub
    Private Sub SActiveBkProg()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            cboDiaBK.IsEnabled = chkActivaProgBK.IsEnabled
            txtHoraBk.IsEnabled = chkActivaProgBK.IsEnabled
        End If
    End Sub
    Private Sub SPuebleCboDias()
        MblnPoblandoCbo = True
        cboDiaBK.Items.Clear()
        cboDiaBK.Items.Add("Domingo")
        cboDiaBK.Items.Add("Lunes")
        cboDiaBK.Items.Add("Martes")
        cboDiaBK.Items.Add("Miercoles")
        cboDiaBK.Items.Add("Jueves")
        cboDiaBK.Items.Add("Viernes")
        cboDiaBK.Items.Add("Sabado")
        MblnPoblandoCbo = False
    End Sub
#End Region
End Class
