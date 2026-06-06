Friend Class ClsModuloContribucion
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriModulosContribucion"
    ' Variables de modulo
    Private McolSectoresModulo As Collection = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Panorama
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Indica si se instancia como un objeto navegable o como un Objeto único.</param>
    ''' <remarks>Si se instancia como un objeto navegable, se crea un datatable que contiene las columnas de
    ''' la llave con las llaves de todos los objetos y queda a la espera de que se indique que objeto abrir.
    ''' Si se instancia como un objeto único, queda a la espera de recibir el valor de los campos de la llave 
    ''' para abrir dicho objeto. </remarks>
    Public Sub New(aenuModoInstanciaObj As EnuModoInstanciaObjDef)
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar 
                    un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = Nothing
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
            lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                                ClsIdModuloShr.SstrNombreCampoBd}
            HcolFiltros.Add(ClsOrionCop.strFiltroUbicacion)
        Else
            hblnEsCreable = False
            hblnEsModificable = False
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        HblnEsAnulable = False
        HblnEsSuprimible = False
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsCentroUtilOriCop, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        HblnEsSuprimible = False
        '
        DrwRegistroActual = adrwObjeto
        DtbTablaColeccion = DrwRegistroActual.Table
    End Sub
#End Region
#Region "Propiedades"
#Region "Propiedades indentificadoras"
    Protected Overrides ReadOnly Property HstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Protected Overrides ReadOnly Property HenuIdClase As EnuIdClasesPanDef
        Get
            Return EnuIdClasesPanDef.enuModuloContribucion
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Módulo Contribución"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & ObjNombreModuloStr.ObjValorPro & Chr(34)
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjContribuyeCuotaAdminBln As New ClsContribuyeCuotaAdminBln(Me)
    Friend ReadOnly Property ObjIdCarpetaModuloShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtilModuloShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdModuloShr As New ClsIdModuloShr(Me)
    Friend ReadOnly Property ObjNombreModuloStr As New ClsNombreModuloStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjContribuyeCuotaAdminBln)
                HcolPropiedades.Add(ObjIdCarpetaModuloShr)
                HcolPropiedades.Add(ObjIdCentroUtilModuloShr)
                HcolPropiedades.Add(ObjIdModuloShr)
                HcolPropiedades.Add(ObjNombreModuloStr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    '
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.sVacie()
        McolSectoresModulo = Nothing
    End Sub
    Protected Overrides Sub SCreeObj(aobjValorLlave() As Object)
        MyBase.SCreeObj(aobjValorLlave)
        ObjIdModuloShr.SValide()
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        GobjPanDat.SControleProcesoObj(True)
        Try
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                SNumereObj()
                ObjIdCarpetaModuloShr.ObjValorPro = GshrIdCarpeta
                ObjIdCentroUtilModuloShr.ObjValorPro = GshrIdCentroUtil
            End If
            MyBase.SActualice(ablnExigeRequeridos)
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            GobjPanDat.SControleProcesoObj(False)
        End Try
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdModuloShr.ToString
        End Get
    End Property
    Protected Overrides Function FblnSuprimio() As Boolean
        Dim lblnSuprimio = FblnEsSuprimible()
        If lblnSuprimio Then
            lblnSuprimio = MyBase.FblnSuprimio()
            If lblnSuprimio AndAlso BlnEsNavegable Then
                GobjPanorama.SRegistreAccionLogApp(HstrNombreClase, "Suprimir Módulo de Contribución " &
                    ObjIdModuloShr.ToString & "-" & ObjNombreModuloStr.ToString)
            End If
        End If
        Return lblnSuprimio
    End Function
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lblnEsSuprimible = MyBase.FblnPermitidoSuprimir()
        If lblnEsSuprimible Then
            Dim lstrCondicion As String = " = " & ObjIdModuloShr.ToString & " AND " &
                    ClsOrionCop.StrFiltroUbicacion
            lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg({SstrNombreTabla},
                    ObjIdModuloShr.StrNombreCampoBD, lstrCondicion, True, False)
        End If
        Return lblnEsSuprimible
    End Function
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lshrIdModulo As Short
            Dim lstrFiltro = ClsOrionCop.strFiltroUbicacion
            lshrIdModulo = ClsPanorama.FobjUltimaIdNumericaObjeto(sstrNombreTabla, ClsIdModuloShr.sstrNombreCampoBd,
                    ObjIdModuloShr.EnuTipoValor, lstrFiltro) + 1
            ObjIdModuloShr.ObjValorPro = lshrIdModulo
        End If
    End Sub
    Friend Sub SAgregueTodosSectores()
        Try
            If IsNothing(McolSectoresModulo) Then
                McolSectoresModulo = ColSectoresModulo
            End If
            Dim lcolSectores As Collection = GobjParametros.ColSectores
            Dim ldtbSectoresModulo = FdtbSectoresModulo()
            For Each lobjSector As ClsSector In lcolSectores
                Dim lshrIdSector As Short = lobjSector.ObjIdSectorShr.ObjValorPro
                If Not McolSectoresModulo.Contains(lshrIdSector.ToString) Then
                    SAdicioneSector(lobjSector, ldtbSectoresModulo)
                End If
            Next
            ClsPanorama.SActualiceCol(McolSectoresModulo)
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Private Sub SAdicioneSector(aobjSector As ClsSector, adtbSectoresModulo As DataTable)
        Dim ldrwNuevoSectorModulo As DataRow = adtbSectoresModulo.NewRow
        Dim lobjSectorModulo = New ClsSectorModulo(Me, ldrwNuevoSectorModulo)
        With lobjSectorModulo
            .SCreeObj(Nothing)
            .ObjIdCarpeta_SectorModuloShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_SectorModuloShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdModulo_SectorModuloShr.ObjValorPro = ObjIdModuloShr.ObjValorPro
            .ObjIdSector_SectorModuloShr.ObjValorPro = aobjSector.ObjIdSectorShr.ObjValorPro
            .ObjTasaContribucionDbl.ObjValorPro = 1
        End With
        Dim lstrKey As String = lobjSectorModulo.ObjIdSector_SectorModuloShr.ToString
        McolSectoresModulo.Add(lobjSectorModulo, lstrKey)
    End Sub
#End Region
#Region "Manejo SectoresModulo"
    Friend ReadOnly Property ColSectoresModulo As Collection
        Get
            If IsNothing(McolSectoresModulo) Then
                McolSectoresModulo = New Collection
                Dim ldtbSectoresModulo = FdtbSectoresModulo()
                For Each ldrwSectorMod As DataRow In ldtbSectoresModulo.Rows
                    Dim lobjSectorModulo As New ClsSectorModulo(Me, ldrwSectorMod)
                    lobjSectorModulo.SLeaValores(True)
                    Dim lstrKey As String = lobjSectorModulo.ObjIdSector_SectorModuloShr.ToString
                    McolSectoresModulo.Add(lobjSectorModulo, lstrKey)
                Next
            End If
            Return McolSectoresModulo
        End Get
    End Property
    Friend Function FdblBaseTotalParticipaModulo(
            aenuTipoBaseCalculo As EnuTipoBaseCalculo) As Double
        Dim ldblBaseTotalCalModulo As Double = 0
        Dim lcolSectores = GobjParametros.ColSectores
        Dim lobjSector As ClsSector
        For Each lobjSectorModulo As ClsSectorModulo In ColSectoresModulo
            lobjSector = lcolSectores(lobjSectorModulo.ObjIdSector_SectorModuloShr.ToString)
            ldblBaseTotalCalModulo +=
                        lobjSector.DblBaseParticipacionSector(aenuTipoBaseCalculo) *
                        lobjSectorModulo.ObjTasaContribucionDbl.ObjValorPro
        Next
        Return ldblBaseTotalCalModulo
    End Function
    Friend Function FdblTasaParticipacionSector(ashrIdSector As Short,
            aenuBaseCalculo As EnuTipoBaseCalculo) As Double
        Dim ldblTasaPartSec = 0.0
        Dim lblnAreaPonderada = GobjParametros.FblnAreaPonderada
        Dim lcolSectores = GobjParametros.ColSectores
        Dim ldblBaseTotalCalModulo = FdblBaseTotalParticipaModulo(aenuBaseCalculo)
        If lcolSectores.Contains(ashrIdSector.ToString) Then
            Dim lobjSector As ClsSector = lcolSectores(ashrIdSector.ToString)
            Dim ldblTasaContribucionSec = FdblTasaContribucionSector(ashrIdSector)
            If ldblBaseTotalCalModulo > 0 Then
                Dim ldblBaseParticiSec As Double
                ldblBaseParticiSec = lobjSector.DblBaseParticipacionSector(
                        aenuBaseCalculo) * ldblTasaContribucionSec
                ldblTasaPartSec = ldblBaseParticiSec / ldblBaseTotalCalModulo
            Else
                Throw New ErrorInesperadoPanLException("Total Area de Contribucion del Modulo " &
                        ObjNombreModuloStr.ObjValorPro & " es igual cero!")
            End If
        End If
        Return ldblTasaPartSec
    End Function
    Private Function FdblTasaContribucionSector(ashrIdSector As Short) As Double
        Dim ldblTasaContr = 0.0
        If IsNothing(McolSectoresModulo) Then
            McolSectoresModulo = ColSectoresModulo
        End If
        If McolSectoresModulo.Contains(ashrIdSector.ToString) Then
            Dim lobjSectModu As ClsSectorModulo = McolSectoresModulo(ashrIdSector.ToString)
            ldblTasaContr = lobjSectModu.ObjTasaContribucionDbl.ObjValorPro
        End If
        Return ldblTasaContr
    End Function
    Friend Function FdtbSectoresModulo() As DataTable
        Dim lstrCamposSelect() As String = {StrCampoCarpeta,
                                            StrCampoCentroUtil,
                                            ClsIdModulo_SectorModuloShr.SstrNombreCampoBd,
                                            ClsIdSector_SectorModuloShr.SstrNombreCampoBd,
                                            "'*' AS NombreSector", ClsTasaContribucionDbl.SstrNombreCampoBd}
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdModulo_SectorModuloShr.SstrNombreCampoBd & " = " & ObjIdModuloShr.ObjValorPro
        Dim lstrIndice(,) As String = {{StrCampoCarpeta, "ASC"},
                                           {StrCampoCentroUtil, "ASC"},
                                           {ClsIdModulo_SectorModuloShr.SstrNombreCampoBd, "ASC"},
                                           {ClsIdSector_SectorModuloShr.SstrNombreCampoBd, "ASC"}}
        Dim ldtbSectoresModulo = ClsPanorama.FdtbDataTable(ClsSectorModulo.SstrNombreTabla,
                    lstrCamposSelect, lstrIndice, lstrFiltro)
        SRepuebleNombresSectores(ldtbSectoresModulo)
        Return ldtbSectoresModulo
    End Function
    Private Sub SRepuebleNombresSectores(adtbSectoresModulo As DataTable)
        Dim lcolSectores As Collection = GobjParametros.ColSectores
        Dim lobjSector As ClsSector
        For Each ldrwSecMod As DataRow In adtbSectoresModulo.Rows
            lobjSector = lcolSectores(ldrwSecMod(ClsIdSector_SectorModuloShr.SstrNombreCampoBd))
            Dim lstrNomSec As String = lobjSector.ObjNombreSectorStr.ObjValorPro
            ldrwSecMod("NombreSector") = lstrNomSec
        Next
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsContribuyeCuotaAdminBln
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ContribuyeCuotaAdmin"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Contribuye a Cuota de Administración"
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HenuTipoValor = EnuTipoValor.enuBoolean
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsIdModuloShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdModuloContribucion"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdModuloContribucion"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Byte.MaxValue,
                        BlnEsRequerido, EnuTipoValor)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsNombreModuloStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Nombre"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Nombre Tipo Predio"
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HshrLongitud = 40
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, HshrLongitud,
                HblnEsRequerido)
        If HblnEsValido Then
            HobjValorNew = HobjValorNew.ToString.ToUpper
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
#End Region
