Friend Class ClsCorreoE
#Region "Definiciones"
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriCorreos"
    Private Const V As Boolean = True
    ' Campos de propiedad
    Private MblnFactAuto As Boolean = False
    Private MdblIdCliente As Double = 0 ' Cero indica todos los clientes posibles
    Private MdtmFecFin As Date = GCDTMFECHANULA
    Private MdtmFecIni As Date = GCDTMFECHANULA
    Private MentDiasCobroPers As Integer = 0
    Private MenuTipoCorreo As EnuTipoCorreoE = EnuTipoCorreoE.None
    Private MstrArchivoExterno As String = String.Empty
    Private MstrAsunto As String = String.Empty
    Private MstrIdPredAgru As String = "***" ' Vacio indica sin predio agrupador
    Private MstrMensaje As String = String.Empty
    ' Variables de validacion
    Private MblnFactAutoOk As Boolean = False
    ' Variables
    Private MstrPrefijoDoc As String = String.Empty
    Private MentIdDoc As Integer = 0
    Private MstrNroDoc As String = String.Empty
    Private ReadOnly Property MobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
    Private ReadOnly Property MobjPredioAgr As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
    Private MobjDocumento As ClsCBObjetoPan = Nothing
    ' Conección
    Private ReadOnly MobjPanDat As New ClsPanoramaDat(GCOBJREGISTRO)
#End Region

#Region "Constructores"
    Sub New()
        '
    End Sub
#End Region

#Region "Procedimientos"
    Friend Sub SVacie()
        MobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, 0})
        MblnFactAuto = False
        MdblIdCliente = -1 ' Cero indica todos los clientes posibles
        MdtmFecFin = GCDTMFECHANULA
        MdtmFecIni = GCDTMFECHANULA
        MstrArchivoExterno = String.Empty
        MentDiasCobroPers = 0
        MstrAsunto = String.Empty
        MstrIdPredAgru = "***" ' Vacio indica sin predio agrupador
        MstrMensaje = String.Empty
    End Sub
#End Region

#Region "Propiedades"
    Friend Shared ReadOnly Property SstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Private Function FcolCamposMail() As Collection
        Dim lcolCamMail As New Collection From {
            "ArchivoExt",
            "Asunto",
            "DiasPersuasivo",
            "Fecha",
            "FechaEnvio",
            "IdCarpeta",
            "IdCentroUtil",
            "IdTerceroCliente",
            "IdTipoCorreo",
            "IdUsuario",
            "Mensaje",
            "NumeroDoc",
            "Ordinal"
        }
        Return lcolCamMail
    End Function
    Friend Property BlnFactAuto As Boolean
        Get
            Return MblnFactAuto
        End Get
        Set(value As Boolean)
            MblnFactAuto = value
        End Set
    End Property
    Friend Property EnuTipoCorreo As EnuTipoCorreoE
        Get
            Return MenuTipoCorreo
        End Get
        Set(value As EnuTipoCorreoE)
            MenuTipoCorreo = value
        End Set
    End Property
    ' Si es string vacio indica sin predio agrupador o a todos
    Friend Property StrIdPredioAgrupador As String
        Get
            Return MstrIdPredAgru
        End Get
        Set(value As String)
            MstrIdPredAgru = value
        End Set
    End Property
    Friend Property StrAsunto As String
        Get
            Return MstrAsunto
        End Get
        Set(value As String)
            MstrAsunto = value
        End Set
    End Property
    Friend Property StrMensaje As String
        Get
            Return MstrMensaje
        End Get
        Set(value As String)
            MstrMensaje = value
        End Set
    End Property
    Friend Property StrArchivoExterno As String
        Get
            Return MstrArchivoExterno
        End Get
        Set(value As String)
            MstrArchivoExterno = value
        End Set
    End Property
    Friend Property DtmFechaIni As Date
        Get
            Return MdtmFecIni
        End Get
        Set(value As Date)
            MdtmFecIni = value
        End Set
    End Property
    Friend Property DtmFechaFin As Date
        Get
            Return MdtmFecFin
        End Get
        Set(value As Date)
            MdtmFecFin = value
        End Set
    End Property
    Friend Property EntDiasCobroPers As Integer
        Get
            Return MentDiasCobroPers
        End Get
        Set(value As Integer)
            MentDiasCobroPers = value
        End Set
    End Property
    ' Si el id del cliente es cero indica a todos los clientes posibles
    Friend Property DblIdCliente As Object
        Get
            Return MdblIdCliente
        End Get
        Set(value As Object)
            Dim lstrValor As String = value.ToString
            If String.IsNullOrEmpty(lstrValor) Then
                MdblIdCliente = 0
            ElseIf IsNumeric(lstrValor) Then
                MdblIdCliente = CType(lstrValor, Double)
            End If
            SAbraCliente()
        End Set
    End Property
    Friend Shared ReadOnly Property StrNombreClase As String
        Get
            Return "Correo Electrónico"
        End Get
    End Property
    Friend ReadOnly Property ObjCliente As ClsCliente
        Get
            Return MobjCliente
        End Get
    End Property
    Friend ReadOnly Property StrNombreCliente As String
        Get
            Dim lstrNomCli = String.Empty
            If Not IsNothing(MobjCliente) AndAlso MobjCliente.BlnExiste Then
                lstrNomCli = MobjCliente.ObjNombreCompletoStr.ToString()
            End If
            Return lstrNomCli
        End Get
    End Property
    Friend ReadOnly Property ObjPredioAgr As ClsPredio
        Get
            Return MobjPredioAgr
        End Get
    End Property
    Friend Property StrNroDocumento As String
        Get
            Return MstrNroDoc
        End Get
        Set(value As String)
            If Not String.IsNullOrEmpty(value) Then
                MstrNroDoc = value
                MstrPrefijoDoc = ClsPanorama.FstrPrefijoDcto(MstrNroDoc)
                MentIdDoc = ClsPanorama.FentIdDcto(MstrNroDoc)
            End If
        End Set
    End Property
#End Region

#Region "Métodos"
    Private Sub SAbraCliente()
        Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, MdblIdCliente}
        MobjCliente.SAbra(lobjValorLlave)
    End Sub
    Friend Function FenuTipoDocOrigen() As EnuTipoDocOri
        Dim lenuTipoDoc As EnuTipoDocOri = EnuTipoDocOri.None
        Select Case EnuTipoCorreo
            Case EnuTipoCorreoE.EnuFac, EnuTipoCorreoE.EnuFactAuto
                lenuTipoDoc = EnuTipoDocOri.EnuFactura
            Case EnuTipoCorreoE.EnuNAA
                lenuTipoDoc = EnuTipoDocOri.EnuNotaCon
            Case EnuTipoCorreoE.EnuNCR
                lenuTipoDoc = EnuTipoDocOri.EnuNotaCr
            Case EnuTipoCorreoE.EnuNDB
                lenuTipoDoc = EnuTipoDocOri.EnuNotaDb
            Case EnuTipoCorreoE.EnuRC, EnuTipoCorreoE.EnuRecibos
                lenuTipoDoc = EnuTipoDocOri.EnuReciboCaja
        End Select
        Return lenuTipoDoc
    End Function
    ''' <summary>
    ''' Indica si hay pendientes por enviar correos en los dos últimos días
    ''' </summary>
    ''' <returns></returns>
    Friend Function FblnHayCorreoPorEnviar() As Boolean
        Dim ldtbClientesAEnviar As DataTable, lblnHay = False
        Dim lstrTabla = ClsCorreoE.SstrNombreTabla
        Dim lstrCampSel As String() = {"DISTINCT ArchivoExt", "Asunto", "IdTipoCorreo", "Mensaje",
                "DiasPersuasivo"}
        Dim lstrFechaIni = "'" & ClsPanoramaDat.FstrFechaNormalizada(
                Date.Today.AddDays(-1).ToString()) & "'"
        Dim lstrFechaFin = "'" & ClsPanoramaDat.FstrFechaNormalizada(
                Date.Now.AddDays(1).ToString()) & "'"
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND FechaEnvio BETWEEN " &
                lstrFechaIni & " AND " & lstrFechaFin & " AND (IdTipoCorreo = " &
                EnuTipoCorreoE.EnuFactAuto &
                " OR IdTipoCorreo = " & EnuTipoCorreoE.EnuArchExt &
                " OR IdTipoCorreo = " & EnuTipoCorreoE.EnuCobroPers &
                " OR IdTipoCorreo = " & EnuTipoCorreoE.EnuSoloMens & ")"
        Dim lstrOrden As String(,) = {{"IdTipoCorreo", "ASC"}}
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSel, lstrOrden, lstrFiltro,
                False, {})
        If ldtbRes.Rows.Count > 0 Then
            For Each ldrwRes As DataRow In ldtbRes.Rows
                EnuTipoCorreo = ClsPanorama.FobjValorCampo(ldrwRes("IdTipoCorreo"),
                        EnuTipoValor.EnuByte)
                StrAsunto = ClsPanorama.FobjValorCampo(ldrwRes("Asunto"), EnuTipoValor.EnuString)
                If EnuTipoCorreo = EnuTipoCorreoE.EnuFactAuto Then
                    StrArchivoExterno = String.Empty
                    EntDiasCobroPers = 0
                ElseIf EnuTipoCorreo = EnuTipoCorreoE.EnuArchExt Then
                    StrArchivoExterno = ClsPanorama.FobjValorCampo(ldrwRes("ArchivoExt"),
                            EnuTipoValor.EnuString)
                    EntDiasCobroPers = 0
                ElseIf EnuTipoCorreo = EnuTipoCorreoE.EnuCobroPers Then
                    StrArchivoExterno = String.Empty
                    EntDiasCobroPers = ClsPanorama.FobjValorCampo(ldrwRes("DiasPersuasivo"),
                            EnuTipoValor.EnuInteger)
                End If
                StrMensaje = ClsPanorama.FobjValorCampo(ldrwRes("Mensaje"), EnuTipoValor.EnuString)
                ldtbClientesAEnviar = FdtbClientesAEnviar(True)
                lblnHay = ldtbClientesAEnviar.Rows.Count > 0
                If EnuTipoCorreo = EnuTipoCorreoE.EnuFactAuto Then
                    Exit For
                End If
                If lblnHay Then Exit For
            Next
        End If
        Return lblnHay
    End Function
    Friend Function FdtbClientesAEnviar(ablnDesdeUltimo As Boolean) As DataTable
        Dim ldblUltCli As Double
        If ablnDesdeUltimo Then
            ldblUltCli = FdblUltimoClienteEnviado(EnuTipoCorreo)
        Else
            ldblUltCli = 0
        End If
        Dim ldtbClientesAEnviar As DataTable
        If EnuTipoCorreo = EnuTipoCorreoE.EnuCobroPers Then
            ldtbClientesAEnviar = ClsOrionCop.FdtbPrediosAgrMorosos(EntDiasCobroPers, ldblUltCli)
        Else
            ldtbClientesAEnviar = FdtbClientesConCorreo(ldblUltCli)
        End If
        Return ldtbClientesAEnviar
    End Function
    Friend Function FdblUltimoClienteEnviado(aenuTipoCorreo As EnuTipoCorreoE) As Double
        Dim lstrTabla = ClsCorreoE.SstrNombreTabla
        Dim lstrCampSel As String() = {"*"}
        Dim lstrOrden As String(,) = {{"Ordinal", "DESC"}}
        Dim lstrFechaLimite = Date.Today.AddDays(-1).ToString()
        lstrFechaLimite = ClsPanoramaDat.FstrFechaHoraNormalizada(lstrFechaLimite)
        Dim lstrExpSql As String
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND IdTipoCorreo = " & aenuTipoCorreo &
                " AND FechaEnvio >= '" & lstrFechaLimite & "'"
        If aenuTipoCorreo = EnuTipoCorreoE.EnuArchExt Then
            Dim lstrArchExt = StrArchivoExterno.Replace("\", "\\")
            lstrFiltro &= " AND ArchivoExt = '" & lstrArchExt & "'"
        ElseIf aenuTipoCorreo = EnuTipoCorreoE.EnuSoloMens Then
            lstrFiltro &= " AND Mensaje = '" & StrMensaje & "'"
        ElseIf aenuTipoCorreo = EnuTipoCorreoE.EnuCobroPers Then
            lstrFiltro &= " AND DiasPersuasivo = " & EntDiasCobroPers
        End If
        lstrExpSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCampSel, lstrOrden,
                    lstrFiltro, {})
        lstrExpSql &= " LIMIT 1"
        Dim ldblUltCli As Double = 0
        Dim ldtbUltimoEnviado As DataTable = ClsPanorama.FdtbDataTable(lstrExpSql)
        If ldtbUltimoEnviado.Rows.Count > 0 Then
            ldblUltCli = ClsPanorama.FobjValorCampo(ldtbUltimoEnviado.Rows(0)(
                    "IdTerceroCliente"), EnuTipoValor.EnuDouble)
        End If
        Return ldblUltCli
    End Function
#End Region

#Region "Validaciones"
    Friend Function FblnEsValidoTipoCorreo(aenuTipoCorreo As EnuTipoCorreoE) As Boolean
        EnuTipoCorreo = aenuTipoCorreo
        Return MenuTipoCorreo <> EnuTipoCorreoE.None
    End Function
    Friend Function FblnEsValidoFactAuto(ablnAuto As Boolean)
        If ablnAuto Then
            MblnFactAutoOk = EnuTipoCorreo = EnuTipoCorreoE.EnuFactAuto
        Else
            MblnFactAutoOk = EnuTipoCorreo <> EnuTipoCorreoE.EnuFactAuto
        End If
        If MblnFactAutoOk Then
            BlnFactAuto = ablnAuto
        End If
        Return MblnFactAutoOk
    End Function
    Friend Function FblnEsValidoArchivoExt(astrArchivoExterno As String,
            ByRef astrMens As String) As Boolean
        MstrArchivoExterno = astrArchivoExterno
        Dim lblnEsValido = MenuTipoCorreo = EnuTipoCorreoE.EnuArchExt
        If lblnEsValido Then
            lblnEsValido = My.Computer.FileSystem.FileExists(MstrArchivoExterno)
            If lblnEsValido Then
                StrArchivoExterno = astrArchivoExterno
            Else
                astrMens = "No se pudo encontrar el Archivo!"
            End If
        End If
        Return lblnEsValido
    End Function
    Friend Function FblnEsValidoIdCliente(adblIdCliente As Double) As Boolean
        MdblIdCliente = adblIdCliente
        Dim lblnEsValido = MenuTipoCorreo <> EnuTipoCorreoE.None
        If lblnEsValido Then
            If EnuTipoCorreo = EnuTipoCorreoE.EnuFac OrElse EnuTipoCorreo =
                    EnuTipoCorreoE.EnuNAA OrElse EnuTipoCorreo = EnuTipoCorreoE.EnuNDB OrElse
                    EnuTipoCorreo = EnuTipoCorreoE.EnuNCR OrElse
                    EnuTipoCorreo = EnuTipoCorreoE.EnuRC Then
                lblnEsValido = ClsPanorama.FblnEsValidoNumero(MdblIdCliente, GCDBLMINTERC,
                        GCDBLMAXTERC, True)
                Dim lobjValorLlave As Object()
                If lblnEsValido Then
                    If DblIdCliente <> MobjCliente.ObjIdClienteDbl.ObjValorPro Then
                        lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, MdblIdCliente}
                        MobjCliente.SAbra(lobjValorLlave)
                        lblnEsValido = MobjCliente.BlnExiste
                    End If
                    If lblnEsValido Then
                        lblnEsValido = MobjCliente.ObjRecibeDocsPorEmailBln.ObjValorPro
                    End If
                End If
            Else
                lblnEsValido = ClsPanorama.FblnEsValidoNumero(MdblIdCliente, GCDBLMINTERC,
                        GCDBLMAXTERC, False, EnuTipoValor.enuDouble)
            End If
        End If
        Return lblnEsValido
    End Function
    Friend Function FblnEsValidoIdPreAgr(astrIdpreAgr As String) As Boolean
        MstrIdPredAgru = astrIdpreAgr
        Dim lblnEsValido = MenuTipoCorreo <> EnuTipoCorreoE.None
        If lblnEsValido Then
            If EnuTipoCorreo = EnuTipoCorreoE.EnuFac OrElse EnuTipoCorreo = EnuTipoCorreoE.EnuNAA OrElse
                    EnuTipoCorreo = EnuTipoCorreoE.EnuNDB OrElse EnuTipoCorreo = EnuTipoCorreoE.EnuNCR OrElse
                    EnuTipoCorreo = EnuTipoCorreoE.EnuRC Then
                lblnEsValido = MobjCliente.BlnExiste
                If lblnEsValido Then
                    If MstrIdPredAgru = GCSTRSINPA Then
                        MstrIdPredAgru = String.Empty
                    End If
                    If MstrIdPredAgru = String.Empty Then
                        lblnEsValido = True
                    Else
                        lblnEsValido = MobjCliente.FblnPredioPropDelCliente(MstrIdPredAgru) OrElse
                                MobjCliente.FblnPredioEsArrendado(MstrIdPredAgru)
                        If lblnEsValido Then
                            Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil,
                                StrIdPredioAgrupador}
                            MobjPredioAgr.SAbra(lobjValorLlave)
                        End If
                    End If
                End If
            End If
        End If
        Return lblnEsValido
    End Function
    Friend Function FblnEsValidoNroDoc(astrIdDcto As String, ByRef astrMens As String) As Boolean
        Dim lblnEsValido = MenuTipoCorreo <> EnuTipoCorreoE.None AndAlso
                Not String.IsNullOrEmpty(astrIdDcto) AndAlso astrIdDcto <> My.Resources.Ninguno
        MstrNroDoc = astrIdDcto
        If lblnEsValido Then
            If EnuTipoCorreo = EnuTipoCorreoE.EnuFac OrElse
                    EnuTipoCorreo = EnuTipoCorreoE.EnuNAA OrElse
                    EnuTipoCorreo = EnuTipoCorreoE.EnuNDB OrElse
                    EnuTipoCorreo = EnuTipoCorreoE.EnuNCR OrElse
                    EnuTipoCorreo = EnuTipoCorreoE.EnuRC Then
                MstrPrefijoDoc = ClsPanorama.FstrPrefijoDcto(astrIdDcto)
                lblnEsValido = ClsPanorama.FblnEsValidoString(MstrPrefijoDoc, 0, 5, True)
                If lblnEsValido Then
                    MentIdDoc = ClsPanorama.FentIdDcto(astrIdDcto)
                    lblnEsValido = ClsPanorama.FblnEsValidoNumero(MentIdDoc, 1,
                            Integer.MaxValue, True)
                End If
                If lblnEsValido Then
                    Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil,
                            MstrPrefijoDoc, MentIdDoc}
                    Select Case EnuTipoCorreo
                        Case EnuTipoCorreoE.EnuFac
                            MobjDocumento = New ClsFactura()
                            MobjDocumento.SAbra(lobjValorLlave)
                            Dim lobjFact As ClsFactura = MobjDocumento
                            lblnEsValido = lobjFact.BlnExiste
                            If lblnEsValido Then
                                If lobjFact.BlnEsFacEle Then
                                    lblnEsValido = Not lobjFact.BlnEstaPorRegEFac
                                    If Not lblnEsValido Then
                                        astrMens = "La Factura no puede ser enviada. No tiene CUFE!"
                                    End If
                                End If
                            End If
                        Case EnuTipoCorreoE.EnuNAA
                            MobjDocumento = New ClsNotaCon()
                            MobjDocumento.SAbra(lobjValorLlave)
                            lblnEsValido = MobjDocumento.BlnExiste
                        Case EnuTipoCorreoE.EnuNCR
                            MobjDocumento = New ClsNotaCr()
                            MobjDocumento.SAbra(lobjValorLlave)
                            lblnEsValido = MobjDocumento.BlnExiste
                        Case EnuTipoCorreoE.EnuNDB
                            MobjDocumento = New ClsNotaDb()
                            MobjDocumento.SAbra(lobjValorLlave)
                            lblnEsValido = MobjDocumento.BlnExiste
                        Case EnuTipoCorreoE.EnuRC
                            MobjDocumento = New ClsReciboCaja()
                            MobjDocumento.SAbra(lobjValorLlave)
                            lblnEsValido = MobjDocumento.BlnExiste
                        Case EnuTipoCorreoE.EnuRecibos
                            lblnEsValido = True
                        Case Else
                            lblnEsValido = True
                    End Select
                End If
            End If
        End If
        Return lblnEsValido
    End Function
    Friend Function FblnEsValidoAsunto(astrAsunto As String, ByRef astrMens As String) As Boolean
        MstrAsunto = astrAsunto
        Dim lblnEsValido = MenuTipoCorreo <> EnuTipoCorreoE.None
        If lblnEsValido Then
            lblnEsValido = ClsPanorama.FblnEsValidoString(MstrAsunto, 3, 100, True)
        End If
        If Not lblnEsValido Then
            astrMens = "El Asunto debe ser un Texto con longitud entre tres y cien Caractéres!"
        End If
        Return lblnEsValido
    End Function
    Friend Function FblnEsValidoMensaje(astrMensMail As String, ByRef astrMens As String) As Boolean
        MstrMensaje = astrMensMail
        Dim lblnEsValido = MenuTipoCorreo <> EnuTipoCorreoE.None
        If lblnEsValido Then
            lblnEsValido = ClsPanorama.FblnEsValidoString(StrMensaje, 3, 500, True)
            If Not lblnEsValido Then
                astrMens = "El mensaje debe ser un texto con longitud entre tres " &
                        "y quinientos caractéres!"
            End If
        End If
        Return lblnEsValido
    End Function
    Friend Function FblnEsValidoMensajeCobroPer(astrMensMail As String,
            ByRef astrMens As String) As Boolean
        MstrMensaje = astrMensMail
        Dim lblnEsValido = MenuTipoCorreo <> EnuTipoCorreoE.None
        If lblnEsValido Then
            lblnEsValido = ClsPanorama.FblnEsValidoString(StrMensaje, 3, 1500, True)
            If Not lblnEsValido Then
                astrMens = "El mensaje debe ser un texto con longitud entre tres y 
                        mil quinientos caractéres!"
            End If
        End If
        Return lblnEsValido
    End Function
    Friend Function FblnEsValidoDiasVen(astrDias As String, ByRef astrMens As String)
        Dim lblnEsValido = ClsPanorama.FblnEsValidoNumero(astrDias, 1, 1000, True,
                EnuTipoValor.EnuInteger)
        If Not lblnEsValido Then
            astrMens = "El valor debe ser un número entero entre uno y mil!"
        Else
            EntDiasCobroPers = CInt(astrDias)
        End If
        Return lblnEsValido
    End Function
    Friend Function FblnEsValFecIni(adtmFecIni As Date, ByRef astrMens As String) As Boolean
        MdtmFecIni = adtmFecIni
        Dim lblnEsValido = MenuTipoCorreo = EnuTipoCorreoE.EnuRecibos
        If lblnEsValido Then
            lblnEsValido = (MdtmFecIni <= Date.Today)
            If Not lblnEsValido Then
                astrMens = "La Fecha inicial debe ser anterior a la Fecha del Día de Hoy"
            End If
        End If
        Return lblnEsValido
    End Function
    Friend Function FblnEsValFecFin(adtmFecFin As Date, ByRef astrMens As String) As Boolean
        MdtmFecFin = adtmFecFin
        Dim lblnEsValido = MenuTipoCorreo = EnuTipoCorreoE.EnuRecibos
        If lblnEsValido Then
            lblnEsValido = FblnEsValFecIni(MdtmFecIni, astrMens)
            If lblnEsValido Then
                lblnEsValido = (DtmFechaFin >= DtmFechaIni) AndAlso (DtmFechaFin <= Date.Today)
                If Not lblnEsValido Then
                    astrMens = "La Fecha final debe estar entre la Fecha Inicial " &
                            "y la Fecha del Día de Hoy"
                End If
            End If
        End If
        Return lblnEsValido
    End Function
    Friend Function FblnHayConInternet(ByRef astrMens As String) As Boolean
        Dim lstrMens = "No hay conexión a internet!"
        Dim lblnHyaInt = FblnHayInternet()
        If Not lblnHyaInt Then
            astrMens = lstrMens
        End If
        Return lblnHyaInt
    End Function
#End Region

#Region "Histórico"
    Friend Function FdtbCorreosEnviados(adblIdCliente As Double, adtmDesde As Date,
            adtmHasta As Date) As DataTable
        Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmDesde) & "'"
        Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmHasta.AddDays(1)) & "'"
        Dim lstrCampoTipo = FstrSelectTipo()
        Dim lstrTablaPri = MCSTRNOMBRETABLA
        Dim lstrTablaSec = ClsCliente.SstrNombreTabla
        Dim lstrCamSelPri As String() = {"*", lstrCampoTipo & " AS Tipo "}
        Dim lstrCamSelSec As String() = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCampRelPri = {StrCampoCarpeta, StrCampoCentroUtil, "IdTerceroCliente"}
        Dim lstrCampRelSec = {StrCampoCarpeta, StrCampoCentroUtil, ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrOrden = {{"Fecha", "DESC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri &
                " AND FechaEnvio BETWEEN " & lstrFechaDesde & " AND " & lstrFechaHasta
        If adblIdCliente > 0 Then
            lstrFiltro &= " AND P.IdTerceroCliente  = " & adblIdCliente
        End If
        Dim ldtbCorreosClie = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamSelPri,
                lstrTablaSec, lstrCamSelSec, lstrCampRelPri, lstrCampRelSec, lstrOrden, False,
                lstrFiltro, Array.Empty(Of String))
        Return ldtbCorreosClie
    End Function
    Private Function FstrSelectTipo() As String
        Dim lstrSelTip = "IIF(IdTipoCorreo = " & EnuTipoCorreoE.EnuSoloMens &
                ", 'Mensaje', IF(IdTipoCorreo = " &
                EnuTipoCorreoE.EnuArchExt & ", 'Archivo', IF(IdTipoCorreo = " &
                EnuTipoCorreoE.EnuFactAuto &
                ", 'Factura Aut.', IF(IdTipoCorreo = " & EnuTipoCorreoE.EnuFac &
                ", 'Factura', IF(IdTipoCorreo = " &
                EnuTipoCorreoE.EnuRC & ", 'Recibo caja', IF(IdTipoCorreo = " &
                EnuTipoCorreoE.EnuNCR &
                ", 'Nota crédito', If(IdTipoCorreo = " & EnuTipoCorreoE.EnuNDB &
                ", 'Nota int. mora', IF(IdTipoCorreo = " &
                EnuTipoCorreoE.EnuNAA & ", 'Nota aplicación anticipo', IF(IdTipoCorreo = " &
                EnuTipoCorreoE.EnuRecibos & ", 'Recibo caja', '')))))))))"
        Return lstrSelTip
    End Function
    Friend Function FdtbHistoricoTipo(aenuTipo As EnuTipoCorreoE, adtmFechaDesde As Date,
            adtmFechaHasta As Date) As DataTable
        Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmFechaDesde) & "'"
        Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmFechaHasta.AddDays(1)) & "'"
        Dim lstrTabla = MCSTRNOMBRETABLA
        Dim lstrCampSel As String() = {"Asunto", "FechaEnvio", "NumeroDoc", "Mensaje"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion &
                " AND FechaEnvio BETWEEN " & lstrFechaDesde & " AND " & lstrFechaHasta &
                " AND IdTipoCorreo = " & aenuTipo
        Dim lstrOrden As String(,) = {{"FechaEnvio", "ASC"}, {"NumeroDoc", "ASC"}}
        Dim ldtbHistTipo = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSel, lstrOrden,
                lstrFiltro, False, {})
        Return ldtbHistTipo
    End Function
#End Region

#Region "Conexion BD"
    Friend Sub SRegistreUltimo()
        Dim lstrArch = String.Empty
        If EnuTipoCorreo = EnuTipoCorreoE.EnuArchExt Then
            lstrArch = StrArchivoExterno
        End If
        SRegistreMens(lstrArch, GCDBLTERCERONULO, String.Empty)
    End Sub
    Friend Sub SRegistreMens(astrArch As String, adblIdCliete As Double,
            astrNumDoc As String)
        Dim lcolDatosMail As New Collection
        Dim lentOrdinal = ClsPanorama.FobjUltimaIdNumericaObjeto(MCSTRNOMBRETABLA, "Ordinal",
                EnuTipoValor.EnuInteger, String.Empty)
        lcolDatosMail.Clear()
        lentOrdinal += 1
        lcolDatosMail.Add(astrArch, "ArchivoExt")
        lcolDatosMail.Add(StrAsunto, "Asunto")
        lcolDatosMail.Add(EntDiasCobroPers, "DiasPersuasivo")
        lcolDatosMail.Add(Date.Now, "Fecha")
        lcolDatosMail.Add(Date.Now, "FechaEnvio")
        lcolDatosMail.Add(GshrIdCarpeta, "IdCarpeta")
        lcolDatosMail.Add(GshrIdCentroUtil, "IdCentroUtil")
        lcolDatosMail.Add(adblIdCliete, "IdTerceroCliente")
        lcolDatosMail.Add(EnuTipoCorreo, "IdTipoCorreo")
        lcolDatosMail.Add(GstrIdUsuario, "IdUsuario")
        lcolDatosMail.Add(StrMensaje, "Mensaje")
        lcolDatosMail.Add(astrNumDoc, "NumeroDoc")
        lcolDatosMail.Add(lentOrdinal, "Ordinal")
        MobjPanDat.SInserteRegistro(MCSTRNOMBRETABLA, FcolCamposMail, lcolDatosMail)
    End Sub
#End Region
End Class