Imports System.Runtime.InteropServices
Imports System.Windows.Input
Imports Microsoft.Office.Interop.Excel
#Region "Definiciones"
<Assembly: CLSCompliant(True)>
#End Region
#Region "Enumeradores"
<Flags()>
Public Enum EnuElementosToolBarDef As Integer
    None = 0
    enuSalir = 1
    enuCerrar = 2
    enuSepSalir = 4
    enuCrear = 8
    enuModificar = 16
    enuSuprimir = 32
    enuAnular = 64
    enuGuardar = 128
    enuRefrescar = 256
    enuSepAccion = 512
    enuEstado = 1024
    enuSepEstado = 2048
    enuCalendario = 4096
    enuCalculadora = 8192
    enuNotas = 16384
    enuMensajes = 32768
    enuSepHerramientas = 65536
    enuAlPrimero = 131072
    enuAlAnterior = 262144
    enuAlSiguiente = 524288
    enuAlUltimo = 1048576
    enuBuscar = 2097152
    enuSepNavegar = 4194304
    enuTercero = 8388608
    enuImprimir = 16777216
    enuReportes = 33554432
    enuCamara = 67108864
    enuSepVarios = 134217728
    enuAyuda = 268435456
End Enum
<Flags()>
Public Enum EnuElementosAdicionalesDef As Integer
    None = 0
    enuBuscar = 1
    enuTercero = 2
    enuImprimir = 4
    enuReportes = 8
    enuCamara = 16
End Enum
Public Enum EnuTamanoIconos As Integer
    EnuPequeño = 0
    EnuMediano
    EnuGrande
End Enum
Friend Enum EnuEstadoAyudaDef As Byte
    EnuOff
    EnuOn
    EnuFrmOn
End Enum
#End Region
Friend Structure StcValidadorWpf
    Implements IEquatable(Of StcValidadorWpf)
    Public Property BlnValido As Boolean
    Public Property CtlControl As Control
    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then Return False
        Dim lblnEsIgual As Boolean = obj.GetType.Name = "stcValidadorWpf"
        If lblnEsIgual Then
            lblnEsIgual = Equals(obj)
        End If
        Return lblnEsIgual
    End Function
    Public Overloads Function Equals(other As StcValidadorWpf) As Boolean Implements IEquatable(Of StcValidadorWpf).Equals
        Dim lblnEsIgual As Boolean = (BlnValido = other.BlnValido)
        If lblnEsIgual Then
            lblnEsIgual = (CtlControl.Name = other.CtlControl.Name)
        End If
        Return lblnEsIgual
    End Function
    Public Shared Operator =(astcValidador1 As StcValidadorWpf, astcValidador2 As StcValidadorWpf) As Boolean
        Return astcValidador1.Equals(astcValidador2)
    End Operator
    Public Shared Operator <>(astcValidador1 As StcValidadorWpf, astcValidador2 As StcValidadorWpf) As Boolean
        Return Not astcValidador1.Equals(astcValidador2)
    End Operator
    Public Overrides Function GetHashCode() As Integer
        Return 0
    End Function
End Structure
#Region "Clases de Comando"
Public NotInheritable Class ClsComandos
    Private Shared mrucCrear As RoutedUICommand = Nothing
    Private Shared mrucModificar As RoutedUICommand = Nothing
    Private Shared mrucGuardar As RoutedUICommand = Nothing
    Private Shared mrucSuprimir As RoutedUICommand = Nothing
    Private Shared mrucAnular As RoutedUICommand = Nothing
    Private Shared mrucRefrescar As RoutedUICommand = Nothing
    Private Shared mrucAlPrimero As RoutedUICommand = Nothing
    Private Shared mrucAlAnterior As RoutedUICommand = Nothing
    Private Shared mrucAlSiguiente As RoutedUICommand = Nothing
    Private Shared mrucAlUltimo As RoutedUICommand = Nothing
    Private Shared mrucBuscar As RoutedUICommand = Nothing
    Private Shared mrucCerrar As RoutedUICommand = Nothing
    Private Shared mrucSalir As RoutedUICommand = Nothing
    Private Shared mrucAceptar As RoutedUICommand = Nothing
    Private Shared mrucCancelar As RoutedUICommand = Nothing
    Private Shared mrucCalendario As RoutedUICommand = Nothing
    Private Shared mrucCalculadora As RoutedUICommand = Nothing
    Private Shared mrucNotas As RoutedUICommand = Nothing
    Private Shared mrucMensajes As RoutedUICommand = Nothing
    Private Shared mrucImprimir As RoutedUICommand = Nothing
    Private Shared mrucAyuda As RoutedUICommand = Nothing

    Private Sub New()
    End Sub

    Public Shared ReadOnly Property RucCrear As RoutedUICommand
        Get
            If mrucCrear Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.N, ModifierKeys.Control)
                }
                mrucCrear = New RoutedUICommand(My.Resources.CrearObj, My.Resources.CrearCmd,
                        GetType(ClsComandos), lcolEntradasDeUsuario)
            End If
            Return mrucCrear
        End Get
    End Property

    Public Shared ReadOnly Property RucModificar As RoutedUICommand
        Get
            If mrucModificar Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.M, ModifierKeys.Control)
                }
                mrucModificar = New RoutedUICommand(My.Resources.ModObj, My.Resources.ModCmd,
                        GetType(ClsComandos), lcolEntradasDeUsuario)
            End If
            Return mrucModificar
        End Get
    End Property

    Public Shared ReadOnly Property RucGuardar As RoutedUICommand
        Get
            If mrucGuardar Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.G, ModifierKeys.Control)
                }
                mrucGuardar = New RoutedUICommand(My.Resources.GuaObj, My.Resources.GuaCmd,
                        GetType(ClsComandos), lcolEntradasDeUsuario)
            End If
            Return mrucGuardar
        End Get
    End Property

    Public Shared ReadOnly Property RucSuprimir As RoutedUICommand
        Get
            If mrucSuprimir Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.Delete, ModifierKeys.Alt)
                }
                mrucSuprimir = New RoutedUICommand(My.Resources.SupObj, My.Resources.SupCmd, GetType(ClsComandos),
                        lcolEntradasDeUsuario)
            End If
            Return mrucSuprimir
        End Get
    End Property

    Public Shared ReadOnly Property RucAnular As RoutedUICommand
        Get
            If mrucAnular Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.A, ModifierKeys.Control)
                }
                mrucAnular = New RoutedUICommand(My.Resources.AnuObj, My.Resources.AnuCmd, GetType(ClsComandos),
                        lcolEntradasDeUsuario)
            End If
            Return mrucAnular
        End Get
    End Property

    Public Shared ReadOnly Property RucRefrescar As RoutedUICommand
        Get
            If mrucRefrescar Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.R, ModifierKeys.Control)
                }
                mrucRefrescar = New RoutedUICommand(My.Resources.RefObj, My.Resources.RefCmd, GetType(ClsComandos),
                        lcolEntradasDeUsuario)
            End If
            Return mrucRefrescar
        End Get
    End Property

    Public Shared ReadOnly Property RucAlPrimero As RoutedUICommand
        Get
            If mrucAlPrimero Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.Home, ModifierKeys.Alt)
                }
                mrucAlPrimero = New RoutedUICommand(My.Resources.IrPriObj, My.Resources.IrPriCmd, GetType(ClsComandos),
                        lcolEntradasDeUsuario)
            End If
            Return mrucAlPrimero
        End Get
    End Property

    Public Shared ReadOnly Property RucAlAnterior As RoutedUICommand
        Get
            If mrucAlAnterior Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.Left, ModifierKeys.Alt)
                }
                mrucAlAnterior = New RoutedUICommand(My.Resources.IrAntObj, My.Resources.IrAntCmd, GetType(ClsComandos),
                        lcolEntradasDeUsuario)
            End If
            Return mrucAlAnterior
        End Get
    End Property

    Public Shared ReadOnly Property RucAlSiguiente As RoutedUICommand
        Get
            If mrucAlSiguiente Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.Right, ModifierKeys.Alt)
                }
                mrucAlSiguiente = New RoutedUICommand(My.Resources.IrSigObj, My.Resources.IrSigCmd, GetType(ClsComandos),
                        lcolEntradasDeUsuario)
            End If
            Return mrucAlSiguiente
        End Get
    End Property

    Public Shared ReadOnly Property RucAlUltimo As RoutedUICommand
        Get
            If mrucAlUltimo Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.End, ModifierKeys.Alt)
                }
                mrucAlUltimo = New RoutedUICommand(My.Resources.IrUltObj, My.Resources.IrUltCmd, GetType(ClsComandos),
                        lcolEntradasDeUsuario)
            End If
            Return mrucAlUltimo
        End Get
    End Property

    Public Shared ReadOnly Property RucBuscar As RoutedUICommand
        Get
            If mrucBuscar Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.B, ModifierKeys.Control)
                }
                mrucBuscar = New RoutedUICommand(My.Resources.IrBusObj, My.Resources.IrBusCmd, GetType(ClsComandos),
                        lcolEntradasDeUsuario)
            End If
            Return mrucBuscar
        End Get
    End Property

    Public Shared ReadOnly Property RucCerrar As RoutedUICommand
        Get
            If mrucCerrar Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.Escape, ModifierKeys.None)
                }
                mrucCerrar = New RoutedUICommand(My.Resources.CierraWin, My.Resources.CierraWinCmd, GetType(ClsComandos),
                        lcolEntradasDeUsuario)
            End If
            Return mrucCerrar
        End Get
    End Property

    Public Shared ReadOnly Property RucSalir As RoutedUICommand
        Get
            If mrucSalir Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.F4, ModifierKeys.Alt)
                }
                mrucSalir = New RoutedUICommand(My.Resources.SaleApp, My.Resources.SaleAppCmd, GetType(ClsComandos),
                        lcolEntradasDeUsuario)
            End If
            Return mrucSalir
        End Get
    End Property

    Public Shared ReadOnly Property RucAceptar As RoutedUICommand
        Get
            If mrucAceptar Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.A, ModifierKeys.Alt)
                }
                mrucAceptar = New RoutedUICommand(My.Resources.AceptaCam, My.Resources.AceptaCamCmd, GetType(ClsComandos),
                        lcolEntradasDeUsuario)
            End If
            Return mrucAceptar
        End Get
    End Property

    Public Shared ReadOnly Property RucCancelar As RoutedUICommand
        Get
            If mrucCancelar Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.C, ModifierKeys.Alt)
                }
                mrucCancelar = New RoutedUICommand(My.Resources.CancelaCam, My.Resources.CancelaCamCmd,
                        GetType(ClsComandos), lcolEntradasDeUsuario)
            End If
            Return mrucCancelar
        End Get
    End Property

    Public Shared ReadOnly Property RucCalendario As RoutedUICommand
        Get
            If mrucCalendario Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.F2, ModifierKeys.None)
                }
                mrucCalendario = New RoutedUICommand(My.Resources.AbreCalendario, My.Resources.AbreCalendarioCmd,
                        GetType(ClsComandos), lcolEntradasDeUsuario)
            End If
            Return mrucCalendario
        End Get
    End Property

    Public Shared ReadOnly Property RucCalculadora As RoutedUICommand
        Get
            If mrucCalculadora Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.F3, ModifierKeys.None)
                }
                mrucCalculadora = New RoutedUICommand(My.Resources.AbreCalc, My.Resources.AbreCalcCmd,
                        GetType(ClsComandos), lcolEntradasDeUsuario)
            End If
            Return mrucCalculadora
        End Get
    End Property

    Public Shared ReadOnly Property RucNotas As RoutedUICommand
        Get
            If mrucNotas Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.F4, ModifierKeys.None)
                }
                mrucNotas = New RoutedUICommand(My.Resources.AbreBloc, My.Resources.AbreBlocCmd,
                        GetType(ClsComandos), lcolEntradasDeUsuario)
            End If
            Return mrucNotas
        End Get
    End Property

    Public Shared ReadOnly Property RucMensajes As RoutedUICommand
        Get
            If mrucMensajes Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.F5, ModifierKeys.None)
                }
                mrucMensajes = New RoutedUICommand(My.Resources.AbreMens, My.Resources.AbreMensCmd,
                        GetType(ClsComandos), lcolEntradasDeUsuario)
            End If
            Return mrucMensajes
        End Get
    End Property

    Public Shared ReadOnly Property RucImprimir As RoutedUICommand
        Get
            If mrucImprimir Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.P, ModifierKeys.Control)
                }
                mrucImprimir = New RoutedUICommand("", "cmdImprimir",
                        GetType(ClsComandos), lcolEntradasDeUsuario)
            End If
            Return mrucImprimir
        End Get
    End Property

    Public Shared ReadOnly Property RucAyuda As RoutedUICommand
        Get
            If mrucAyuda Is Nothing Then
                Dim lcolEntradasDeUsuario As New InputGestureCollection From {
                    New KeyGesture(Key.F1, ModifierKeys.None)
                }
                mrucAyuda = New RoutedUICommand(My.Resources.AbreAyuda, My.Resources.AbreAyudaCmd,
                        GetType(ClsComandos), lcolEntradasDeUsuario)
            End If
            Return mrucAyuda
        End Get
    End Property
End Class
#End Region
'Clase utilizada para poblar el ListView
Friend Class ClsViewItemPar
    Public Property StrPropiedad As String = String.Empty
    Public Property StrValor As String = String.Empty
    Public Sub New(astrPropiedad As String, astrValor As String)
        StrPropiedad = astrPropiedad
        StrValor = astrValor
    End Sub
End Class
Friend Class ClsNovedad
    Public Property FechaNovedad As Date = GCDTMFECHANULA
    Public Property IdCta As String = String.Empty
    Public Property IdTerceroCliente As Double = 0.0
    Public Property IdPredioAgrupador As String = String.Empty
    Public Property Detalle As String = String.Empty
    Public Property TipoDoc As String = String.Empty
    Public Property Debito As Decimal = 0D
    Public Property Credito As Decimal = 0D
    Public Property Saldo As Decimal = 0D
End Class