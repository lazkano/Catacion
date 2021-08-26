Public Class DistribucionCafe

    Private _CodigoArticulo As String
    Public Property CodigoArticulo() As String
        Get
            Return _CodigoArticulo
        End Get
        Set(ByVal value As String)
            _CodigoArticulo = value
        End Set
    End Property

    Private _Recibo As String
    Public Property Recibo() As String
        Get
            Return _Recibo
        End Get
        Set(ByVal value As String)
            _Recibo = value
        End Set
    End Property

    Private _CantidadQQ As Double
    Public Property CantidadQQ() As Double
        Get
            Return _CantidadQQ
        End Get
        Set(ByVal value As Double)
            _CantidadQQ = value
        End Set
    End Property

    Private _CodigoFinca As String
    Public Property CodigoFinca() As String
        Get
            Return _CodigoFinca
        End Get
        Set(ByVal value As String)
            _CodigoFinca = value
        End Set
    End Property

    Private _CantidadRequerida As Double
    Public Property CantidadRequerida() As Double
        Get
            Return _CantidadRequerida
        End Get
        Set(ByVal value As Double)
            _CantidadRequerida = value
        End Set
    End Property

    Private _DescripcionArticulo As String
    Public Property DescripcionArticulo() As String
        Get
            Return _DescripcionArticulo
        End Get
        Set(ByVal value As String)
            _DescripcionArticulo = value
        End Set
    End Property

    Private _Cantidad As Double
    Public Property Cantidad() As Double
        Get
            Return _Cantidad
        End Get
        Set(ByVal value As Double)
            _Cantidad = value
        End Set
    End Property

    Private _TipoCafe As String
    Public Property TipoCafe() As String
        Get
            Return _TipoCafe
        End Get
        Set(ByVal value As String)
            _TipoCafe = value
        End Set
    End Property

    Private _Escala As String
    Public Property Escala() As String
        Get
            Return _Escala
        End Get
        Set(ByVal value As String)
            _Escala = value
        End Set
    End Property

    Private _EscalaRechazo As String
    Public Property EscalaRechazo() As String
        Get
            Return _EscalaRechazo
        End Get
        Set(ByVal value As String)
            _EscalaRechazo = value
        End Set
    End Property

    Private _CalidadCafe As String
    Public Property CalidadCafe() As String
        Get
            Return _CalidadCafe
        End Get
        Set(ByVal value As String)
            _CalidadCafe = value
        End Set
    End Property

    Private _Consumido As Double
    Public Property Consumido() As Double
        Get
            Return _Consumido
        End Get
        Set(ByVal value As Double)
            _Consumido = value
        End Set
    End Property

    Private _Disponible As Double
    Public Property Disponible() As Double
        Get
            Return _Disponible
        End Get
        Set(ByVal value As Double)
            _Disponible = value
        End Set
    End Property

    Private _CodigoUnidadMedida As String
    Public Property CodigoUnidadMedida() As String
        Get
            Return _CodigoUnidadMedida
        End Get
        Set(ByVal value As String)
            _CodigoUnidadMedida = value
        End Set
    End Property

    Private _DescripcionUnidadMedida As String
    Public Property DescripcionUnidadMedida() As String
        Get
            Return _DescripcionUnidadMedida
        End Get
        Set(ByVal value As String)
            _DescripcionUnidadMedida = value
        End Set
    End Property

    Private _CodigoAlmacen As String
    Public Property CodigoAlmacen() As String
        Get
            Return _CodigoAlmacen
        End Get
        Set(ByVal value As String)
            _CodigoAlmacen = value
        End Set
    End Property

    Private _MetodoEmision As String
    Public Property MetodoEmision() As String
        Get
            Return _MetodoEmision
        End Get
        Set(ByVal value As String)
            _MetodoEmision = value
        End Set
    End Property

    Private _CantidadPendiente As Double
    Public Property CantidadPendiente() As Double
        Get
            Return _CantidadPendiente
        End Get
        Set(ByVal value As Double)
            _CantidadPendiente = value
        End Set
    End Property

    Private _CantidadAlmacen As Double
    Public Property CantidadAlmacen() As Double
        Get
            Return _CantidadAlmacen
        End Get
        Set(ByVal value As Double)
            _CantidadAlmacen = value
        End Set
    End Property

    Private _SectorRuta As String
    Public Property SectorRuta() As String
        Get
            Return _SectorRuta
        End Get
        Set(ByVal value As String)
            _SectorRuta = value
        End Set
    End Property

    Private _RendimientoNeto As Double
    Public Property RendimientoNeto() As Double
        Get
            Return _RendimientoNeto
        End Get
        Set(ByVal value As Double)
            _RendimientoNeto = value
        End Set
    End Property

    Private _RendimientoBruto As Double
    Public Property RendimientoBruto() As Double
        Get
            Return _RendimientoBruto
        End Get
        Set(ByVal value As Double)
            _RendimientoBruto = value
        End Set
    End Property

    Private _CantidadSacos As String
    Public Property CantidadSacos() As String
        Get
            Return _CantidadSacos
        End Get
        Set(ByVal value As String)
            _CantidadSacos = value
        End Set
    End Property

    Private _Vueltas As String
    Public Property Vueltas() As String
        Get
            Return _Vueltas
        End Get
        Set(ByVal value As String)
            _Vueltas = value
        End Set
    End Property

    Private _Recuperacion As String
    Public Property Recuperacion() As String
        Get
            Return _Recuperacion
        End Get
        Set(ByVal value As String)
            _Recuperacion = value
        End Set
    End Property
End Class
