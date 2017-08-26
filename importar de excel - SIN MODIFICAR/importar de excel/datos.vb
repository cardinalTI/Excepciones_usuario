Public Class datos

    Private m_nombre As String
    Private m_correo As String
    Private m_rfc As String

    Private m_mensaje As String
    Private m_mensaje1 As String
    Private m_mensaje3 As String

    Private m_numero As String
    Private m_mes As String
    Private m_dia As String
    Private m_numero1 As String
    Private m_mes1 As String
    Private m_dia1 As String


    Private m_FechaNomina As String
    Private m_IdNomina As String
    Private m_Idcentro As String
    Private m_nombrec As String

    Property mensaje() As String
        Get
            Return Me.m_mensaje
        End Get
        Set(ByVal value As String)
            Me.m_mensaje = value
        End Set
    End Property

    Property mensaje1() As String
        Get
            Return Me.m_mensaje1
        End Get
        Set(ByVal value As String)
            Me.m_mensaje1 = value
        End Set
    End Property

    Property mensaje3() As String
        Get
            Return Me.m_mensaje3
        End Get
        Set(ByVal value As String)
            Me.m_mensaje3 = value
        End Set
    End Property

    Property nombre() As String
        Get
            Return Me.m_nombre
        End Get
        Set(ByVal value As String)
            Me.m_nombre = value
        End Set
    End Property

    Property correo() As String
        Get
            Return Me.m_correo
        End Get
        Set(ByVal value As String)
            Me.m_correo = value
        End Set
    End Property

    Property rfc() As String
        Get
            Return Me.m_rfc
        End Get
        Set(ByVal value As String)
            Me.m_rfc = value
        End Set
    End Property

    Property numero() As String
        Get
            Return Me.m_numero
        End Get
        Set(ByVal value As String)
            Me.m_numero = value
        End Set
    End Property

    Property mes() As String
        Get
            Return Me.m_mes
        End Get
        Set(ByVal value As String)
            Me.m_mes = value
        End Set
    End Property

    Property dia() As String
        Get
            Return Me.m_dia
        End Get
        Set(ByVal value As String)
            Me.m_dia = value
        End Set
    End Property


    Property numero1() As String
        Get
            Return Me.m_numero1
        End Get
        Set(ByVal value As String)
            Me.m_numero1 = value
        End Set
    End Property

    Property mes1() As String
        Get
            Return Me.m_mes1
        End Get
        Set(ByVal value As String)
            Me.m_mes1 = value
        End Set
    End Property

    Property dia1() As String
        Get
            Return Me.m_dia1
        End Get
        Set(ByVal value As String)
            Me.m_dia1 = value
        End Set
    End Property


    ''prima vacacional
    Private m_numerop As String
    Property numerop() As String
        Get
            Return Me.m_numerop
        End Get
        Set(ByVal value As String)
            Me.m_numerop = value
        End Set
    End Property

    Private m_diasp As String
    Property diap() As String
        Get
            Return Me.m_diasp
        End Get
        Set(ByVal value As String)
            Me.m_diasp = value
        End Set
    End Property
    Private m_totalp As String
    Property totalp() As Double
        Get
            Return Me.m_totalp
        End Get
        Set(ByVal value As Double)
            Me.m_totalp = value
        End Set
    End Property

    Private m_conceptop As String
    Property conceptop() As String
        Get
            Return Me.m_conceptop
        End Get
        Set(ByVal value As String)
            Me.m_conceptop = value
        End Set
    End Property
    Private m_nombrep As String
    Property nombrep() As String
        Get
            Return Me.m_nombrep
        End Get
        Set(ByVal value As String)
            Me.m_nombrep = value
        End Set
    End Property


    ''nuevo

    Property IdNomina() As String
        Get
            Return Me.m_IdNomina
        End Get
        Set(ByVal value As String)
            Me.m_IdNomina = value
        End Set
    End Property

    Property FechaNomina() As String
        Get
            Return Me.m_FechaNomina
        End Get
        Set(ByVal value As String)
            Me.m_FechaNomina = value
        End Set
    End Property


    Property Idcentro() As String
        Get
            Return Me.m_Idcentro
        End Get
        Set(ByVal value As String)
            Me.m_Idcentro = value
        End Set
    End Property

    Property nombrec() As String
        Get
            Return Me.m_nombrec
        End Get
        Set(ByVal value As String)
            Me.m_nombrec = value
        End Set
    End Property
End Class
