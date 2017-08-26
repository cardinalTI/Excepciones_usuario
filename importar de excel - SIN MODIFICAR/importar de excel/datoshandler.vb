Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO

Public Class reportehandler
    Private m_con As String
    Private m_connODBC As SqlConnection
    Dim centro As String
    Private mensaje3 As String = ""
    Private m_Conn As String
    Private m_ConnODBC2009 As OdbcConnection
    Private _conexiones As String

  

    

    Public Function regresadatos() As ArrayList



        'If centrouno = "4310" Then
        '    centro = "7"
        'End If

        Dim trODBC As SqlTransaction
        Try
            Dim cadenaODBC As String

            cadenaODBC = Me.m_con

            Dim OdbcDr As SqlDataReader
            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()
            'Me.m_connODBC = New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
            Me.m_connODBC = New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
            ' Me.m_connODBC = New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
            '  Me.m_connODBC = New SqlConnection("Server=(local)\sqlexpress10;Database=incidencias;integrated security=true")
            Me.m_connODBC.Open()
            trODBC = Me.m_connODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New SqlCommand("", Me.m_connODBC, trODBC)
            With strQuery
                .Append("SELECT DISTINCT NOEMPLEADO as numero, '12' as dia,0 as falta FROM TBLDETALLESINCIDENCIAS ")
                .Append("WHERE NOEMPLEADO NOT IN (SELECT NOEMPLEADO FROM TBLDETALLESINCIDENCIAS WHERE CLAVE = '12' ) ")
                .Append(" UNION ")
                .Append("SELECT DISTINCT NOEMPLEADO,CLAVE,CANTIDAD FROM TBLDETALLESINCIDENCIAS ")
                .Append("WHERE CLAVE = '12' ")

            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            OdbcDr = commODBC.ExecuteReader()
            While OdbcDr.Read()
                Dim c As New datos
                c.numero = OdbcDr("numero")
                c.mes = OdbcDr("falta")
                c.dia = OdbcDr("dia")


                c.mensaje1 = "|1|" + c.numero + "," + "12" + "," + c.mes
                'Me.mensaje3 = c.mensaje1 + c.mensaje1
                'MsgBox(c.mensaje)
                arreDatos.Add(c)
            End While
            Me.m_connODBC.Close()
            Return Me.Obtendetalle(arreDatos)
        Catch ex As Exception
            Try
                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_connODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try


    End Function

    ''abril
    Public Function regresadatosm(ByVal dias As String) As ArrayList



        'If centrouno = "4310" Then
        '    centro = "7"
        'End If

        Dim trODBC As SqlTransaction
        Try
            Dim cadenaODBC As String

            cadenaODBC = Me.m_con

            Dim OdbcDr As SqlDataReader
            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()
            'Me.m_connODBC = New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
            Me.m_connODBC = New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
            ' Me.m_connODBC = New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
            '  Me.m_connODBC = New SqlConnection("Server=(local)\sqlexpress10;Database=incidencias;integrated security=true")
            Me.m_connODBC.Open()
            trODBC = Me.m_connODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New SqlCommand("", Me.m_connODBC, trODBC)
            With strQuery
                .Append("SELECT DISTINCT NOEMPLEADO as numero, '12' as dia,0 as falta FROM TBLDETALLESINCIDENCIAS ")
                .Append("WHERE NOEMPLEADO NOT IN (SELECT NOEMPLEADO FROM TBLDETALLESINCIDENCIAS WHERE CLAVE = '12' ) ")
                .Append(" UNION ")
                .Append("SELECT DISTINCT NOEMPLEADO,CLAVE,CANTIDAD FROM TBLDETALLESINCIDENCIAS ")
                .Append("WHERE CLAVE = '12' ")

            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            OdbcDr = commODBC.ExecuteReader()
            While OdbcDr.Read()
                Dim c As New datos
                c.numero = OdbcDr("numero")
                c.mes = OdbcDr("falta")
                c.dia = OdbcDr("dia")


                c.mensaje1 = "|1|" + c.numero + "," + dias + "," + c.mes
                'Me.mensaje3 = c.mensaje1 + c.mensaje1
                'MsgBox(c.mensaje)
                arreDatos.Add(c)
            End While
            Me.m_connODBC.Close()
            Return Me.Obtendetalle(arreDatos)
        Catch ex As Exception
            Try
                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_connODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try


    End Function
    ''abril

    ''julio

    Public Function regresadatosaicel(ByVal dias As String) As ArrayList



        'If centrouno = "4310" Then
        '    centro = "7"
        'End If

        Dim trODBC As SqlTransaction
        Try
            Dim cadenaODBC As String

            cadenaODBC = Me.m_con

            Dim OdbcDr As SqlDataReader
            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()
            'Me.m_connODBC = New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
            Me.m_connODBC = New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
            ' Me.m_connODBC = New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
            '  Me.m_connODBC = New SqlConnection("Server=(local)\sqlexpress10;Database=incidencias;integrated security=true")
            Me.m_connODBC.Open()
            trODBC = Me.m_connODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New SqlCommand("", Me.m_connODBC, trODBC)
            With strQuery
                .Append("SELECT DISTINCT NOEMPLEADO as numero, '12' as dia,0 as falta FROM TBLDETALLESINCIDENCIAS ")
                .Append("WHERE NOEMPLEADO NOT IN (SELECT NOEMPLEADO FROM TBLDETALLESINCIDENCIAS WHERE CLAVE = '12' ) ")
                .Append(" UNION ")
                .Append("SELECT DISTINCT NOEMPLEADO,CLAVE,CANTIDAD FROM TBLDETALLESINCIDENCIAS ")
                .Append("WHERE CLAVE = '12' ")

            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            OdbcDr = commODBC.ExecuteReader()
            While OdbcDr.Read()
                Dim c As New datos
                c.numero = OdbcDr("numero")
                c.mes = OdbcDr("falta")
                c.dia = OdbcDr("dia")


                c.mensaje1 = "|1|" + c.numero + "," + dias + "," + c.mes
                'Me.mensaje3 = c.mensaje1 + c.mensaje1
                'MsgBox(c.mensaje)
                arreDatos.Add(c)
            End While
            Me.m_connODBC.Close()
            Return Me.Obtendetalle(arreDatos)
        Catch ex As Exception
            Try
                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_connODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try


    End Function

    ''julio





    Private mensaje As String = ""

    ''nuevo

    Public Function Obtendetalle(ByVal arre As ArrayList) As ArrayList
        Dim cadena As String

        'Dim tr As FbTransaction
        For i As Integer = 0 To arre.Count - 1
            Dim trODBC As SqlTransaction
            Try
                Dim cadenaODBC As String
                Dim OdbcDr As SqlDataReader
                cadenaODBC = Me.m_con
                Dim strQuery As New System.Text.StringBuilder()
                'Me.m_conn.Open()
                'Me.m_connODBC = New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
                Me.m_connODBC = New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")

                ' Me.m_connODBC = New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
                Me.m_connODBC.Open()
                trODBC = Me.m_connODBC.BeginTransaction(IsolationLevel.Serializable)
                Dim commODBC As New SqlCommand("", Me.m_connODBC, trODBC)

                With strQuery
                    .Remove(0, .Length)
                    .Append("SELECT NoEmpleado as Nempleado, cantidad as cantidad, Clave as clave FROM TBLDETALLESINCIDENCIAS ")
                    .Append("where NOEMPLEADO = " & CType(arre(i), datos).numero)
                    .Append("and CLAVE != '12' ")

                End With

                commODBC.CommandText = strQuery.ToString
                OdbcDr = commODBC.ExecuteReader()

                While OdbcDr.Read()

                    Dim c As New datos
                    'CType(arre(i), datos).mes1 = OdbcDr("Nempleado".ToString)
                    CType(arre(i), datos).dia1 = OdbcDr("clave".ToString)
                    CType(arre(i), datos).numero1 = OdbcDr("cantidad".ToString)
                    CType(arre(i), datos).mensaje = "|1.1|" + CType(arre(i), datos).dia1 + "," + CType(arre(i), datos).numero1

                    CType(arre(i), datos).mensaje3 = Me.mensaje

                    Dim uno As String
                    Dim dos As String
                    Dim tres As String


                    'uno = CType(arre(i), datos).mes1
                    dos = CType(arre(i), datos).dia1
                    tres = CType(arre(i), datos).numero1
                    ' CType(arre(i), datos).mensaje = "|1.1|" + uno + "," + dos + "," + tres
                    CType(arre(i), datos).mensaje = "|1.1|" + dos + "," + tres
                    Me.mensaje = Me.mensaje + vbCrLf + CType(arre(i), datos).mensaje

                    'c.mes = OdbcDr("falta")
                    ' c.dia = OdbcDr("dia")
                    ' c.mensaje = "|1|" + c.numero + "," + "14" + "," + c.mes
                    CType(arre(i), datos).mensaje3 = Me.mensaje

                End While
                Dim d As New datos
                d.mensaje = Me.mensaje
                Me.mensaje = ""

                'Me.m_conn.Close()
                Me.m_connODBC.Close()

            Catch ex As Exception
                Try
                    trODBC.Rollback()
                Catch ex1 As Exception
                    MsgBox(ex.Message)
                End Try
            Finally
                Try
                    Me.m_connODBC.Close()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End Try
        Next
        Return arre


    End Function


    ''nuevo marzo

    'Public Function regresadatos(ByVal centrouno As String) As ArrayList



    '    If centrouno = "4310" Then
    '        centro = "7"
    '    End If

    '    Dim trODBC As SqlTransaction
    '    Try
    '        Dim cadenaODBC As String

    '        cadenaODBC = Me.m_con

    '        Dim OdbcDr As SqlDataReader
    '        Dim arreDatos As New ArrayList
    '        Dim strQuery As New System.Text.StringBuilder()
    '        'Me.m_conn.Open()
    '        Me.m_connODBC = New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
    '        'Me.m_connODBC = New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
    '        ' Me.m_connODBC = New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
    '        '  Me.m_connODBC = New SqlConnection("Server=(local)\sqlexpress10;Database=incidencias;integrated security=true")
    '        Me.m_connODBC.Open()
    '        trODBC = Me.m_connODBC.BeginTransaction(IsolationLevel.Serializable)
    '        Dim commODBC As New SqlCommand("", Me.m_connODBC, trODBC)
    '        With strQuery
    '            .Append("SELECT DISTINCT NOEMPLEADO as numero, " & centro & " as dia,0 as falta FROM TBLDETALLESINCIDENCIAS ")
    '            .Append("WHERE NOEMPLEADO NOT IN (SELECT NOEMPLEADO FROM TBLDETALLESINCIDENCIAS WHERE CLAVE = " & centro & " ) ")
    '            .Append(" UNION ")
    '            .Append("SELECT DISTINCT NOEMPLEADO,CLAVE,CANTIDAD FROM TBLDETALLESINCIDENCIAS ")
    '            .Append("WHERE CLAVE = " & centro)

    '        End With
    '        'comm.CommandText = strQuery.ToString
    '        commODBC.CommandText = strQuery.ToString
    '        'dr = com.ExecuteReader()
    '        OdbcDr = commODBC.ExecuteReader()
    '        While OdbcDr.Read()
    '            Dim c As New datos
    '            c.numero = OdbcDr("numero")
    '            c.mes = OdbcDr("falta")
    '            c.dia = OdbcDr("dia")


    '            c.mensaje1 = "|1|" + c.numero + "," + centro + "," + c.mes
    '            'Me.mensaje3 = c.mensaje1 + c.mensaje1
    '            'MsgBox(c.mensaje)
    '            arreDatos.Add(c)
    '        End While
    '        Me.m_connODBC.Close()
    '        Return Me.Obtendetalle(arreDatos)
    '    Catch ex As Exception
    '        Try
    '            trODBC.Rollback()
    '        Catch ex1 As Exception
    '            MsgBox(ex.Message)
    '        End Try
    '    Finally
    '        Try
    '            Me.m_connODBC.Close()
    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        End Try
    '    End Try


    'End Function



    'Private mensaje As String = ""

    ''nuevo

    'Public Function Obtendetalle(ByVal arre As ArrayList) As ArrayList
    '    Dim cadena As String

    '    'Dim tr As FbTransaction
    '    For i As Integer = 0 To arre.Count - 1
    '        Dim trODBC As SqlTransaction
    '        Try
    '            Dim cadenaODBC As String
    '            Dim OdbcDr As SqlDataReader
    '            cadenaODBC = Me.m_con
    '            Dim strQuery As New System.Text.StringBuilder()
    '            'Me.m_conn.Open()
    '            Me.m_connODBC = New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
    '            'Me.m_connODBC = New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")

    '            ' Me.m_connODBC = New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
    '            Me.m_connODBC.Open()
    '            trODBC = Me.m_connODBC.BeginTransaction(IsolationLevel.Serializable)
    '            Dim commODBC As New SqlCommand("", Me.m_connODBC, trODBC)

    '            With strQuery
    '                .Remove(0, .Length)
    '                .Append("SELECT NoEmpleado as Nempleado, cantidad as cantidad, Clave as clave FROM TBLDETALLESINCIDENCIAS ")
    '                .Append("where NOEMPLEADO = " & CType(arre(i), datos).numero)
    '                .Append("AND CLAVE != " & centro & " and CLAVE != '12' ")

    '            End With

    '            commODBC.CommandText = strQuery.ToString
    '            OdbcDr = commODBC.ExecuteReader()

    '            While OdbcDr.Read()

    '                Dim c As New datos
    '                'CType(arre(i), datos).mes1 = OdbcDr("Nempleado".ToString)
    '                CType(arre(i), datos).dia1 = OdbcDr("clave".ToString)
    '                CType(arre(i), datos).numero1 = OdbcDr("cantidad".ToString)
    '                CType(arre(i), datos).mensaje = "|1.1|" + CType(arre(i), datos).dia1 + "," + CType(arre(i), datos).numero1

    '                CType(arre(i), datos).mensaje3 = Me.mensaje

    '                Dim uno As String
    '                Dim dos As String
    '                Dim tres As String


    '                'uno = CType(arre(i), datos).mes1
    '                dos = CType(arre(i), datos).dia1
    '                tres = CType(arre(i), datos).numero1
    '                ' CType(arre(i), datos).mensaje = "|1.1|" + uno + "," + dos + "," + tres
    '                CType(arre(i), datos).mensaje = "|1.1|" + dos + "," + tres
    '                Me.mensaje = Me.mensaje + vbCrLf + CType(arre(i), datos).mensaje

    '                'c.mes = OdbcDr("falta")
    '                ' c.dia = OdbcDr("dia")
    '                ' c.mensaje = "|1|" + c.numero + "," + "14" + "," + c.mes
    '                CType(arre(i), datos).mensaje3 = Me.mensaje

    '            End While
    '            Dim d As New datos
    '            d.mensaje = Me.mensaje
    '            Me.mensaje = ""

    '            'Me.m_conn.Close()
    '            Me.m_connODBC.Close()

    '        Catch ex As Exception
    '            Try
    '                trODBC.Rollback()
    '            Catch ex1 As Exception
    '                MsgBox(ex.Message)
    '            End Try
    '        Finally
    '            Try
    '                Me.m_connODBC.Close()
    '            Catch ex As Exception
    '                MsgBox(ex.Message)
    '            End Try
    '        End Try
    '    Next
    '    Return arre


    'End Function

    ''nuevo marzo



    ''nuevo






    Function gridatxt(ByVal Grid As DataGridView) As Boolean
        Dim texto As StreamWriter
        Dim escribo As String
        Dim filas As Integer
        Dim columnas As Integer
        Dim titulo As String = ""
        Dim palabra As String
        Dim letras As Integer
        Dim nuevo As String
        filas = Grid.RowCount - 1
        columnas = Grid.ColumnCount - 1

        'texto = New StreamWriter("C:\Users\Lider2\Desktop\incidencias.txt", True)
        texto = New StreamWriter("c:\exportaciones\exepciones.txt", True)
        'texto = New StreamWriter("c:\Users\Soporte\Desktop\incidencias.txt", True)
        Dim tamaño(columnas) As Integer
        For i = 0 To columnas
            tamaño(i) = 0
        Next
        For a = 0 To filas
            Grid.CurrentCell = Grid.Rows(a).Cells(0)
            For b = 0 To columnas
                titulo = Grid.Columns(b).Name
                If IsDBNull(Grid.CurrentRow.Cells.Item(titulo).Value) Then
                    'palabra = "NULL"
                    palabra = vbCrLf

                Else
                    palabra = Grid.CurrentRow.Cells.Item(titulo).Value
                End If
                letras = palabra.Length
                If letras > tamaño(b) Then
                    tamaño(b) = letras

                End If
            Next
        Next
        For a = 0 To filas
            escribo = ""
            Grid.CurrentCell = Grid.Rows(a).Cells(0)
            For b = 0 To columnas
                titulo = Grid.Columns(b).Name
                If b = 0 Then
                    If IsDBNull(Grid.CurrentRow.Cells.Item(titulo).Value) Then
                        'escribo = "NULL"
                        palabra = vbCrLf

                    Else
                        escribo = Grid.CurrentRow.Cells.Item(titulo).Value
                        letras = escribo.Length
                    End If
                Else
                    If IsDBNull(Grid.CurrentRow.Cells.Item(titulo).Value) Then
                        'escribo = "NULL"
                        palabra = vbCrLf
                    Else
                        nuevo = Grid.CurrentRow.Cells.Item(titulo).Value
                        letras = nuevo.Length
                        Do While letras < tamaño(b)
                            'nuevo = nuevo & ""
                            letras = letras + 1
                        Loop
                        escribo = escribo & "" & nuevo
                    End If
                End If
            Next
            texto.WriteLine(escribo)
        Next
        texto.Close()
    End Function

    'nuevo abril

    Public Function ObtenNominas() As ArrayList
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction
        Try
            Dim cadenaODBC As String

            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
           ";PWD=ata8244;DBNAME=192.168.2.83" & _
               ":C:\microsip datos\MORGET.FDB"

            Dim OdbcDr As OdbcDataReader
            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()
            Me.m_ConnODBC2009 = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC2009.Open()
            trODBC = Me.m_ConnODBC2009.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC2009, trODBC)
            With strQuery
                .Remove(0, .Length)
                .Append("select * from nominas")

            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            OdbcDr = commODBC.ExecuteReader()
            While OdbcDr.Read()
                Dim n As New datos
                n.IdNomina = OdbcDr("NOMINA_ID")
                n.FechaNomina = OdbcDr("FECHA_PAGO")
                arreDatos.Add(n)
            End While
            'Me.m_conn.Close()
            Me.m_ConnODBC2009.Close()
            Return arreDatos
        Catch ex As Exception
            Try
                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox("Error No Controlado 55H: " & ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC2009.Close()
            Catch ex As Exception
                MsgBox("Error No Controlado 61H: " & ex.Message)
            End Try
        End Try
    End Function

    Public Function ObtenCentro() As ArrayList
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction
        Try
            Dim cadenaODBC As String

            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
           ";PWD=ata8244;DBNAME=192.168.2.83" & _
               ":C:\microsip datos\MORGET.FDB"

            Dim OdbcDr As OdbcDataReader
            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()
            Me.m_ConnODBC2009 = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC2009.Open()
            trODBC = Me.m_ConnODBC2009.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC2009, trODBC)
            With strQuery
                .Remove(0, .Length)
                .Append("select * from frecuencias_pago ")

            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            OdbcDr = commODBC.ExecuteReader()
            While OdbcDr.Read()
                Dim n As New datos
                n.Idcentro = OdbcDr("frepag_id")
                n.nombrec = OdbcDr("nombre")
                arreDatos.Add(n)
            End While
            'Me.m_conn.Close()
            Me.m_ConnODBC2009.Close()
            Return arreDatos
        Catch ex As Exception
            Try
                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox("Error No Controlado 55H: " & ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC2009.Close()
            Catch ex As Exception
                MsgBox("Error No Controlado 61H: " & ex.Message)
            End Try
        End Try
    End Function
    'nuevo abril 





End Class
