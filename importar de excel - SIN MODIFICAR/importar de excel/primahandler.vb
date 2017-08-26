Imports System.Data.Odbc
Public Class primahandler
    Private m_con As String
    Private m_connODBC As OdbcConnection

    Public Sub New(ByVal conexion As String)
        Me.m_con = conexion
    End Sub


    Public Function Obtenprima(ByVal fechai As Date, ByVal fechaf As Date) As ArrayList
        Dim año, mes, dia, año2, mes2, dia2 As String
        Dim inicial As String
        Dim final As String

        año = fechai.Year.ToString
        mes = fechai.Month.ToString
        dia = fechai.Day.ToString

        If mes.Length = 1 Then
            mes = "0" & mes
        End If

        año2 = fechaf.Year.ToString
        mes2 = fechaf.Month.ToString
        dia2 = fechaf.Day.ToString

        If mes2.Length = 1 Then
            mes2 = "0" & mes2
        End If

        inicial = "'" + mes + "-" + dia + "-" + año + "'"
        final = "'" + mes2 + "-" + dia2 + "-" + año2 + "'"

        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction
        If mes <> mes2 Then
            ''meses diferentes
            Try
                Dim cadenaODBC As String

                cadenaODBC = Me.m_con

                Dim OdbcDr As OdbcDataReader
                Dim arreDatos As New ArrayList
                Dim strQuery As New System.Text.StringBuilder()
                'Me.m_conn.Open()
                Me.m_connODBC = New OdbcConnection(cadenaODBC)
                Me.m_connODBC.Open()
                trODBC = Me.m_connODBC.BeginTransaction(IsolationLevel.Serializable)
                Dim commODBC As New OdbcCommand("", Me.m_connODBC, trODBC)
                With strQuery
                    .Remove(0, .Length)
                    ' .Append("select numero,fecha_ingreso,salario_diario from empleados where estatus = 'A' and fecha_ingreso  between " & inicial & "and " & final)
                    .Append("select numero,nombres,fecha_ingreso,salario_diario from empleados ")
                    .Append(" where estatus = 'A' ")
                    .Append("and  extract(month from fecha_ingreso) = extract(month from date " & inicial & " ) ")
                    .Append(" and extract(day from fecha_ingreso) > extract(day from date  " & inicial & ") ")
                    .Append("union ")
                    .Append("select numero,nombres,fecha_ingreso,salario_diario from empleados ")
                    .Append(" where estatus = 'A' ")
                    .Append(" and  extract(month from fecha_ingreso) = extract(month from date  " & final & ") ")
                    .Append("and extract(day from fecha_ingreso) < extract(day from date  " & final & ")")

                End With
                'comm.CommandText = strQuery.ToString
                commODBC.CommandText = strQuery.ToString
                'dr = com.ExecuteReader()
                OdbcDr = commODBC.ExecuteReader()
                While OdbcDr.Read()
                    Dim c As New datos
                    c.numerop = OdbcDr("numero")
                    c.diap = OdbcDr("fecha_ingreso")
                    c.conceptop = "primav"
                    c.nombrep = OdbcDr("nombres")
                    c.totalp = OdbcDr("salario_diario")
                    'c.idalmacen = OdbcDr("ALMACEN_ID")
                    'c.almacen = OdbcDr("NOMBRE")
                    arreDatos.Add(c)
                End While
                'Me.m_conn.Close()
                Me.m_connODBC.Close()
                Return arreDatos
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

            ''meses iguales
        Else
            Try
                Dim cadenaODBC As String

                cadenaODBC = Me.m_con

                Dim OdbcDr As OdbcDataReader
                Dim arreDatos As New ArrayList
                Dim strQuery As New System.Text.StringBuilder()
                'Me.m_conn.Open()
                Me.m_connODBC = New OdbcConnection(cadenaODBC)
                Me.m_connODBC.Open()
                trODBC = Me.m_connODBC.BeginTransaction(IsolationLevel.Serializable)
                Dim commODBC As New OdbcCommand("", Me.m_connODBC, trODBC)
                With strQuery
                    .Remove(0, .Length)

                    .Append("select nombres,fecha_ingreso from empleados ")
                    .Append("where estatus = 'A' ")
                    .Append("and  extract(month from fecha_ingreso) = extract(month from date  " & inicial & ") ")
                    .Append("and extract(day from fecha_ingreso) between extract(day from date  " & inicial & ") and extract(day from date  " & inicial & ") ")
                    .Append("order by fecha_ingreso ")


                End With
                'comm.CommandText = strQuery.ToString
                commODBC.CommandText = strQuery.ToString
                'dr = com.ExecuteReader()
                OdbcDr = commODBC.ExecuteReader()
                While OdbcDr.Read()
                    Dim c As New datos
                    c.numerop = OdbcDr("numero")
                    c.diap = OdbcDr("fecha_ingreso")
                    c.conceptop = "primav"
                    c.totalp = OdbcDr("salario_diario")
                    'c.idalmacen = OdbcDr("ALMACEN_ID")
                    'c.almacen = OdbcDr("NOMBRE")
                    arreDatos.Add(c)
                End While
                'Me.m_conn.Close()
                Me.m_connODBC.Close()
                Return arreDatos
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

        End If

    End Function
End Class
