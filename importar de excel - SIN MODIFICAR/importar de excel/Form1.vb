Imports System.Data.OleDb
Imports System.Data.SqlClient
'Imports FirebirdSql.Data.FirebirdClient
Imports System.Data.Odbc

Public Class Form1
    Private m_con As String
    Private m_ConnODBC As OdbcConnection
    Private ccalculos As primahandler
    Private ccentros As reportehandler
    Private cdatos As reportehandler
    Private arreDatos As ArrayList
    Private arreprima As ArrayList
    Private arrecentro As New ArrayList
    Private arrefecha As New ArrayList
    Public logtexto As String
    Dim texto1 = "Dias festivos"
    Dim texto2 = "Prima Dominical"
    Dim texto3 = "Fonacot"
    Dim texto4 = "Faltas"
    Dim texto5 = "Bonos"
    Dim texto6 = "Prima vacacional"
    Dim texto7 = "Salario Retroactivo"
    Dim texto8 = "Infonavit"
    Dim texto9 = "Ajuste Infonavit"
    Public DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
    Public Sub Exportar_Excel(ByVal dgv As DataGridView, ByVal pth As String)

        Dim xlApp As Object = CreateObject("Excel.Application")
        'crear una nueva hoja de calculo 
        Dim xlWB As Object = xlApp.WorkBooks.add
        Dim xlWS As Object = xlWB.WorkSheets(1)

        'exportamos los caracteres de las columnas 
        For c As Integer = 0 To DataGridView1.Columns.Count - 1
            xlWS.cells(1, c + 1).value = DataGridView1.Columns(c).HeaderText
        Next
        'exportamos las cabeceras de columnas 
        For r As Integer = 0 To DataGridView1.RowCount - 1
            For c As Integer = 0 To DataGridView1.Columns.Count - 1
                xlWS.cells(r + 2, c + 1).value = DataGridView1.Item(c, r).Value
            Next
        Next
        'guardamos la hoja de calculo en la ruta especificada 
        xlWB.saveas(pth)
        xlWS = Nothing
        xlWB = Nothing
        xlApp.quit()
        xlApp = Nothing
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False

            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        Try
            Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
            Dim cnConex As New OleDbConnection(stConexion)
            Dim Cmd As New OleDbCommand("Select *From [DIA FESTIVO$]")
            Dim Ds As New DataSet
            Dim Da As New OleDbDataAdapter
            Dim Dt As New DataTable
            cnConex.Open()
            Cmd.Connection = cnConex
            Da.SelectCommand = Cmd
            Da.Fill(Ds)
            Dt = Ds.Tables(0)
            Me.DataGridView1.Columns.Clear()

            Me.DataGridView1.DataSource = Dt
            Me.DataGridView1.AutoGenerateColumns = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub



    ''primer boton

    Public Sub Agregausuario1(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")

        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                 "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "18"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            'MsgBox(ex.Message.ToString)
            MsgBox("Error al conectar a la base de datos")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try
    End Sub

    ''segundo boton

    Public Sub Agregausuario2(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "020"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub FONACOT(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "030"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox(ex.Message.ToString)
            MsgBox("Error al conectar a la base de datos")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub


    ''tercer boton

    Public Sub Agregausuario3(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "12"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)

                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            ' MsgBox(ex.Message.ToString)
            MsgBox("Error al conectar a la base de datos")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    ''cuarto boton

    Public Sub Agregausuario4(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                 "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "19"



        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)

                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            'MsgBox(ex.Message.ToString)
            MsgBox("Error al conectar a la base de datos")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    ''quinto prima vacacional

    Public Sub Agregausuario5(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                 "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "024"



        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)

                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            'MsgBox(ex.Message.ToString)
            MsgBox("Error al conectar a la base de datos")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub
    ''sexto horas extras

    Public Sub Agregausuario6(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                 "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "021"



        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)

                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            'MsgBox(ex.Message.ToString)
            MsgBox("Error al conectar a la base de datos")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub
    ''nuevo Infonavit


    Public Sub AgregarInfonavit(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                 "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "025"



        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)

                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            'MsgBox(ex.Message.ToString)
            MsgBox("Error al conectar a la base de datos")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub AgregarInfonavit2(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                 "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "026"



        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)

                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            'MsgBox(ex.Message.ToString)
            MsgBox("Error al conectar a la base de datos")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    ''DESCANSO LABORADO1

    Public Sub DESCANSO1(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "032"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub
    ''DESCANSO LABORADO2

    Public Sub DESCANSO2(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "033"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub ingresosx(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "022"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    'isr

    Public Sub isr(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "002"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    ''subsidio

    Public Sub subsidio(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "017"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button2.Enabled = False

        TextBox1.Text = "Paso 1: Generar la Prima Vacacional en la pestaña (prima vacacional)." + vbCrLf +
            "Paso 2: Revisar que el formato de Excel sea el correcto." + vbCrLf +
            "Paso 3: cargar el documento de Excel correspondiente. " + vbCrLf +
            "Precionar el boton (Obtener datos de Excel). " + vbCrLf + vbCrLf +
            "Paso 4: Una vez terminado el proceso precionar el boton (Generar archivo de texto)." + vbCrLf + vbCrLf +
            "NOTA: El TXT sera guardado en la ruta C:\exportaciones con el nombre exepciones.txt"


        'Dim conexion As String = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        '      ";PWD=ata8244;DBNAME=192.168.2.82" & _
        '   ":C:\microsip datos\NEXTEL.FDB"

        Dim conexion As String = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            ";PWD=ata8244;DBNAME=192.168.2.83" & _
                ":C:\microsip datos\NEXTEL.FDB"


        Dim conexiones As String = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
           ";PWD=ata8244;DBNAME=192.168.2.83" & _
               ":C:\microsip datos\NEXTEL.FDB"



        'Me.ccalculos = New primahandler(conexion)

        'Me.cdatos = New reportehandler()
        'Me.Muestrafecha()
        'Me.Muestracentro()


    End Sub

    Private Sub Muestrafecha()
        Try
            Me.arrefecha = Me.cdatos.ObtenNominas()
            Me.Combofecha.DataSource = Nothing
            Me.Combofecha.Items.Clear()
            If Me.arrefecha.Count > 0 Then
                With Me.Combofecha
                    .DisplayMember = "FechaNomina"
                    .DataSource = Me.arrefecha
                    .ValueMember = "IdNomina"
                End With
            End If
        Catch ex As Exception
            MsgBox("Error No Controlado 33: " & ex.Message, MsgBoxStyle.Critical, "Sistema")
        End Try
    End Sub

    Private Sub Muestracentro()
        Try
            Me.arrecentro = Me.cdatos.ObtenCentro()
            Me.combocentro.DataSource = Nothing
            Me.combocentro.Items.Clear()
            If Me.arrecentro.Count > 0 Then
                With Me.combocentro
                    .DisplayMember = "nombrec"
                    .DataSource = Me.arrecentro
                    .ValueMember = "Idcentro"
                End With
            End If
        Catch ex As Exception
            MsgBox("Error No Controlado 33: " & ex.Message, MsgBoxStyle.Critical, "Sistema")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        Try
            Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
            Dim cnConex As New OleDbConnection(stConexion)
            Dim Cmd As New OleDbCommand("Select [ EMPLEADO],[NOMBRE],[DIAS PRIMA DOMINICAL] From [PRIMA DOMINICAL$]")
            Dim Ds As New DataSet
            Dim Da As New OleDbDataAdapter
            Dim Dt As New DataTable
            cnConex.Open()
            Cmd.Connection = cnConex
            Da.SelectCommand = Cmd
            Da.Fill(Ds)
            Dt = Ds.Tables(0)
            Me.DataGridView2.Columns.Clear()
            Me.DataGridView2.DataSource = Dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        Try
            Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
            Dim cnConex As New OleDbConnection(stConexion)
            Dim Cmd As New OleDbCommand("Select  [EMPLEADO],[NOMBRE],[FALTAS] From [FALTAS$]")
            Dim Ds As New DataSet
            Dim Da As New OleDbDataAdapter
            Dim Dt As New DataTable
            cnConex.Open()
            Cmd.Connection = cnConex
            Da.SelectCommand = Cmd
            Da.Fill(Ds)
            Dt = Ds.Tables(0)
            Me.DataGridView3.Columns.Clear()
            Me.DataGridView3.DataSource = Dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        Button5.Enabled = True
    End Sub

    'Public Sub arreglo()
    '    Dim sArray As New ArrayList

    '    sArray.Add("BONOS")
    '    sArray.Add("FALTAS")
    '    sArray.Add("PRIMA DOMINICAL")
    '    sArray.Add("FONACOT")
    'End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'Dim sArray As New ArrayList

        'sArray.Add("BONOS")
        'sArray.Add("FALTAS")
        'sArray.Add("PRIMA DOMINICAL")
        'sArray.Add("FONACOT")

        Dim sArray As String() = {"BONOS", "FALTAS", "PRIMA DOMINICAL", "FONACOT", "DIAS FESTIVOS", "PRIMA VACACIONAL", "SALARIO RETROACTIVO", "INFONAVIT", "AJUSTE INFONAVIT", "DESCANSO LABORADO EX", "DESCANSO LABORADO GR", "OTROS INGRESOS EXENTOS", "ISR", "SUBSIDIO PARA EL EMPLEO"}


        '    Dim sArray As New List(Of String) _
        'From {"BONOS", "FALTAS", "PRIMA DOMINICAL", "FONACOT"}

        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        ''FOR EACH
        For Each item As String In sArray
            Try
                Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
                Dim cnConex As New OleDbConnection(stConexion)
                Dim Cmd As New OleDbCommand("Select * From [" & item & "$]")
                ' Dim Cmd As New OleDbCommand("Select *From [BONOS$]")
                Dim Ds As New DataSet
                Dim Da As New OleDbDataAdapter
                Dim Dt As New DataTable
                cnConex.Open()
                Cmd.Connection = cnConex
                Da.SelectCommand = Cmd
                Da.Fill(Ds)
                Dt = Ds.Tables(0)
                ''if
                If item = "BONOS" Then
                    Me.DataGridView4.Columns.Clear()
                    Me.DataGridView4.DataSource = Dt
                End If
                If item = "FALTAS" Then
                    Me.DataGridView3.Columns.Clear()
                    Me.DataGridView3.DataSource = Dt
                End If
                If item = "PRIMA DOMINICAL" Then
                    Me.DataGridView2.Columns.Clear()
                    Me.DataGridView2.DataSource = Dt
                End If
                If item = "FONACOT" Then
                    Me.DataGridView8.Columns.Clear()
                    Me.DataGridView8.DataSource = Dt
                End If
                If item = "DIAS FESTIVOS" Then
                    Me.DataGridView1.Columns.Clear()
                    Me.DataGridView1.DataSource = Dt
                End If
                If item = "PRIMA VACACIONAL" Then
                    Me.datagridpria.Columns.Clear()
                    Me.datagridpria.DataSource = Dt
                End If

                If item = "SALARIO RETROACTIVO" Then
                    Me.DataGridView7.Columns.Clear()
                    Me.DataGridView7.DataSource = Dt
                End If
                If item = "INFONAVIT" Then
                    Me.DataGridView9.Columns.Clear()
                    Me.DataGridView9.DataSource = Dt
                End If
                If item = "AJUSTE INFONAVIT" Then
                    Me.DataGridView10.Columns.Clear()
                    Me.DataGridView10.DataSource = Dt
                End If

                If item = "DESCANSO LABORADO EX" Then
                    Me.DataGridView11.Columns.Clear()
                    Me.DataGridView11.DataSource = Dt
                End If

                If item = "DESCANSO LABORADO GR" Then
                    Me.DataGridView12.Columns.Clear()
                    Me.DataGridView12.DataSource = Dt
                End If

                If item = "OTROS INGRESOS EXENTOS" Then
                    Me.DataGridView13.Columns.Clear()
                    Me.DataGridView13.DataSource = Dt
                End If

                If item = "ISR" Then
                    Me.DataGridView14.Columns.Clear()
                    Me.DataGridView14.DataSource = Dt
                End If

                If item = "SUBSIDIO PARA EL EMPLEO" Then
                    Me.DataGridView15.Columns.Clear()
                    Me.DataGridView15.DataSource = Dt
                End If

            Catch ex As Exception
                ' MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                MsgBox("Ingresa el formato correcto del archivo de Excel")
            End Try
        Next
        If Me.DataGridView4.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de BONOS", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'fila.DefaultCellStyle.BackColor = Color.Red
            'e.CellStyle.BackColor = Color.Coral
            DataGridView4.DefaultCellStyle.BackColor = Color.Red


        End If
        If Me.DataGridView3.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de FALTAS ", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
            DataGridView3.DefaultCellStyle.BackColor = Color.Red
        End If
        If Me.DataGridView2.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de PRIMA DOMINICAL", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
            DataGridView2.DefaultCellStyle.BackColor = Color.Red
        End If
        If Me.DataGridView8.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de FONACOT", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
            DataGridView8.DefaultCellStyle.BackColor = Color.Red
        End If
        If Me.DataGridView1.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de DIAS FESTVOS", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
            DataGridView1.DefaultCellStyle.BackColor = Color.Red
        End If
        'If Me.DataGridView6.Rows.Count = 0 Then
        '    MessageBox.Show("Ingrese datos en el grid.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'End If
        If Me.DataGridView7.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de SALARIO RETROACTIVO ", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
            DataGridView7.DefaultCellStyle.BackColor = Color.Red
        End If
        If Me.DataGridView9.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de INFONAVIT ", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
            DataGridView9.DefaultCellStyle.BackColor = Color.Red
        End If
        If Me.DataGridView10.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de AJUSTE INFONAVIT", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
            DataGridView10.DefaultCellStyle.BackColor = Color.Red
        End If
        If Me.DataGridView11.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de DESCANSO LABORADO  EX", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
            DataGridView11.DefaultCellStyle.BackColor = Color.Red
        End If
        If Me.DataGridView12.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de DESCANSO LABORADO GR", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
            DataGridView12.DefaultCellStyle.BackColor = Color.Red
        End If
        If Me.datagridpria.Rows.Count = 0 Then
            MessageBox.Show("Aun no se ha revizado si existe prima vacacional", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            datagridpria.DefaultCellStyle.BackColor = Color.Red
        End If

        If Me.DataGridView13.Rows.Count = 0 Then
            MessageBox.Show("Aun no se ha revizado si existe OTROS INGRESOS EXENTOS", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            DataGridView13.DefaultCellStyle.BackColor = Color.Red
        End If

        If Me.DataGridView14.Rows.Count = 0 Then
            MessageBox.Show("Aun no se ha revizado si existe ISR", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            DataGridView14.DefaultCellStyle.BackColor = Color.Red
        End If

        If Me.DataGridView15.Rows.Count = 0 Then
            MessageBox.Show("Aun no se ha revizado si existe SUBSIDIO PARA EL EMPLEO", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            DataGridView15.DefaultCellStyle.BackColor = Color.Red
        End If


        Button2.Enabled = True
    End Sub

    Public Sub delete()


        ' Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")



        Dim consulta As String
        consulta = ("delete from tblDetallesIncidencias")


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                End With
                DBCon.Open()

                comm.ExecuteNonQuery()

            End Using

        Catch ex As SqlException

            ' MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            MsgBox("Error al conectar con la base de datos")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try
    End Sub


    'cambio de estatus para empleados


    ''desactivar

    Public Sub desactivar()
        Dim trODBC As OdbcTransaction
        Dim cadenaODBC As String
        Try

            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" &
          ";PWD=ata8244;DBNAME=192.168.2.83" &
       ":C:\microsip datos\AICEL.FDB"

            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("update empleados set estatus = 'I' where frepag_id = 320")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Sub


    ''activar
    Public Sub estatusempleado(empleado)
        Dim trODBC As OdbcTransaction
        Dim cadenaODBC As String
        Try

            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" &
          ";PWD=ata8244;DBNAME=192.168.2.83" &
       ":C:\microsip datos\AICEL.FDB"

            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("update empleados set estatus = 'A' where numero = " & empleado & " ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Sub





    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ''DIAS FESTIVOS
        My.Computer.FileSystem.DeleteFile("c:\exportaciones\exepciones.txt")
        Me.delete()
        Dim contador As Integer = 0

        Try

            For i As Integer = 0 To Me.DataGridView1.Rows.Count - 1
                With Me.DataGridView1.Rows(i)

                    Me.Agregausuario1(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 18)
                    'TextBox2.Text = logtexto + texto1 + " ok..." + vbCrLf

                    'contador = +1
                End With
            Next
            ' MsgBox("Los archivos fueron almacenados correctamente")
            'MsgBox("El total de registros fueron " & contador)
        Catch ex As Exception
            'MsgBox("Error No Controlado 351: " & ex.Message)
            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try


        ''PRIMA DOMINICAL

        Try
            'Dim texto = "Prima dominical"
            For i As Integer = 0 To Me.DataGridView2.Rows.Count - 1
                With Me.DataGridView2.Rows(i)

                    Me.Agregausuario2(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "020".ToString)


                    'contador = +1
                End With
            Next
            '   MsgBox("Los archivos fueron almacenados correctamente")
            'MsgBox("El total de registros fueron " & contador)
        Catch ex As Exception
            ' MsgBox("Error No Controlado 351: " & ex.Message)
            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        ''FONACOT

        Try

            For i As Integer = 0 To Me.DataGridView8.Rows.Count - 1
                With Me.DataGridView8.Rows(i)

                    Me.FONACOT(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "030".ToString)
                    'logtexto = logtexto + texto3 + " ok..." + vbCrLf
                    'TextBox2.Text = logtexto

                    'contador = +1
                End With
            Next
            '   MsgBox("Los archivos fueron almacenados correctamente")
            'MsgBox("El total de registros fueron " & contador)
        Catch ex As Exception
            ' MsgBox("Error No Controlado 351: " & ex.Message)
            MsgBox("Ingresa el formato correcto del archivo de Excel")

        End Try


        ''FALTAS

        Try

            For i As Integer = 0 To Me.DataGridView3.Rows.Count - 1
                With Me.DataGridView3.Rows(i)

                    Me.Agregausuario3(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 12)
                    'logtexto = logtexto + texto4 + " ok..." + vbCrLf
                    'TextBox2.Text = logtexto
                    'contador = +1
                End With
            Next
            '  MsgBox("Los archivos fueron almacenados correctamente")
            'MsgBox("El total de registros fueron " & contador)
        Catch ex As Exception
            ' MsgBox("Error No Controlado 351: " & ex.Message)
            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        ''BONOS

        Try

            For i As Integer = 0 To Me.DataGridView4.Rows.Count - 1
                With Me.DataGridView4.Rows(i)

                    Me.Agregausuario4(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 19)
                    'logtexto = logtexto + texto5 + " ok..." + vbCrLf
                    'TextBox2.Text = logtexto
                    'contador = +1
                End With
            Next
            ' MsgBox("Los archivos fueron almacenados correctamente")
            'MsgBox("El total de registros fueron " & contador)
        Catch ex As Exception
            ' MsgBox("Error No Controlado 351: " & ex.Message)
            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        ''NUEVO

        ''PRIMA VACACIONAL  datagridpria

        'Try

        '    For i As Integer = 0 To Me.DataGridView6.Rows.Count - 1
        '        With Me.DataGridView6.Rows(i)

        '            Me.Agregausuario5(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "024")
        '            'logtexto = logtexto + texto6 + " ok..." + vbCrLf
        '            'TextBox2.Text = logtexto
        '            'contador = +1
        '        End With
        '    Next
        '    ' MsgBox("Los archivos fueron almacenados correctamente")
        '    'MsgBox("El total de registros fueron " & contador)
        'Catch ex As Exception
        '    ' MsgBox("Error No Controlado 351: " & ex.Message)
        '    MsgBox("Ingresa el formato correcto del archivo de Excel")
        'End Try
        Try

            For i As Integer = 0 To Me.datagridpria.Rows.Count - 1
                With Me.datagridpria.Rows(i)

                    Me.Agregausuario5(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "024")
                    'logtexto = logtexto + texto6 + " ok..." + vbCrLf
                    'TextBox2.Text = logtexto
                    'contador = +1
                End With
            Next
            ' MsgBox("Los archivos fueron almacenados correctamente")
            'MsgBox("El total de registros fueron " & contador)
        Catch ex As Exception
            ' MsgBox("Error No Controlado 351: " & ex.Message)
            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        ''SALARIO R

        Try

            For i As Integer = 0 To Me.DataGridView7.Rows.Count - 1
                With Me.DataGridView7.Rows(i)

                    Me.Agregausuario6(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "021")
                    'logtexto = logtexto + texto7 + " ok..." + vbCrLf
                    'TextBox2.Text = logtexto
                    'contador = +1
                End With
            Next
            ' MsgBox("Los archivos fueron almacenados correctamente")
            'MsgBox("El total de registros fueron " & contador)
        Catch ex As Exception
            ' MsgBox("Error No Controlado 351: " & ex.Message)
            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        ''INFONAVIT1
        Try

            For i As Integer = 0 To Me.DataGridView9.Rows.Count - 1
                With Me.DataGridView9.Rows(i)

                    Me.AgregarInfonavit(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "025")
                    'logtexto = logtexto + texto8 + " ok..." + vbCrLf
                    'TextBox2.Text = logtexto
                    ''contador = +1
                End With
            Next
            ' MsgBox("Los archivos fueron almacenados correctamente")
            'MsgBox("El total de registros fueron " & contador)
        Catch ex As Exception
            ' MsgBox("Error No Controlado 351: " & ex.Message)
            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        ''INFONAVIT2
        Try

            For i As Integer = 0 To Me.DataGridView10.Rows.Count - 1
                With Me.DataGridView10.Rows(i)

                    Me.AgregarInfonavit2(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "026")
                    'logtexto = logtexto + texto9 + " ok..." + vbCrLf
                    'TextBox2.Text = logtexto
                    'contador = +1
                End With
            Next
            ' MsgBox("Los archivos fueron almacenados correctamente")
            'MsgBox("El total de registros fueron " & contador)
        Catch ex As Exception
            ' MsgBox("Error No Controlado 351: " & ex.Message)
            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        '' DESCANSO LABORADO 1

        Try
            'Dim texto = "DESCANSO LABORADO EX"
            For i As Integer = 0 To Me.DataGridView11.Rows.Count - 1
                With Me.DataGridView11.Rows(i)

                    Me.DESCANSO1(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "032".ToString)


                    'contador = +1
                End With
            Next
            '   MsgBox("Los archivos fueron almacenados correctamente")
            'MsgBox("El total de registros fueron " & contador)
        Catch ex As Exception
            ' MsgBox("Error No Controlado 351: " & ex.Message)
            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        ''DESCANSO LABORADO 2

        Try
            'Dim texto = "DESCANSO LABORADO EX"
            For i As Integer = 0 To Me.DataGridView12.Rows.Count - 1
                With Me.DataGridView12.Rows(i)

                    Me.DESCANSO2(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "033".ToString)


                    'contador = +1
                End With
            Next
            '   MsgBox("Los archivos fueron almacenados correctamente")
            'MsgBox("El total de registros fueron " & contador)
        Catch ex As Exception
            ' MsgBox("Error No Controlado 351: " & ex.Message)
            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        Try
            'Dim texto = "otros ingresos excentos"
            For i As Integer = 0 To Me.DataGridView13.Rows.Count - 1
                With Me.DataGridView13.Rows(i)

                    Me.ingresosx(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "022".ToString)


                    'contador = +1
                End With
            Next
            '   MsgBox("Los archivos fueron almacenados correctamente")
            'MsgBox("El total de registros fueron " & contador)
        Catch ex As Exception
            ' MsgBox("Error No Controlado 351: " & ex.Message)
            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        ''isr
        Try
            'Dim texto = "ISR"
            For i As Integer = 0 To Me.DataGridView14.Rows.Count - 1
                With Me.DataGridView14.Rows(i)

                    Me.isr(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "002".ToString)


                    'contador = +1
                End With
            Next
            '   MsgBox("Los archivos fueron almacenados correctamente")
            'MsgBox("El total de registros fueron " & contador)
        Catch ex As Exception
            ' MsgBox("Error No Controlado 351: " & ex.Message)
            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try


        'SUBSIDIO PARA EL EMPLEO

        Try
            'Dim texto = "ISR"
            For i As Integer = 0 To Me.DataGridView15.Rows.Count - 1
                With Me.DataGridView15.Rows(i)

                    Me.subsidio(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "017".ToString)


                    'contador = +1
                End With
            Next
            '   MsgBox("Los archivos fueron almacenados correctamente")
            'MsgBox("El total de registros fueron " & contador)
        Catch ex As Exception
            ' MsgBox("Error No Controlado 351: " & ex.Message)
            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        MsgBox("Los archivos fueron almacenados correctamente")
        Muestradatos()
    End Sub



    Private Sub Muestradatos()


        Dim c As New reportehandler

        arreDatos = New ArrayList

        arreDatos = c.regresadatos()
        Try
            For i As Integer = 0 To Me.arreDatos.Count - 1

                With CType(Me.arreDatos(i), datos)

                    'If .mensaje3 <> "" Then

                    Me.DataSet11.datos.AdddatosRow(.mensaje1, .mensaje3)

                    'End If
                    Me.DataGridView5.DataSource = DataSet11.Tables(0).DefaultView
                End With

            Next

        Catch ex As Exception
            MsgBox("Error No Controlado: " & ex.Message, MsgBoxStyle.Critical, "REPORTE")
        End Try

        c.gridatxt(DataGridView5)
        MsgBox("El .txt Fue creado correctamente.")
    End Sub


    ''abril
    Private Sub Muestradatosm()


        Dim c As New reportehandler

        arreDatos = New ArrayList

        arreDatos = c.regresadatosm(cbxcentro.Text)
        Try
            For i As Integer = 0 To Me.arreDatos.Count - 1

                With CType(Me.arreDatos(i), datos)

                    'If .mensaje3 <> "" Then

                    Me.DataSet11.datos.AdddatosRow(.mensaje1, .mensaje3)

                    'End If
                    Me.DataGridView5.DataSource = DataSet11.Tables(0).DefaultView
                End With

            Next

        Catch ex As Exception
            MsgBox("Error No Controlado: " & ex.Message, MsgBoxStyle.Critical, "REPORTE")
        End Try

        c.gridatxt(DataGridView5)
        MsgBox("El .txt Fue creado correctamente.")
    End Sub
    ''abril

    ''uphrtitloli


    Private Sub Muestradatoshup()


        Dim c As New reportehandler

        arreDatos = New ArrayList

        arreDatos = c.regresadatosm(ComboBox1.Text)
        Try
            For i As Integer = 0 To Me.arreDatos.Count - 1

                With CType(Me.arreDatos(i), datos)

                    'If .mensaje3 <> "" Then

                    Me.DataSet11.datos.AdddatosRow(.mensaje1, .mensaje3)

                    'End If
                    Me.DataGridView5.DataSource = DataSet11.Tables(0).DefaultView
                End With

            Next

        Catch ex As Exception
            MsgBox("Error No Controlado: " & ex.Message, MsgBoxStyle.Critical, "REPORTE")
        End Try

        c.gridatxt(DataGridView5)
        MsgBox("El .txt Fue creado correctamente.")
    End Sub

    ''hupetitloli


    ''abril
    Private Sub Muestradatosaicel()


        Dim c As New reportehandler

        arreDatos = New ArrayList

        arreDatos = c.regresadatosaicel(CBXaicel.Text)
        Try
            For i As Integer = 0 To Me.arreDatos.Count - 1

                With CType(Me.arreDatos(i), datos)

                    'If .mensaje3 <> "" Then

                    Me.DataSet11.datos.AdddatosRow(.mensaje1, .mensaje3)

                    'End If
                    Me.DataGridView5.DataSource = DataSet11.Tables(0).DefaultView
                End With

            Next

        Catch ex As Exception
            MsgBox("Error No Controlado: " & ex.Message, MsgBoxStyle.Critical, "REPORTE")
        End Try

        c.gridatxt(DataGridView5)
        MsgBox("El .txt Fue creado correctamente.")
    End Sub
    ''abril

    Private Sub Button7_Click(sender As System.Object, e As System.EventArgs)
        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        Try
            Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
            Dim cnConex As New OleDbConnection(stConexion)
            Dim Cmd As New OleDbCommand("Select *From [PRIMA VACACIONAL$]")
            Dim Ds As New DataSet
            Dim Da As New OleDbDataAdapter
            Dim Dt As New DataTable
            cnConex.Open()
            Cmd.Connection = cnConex
            Da.SelectCommand = Cmd
            Da.Fill(Ds)
            Dt = Ds.Tables(0)
            Me.DataGridView6.Columns.Clear()
            Me.DataGridView6.DataSource = Dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        Button5.Enabled = True
    End Sub

    Private Sub Button8_Click(sender As System.Object, e As System.EventArgs)
        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        Try
            Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
            Dim cnConex As New OleDbConnection(stConexion)
            Dim Cmd As New OleDbCommand("Select *From [SALARIO RETROACTIVO$]")
            Dim Ds As New DataSet
            Dim Da As New OleDbDataAdapter
            Dim Dt As New DataTable
            cnConex.Open()
            Cmd.Connection = cnConex
            Da.SelectCommand = Cmd
            Da.Fill(Ds)
            Dt = Ds.Tables(0)
            Me.DataGridView7.Columns.Clear()
            Me.DataGridView7.DataSource = Dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub Button9_Click(sender As System.Object, e As System.EventArgs) Handles Button9.Click
        muestra()
        Button2.Enabled = True
    End Sub

    Private Sub muestra()
        Dim fechasi = DateTimePicker1.Value
        Dim fechasf = DateTimePicker2.Value
        Dim tipouno As Double

        Dim año, mes, dia, año2, mes2, dia2 As String
        Dim inicial1 As String
        Dim final1 As String

        año = fechasi.Year.ToString
        mes = fechasi.Month.ToString
        dia = fechasi.Day.ToString

        If mes.Length = 1 Then
            mes = "0" & mes
        End If

        Dim nombrem As String
        nombrem = MonthName(mes)

        año2 = fechasf.Year.ToString
        mes2 = fechasf.Month.ToString
        dia2 = fechasf.Day.ToString

        If mes2.Length = 1 Then
            mes2 = "0" & mes2
        End If

        Dim nombrem2 As String
        nombrem2 = MonthName(mes2)
        inicial1 = dia + "/" + nombrem + "/" + año
        final1 = dia2 + "/" + nombrem2 + "/" + año2

        Me.arreprima = Me.ccalculos.Obtenprima(fechasi, fechasf)
        Try
            Me.Dataprima1.Clear()
            For i As Integer = 0 To Me.arreprima.Count - 1
                With CType(Me.arreprima(i), datos)

                    tipouno = fecha(.diap, DateTimePicker2.Value.Date, .totalp)

                    Me.Dataprima1.primav.AddprimavRow(.numerop, .nombrep, tipouno)
                    Me.datagridpria.DataSource = Dataprima1.Tables(0).DefaultView

                    'Me.datagridpria(Me.Dataprima1.)

                End With
            Next
        Catch ex As Exception

        End Try
    End Sub
    Public Function fecha(ByVal fecha1 As String, ByVal fecha2 As String, ByVal totalp As Double) As Double


        Dim wD As Long
        Dim uno As Double


        Dim fechai As DateTime = DateTime.Parse(fecha1)
        Dim fechaf As Date = DateTimePicker2.Value.Date

        wD = DateDiff(DateInterval.Day, fechai, fechaf)

        '  MsgBox(wD)

        If wD < 366 Then

            uno = (totalp * 6) * 0.25

        End If

        If wD >= 367 And wD <= 731 Then

            uno = (totalp * 8) * 0.25
        End If

        If wD >= 1096 And wD <= 1461 Then

            uno = (totalp * 10) * 0.25
        End If
        Return uno
    End Function

    Private Sub Button10_Click(sender As System.Object, e As System.EventArgs)
        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        Try
            Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
            Dim cnConex As New OleDbConnection(stConexion)
            Dim Cmd As New OleDbCommand("Select [EMPLEADO],[NOMBRE],[INFONACOT] From [FONACOT$]")
            Dim Ds As New DataSet
            Dim Da As New OleDbDataAdapter
            Dim Dt As New DataTable
            cnConex.Open()
            Cmd.Connection = cnConex
            Da.SelectCommand = Cmd
            Da.Fill(Ds)
            Dt = Ds.Tables(0)
            Me.DataGridView8.Columns.Clear()
            Me.DataGridView8.DataSource = Dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        Button2.Enabled = True
    End Sub

    Private Sub Button11_Click(sender As System.Object, e As System.EventArgs)
        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        Try
            Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
            Dim cnConex As New OleDbConnection(stConexion)
            Dim Cmd As New OleDbCommand("Select *From [INFONAVIT$]")
            Dim Ds As New DataSet
            Dim Da As New OleDbDataAdapter
            Dim Dt As New DataTable
            cnConex.Open()
            Cmd.Connection = cnConex
            Da.SelectCommand = Cmd
            Da.Fill(Ds)
            Dt = Ds.Tables(0)
            Me.DataGridView8.Columns.Clear()
            Me.DataGridView8.DataSource = Dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        Button5.Enabled = True
    End Sub

    Private Sub Button12_Click(sender As System.Object, e As System.EventArgs)
        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        Try
            Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
            Dim cnConex As New OleDbConnection(stConexion)
            Dim Cmd As New OleDbCommand("Select *From [AJUSTE_INFONAVIT$]")
            Dim Ds As New DataSet
            Dim Da As New OleDbDataAdapter
            Dim Dt As New DataTable
            cnConex.Open()
            Cmd.Connection = cnConex
            Da.SelectCommand = Cmd
            Da.Fill(Ds)
            Dt = Ds.Tables(0)
            Me.DataGridView8.Columns.Clear()
            Me.DataGridView8.DataSource = Dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        Button5.Enabled = True
    End Sub

    Private Sub Button1_Click_1(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'Dim sArray As New ArrayList

        'sArray.Add("BONOS")
        'sArray.Add("FALTAS")
        'sArray.Add("PRIMA DOMINICAL")
        'sArray.Add("FONACOT")

        Dim sArray As String() = {"Pension Alimenticia", "Faltas", "Subsidio para el empleo", "Seguro social", "Descuento Infonavit", "Descuento Infonavit periodos an", "Otros Ingresos", "Ausentismo", "Haberes del retiro", "ISR", "Otros Descuentos", "Otro Descuento", "Ajuste en Subsidio"}


        '    Dim sArray As New List(Of String) _
        'From {"BONOS", "FALTAS", "PRIMA DOMINICAL", "FONACOT"}

        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        ''FOR EACH
        For Each item As String In sArray
            Try
                Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
                Dim cnConex As New OleDbConnection(stConexion)
                Dim Cmd As New OleDbCommand("Select * From [" & item & "$]")
                ' Dim Cmd As New OleDbCommand("Select *From [BONOS$]")
                Dim Ds As New DataSet
                Dim Da As New OleDbDataAdapter
                Dim Dt As New DataTable
                cnConex.Open()
                Cmd.Connection = cnConex
                Da.SelectCommand = Cmd
                Da.Fill(Ds)
                Dt = Ds.Tables(0)
                ''if


                If item = "Ajuste en Subsidio" Then
                    Me.DGVajuste.Columns.Clear()
                    Me.DGVajuste.DataSource = Dt
                End If
                If item = "Pension Alimenticia" Then
                    Me.dgvpensionesm.Columns.Clear()
                    Me.dgvpensionesm.DataSource = Dt
                End If


                If item = "Subsidio para el empleo" Then
                    Me.DGVsubsidio.Columns.Clear()
                    Me.DGVsubsidio.DataSource = Dt
                End If
                If item = "Seguro social" Then
                    Me.dgvsegurosocial.Columns.Clear()
                    Me.dgvsegurosocial.DataSource = Dt
                End If
                If item = "Descuento Infonavit" Then
                    Me.dgvdescuentoinfonavit.Columns.Clear()
                    Me.dgvdescuentoinfonavit.DataSource = Dt
                End If
                If item = "Descuento Infonavit periodos an" Then
                    Me.dgvdescuentoinfonavitperiodo.Columns.Clear()
                    Me.dgvdescuentoinfonavitperiodo.DataSource = Dt
                End If
                If item = "Otros Ingresos" Then
                    Me.dgvotrosingresos.Columns.Clear()
                    Me.dgvotrosingresos.DataSource = Dt
                End If
                If item = "Ausentismo" Then
                    Me.dgvausentismo.Columns.Clear()
                    Me.dgvausentismo.DataSource = Dt
                End If

                If item = "Haberes del retiro" Then
                    Me.dgvhaberes.Columns.Clear()
                    Me.dgvhaberes.DataSource = Dt
                End If
                If item = "Faltas" Then
                    Me.dgvfaltas.Columns.Clear()
                    Me.dgvfaltas.DataSource = Dt
                End If

                ''nuevos
                If item = "ISR" Then
                    Me.dgvisr2.Columns.Clear()
                    Me.dgvisr2.DataSource = Dt
                End If

                If item = "Otros Descuentos" Then
                    Me.dgvotrosdescuentos2.Columns.Clear()
                    Me.dgvotrosdescuentos2.DataSource = Dt
                End If

                If item = "Otro Descuento" Then
                    Me.dgvotrodesc2.Columns.Clear()
                    Me.dgvotrodesc2.DataSource = Dt
                End If




            Catch ex As Exception
                ' MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                MsgBox("Ingresa el formato correcto del archivo de Excel")
            End Try

        Next
        If Me.DGVajuste.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene Ajuste en Subsidio para el empleo", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If


        If Me.dgvpensionesm.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Pensiones Alimenticias", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        If Me.DGVsubsidio.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Subsidio para el empleo", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        If Me.dgvsegurosocial.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Seguro social ", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
        If Me.dgvdescuentoinfonavit.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Descuento InfonavitL", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
        If Me.dgvdescuentoinfonavitperiodo.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Descuento Infonavit periodos an", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
        If Me.dgvotrosingresos.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Otros Ingresos", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        If Me.dgvausentismo.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Ausentismo ", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
        If Me.dgvhaberes.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Haberes del retiro ", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        If Me.dgvfaltas.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de faltas ", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        ''nuevo

        If Me.dgvisr2.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de ISR ", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
        If Me.dgvotrosdescuentos2.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Otros descuentos ", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
        If Me.dgvotrodesc2.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Otro descuento ", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If




        Button2.Enabled = True
    End Sub

    Private Sub Button3_Click_1(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        My.Computer.FileSystem.DeleteFile("c:\exportaciones\exepciones.txt")
        Me.delete()
        Dim contador As Integer = 0


        ''ajuste 071

        Try

            For i As Integer = 0 To Me.DGVajuste.Rows.Count - 1
                With Me.DGVajuste.Rows(i)

                    Me.GVajuste(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "12")

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try


        Try

            For i As Integer = 0 To Me.dgvpensionesm.Rows.Count - 1
                With Me.dgvpensionesm.Rows(i)

                    Me.gvpensionesm(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "12")

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try


        Try
            'If cbxcentro.Text = "4310" Then
            '    For i As Integer = 0 To Me.dgvfaltas.Rows.Count - 1
            '        With Me.dgvfaltas.Rows(i)

            '            Me.gvfaltas(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "7")

            '        End With
            '    Next

            'End If

            For i As Integer = 0 To Me.dgvfaltas.Rows.Count - 1
                With Me.dgvfaltas.Rows(i)

                    Me.gvfaltas(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "12")

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        Try

            For i As Integer = 0 To Me.DGVsubsidio.Rows.Count - 1
                With Me.DGVsubsidio.Rows(i)

                    Me.GVsubsidio(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 18)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        ''
        Try

            For i As Integer = 0 To Me.dgvsegurosocial.Rows.Count - 1
                With Me.dgvsegurosocial.Rows(i)

                    Me.gvsegurosocial(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 18)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        ''
        Try

            For i As Integer = 0 To Me.dgvdescuentoinfonavit.Rows.Count - 1
                With Me.dgvdescuentoinfonavit.Rows(i)

                    Me.gvdescuentoinfonavit(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 18)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        ''
        Try

            For i As Integer = 0 To Me.dgvdescuentoinfonavitperiodo.Rows.Count - 1
                With Me.dgvdescuentoinfonavitperiodo.Rows(i)

                    Me.gvdescuentoinfonavitperiodo(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 18)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        ''
        Try

            For i As Integer = 0 To Me.dgvotrosingresos.Rows.Count - 1
                With Me.dgvotrosingresos.Rows(i)

                    Me.gvotrosingresos(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 18)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        ''
        Try

            For i As Integer = 0 To Me.dgvausentismo.Rows.Count - 1
                With Me.dgvausentismo.Rows(i)

                    Me.gvausentismo(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 18)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        ''
        Try

            For i As Integer = 0 To Me.dgvhaberes.Rows.Count - 1
                With Me.dgvhaberes.Rows(i)

                    Me.gvhaberes(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 18)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        ''NUEVOS
        Try

            For i As Integer = 0 To Me.dgvisr2.Rows.Count - 1
                With Me.dgvisr2.Rows(i)

                    Me.gvisr2(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 201)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        Try

            For i As Integer = 0 To Me.dgvotrosdescuentos2.Rows.Count - 1
                With Me.dgvotrosdescuentos2.Rows(i)

                    Me.gvotrosdescuentos2(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 108)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        Try

            For i As Integer = 0 To Me.dgvotrodesc2.Rows.Count - 1
                With Me.dgvotrodesc2.Rows(i)

                    Me.gvotrodesc2(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 112)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try


        'Muestradatos()
        Muestradatosm()
    End Sub

    Public Sub GVsubsidio(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "117"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub gvsegurosocial(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "204"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub gvdescuentoinfonavit(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        ' Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "025"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub gvdescuentoinfonavitperiodo(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "026"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub gvotrosingresos(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "118"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub gvausentismo(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "109"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub gvhaberes(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "111"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub gvfaltas(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = id


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    ''NUEVO
    Public Sub gvisr2(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "205"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub
    Public Sub gvotrosdescuentos2(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "108"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub
    Public Sub gvotrodesc2(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "112"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub


    ''UPHETILOLI


    Public Sub GVH046(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source=192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "108"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub GVH002(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "205"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub GVH004(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "206"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    ''UPHETILOLI

    ''ajuste

    Public Sub gvajuste(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " +
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "071"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    ''ajuste

    Public Sub gvpensionesm(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "113"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Private Sub Button4_Click_1(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        'Dim sArray As New ArrayList

        'sArray.Add("BONOS")
        'sArray.Add("FALTAS")
        'sArray.Add("PRIMA DOMINICAL")
        'sArray.Add("FONACOT")

        Dim sArray As String() = {"seguro Social", "Pago por crédito de vivienda", "INFONACOT", "Subsidios por incapacidad", "Otros", "Pensión alimenticia", "Prima vacacional exenta", "Prima Vac Grav", "Subsidio por incapacidad"}


        '    Dim sArray As New List(Of String) _
        'From {"BONOS", "FALTAS", "PRIMA DOMINICAL", "FONACOT"}

        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        ''FOR EACH
        For Each item As String In sArray
            Try
                Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
                Dim cnConex As New OleDbConnection(stConexion)
                Dim Cmd As New OleDbCommand("Select * From [" & item & "$]")
                ' Dim Cmd As New OleDbCommand("Select *From [BONOS$]")
                Dim Ds As New DataSet
                Dim Da As New OleDbDataAdapter
                Dim Dt As New DataTable
                cnConex.Open()
                Cmd.Connection = cnConex
                Da.SelectCommand = Cmd
                Da.Fill(Ds)
                Dt = Ds.Tables(0)
                ''if
                If item = "Pago por crédito de vivienda" Then
                    Me.dgvmpagocredito.Columns.Clear()
                    Me.dgvmpagocredito.DataSource = Dt
                End If
                If item = "INFONACOT" Then
                    Me.dgvminfonacot.Columns.Clear()
                    Me.dgvminfonacot.DataSource = Dt
                End If

                If item = "Otros" Then
                    Me.dgvmotros.Columns.Clear()
                    Me.dgvmotros.DataSource = Dt
                End If
                If item = "Prima vacacional exenta" Then
                    Me.dgvprimavex.Columns.Clear()
                    Me.dgvprimavex.DataSource = Dt
                End If
                If item = "Prima Vac Grav" Then
                    Me.dgvmprimavacgr.Columns.Clear()
                    Me.dgvmprimavacgr.DataSource = Dt
                End If
                If item = "Subsidio por incapacidad" Then
                    Me.dgvmsubinc2.Columns.Clear()
                    Me.dgvmsubinc2.DataSource = Dt
                End If
                If item = "Pensión alimenticia" Then
                    Me.dgvmpension.Columns.Clear()
                    Me.dgvmpension.DataSource = Dt
                End If
                If item = "seguro Social" Then
                    Me.dgvmseguro.Columns.Clear()
                    Me.dgvmseguro.DataSource = Dt
                End If



            Catch ex As Exception
                ' MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                MsgBox("Ingresa el formato correcto del archivo de Excel")
            End Try
        Next

        If Me.dgvmseguro.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de IMSS", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        If Me.dgvmpagocredito.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Pago por crédito de vivienda", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        If Me.dgvminfonacot.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de INFONACOT ", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        If Me.dgvmotros.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Otros", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
        If Me.dgvmpension.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Pensión alimenticia", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        ''
        If Me.dgvprimavex.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Prima vacacional exenta", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
        If Me.dgvmprimavacgr.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Prima Vac Grav", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
        If Me.dgvmsubinc2.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Subsidio por incapacidad", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        My.Computer.FileSystem.DeleteFile("c:\exportaciones\exepciones.txt")
        Me.delete()
        Dim contador As Integer = 0

        Try

            For i As Integer = 0 To Me.dgvmseguro.Rows.Count - 1
                With Me.dgvmseguro.Rows(i)

                    Me.gvmseguro(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 203)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try


        Try

            For i As Integer = 0 To Me.dgvmpagocredito.Rows.Count - 1
                With Me.dgvmpagocredito.Rows(i)

                    Me.gvmpagocredito(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 10)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        Try

            For i As Integer = 0 To Me.dgvminfonacot.Rows.Count - 1
                With Me.dgvminfonacot.Rows(i)

                    Me.gvminfonacot(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 11)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        'Try

        '    For i As Integer = 0 To Me.dgvmsubsidioinc.Rows.Count - 1
        '        With Me.dgvmsubsidioinc.Rows(i)

        '            Me.gvmsubsidioinc(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 90)

        '        End With
        '    Next

        'Catch ex As Exception

        '    MsgBox("Ingresa el formato correcto del archivo de Excel")
        'End Try
        Try

            For i As Integer = 0 To Me.dgvmotros.Rows.Count - 1
                With Me.dgvmotros.Rows(i)

                    Me.gvmotros(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 4)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        Try

            For i As Integer = 0 To Me.dgvmpension.Rows.Count - 1
                With Me.dgvmpension.Rows(i)

                    Me.gvmpension(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 4)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        ''

        Try

            For i As Integer = 0 To Me.dgvprimavex.Rows.Count - 1
                With Me.dgvprimavex.Rows(i)

                    Me.gvprimavex(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "1pv")

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        Try

            For i As Integer = 0 To Me.dgvmprimavacgr.Rows.Count - 1
                With Me.dgvmprimavacgr.Rows(i)

                    Me.gvmprimavacgr(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 91)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        Try

            For i As Integer = 0 To Me.dgvmsubinc2.Rows.Count - 1
                With Me.dgvmsubinc2.Rows(i)

                    Me.gvmsubinc2(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 92)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        Muestradatos()
    End Sub


    ''cinisal

    Public Sub gvmseguro(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "203"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub


    Public Sub gvmpagocredito(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "010"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub


    Public Sub gvminfonacot(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "011"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub


    Public Sub gvmsubsidioinc(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "090"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub


    Public Sub gvmotros(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        ' Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "004"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub


    Public Sub gvmpension(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "007"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub


    Public Sub gvprimavex(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "1pv"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub gvmprimavacgr(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "091"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub gvmsubinc2(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "092"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub


    'wipsi
    Public Sub gvpensionaw(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "2pan"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub gvhaberesw(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "109"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub gvcontrpw(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "108"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub gvotrosw(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "004"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub


    Public Sub gvotrosdescw(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "005"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub


    'wipsi

    ''aicel


    Public Sub GVaicelhaberes(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "109"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub GVaicelotros(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "004"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub GVaicelpension(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "007"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    Public Sub GVaicelodescuentos(ByVal numero As String, ByVal nombre As String, ByVal dia As String, ByVal id As String)
        'Dim DBCon As SqlConnection
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= incidencias;user id=sicossadmi;password=ipp2012;")
        'DBCon = New SqlConnection(New SqlConnection(cn))

        Dim consulta As String
        consulta = "insert into tblDetallesIncidencias (NoEmpleado,Nombre,Cantidad,Clave) " + _
                "values (@campo1,@campo2,@campo3,@campo4)"

        Dim numerou As New SqlParameter("@campo1", DbType.String)
        numerou.Value = numero
        Dim nombreu As New SqlParameter("@campo2", DbType.String)
        nombreu.Value = nombre
        Dim diau As New SqlParameter("@campo3", DbType.String)
        diau.Value = dia
        Dim idu As New SqlParameter("@campo4", DbType.String)
        idu.Value = "005"


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                    .Parameters.Add(numerou)
                    .Parameters.Add(nombreu)
                    .Parameters.Add(diau)
                    .Parameters.Add(idu)


                End With
                DBCon.Open()
                ' MsgBox("Conexion realizada satsfactoriamente")
                comm.ExecuteNonQuery()

            End Using
            '   MsgBox("Conecxion realizada satsfactoriamente")
        Catch ex As SqlException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox("Error al conectar a la base de datos")
            ' MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Sub

    ''aicel

    Public Function actualizar()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String

            '         Dim cadenaODBC As String = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '   ";PWD=ata8244;DBNAME=192.168.2.82" & _
            '":C:\microsip datos\NEXTEL.FDB"



            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
                                ";PWD=ata8244;DBNAME=192.168.2.83" & _
                             ":C:\microsip datos\GRUPO CONISAL.FDB"

            'cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '          ";PWD=ata8244;DBNAME=192.168.2.82" & _
            '       ":C:\microsip datos\NEXTEL.FDB"


            '    cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '   ";PWD=ata8244;DBNAME=189.190.172.169" & _
            '":C:\microsip datos\GRUPO CONISAL.FDB"

            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 116, ID_INTERNO = 9 where CONCEPTO_NO_ID = 8740 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    Public Function actualizarempleado()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String

            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
                              ";PWD=ata8244;DBNAME=192.168.2.83" & _
                           ":C:\microsip datos\GRUPO CONISAL.FDB"

            'cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '          ";PWD=ata8244;DBNAME=192.168.2.82" & _
            '       ":C:\microsip datos\NEXTEL.FDB"


            '    cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '   ";PWD=ata8244;DBNAME=189.190.172.169" & _
            '":C:\microsip datos\GRUPO CONISAL.FDB"

            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("update excep_empleados set dias_hrs_pagar = 15 where dias_hrs_pagar = 12 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function


    Public Function actualizarit()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String


            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
                                ";PWD=ata8244;DBNAME=192.168.2.83" & _
                             ":C:\microsip datos\IT RESOURCES TELECOM SA DE CV.FDB"

            'cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '          ";PWD=ata8244;DBNAME=192.168.2.82" & _
            '       ":C:\microsip datos\NEXTEL.FDB"


            '    cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '   ";PWD=ata8244;DBNAME=189.190.172.169" & _
            '":C:\microsip datos\IT RESOURCES TELECOM SA DE CV.FDB"

            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 116, ID_INTERNO = 9 where CONCEPTO_NO_ID = 9186 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function


    Public Function actualizarempleadoit()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String

            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
                             ";PWD=ata8244;DBNAME=192.168.2.83" & _
                          ":C:\microsip datos\IT RESOURCES TELECOM SA DE CV.FDB"

            'cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '          ";PWD=ata8244;DBNAME=192.168.2.82" & _
            '       ":C:\microsip datos\NEXTEL.FDB"


            '    cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '   ";PWD=ata8244;DBNAME=189.190.172.169" & _
            '":C:\microsip datos\IT RESOURCES TELECOM SA DE CV.FDB"

            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("update excep_empleados set dias_hrs_pagar = 15 where dias_hrs_pagar = 12 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function


    Public Function actualizarempleadowipsi()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String

            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
                             ";PWD=ata8244;DBNAME=192.168.2.83" & _
                          ":C:\microsip datos\WIPSI A C.FDB"

            'cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '          ";PWD=ata8244;DBNAME=192.168.2.82" & _
            '       ":C:\microsip datos\NEXTEL.FDB"


            '    cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '   ";PWD=ata8244;DBNAME=189.190.172.169" & _
            '":C:\microsip datos\GRUPO CONISAL.FDB"

            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("update excep_empleados set dias_hrs_pagar = 15 where dias_hrs_pagar = 7 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function


    ''aicel

    Public Function actualizaraicel()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String

            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
                             ";PWD=ata8244;DBNAME=192.168.2.83" & _
                          ":C:\microsip datos\AICEL.FDB"

            'cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '          ";PWD=ata8244;DBNAME=192.168.2.82" & _
            '       ":C:\microsip datos\NEXTEL.FDB"


            '    cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '   ";PWD=ata8244;DBNAME=189.190.172.169" & _
            '":C:\microsip datos\GRUPO CONISAL.FDB"

            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("update excep_empleados set dias_hrs_pagar = 0 where dias_hrs_pagar = 7 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function


    ''aicel
    ''actualizar abril morget

    Public Function actualizarmorget()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction

        Dim fecha As Date = Convert.ToDateTime(Combofecha.Text)
        Dim dia As String
        Dim mes As String
        Dim ano As String

        dia = fecha.Day
        mes = fecha.Month
        ano = fecha.Year

        Dim fechados As String = "'" & dia & "." & mes & "." & ano & "'"




        Try
            Dim cadenaODBC As String

            '         Dim cadenaODBC As String = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '   ";PWD=ata8244;DBNAME=192.168.2.82" & _
            '":C:\microsip datos\NEXTEL.FDB"



            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.82" & _
      ":C:\microsip datos\MORGET.FDB"

            'cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '          ";PWD=ata8244;DBNAME=192.168.2.82" & _
            '       ":C:\microsip datos\NEXTEL.FDB"


            '    cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '   ";PWD=ata8244;DBNAME=189.190.172.169" & _
            '":C:\microsip datos\GRUPO CONISAL.FDB"

            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("update excep_empleados_det set CONCEPTO_NO_ID = 116, ID_INTERNO = 9 where CONCEPTO_NO_ID = 7235 ")
                .Append("and exists( ")
                .Append("select fp.nombre, n.frepag_id, exp.nomina_id,n.fecha,em.nombre_completo, em.empleado_id,exp.excep_emp_id, concepto_no_id, exd.cuota  from empleados  em ")
                .Append("inner join excep_empleados exp ")
                .Append("on em.empleado_id = exp.empleado_id ")
                .Append("inner join excep_empleados_det   exd ")
                .Append("on exp.excep_emp_id =  exd.excep_emp_id ")
                .Append("inner join nominas n ")
                .Append("on exp.nomina_id = n.nomina_id ")
                .Append("inner join frecuencias_pago fp ")
                .Append("on n.frepag_id = fp.frepag_id ")
                .Append("where n.fecha = " & fechados & "  and fp.nombre = " & "'" & combocentro.Text & "'" & " )")



            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    Public Function actualizarmorgettres()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction

        Dim fecha As Date = Convert.ToDateTime(Combofecha.Text)
        Dim dia As String
        Dim mes As String
        Dim ano As String

        dia = fecha.Day
        mes = fecha.Month
        ano = fecha.Year

        Dim fechados As String = "'" & dia & "." & mes & "." & ano & "'"




        Try
            Dim cadenaODBC As String

            '         Dim cadenaODBC As String = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '   ";PWD=ata8244;DBNAME=192.168.2.82" & _
            '":C:\microsip datos\NEXTEL.FDB"



            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.82" & _
      ":C:\microsip datos\MORGET.FDB"

            'cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '          ";PWD=ata8244;DBNAME=192.168.2.82" & _
            '       ":C:\microsip datos\NEXTEL.FDB"


            '    cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '   ";PWD=ata8244;DBNAME=189.190.172.169" & _
            '":C:\microsip datos\GRUPO CONISAL.FDB"

            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("update excep_empleados_det set CONCEPTO_NO_ID = 156, ID_INTERNO = 7 where CONCEPTO_NO_ID = 7385 ")
                .Append("and exists( ")
                .Append("select fp.nombre, n.frepag_id, exp.nomina_id,n.fecha,em.nombre_completo, em.empleado_id,exp.excep_emp_id, concepto_no_id, exd.cuota  from empleados  em ")
                .Append("inner join excep_empleados exp ")
                .Append("on em.empleado_id = exp.empleado_id ")
                .Append("inner join excep_empleados_det   exd ")
                .Append("on exp.excep_emp_id =  exd.excep_emp_id ")
                .Append("inner join nominas n ")
                .Append("on exp.nomina_id = n.nomina_id ")
                .Append("inner join frecuencias_pago fp ")
                .Append("on n.frepag_id = fp.frepag_id ")
                .Append("where n.fecha = " & fechados & "  and fp.nombre = " & "'" & combocentro.Text & "'" & " )")



            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function


    Public Function actualizarmorgetDOS()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String

            '         Dim cadenaODBC As String = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '   ";PWD=ata8244;DBNAME=192.168.2.82" & _
            '":C:\microsip datos\NEXTEL.FDB"



            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.82" & _
      ":C:\microsip datos\MORGET.FDB"

            'cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '          ";PWD=ata8244;DBNAME=192.168.2.82" & _
            '       ":C:\microsip datos\NEXTEL.FDB"


            '    cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '   ";PWD=ata8244;DBNAME=189.190.172.169" & _
            '":C:\microsip datos\GRUPO CONISAL.FDB"

            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 156, ID_INTERNO = 7 where CONCEPTO_NO_ID = 7385 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    ''upetiloli

    Public Function actualizarhsemanalISR()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String

         


            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME=192.168.2.83" & _
       ":C:\microsip datos\UPHETILOLI 2.FDB"



            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 115, ID_INTERNO = 8 where CONCEPTO_NO_ID = 479 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    ''upetiloli


    ''actualizar morget semanal
    Public Function actualizarmsemanalseguro()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String





            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\1 MORGET SEMANAL.FDB"



            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 116, ID_INTERNO = 9 where CONCEPTO_NO_ID = 1850 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function


    Public Function actualizarmsemanalsubsidio()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String





            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\1 MORGET SEMANAL.FDB"



            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 156, ID_INTERNO = 7 where CONCEPTO_NO_ID = 1852 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function


    Public Function actualizarmsemanalISR()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String





            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\1 MORGET SEMANAL.FDB"



            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 115, ID_INTERNO = 8 where CONCEPTO_NO_ID = 1851 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    ''actualizar morget semanal


    ''actualizar morget catorcenal

    Public Function actualizarmcatorcenalseguro()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String





            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\2 MORGET CATORCENAL.FDB"



            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 116, ID_INTERNO = 9 where CONCEPTO_NO_ID = 1848 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    Public Function actualizarmcatorcenalsubsidio()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String





            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\2 MORGET CATORCENAL.FDB"



            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 156, ID_INTERNO = 7 where CONCEPTO_NO_ID = 1850 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    Public Function actualizarmcatorcenalISR()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String





            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\2 MORGET CATORCENAL.FDB"



            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 115, ID_INTERNO = 8 where CONCEPTO_NO_ID = 1849 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function
    ''actualizar morget catorcenal


    ''actualizar morget quincenal

    Public Function actualizarmquincenalseguro()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String





            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\3  MORGET QUINCENAL.FDB"



            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 116, ID_INTERNO = 9 where CONCEPTO_NO_ID = 1848 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    Public Function actualizarmquincenalsubsidio()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String





            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\3  MORGET QUINCENAL.FDB"



            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 156, ID_INTERNO = 7 where CONCEPTO_NO_ID = 1850 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    Public Function actualizarmquincenalISR()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String





            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\3  MORGET QUINCENAL.FDB"



            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 115, ID_INTERNO = 8 where CONCEPTO_NO_ID = 1849 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    ''actualizar morget quincenal


    ''actualizar morget mensual

    Public Function actualizarmmensualseguro()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String





            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\4 MORGET MENSUAL.FDB"



            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 116, ID_INTERNO = 9 where CONCEPTO_NO_ID = 1848 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    Public Function actualizarmmensualsubsidio()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String





            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\4 MORGET MENSUAL.FDB"



            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 156, ID_INTERNO = 7 where CONCEPTO_NO_ID = 1850 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    Public Function actualizarmmensualISR()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction



        Try
            Dim cadenaODBC As String





            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\4 MORGET MENSUAL.FDB"



            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 115, ID_INTERNO = 8 where CONCEPTO_NO_ID = 1849 ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    ''actualizar morget mensual


    ''actualizar abril morget


    Private Sub Button7_Click_1(sender As System.Object, e As System.EventArgs) Handles Button7.Click
        actualizar()
        actualizarempleado()
        actualizarit()
        actualizarempleadoit()

        MessageBox.Show("Actualización correcta")

    End Sub

    Private Sub btnactualizarm_Click(sender As System.Object, e As System.EventArgs) Handles btnactualizarm.Click
        '  actualizarmorget()
        ' actualizarmorgetDOS()
        'actualizarmorgettres()

        actualizarmsemanalseguro()
        actualizarmsemanalsubsidio()
        actualizarmsemanalISR()


        MessageBox.Show("Actualizacion Terminada")

    End Sub

    Private Sub Button8_Click_1(sender As System.Object, e As System.EventArgs) Handles Button8.Click

        'sArray.Add("BONOS")
        'sArray.Add("FALTAS")
        'sArray.Add("PRIMA DOMINICAL")
        'sArray.Add("FONACOT")

        Dim sArray As String() = {"Pension Alimenticia", "Haberes de retiro", "Contraprestacion", "Otros", "Otros Descuentos"}


        '    Dim sArray As New List(Of String) _
        'From {"BONOS", "FALTAS", "PRIMA DOMINICAL", "FONACOT"}

        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        ''FOR EACH
        For Each item As String In sArray
            Try
                Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
                Dim cnConex As New OleDbConnection(stConexion)
                Dim Cmd As New OleDbCommand("Select * From [" & item & "$]")
                ' Dim Cmd As New OleDbCommand("Select *From [BONOS$]")
                Dim Ds As New DataSet
                Dim Da As New OleDbDataAdapter
                Dim Dt As New DataTable
                cnConex.Open()
                Cmd.Connection = cnConex
                Da.SelectCommand = Cmd
                Da.Fill(Ds)
                Dt = Ds.Tables(0)
                ''if
                If item = "Pension Alimenticia" Then
                    Me.dgvpensionaw.Columns.Clear()
                    Me.dgvpensionaw.DataSource = Dt
                End If


                If item = "Haberes de retiro" Then
                    Me.dgvhaberesw.Columns.Clear()
                    Me.dgvhaberesw.DataSource = Dt
                End If
                If item = "Contraprestacion" Then
                    Me.dgvcontrpw.Columns.Clear()
                    Me.dgvcontrpw.DataSource = Dt
                End If

                If item = "Otros" Then
                    Me.otrosw.Columns.Clear()
                    Me.otrosw.DataSource = Dt
                End If
                If item = "Otros Descuentos" Then
                    Me.dgvotrosdescw.Columns.Clear()
                    Me.dgvotrosdescw.DataSource = Dt
                End If





            Catch ex As Exception
                ' MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                MsgBox("Ingresa el formato correcto del archivo de Excel")
            End Try
        Next

        If Me.dgvpensionaw.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Pensiones alimenticias", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        If Me.dgvhaberesw.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Haberes", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        If Me.dgvcontrpw.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Contraprestaciones", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        If Me.otrosw.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Otros ", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        If Me.dgvotrosdescw.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene exepciones de Otros Descuentos", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Private Sub Button10_Click_1(sender As System.Object, e As System.EventArgs) Handles Button10.Click
        My.Computer.FileSystem.DeleteFile("c:\exportaciones\exepciones.txt")
        Me.delete()
        Dim contador As Integer = 0

        Try

            For i As Integer = 0 To Me.dgvpensionaw.Rows.Count - 1
                With Me.dgvpensionaw.Rows(i)

                    Me.gvpensionaw(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 203)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try


        Try

            For i As Integer = 0 To Me.dgvhaberesw.Rows.Count - 1
                With Me.dgvhaberesw.Rows(i)

                    Me.gvhaberesw(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 203)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        Try

            For i As Integer = 0 To Me.dgvcontrpw.Rows.Count - 1
                With Me.dgvcontrpw.Rows(i)

                    Me.gvcontrpw(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 203)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        Try

            For i As Integer = 0 To Me.otrosw.Rows.Count - 1
                With Me.otrosw.Rows(i)

                    Me.gvotrosw(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 203)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        Try

            For i As Integer = 0 To Me.dgvotrosdescw.Rows.Count - 1
                With Me.dgvotrosdescw.Rows(i)

                    Me.gvotrosdescw(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 203)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        Muestradatos()

    End Sub

    Private Sub Button11_Click_1(sender As System.Object, e As System.EventArgs) Handles Button11.Click
        actualizarempleadowipsi()


    End Sub

    Private Sub btnmcato_Click(sender As System.Object, e As System.EventArgs) Handles btnmcato.Click
        actualizarmcatorcenalseguro()
        actualizarmcatorcenalsubsidio()
        actualizarmcatorcenalISR()
        MessageBox.Show("Actualizacion Terminada")
    End Sub

    Private Sub btnmquinc_Click(sender As System.Object, e As System.EventArgs) Handles btnmquinc.Click
        actualizarmquincenalseguro()
        actualizarmquincenalsubsidio()
        actualizarmquincenalISR()
        MessageBox.Show("Actualizacion Terminada")

    End Sub

    Private Sub btnmmensual_Click(sender As System.Object, e As System.EventArgs) Handles btnmmensual.Click
        actualizarmmensualseguro()
        actualizarmmensualsubsidio()
        actualizarmmensualISR()
        MessageBox.Show("Actualizacion Terminada")
    End Sub


    Public Function actualizarfoldurseguro()
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction

        Try
            Dim cadenaODBC As String


            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.82" & _
      ":C:\microsip datos\GRUPO CONISAL.FDB"

            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("Update excep_empleados_det set CONCEPTO_NO_ID = 116, ID_INTERNO = 9 where CONCEPTO_NO_ID = 8740 ")


            End With

            commODBC.CommandText = strQuery.ToString
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    Private Sub btnfoldur_Click(sender As System.Object, e As System.EventArgs) Handles btnfoldur.Click
        ' actualizarfoldurseguro()
    End Sub

    Private Sub BTNaicelex_Click(sender As System.Object, e As System.EventArgs) Handles BTNaicelex.Click
        If CheckBox1.Checked = True Then
            Dim sArray As String() = {"Haberes de retiro", "Otros", "Pension Alimenticia", "Otros Descuentos", "Empleados"}

            Dim stRuta As String = ""
            Dim openFD As New OpenFileDialog()
            With openFD
                .Title = "Seleccionar archivos"
                .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
                .Multiselect = False
                .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    stRuta = .FileName
                End If
            End With

            For Each item As String In sArray
                Try
                    Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
                    Dim cnConex As New OleDbConnection(stConexion)
                    Dim Cmd As New OleDbCommand("Select * From [" & item & "$]")
                    Dim Ds As New DataSet
                    Dim Da As New OleDbDataAdapter
                    Dim Dt As New DataTable
                    cnConex.Open()
                    Cmd.Connection = cnConex
                    Da.SelectCommand = Cmd
                    Da.Fill(Ds)
                    Dt = Ds.Tables(0)
                    ''if
                    If item = "Haberes de retiro" Then
                        Me.DGVaicelhaberes.Columns.Clear()
                        Me.DGVaicelhaberes.DataSource = Dt
                    End If


                    If item = "Otros" Then
                        Me.DGVaicelotros.Columns.Clear()
                        Me.DGVaicelotros.DataSource = Dt
                    End If
                    If item = "Pension Alimenticia" Then
                        Me.DGVaicelpension.Columns.Clear()
                        Me.DGVaicelpension.DataSource = Dt
                    End If

                    If item = "Otros Descuentos" Then
                        Me.DGVaicelodescuentos.Columns.Clear()
                        Me.DGVaicelodescuentos.DataSource = Dt
                    End If

                    If item = "Empleados" Then
                        Me.DataGridView16.Columns.Clear()
                        Me.DataGridView16.DataSource = Dt
                    End If


                Catch ex As Exception
                    ' MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                    MsgBox("Ingresa el formato correcto del archivo de Excel")
                End Try
            Next

            If Me.DataGridView16.Rows.Count = 0 Then
                MessageBox.Show("El excel no contiene el listado de Empleados", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If

            If Me.DGVaicelhaberes.Rows.Count = 0 Then
                MessageBox.Show("El excel no contiene exepciones de Haberes de retiro", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If

            If Me.DGVaicelotros.Rows.Count = 0 Then
                MessageBox.Show("El excel no contiene exepciones de Otros", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If

            If Me.DGVaicelpension.Rows.Count = 0 Then
                MessageBox.Show("El excel no contiene exepciones de Pension Alimenticia", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If

            If Me.DGVaicelodescuentos.Rows.Count = 0 Then
                MessageBox.Show("El excel no contiene exepciones de Otros Descuentos ", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If
        Else
            Dim sArray As String() = {"Haberes de retiro", "Otros", "Pension Alimenticia", "Otros Descuentos"}

            Dim stRuta As String = ""
            Dim openFD As New OpenFileDialog()
            With openFD
                .Title = "Seleccionar archivos"
                .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
                .Multiselect = False
                .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    stRuta = .FileName
                End If
            End With

            For Each item As String In sArray
                Try
                    Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
                    Dim cnConex As New OleDbConnection(stConexion)
                    Dim Cmd As New OleDbCommand("Select * From [" & item & "$]")
                    Dim Ds As New DataSet
                    Dim Da As New OleDbDataAdapter
                    Dim Dt As New DataTable
                    cnConex.Open()
                    Cmd.Connection = cnConex
                    Da.SelectCommand = Cmd
                    Da.Fill(Ds)
                    Dt = Ds.Tables(0)
                    ''if
                    If item = "Haberes de retiro" Then
                        Me.DGVaicelhaberes.Columns.Clear()
                        Me.DGVaicelhaberes.DataSource = Dt
                    End If


                    If item = "Otros" Then
                        Me.DGVaicelotros.Columns.Clear()
                        Me.DGVaicelotros.DataSource = Dt
                    End If
                    If item = "Pension Alimenticia" Then
                        Me.DGVaicelpension.Columns.Clear()
                        Me.DGVaicelpension.DataSource = Dt
                    End If

                    If item = "Otros Descuentos" Then
                        Me.DGVaicelodescuentos.Columns.Clear()
                        Me.DGVaicelodescuentos.DataSource = Dt
                    End If


                Catch ex As Exception
                    ' MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                    MsgBox("Ingresa el formato correcto del archivo de Excel")
                End Try
            Next

            If Me.DGVaicelhaberes.Rows.Count = 0 Then
                MessageBox.Show("El excel no contiene exepciones de Haberes de retiro", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If

            If Me.DGVaicelotros.Rows.Count = 0 Then
                MessageBox.Show("El excel no contiene exepciones de Otros", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If

            If Me.DGVaicelpension.Rows.Count = 0 Then
                MessageBox.Show("El excel no contiene exepciones de Pension Alimenticia", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If

            If Me.DGVaicelodescuentos.Rows.Count = 0 Then
                MessageBox.Show("El excel no contiene exepciones de Otros Descuentos ", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If
        End If

    End Sub

    Private Sub BTNaicelgen_Click(sender As System.Object, e As System.EventArgs) Handles BTNaicelgen.Click
        My.Computer.FileSystem.DeleteFile("c:\exportaciones\exepciones.txt")
        Me.delete()

        ''desactivar empleados

        If CheckBox1.Checked = True Then

            Me.desactivar()
            For i As Integer = 0 To Me.DataGridView16.Rows.Count - 1

                With Me.DataGridView16.Rows(i)

                    Me.estatusempleado(.Cells(0).Value)
                End With
            Next
        End If

        Dim contador As Integer = 0

        Try

            For i As Integer = 0 To Me.DGVaicelhaberes.Rows.Count - 1
                With Me.DGVaicelhaberes.Rows(i)

                    Me.GVaicelhaberes(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 203)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        Try

            For i As Integer = 0 To Me.DGVaicelotros.Rows.Count - 1
                With Me.DGVaicelotros.Rows(i)

                    Me.GVaicelotros(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 203)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        Try

            For i As Integer = 0 To Me.DGVaicelpension.Rows.Count - 1
                With Me.DGVaicelpension.Rows(i)

                    Me.GVaicelpension(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 203)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        Try

            For i As Integer = 0 To Me.DGVaicelodescuentos.Rows.Count - 1

                With Me.DGVaicelodescuentos.Rows(i)

                    Me.GVaicelodescuentos(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, 203)

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        Muestradatosaicel()
    End Sub

    Private Sub BTNaicelact_Click(sender As System.Object, e As System.EventArgs) Handles BTNaicelact.Click
        actualizaraicel()
        MsgBox("Actualización correcta")
    End Sub

    Private Sub TabPage7_Click(sender As System.Object, e As System.EventArgs) Handles TabPage7.Click

    End Sub

    Private Sub Button12_Click_1(sender As System.Object, e As System.EventArgs) Handles Button12.Click

        'Dim sArray As New ArrayList

        'sArray.Add("BONOS")
        'sArray.Add("FALTAS")
        'sArray.Add("PRIMA DOMINICAL")
        'sArray.Add("FONACOT")

        Dim sArray As String() = {"Ingresos Asimilados a Salarios", "ISR", "Otros"}


        '    Dim sArray As New List(Of String) _
        'From {"BONOS", "FALTAS", "PRIMA DOMINICAL", "FONACOT"}

        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        ''FOR EACH
        For Each item As String In sArray
            Try
                Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
                Dim cnConex As New OleDbConnection(stConexion)
                Dim Cmd As New OleDbCommand("Select * From [" & item & "$]")
                ' Dim Cmd As New OleDbCommand("Select *From [BONOS$]")
                Dim Ds As New DataSet
                Dim Da As New OleDbDataAdapter
                Dim Dt As New DataTable
                cnConex.Open()
                Cmd.Connection = cnConex
                Da.SelectCommand = Cmd
                Da.Fill(Ds)
                Dt = Ds.Tables(0)
                ''if

                If item = "Ingresos Asimilados a Salarios" Then
                    Me.DGVH046.Columns.Clear()
                    Me.DGVH046.DataSource = Dt
                End If


               
                If item = "ISR" Then
                    Me.DGVH002.Columns.Clear()
                    Me.DGVH002.DataSource = Dt
                End If
                If item = "Otros" Then
                    Me.DGVH004.Columns.Clear()
                    Me.DGVH004.DataSource = Dt
                End If
              
                ''nuevos
              

            Catch ex As Exception
                ' MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                MsgBox("Ingresa el formato correcto del archivo de Excel")
            End Try

        Next

        If Me.DGVH046.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene Ingresos Asimilados a Salarios.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        If Me.DGVH002.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene ISR", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        If Me.DGVH004.Rows.Count = 0 Then
            MessageBox.Show("El excel no contiene OTROS ", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If


        Button2.Enabled = True
    End Sub



    Private Sub Button13_Click(sender As System.Object, e As System.EventArgs) Handles Button13.Click
        My.Computer.FileSystem.DeleteFile("c:\exportaciones\exepciones.txt")
        Me.delete()
        Dim contador As Integer = 0



        Try

            For i As Integer = 0 To Me.DGVH046.Rows.Count - 1
                With Me.DGVH046.Rows(i)

                    Me.GVH046(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "12")

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try

        Try

            For i As Integer = 0 To Me.DGVH002.Rows.Count - 1
                With Me.DGVH002.Rows(i)

                    Me.GVH002(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "12")

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try
        Try

            For i As Integer = 0 To Me.DGVH004.Rows.Count - 1
                With Me.DGVH004.Rows(i)

                    Me.GVH004(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, "12")

                End With
            Next

        Catch ex As Exception

            MsgBox("Ingresa el formato correcto del archivo de Excel")
        End Try


        Muestradatoshup()
    End Sub

    Private Sub Button14_Click(sender As System.Object, e As System.EventArgs) Handles Button14.Click
        actualizarhsemanalISR()

    End Sub

    Private Sub CBXaicel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBXaicel.SelectedIndexChanged

    End Sub
End Class
