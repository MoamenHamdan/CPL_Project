Imports MySql.Data.MySqlClient ' Import MySQL library for database connection and operations

Module config
    ' Database connection configuration
    Dim server As String = "localhost"
    Dim user As String = "root"
    Dim db As String = "db_budget"
    Dim pass As String = ""
    Dim strconection As String = "server=" & server & " ;user id=" & user & ";password=" & pass & ";database=" & db & ";sslMode=none"

    ' Command and data adapter objects for executing queries
    Dim cmd As MySqlCommand
    Dim da As MySqlDataAdapter

    ' Data table to store query results
    Public dt As DataTable

    ' Query string for SQL commands
    Public sql As String

    ' Create and return a new MySQL connection object
    Private Function mysqlcon() As MySqlConnection
        Return New MySqlConnection(strconection)
    End Function

    ' Public MySQL connection object
    Public con As MySqlConnection = mysqlcon()

    ' Executes a given SQL query (INSERT, UPDATE, DELETE)
    Public Sub executeQuery(sql As String)
        Try
            con.Open() ' Open the database connection
            cmd = New MySqlCommand ' Initialize command object
            With cmd
                .Connection = con ' Assign connection
                .CommandText = sql ' Assign query text
                .ExecuteNonQuery() ' Execute the query
            End With
        Catch ex As Exception
            MsgBox(ex.Message) ' Show error message in case of failure
        Finally
            con.Close() ' Close the connection
        End Try
    End Sub

    ' Loads query results into a DataGridView
    Public Sub loadResultList(sql As String, dtg As DataGridView)
        Try
            cmd = New MySqlCommand ' Initialize command object
            With cmd
                .Connection = con ' Assign connection
                .CommandText = sql ' Assign query text
            End With
            da = New MySqlDataAdapter ' Initialize data adapter
            da.SelectCommand = cmd ' Set the command to adapter
            dt = New DataTable ' Create a new DataTable
            da.Fill(dt) ' Fill the DataTable with results
            dtg.DataSource = dt ' Bind the DataTable to DataGridView
            dtg.Columns(0).Visible = False ' Hide the first column (ID or primary key)
        Catch ex As Exception
            MsgBox(ex.Message) ' Show error message in case of failure
        Finally
            con.Close() ' Close the connection
            da.Dispose() ' Dispose of the adapter
        End Try
    End Sub

    ' Loads a single result count from a query
    Public Function loadSingleResult(sql As String)
        Dim maxrow = 0 ' Variable to hold the count of rows
        Try
            cmd = New MySqlCommand ' Initialize command object
            With cmd
                .Connection = con ' Assign connection
                .CommandText = sql ' Assign query text
            End With
            da = New MySqlDataAdapter ' Initialize data adapter
            da.SelectCommand = cmd ' Set the command to adapter
            dt = New DataTable ' Create a new DataTable
            da.Fill(dt) ' Fill the DataTable with results
            maxrow = dt.Rows.Count ' Get the row count
        Catch ex As Exception
            MsgBox(ex.Message) ' Show error message in case of failure
        Finally
            con.Close() ' Close the connection
            da.Dispose() ' Dispose of the adapter
        End Try
        Return maxrow ' Return the row count
    End Function
End Module
