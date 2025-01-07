Public Class Form1

    ' Variables for tracking row count, balance, and selected record ID
    Dim maxrow As Integer
    Dim balance As Double = 0
    Dim id As Integer = 0

    ' Event handler for the Save button
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim money As Double

        ' Validate that the transaction type is selected
        If cboType.Text = "Select" Then
            MsgBox("Pls choose the transaction type", MsgBoxStyle.Exclamation)
            Return
        End If

        ' Validate that money and remarks fields are not empty
        If txtMoney.Text = "" Then
            MsgBox("Money cannot be null", MsgBoxStyle.Exclamation)
            Return
        End If
        If txtRemarks.Text = "" Then
            MsgBox("Remarks cannot be null", MsgBoxStyle.Exclamation)
            Return
        End If

        ' Check if the record exists in the database
        sql = "SELECT * FROM `tblbudget` WHERE BugetID = " & id

        maxrow = loadSingleResult(sql)

        If maxrow > 0 Then
            ' If record exists, confirm update
            If MessageBox.Show("Do you want to update this record?", "Update",
                               MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                ' Calculate updated balance based on transaction type
                If cboType.Text = "Received" Then
                    money = Double.Parse(txtMoney.Text) - Double.Parse(dt.Rows(0).Item(1))
                    balance = Double.Parse(dt.Rows(0).Item(3)) + money
                Else
                    money = Double.Parse(dt.Rows(0).Item(2)) - Double.Parse(txtMoney.Text)
                    balance = Double.Parse(dt.Rows(0).Item(3)) + money
                End If

                ' Update subsequent records' balances in the database
                Try
                    sql = "SELECT BugetID,BudgetBalance FROM `tblbudget` WHERE BugetID > " & id & " ORDER BY `BugetID` asc"
                    loadSingleResult(sql)
                    For Each r As DataRow In dt.Rows
                        sql = "UPDATE tblbudget SET BudgetBalance = BudgetBalance + '" & money & "' WHERE BugetID = " & r.Item(0)
                        executeQuery(sql)
                    Next
                Catch ex As Exception
                End Try

                ' Update the selected record in the database based on transaction type
                Select Case cboType.Text
                    Case "Received"
                        sql = "UPDATE `tblbudget` SET `BudgetIn`='" & txtMoney.Text & "',`BudgetBalance`= '" & balance & "',
                                `Remarks`='" & txtRemarks.Text & "',`TrasactionDate`='" & dtpTransDate.Text & "',`Type`='" & cboType.Text & "' 
                                WHERE `BugetID`=" & id
                        executeQuery(sql)
                    Case "Withdraw"
                        sql = "UPDATE `tblbudget` SET `BudgetOut`='" & txtMoney.Text & "',`BudgetBalance`= '" & balance & "',
                                `Remarks`='" & txtRemarks.Text & "',`TrasactionDate`='" & dtpTransDate.Text & "',`Type`='" & cboType.Text & "' 
                                WHERE `BugetID`=" & id
                        executeQuery(sql)
                End Select
            End If
        Else
            ' If record does not exist, confirm save
            If MessageBox.Show("Do you want to Save this record?", "Save", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                Select Case cboType.Text
                    Case "Received"
                        ' Calculate new balance
                        sql = "SELECT BudgetBalance FROM `tblbudget` ORDER BY `BugetID` DESC"
                        maxrow = loadSingleResult(sql)
                        If maxrow > 0 Then
                            balance = Double.Parse(dt.Rows(0).Item(0)) + Double.Parse(txtMoney.Text)
                        Else
                            balance = txtMoney.Text
                        End If

                        ' Insert new record for "Received" transaction
                        sql = "INSERT INTO `tblbudget`(`BudgetIn`, `BudgetOut`, `BudgetBalance`, `Remarks`, `TrasactionDate`,Type) 
                        VALUES ('" & txtMoney.Text & "',0,'" & balance & "','" & txtRemarks.Text & "','" & dtpTransDate.Text & "',
                            '" & cboType.Text & "')"
                        executeQuery(sql)
                    Case "Withdraw"
                        ' Validate sufficient balance for "Withdraw" transaction
                        sql = "SELECT BudgetBalance FROM `tblbudget` ORDER BY `BugetID` DESC"
                        maxrow = loadSingleResult(sql)
                        If maxrow > 0 Then
                            balance = Double.Parse(dt.Rows(0).Item(0)) - Double.Parse(txtMoney.Text)
                        Else
                            MsgBox("transaction cannot be proccess", MsgBoxStyle.Exclamation)
                            Return
                        End If

                        ' Insert new record for "Withdraw" transaction
                        sql = "INSERT INTO `tblbudget`(`BudgetIn`, `BudgetOut`, `BudgetBalance`, `Remarks`, `TrasactionDate`,Type) 
                        VALUES (0,'" & txtMoney.Text & "','" & balance & "','" & txtRemarks.Text & "','" & dtpTransDate.Text & "',
                            '" & cboType.Text & "')"
                        executeQuery(sql)
                End Select
            End If
        End If

        ' Clear form fields and refresh the data
        clear()
    End Sub

    ' Clears input fields and reloads the data grid
    Private Sub clear()
        id = 0
        txtMoney.Text = 0
        txtRemarks.Clear()
        cboType.Text = "Select"
        dtpTransDate.Text = Now()

        ' Load all records into the data grid
        sql = "SELECT `BugetID`,`TrasactionDate` as 'Date', `BudgetIn` as 'Recieved', `BudgetOut` as 'Withdraw', `BudgetBalance` as 'Balance', `Remarks`,  `Type` FROM `tblbudget` ORDER BY BugetID ASC"
        loadResultList(sql, dtglist)

        ' Recalculate totals and display them
        Dim recieve, withdraw, bal As Double
        For i As Integer = 0 To dtglist.RowCount - 1
            recieve += dtglist.Rows(i).Cells(2).Value
            withdraw += dtglist.Rows(i).Cells(3).Value
        Next
        bal = recieve - withdraw
        txtRecieved.Text = recieve.ToString("")
        txtWidthraw.Text = withdraw.ToString("N2")
        txtBalance.Text = bal.ToString("N2")
    End Sub

    ' Form load event to initialize data
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clear()
    End Sub

    ' Select all text when clicking on the Money field
    Private Sub txtMoney_Click(sender As Object, e As EventArgs) Handles txtMoney.Click
        txtMoney.SelectAll()
    End Sub

    ' Select all text when clicking on the Remarks field
    Private Sub txtRemarks_Click(sender As Object, e As EventArgs) Handles txtRemarks.Click
        txtRemarks.SelectAll()
    End Sub

    ' Search records based on a date range
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        sql = "SELECT `BugetID`,`TrasactionDate` as 'Date', `BudgetIn` as 'Received', `BudgetOut` as 'Withdraw', `BudgetBalance` as 'Balance', `Remarks`,  `Type` 
                FROM `tblbudget` WHERE DATE(TrasactionDate) BETWEEN '" & dtpfrom.Text & "' AND '" & dtpto.Text & "' ORDER BY BugetID ASC"
        loadResultList(sql, dtglist)

        ' Recalculate totals for the filtered data
        Dim recieve, withdraw, bal As Double
        For i As Integer = 0 To dtglist.RowCount - 1
            recieve += dtglist.Rows(i).Cells(2).Value
            withdraw += dtglist.Rows(i).Cells(3).Value
        Next
        bal = recieve - withdraw
        txtRecieved.Text = recieve.ToString("N2")
        txtWidthraw.Text = withdraw.ToString("N2")
        txtBalance.Text = bal.ToString("N2")
    End Sub

    ' Clear all filters and refresh data
    Private Sub btnclear_Click(sender As Object, e As EventArgs) Handles btnclear.Click
        clear()
    End Sub

    ' Delete the selected record
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Dim money As Double

        ' Determine the amount based on transaction type
        If dtglist.CurrentRow.Cells(6).Value = "Received" Then
            money = dtglist.CurrentRow.Cells(2).Value
        Else
            money = dtglist.CurrentRow.Cells(3).Value
        End If

        ' Confirm deletion
        If MessageBox.Show("Are you sure you want to delete this record?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
            sql = "SELECT BugetID,BudgetBalance FROM `tblbudget` 
                WHERE BugetID > " & dtglist.CurrentRow.Cells(0).Value & " ORDER BY `BugetID` asc"
            maxrow = loadSingleResult(sql)

            ' Update subsequent balances after deletion
            If maxrow > 0 Then
                For Each r As DataRow In dt.Rows
                    sql = "UPDATE tblbudget SET BudgetBalance = BudgetBalance - '" & money & "' WHERE BugetID = " & r.Item(0)
                    executeQuery(sql)
                Next
            End If

            ' Delete the record from the database
            sql = "DELETE FROM tblbudget WHERE BugetID = " & dtglist.CurrentRow.Cells(0).Value
            executeQuery(sql)
        End If

        ' Clear form fields and refresh data
        clear()
    End Sub

    ' Select the clicked record for editing
    Private Sub dtglist_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dtglist.CellClick
        id = dtglist.CurrentRow.Cells(0).Value
        dtpTransDate.Text = dtglist.CurrentRow.Cells(1).Value
        txtMoney.Text = If(dtglist.CurrentRow.Cells(2).Value = 0, dtglist.CurrentRow.Cells(3).Value, dtglist.CurrentRow.Cells(2).Value)
        cboType.Text = dtglist.CurrentRow.Cells(6).Value
        txtRemarks.Text = dtglist.CurrentRow.Cells(5).Value
    End Sub

    Private Sub txtMoney_TextChanged(sender As Object, e As EventArgs) Handles txtMoney.TextChanged

    End Sub
End Class
