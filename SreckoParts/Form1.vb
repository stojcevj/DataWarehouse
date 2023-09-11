Imports System.Data.OleDb

Public Class Form1
    Private conString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "/Database.mdb; Jet OLEDB:Database Password = Srecko123"
    ReadOnly con As OleDbConnection = New OleDbConnection(conString)
    Dim cmd As OleDbCommand
    Dim adapter As OleDbDataAdapter
    ReadOnly dt As DataTable = New DataTable()
    Dim sortColumn As Integer = -1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.SetupListView()
        Retrieve()
        ListView1.Sorting = SortOrder.Ascending
        ListView1.BackColor = Color.LightYellow
        ListView1.ForeColor = Color.Blue
    End Sub

    Public Sub SetupListView()

        ListView1.View = View.Details
        ListView1.FullRowSelect = True
        ListView1.Columns.Add("ID", 40)
        ListView1.Columns.Add("NAME", 150)
        ListView1.Columns.Add("LOCATION", 150)
        ListView1.Columns.Add("CATEGORY", 150)
        ListView1.Columns.Add("SUBCATEGORY", 150)
        ListView1.Columns.Add("VALUE", 100)
        ListView1.Columns.Add("UOM", 80)
        ListView1.Columns.Add("COUNT", 90)
        ListView1.Columns.Add("PACKAGE", 150)
        ListView1.Columns.Add("DESCRIPTION", 400)
        ListView1.Columns.Add("UNIT_PRICE", 100)
        ListView1.Columns.Add("INFO", 200)
        ListView1.Columns.Add("IMG", 100)
        ListView1.Columns.Add("S_NAME", 150)
        ListView1.Columns.Add("S_ADDRESS", 150)
        ListView1.Columns.Add("S_PHONE", 150)
        ListView1.Columns.Add("S_PRICE", 150)
    End Sub


    Public Sub Add()

        Const SQL As String = "INSERT INTO Table1 ([NAME],[LOCATION],[CATEGORY],[SUBCATEGORY],[VALUE],[UOM],[COUNT],[PACKAGE],[DESCRIPTION],[UNIT PRICE],[INFO],[IMG],[S_NAME],[S_ADDRESS],[S_PHONE],[S_PRICE]) VALUES(@name,@location,@category,@subcategory,@value,@uom,@count,@package,@description,@unitprice,@info,@img,@s_name,@s_address,@s_phone,@s_price)"
        cmd = New OleDbCommand(SQL, con)
        cmd.Parameters.AddWithValue("@name", nametxt.Text)
        cmd.Parameters.AddWithValue("@location", locationtxt.Text)
        cmd.Parameters.AddWithValue("@category", categorytxt.Text)
        cmd.Parameters.AddWithValue("@subcategory", subcategorytxt.Text)
        cmd.Parameters.AddWithValue("@value", valuetxt.Text)
        cmd.Parameters.AddWithValue("@uom", uomtxt.Text)
        cmd.Parameters.AddWithValue("@count", countxt.Text)
        cmd.Parameters.AddWithValue("@package", packagetxt.Text)
        cmd.Parameters.AddWithValue("@description", descriptiontxt.Text)
        cmd.Parameters.AddWithValue("@unitprice", unittxt.Text)
        cmd.Parameters.AddWithValue("@info", infotxt.Text)
        cmd.Parameters.AddWithValue("@img", imgtxt.Text)
        cmd.Parameters.AddWithValue("@s_name", s_nametxt.Text)
        cmd.Parameters.AddWithValue("@s_address", s_addresstxt.Text)
        cmd.Parameters.AddWithValue("@s_phone", s_phonetxt.Text)
        cmd.Parameters.AddWithValue("@s_price", s_pricetxt.Text)

        Try

            con.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                MsgBox("Successfully Inserted")
                CleartextBoxes()
            End If
            con.Close()
            Retrieve()

        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()
        End Try


    End Sub


    Private Sub Populate(ByVal id As String, ByVal name As String, ByVal location As String, ByVal category As String, ByVal subcategory As String, ByVal value As String, ByVal uom As String, ByVal count As String, ByVal package As String, ByVal description As String, ByVal unitprice As String, ByVal info As String, ByVal img As String, ByVal s_name As String, ByVal s_address As String, ByVal s_phone As String, ByVal s_price As String)

        Dim row As String() = New String() {id, name, location, category, subcategory, value, uom, count, package, description, unitprice, info, img, s_name, s_address, s_phone, s_price}
        Dim items As ListViewItem = New ListViewItem(row)
        ListView1.Items.Add(items)

    End Sub

    Private Sub CSC()

        Dim Item As ListViewItem

        For Each Item In ListView1.Items

            If Item.SubItems(7).Text < 4 Then
                Item.SubItems(0).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(1).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(2).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(3).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(4).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(5).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(6).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(7).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(8).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(9).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(10).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(11).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(12).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(13).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(14).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(15).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(16).ForeColor = Color.FromArgb(255, 51, 51)

            End If

            Item.UseItemStyleForSubItems = False

        Next

    End Sub

    Private Sub Retrieve()

        Dim sql As String = "SELECT * FROM Table1 "
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next

            dt.Rows.Clear()
            con.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()
        End Try

        Dim Item As ListViewItem

        For Each Item In ListView1.Items

            If Item.SubItems(7).Text < 4 Then
                Item.SubItems(0).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(1).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(2).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(3).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(4).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(5).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(6).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(7).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(8).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(9).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(10).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(11).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(12).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(13).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(14).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(15).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(16).ForeColor = Color.FromArgb(255, 51, 51)

            End If
            Item.UseItemStyleForSubItems = False

        Next



    End Sub
    Private Sub UpdateLV(ByVal id As String, ByVal name As String, ByVal location As String, ByVal category As String, ByVal subcategory As String, ByVal value As String, ByVal uom As String, ByVal count As String, ByVal package As String, ByVal description As String, ByVal unitprice As String, ByVal info As String, ByVal img As String, ByVal s_name As String, ByVal s_address As String, ByVal s_phone As String, ByVal s_price As String)

        Dim sql As String = "UPDATE Table1 SET NAME='" + name + "',[LOCATION]='" + location + "',[CATEGORY]='" + category + "',[SUBCATEGORY]='" + subcategory + "',[VALUE]='" + value + "',[UOM]='" + uom + "',[COUNT]='" + count + "',[PACKAGE]='" + package + "',[DESCRIPTION]='" + description + "',[UNIT PRICE]='" + unitprice + "',[INFO]='" + info + "',[IMG]='" + img + "',[S_NAME]='" + s_name + "',[S_ADDRESS]='" + s_address + "',[S_PHONE]='" + s_phone + "',[S_PRICE]='" + s_price + "' WHERE ID=" + id + ""
        cmd = New OleDbCommand(sql, con)

        Try

            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.UpdateCommand = con.CreateCommand()
            adapter.UpdateCommand.CommandText = sql

            If (adapter.UpdateCommand.ExecuteNonQuery() > 0) Then
                MsgBox("Succesfully Updated")
                CleartextBoxes()
            End If
            con.Close()
            ListView1.Refresh()
            ListView1.Items.Clear()
            Retrieve()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

    End Sub

    Private Sub Delete(ByVal id As String)

        Dim sql As String = "DELETE FROM Table1 WHERE ID=" + id + ""
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.DeleteCommand = con.CreateCommand()
            adapter.DeleteCommand.CommandText = sql

            If MessageBox.Show("Are you sure to permanently delete this?", "DELETE", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.OK Then
                If cmd.ExecuteNonQuery() > 0 Then
                    CleartextBoxes()
                    MsgBox("Succesfully deleted")
                End If
            End If
            con.Close()
            ListView1.Items.Clear()
            Retrieve()

        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try


    End Sub

    Private Sub CleartextBoxes()

        nametxt.Text = ""
        locationtxt.Text = ""
        categorytxt.Text = ""
        subcategorytxt.Text = ""
        valuetxt.Text = ""
        uomtxt.Text = ""
        countxt.Text = ""
        packagetxt.Text = ""
        descriptiontxt.Text = ""
        unittxt.Text = ""
        infotxt.Text = ""
        imgtxt.Text = ""
        s_nametxt.Text = ""
        s_addresstxt.Text = ""
        s_phonetxt.Text = ""
        s_pricetxt.Text = ""

    End Sub

    Private Sub addbtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles addbtn.Click
        addbtn.BackColor = Color.Khaki

        Add()
        ListView1.Items.Clear()
        Retrieve()

        If (addbtn.BackColor = Color.Khaki) Then

            addbtn.BackColor = Color.White

        End If
    End Sub

    Private Sub retrievebtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles retrievebtn.Click
        retrievebtn.BackColor = Color.Khaki


        ListView1.Items.Clear()
        Retrieve()
        ListView1.ForeColor = Color.Blue

        If (retrievebtn.BackColor = Color.Khaki) Then

            retrievebtn.BackColor = Color.White

        End If
    End Sub

    Private Sub updatebtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles updatebtn.Click

        updatebtn.BackColor = Color.Khaki

        Dim selectedIndex As Int32 = ListView1.SelectedIndices(0)

        If Not selectedIndex = -1 Then
            Dim id As String = ListView1.SelectedItems(0).SubItems(0).Text
            ' Dim id as Int32 = Convert.ToInt32(selected)
            UpdateLV(id, nametxt.Text, locationtxt.Text, categorytxt.Text, subcategorytxt.Text, valuetxt.Text, uomtxt.Text, countxt.Text, packagetxt.Text, descriptiontxt.Text, unittxt.Text, infotxt.Text, imgtxt.Text, s_nametxt.Text, s_addresstxt.Text, s_phonetxt.Text, s_pricetxt.Text)
            ListView1.Items.Clear()
            Retrieve()
        End If

        If (updatebtn.BackColor = Color.Khaki) Then

            updatebtn.BackColor = Color.White

        End If
    End Sub

    Private Sub deletebtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deletebtn.Click
        deletebtn.BackColor = Color.Khaki

        Dim selectedIndex As Int32 = ListView1.SelectedIndices(0)

        If Not selectedIndex = -1 Then
            Dim id As String = ListView1.SelectedItems(0).SubItems(0).Text
            'Dim id as Int32 = Convert.ToInt32(selected)
            Delete(id)
        End If

        If (deletebtn.BackColor = Color.Khaki) Then

            deletebtn.BackColor = Color.White

        End If
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        Try
            Dim selectedIndex As Int32 = ListView1.SelectedIndices(0)

            If Not selectedIndex = -1 Then

                If ListView1.SelectedItems(0).SubItems(0).Text IsNot Nothing Then

                    Dim name As String = ListView1.SelectedItems(0).SubItems(1).Text
                    Dim location As String = ListView1.SelectedItems(0).SubItems(2).Text
                    Dim category As String = ListView1.SelectedItems(0).SubItems(3).Text
                    Dim subcatergory As String = ListView1.SelectedItems(0).SubItems(4).Text
                    Dim value As String = ListView1.SelectedItems(0).SubItems(5).Text
                    Dim uom As String = ListView1.SelectedItems(0).SubItems(6).Text
                    Dim count As String = ListView1.SelectedItems(0).SubItems(7).Text
                    Dim package As String = ListView1.SelectedItems(0).SubItems(8).Text
                    Dim description As String = ListView1.SelectedItems(0).SubItems(9).Text
                    Dim unitprice As String = ListView1.SelectedItems(0).SubItems(10).Text
                    Dim info As String = ListView1.SelectedItems(0).SubItems(11).Text
                    Dim img As String = ListView1.SelectedItems(0).SubItems(12).Text
                    Dim s_name As String = ListView1.SelectedItems(0).SubItems(13).Text
                    Dim s_address As String = ListView1.SelectedItems(0).SubItems(14).Text
                    Dim s_phone As String = ListView1.SelectedItems(0).SubItems(15).Text
                    Dim s_price As String = ListView1.SelectedItems(0).SubItems(16).Text
                    nametxt.Text = name
                    locationtxt.Text = location
                    categorytxt.Text = category
                    subcategorytxt.Text = subcatergory
                    valuetxt.Text = value
                    uomtxt.Text = uom
                    countxt.Text = count
                    packagetxt.Text = package
                    descriptiontxt.Text = description
                    unittxt.Text = unitprice
                    infotxt.Text = info
                    imgtxt.Text = img
                    s_nametxt.Text = s_name
                    s_addresstxt.Text = s_address
                    s_phonetxt.Text = s_phone
                    s_pricetxt.Text = s_price
                    AxAcroPDF1.src = Application.StartupPath & "/Datasheets/" & infotxt.Text & ".pdf"
                    PictureBox1.ImageLocation = Application.StartupPath & "\Images\" & imgtxt.Text & ".jpg"



                End If

            End If

            If (countxt.Text <= 3) Then

                countxt.BackColor = Color.Red
                countxt.ForeColor = Color.White
            Else
                countxt.BackColor = Color.White
                countxt.ForeColor = Color.Black

            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim sql As String = "SELECT * FROM Table1 WHERE NAME like '%" & TextBox1.Text & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next

            dt.Rows.Clear()
            con.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim Item As ListViewItem

        For Each Item In ListView1.Items

            If Item.SubItems(7).Text < 4 Then
                Item.SubItems(0).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(1).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(2).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(3).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(4).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(5).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(6).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(7).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(8).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(9).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(10).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(11).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(12).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(13).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(14).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(15).ForeColor = Color.FromArgb(255, 51, 51)
                Item.SubItems(16).ForeColor = Color.FromArgb(255, 51, 51)

            End If
            Item.UseItemStyleForSubItems = False

        Next

    End Sub

    Private Sub ListView1_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles ListView1.ColumnClick

        If e.Column <> sortColumn Then


            sortColumn = e.Column


            ListView1.Sorting = SortOrder.Ascending

        Else

            If ListView1.Sorting = SortOrder.Ascending Then
                ListView1.Sorting = SortOrder.Descending
            Else
                ListView1.Sorting = SortOrder.Ascending
            End If
        End If


        Me.ListView1.ListViewItemSorter = New ListViewItemComparer(e.Column, ListView1.Sorting)


        ListView1.Sort()
    End Sub

    Private Sub fororder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fororder.Click

        Dim sql As String = "SELECT * FROM Table1 WHERE COUNT <= 3"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()
            ListView1.ForeColor = Color.Red

        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

    End Sub


    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Form2.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ulr As String = "https://www.google.com/maps/search/" & s_addresstxt.Text
        Process.Start(ulr)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        ContextMenuStrip1.Show(Button4, 0, Button4.Height)

    End Sub


    Private Sub METERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles METERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "AMP" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "METER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub AMPToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AMPToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "AMP" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub AUDIOToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AUDIOToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "AUDIO" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub BULBSBNCToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BULBSBNCToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "BNC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub BELTSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BELTSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "BELTS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub BULBSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BULBSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "BULBS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SOCKETToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SOCKETToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "BULBS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "SOCKET" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub LIGHTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LIGHTToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "BULBS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "LIGHT" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub NEONToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NEONToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "BULBS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "NEON" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub BALLSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BALLSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "BEARING" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "BALLS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub BEARINGToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BEARINGToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "BEARING" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub BANANAToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BANANAToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "BANANA" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub PLUGToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PLUGToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "BANANA" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "PLUG" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub BATTERIESToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BATTERIESToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "BATTERIES" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub HOLDERSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HOLDERSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "BATTERIES" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "HOLDERS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub CAPACITORToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CAPACITORToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "CAPACITOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub FILMPToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FILMPToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "CAPACITOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "FILM" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub TANTALToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TANTALToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "CAPACITOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "TANTAL" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub ELKOToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ELKOToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "CAPACITOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "ELKO" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub CERAMICToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CERAMICToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "CAPACITOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "CERAMIC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub TRIMToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TRIMToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "CAPACITOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "TRIM" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub FEEDTROUGHTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FEEDTROUGHTToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "CAPACITOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "FEED TROUGHT" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub CRISTALToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CRISTALToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "CRYSTAL" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub CROCODILEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CROCODILEToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "CROCODILE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub CLIPSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CLIPSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "CROCODILE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "CLIPS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub DIODEToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DIODEToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "DIODE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub ZENERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZENERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "DIODE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "ZENER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub DIACToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DIACToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "DIAC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub DOTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DOTToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "DOT-MATRIX" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub DISPLAYToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DISPLAYToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "DISPLAY" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SEGMENTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SEGMENTToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "DISPLAY" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "7 SEGMENT" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub ENDSLEEVESToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ENDSLEEVESToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "END SLEEVES" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub FASTENERToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FASTENERToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "FASTENER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()

    End Sub

    Private Sub PLASTICToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PLASTICToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "FASTENER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "PLASTIC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub FUSEToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FUSEToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "FUSE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SCOKETToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SCOKETToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "FUSE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "SOCKET" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub FILTERToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FILTERToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "FILTER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub FERRITEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FERRITEToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "FILTER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "FERRITE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()

    End Sub

    Private Sub FERRITEToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FERRITEToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "FERRITE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()

    End Sub

    Private Sub GYROSCOPEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GYROSCOPEToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "GYRO SCOPE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()

    End Sub

    Private Sub GDTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GDTToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "GDT" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub GEARSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GEARSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "GEARS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub HEATSINKToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HEATSINKToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "HEAT SINK" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub HALLEFECTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HALLEFECTToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "HALL EFFECT" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub HOURToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HOURToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "HOUR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub METERToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles METERToolStripMenuItem4.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "HOUR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "METER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub ICToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ICToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SPECIALFUNCTIONToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SPECIALFUNCTIONToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "SPECIAL FUNCTION" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SPECIALToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SPECIALToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "SPECIAL" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub LINEDRIVERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LINEDRIVERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "LINE DRIVER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub ANALOGDIGITALToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ANALOGDIGITALToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "ANALOG-DIGITAL" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub COUNTERCOMPARERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles COUNTERCOMPARERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "COUNTER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub COMPAPRERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles COMPAPRERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "COMPERATOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub TIMERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TIMERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "TIMER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub DIGITALANALOGToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DIGITALANALOGToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "DIGITAL-ANALOG" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub AMPLIFIERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AMPLIFIERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "AMPLIFIER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub TTLSERIEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TTLSERIEToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "74 TTL SERIE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub LSSERIEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LSSERIEToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "74 LS SERIE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub HCTSERIEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HCTSERIEToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "74 HCT SERIE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub HCSERIEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HCSERIEToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "74 HC SERIE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        CSC()

       



    End Sub

    Private Sub CMOSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMOSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "4000 CMOS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub IGBTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IGBTToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "IGBT" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub INDUCTORToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles INDUCTORToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "INDUCTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub KNOBSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KNOBSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "KNOBS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub LENSESToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LENSESToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "LENSES" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub LDRToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LDRToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "LDR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub LEDToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LEDToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "LED" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub DRIVERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DRIVERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "LED" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "DRIVER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub HOLDERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HOLDERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "LED" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "HOLDER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub IREMITTERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IREMITTERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "LED" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "IR EMITTER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub MOSFETToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MOSFETToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "MOSFET" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub NPNToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NPNToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "MOSFET" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "NPN" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub PNPToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PNPToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "MOSFET" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "PNP" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub PAIRSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PAIRSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "MOSFET" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "PAIRS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub MICROCHIPToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MICROCHIPToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "MICROCHIP" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub MICROPHONEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MICROPHONEToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "MICROPHONE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub MOTORToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MOTORToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "MOTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub MINIToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MINIToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "MOTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "MINI" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub VRUSHToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VRUSHToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "MOTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "BRUSH" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SPEEDCONTROLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SPEEDCONTROLToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "MOTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "SPEED CONTROL" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SMALLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SMALLToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "MOTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "SMALL" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SHAFTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SHAFTToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "MOTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "SHAFT" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub MAGNETSToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MAGNETSToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "MAGNETS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub ELECTROToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ELECTROToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "MAGNETS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "ELECTRO" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub NITINOLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NITINOLToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "NITINOL" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub OSCILATORToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OSCILATORToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "OSCILATOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub OPAMPSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OPAMPSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "OP-AMPS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub OPTICToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OPTICToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "OPTIC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub FIBERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FIBERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "OPTIC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "FIBER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub OPTICALToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OPTICALToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "OPTICAL" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SENSORToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SENSORToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "OPTICAL" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "SENSOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub POTENTIOMETERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POTENTIOMETERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "POTENTIOMETER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub PRISMToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PRISMToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "PRISM" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub PIEZOBUZZERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PIEZOBUZZERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "PIEZO BUZZER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub PELTIERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PELTIERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "PELTIER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub PCBToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PCBToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "PCB" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub PINSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PINSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "PCB" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "PINS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub POWERToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POWERToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "POWER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub MATERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MATERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "POWER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "METER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub PLUGToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PLUGToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "POWER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "PLUG" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub ADAPTERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ADAPTERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "POWER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "ADAPTER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SOCKETToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SOCKETToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "POWER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "SOCKET" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub PROGRAMMINGToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PROGRAMMINGToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "PROGRAMMING" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub RESISTORToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RESISTORToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "RESISTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub VARIABLEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VARIABLEToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "RESISTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "VARIABLE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub HIGHCURRENTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HIGHCURRENTToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "RESISTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "HIGH CURRENT" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub RESONATORToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RESONATORToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "RESONATOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub RFToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RFToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "RF" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub RUSSIANToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RUSSIANToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "RUSSIAN COMPONENTS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub RELAYToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RELAYToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "RELAY" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub REEDToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles REEDToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "RELAY" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "REED" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub RUBBERToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RUBBERToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "RUBBER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub GROMMETSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GROMMETSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "RUBBER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "GROMMET" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub LEGSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LEGSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "RUBBER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "LEGS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SWITCHToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SWITCHToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SWITCH" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub ROTARYToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ROTARYToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SWITCH" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "ROTARY" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub THERMOToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles THERMOToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SWITCH" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "THERMO" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub TILTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TILTToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SWITCH" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "TILT" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SLIDEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SLIDEToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SWITCH" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "SLIDE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub MICROToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MICROToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SWITCH" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "MICRO" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub DIPToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DIPToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SWITCH" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "DIP" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub ROCKERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ROCKERToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SWITCH" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "ROCKER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub PUSHBUTTONToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PUSHBUTTONToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SWITCH" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "PUSH BUTTON" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub TOGGLEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TOGGLEToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SWITCH" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "TOGGLE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SENSORToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SENSORToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SENSOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SOLARToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SOLARToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SOLAR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SOLDERBRIDGEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SOLDERBRIDGEToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SOLDER BRIDGE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SPECIALToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SPECIALToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SPECIAL" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SHUNTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SHUNTToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SHUNT" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SHRINKTUBEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SHRINKTUBEToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SHRINK TUBE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SPRINGSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SPRINGSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SPRINGS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SPACERToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SPACERToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SPACER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub METALToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles METALToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "SPACER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "METAL" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub TRANSISTORToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TRANSISTORToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TRANSISTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub NPNToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NPNToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TRANSISTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "NPN" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub PNPToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PNPToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TRANSISTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "PNP" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub PAIRSToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PAIRSToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TRANSISTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "PAIRS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub JFETToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles JFETToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TRANSISTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "J-FET" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub INSULATORToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles INSULATORToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TRANSISTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "INSULATOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub SCRToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SCRToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TRIAC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "SCR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub TRIACToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TRIACToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TRIAC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try


    End Sub

    Private Sub TVSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TVSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TVS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
    End Sub

    Private Sub TERMINALSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TERMINALSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TERMINALS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
    End Sub

    Private Sub TUBEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TUBEToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TUBE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
    End Sub

    Private Sub SOCKETToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SOCKETToolStripMenuItem2.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TUBE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "SOCKET" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub COILToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles COILToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TUBE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "COIL" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub THERMISTORToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles THERMISTORToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "THERMISTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub NTCToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NTCToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "THERMISTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "NTC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub PTCToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PTCToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "THERMISTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "PTC" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub TEMPToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TEMPToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TEMP" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
    End Sub

    Private Sub METERToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles METERToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TEMP" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "METER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub TACHOToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TACHOToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TACHO" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
    End Sub

    Private Sub METERToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles METERToolStripMenuItem2.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "TACHO" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "METER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
        CSC()
    End Sub

    Private Sub VOLTToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VOLTToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "VOLT" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
    End Sub

    Private Sub METERToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles METERToolStripMenuItem3.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "VOLT" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "METER" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
    End Sub

    Private Sub REGULATORToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles REGULATORToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "VOLT" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "REGULATOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
    End Sub

    Private Sub WIREToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WIREToolStripMenuItem1.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "WIRE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

    End Sub

    Private Sub NUTSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUTSToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "WIRE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "NUTS" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
    End Sub

    Private Sub CONNECTORToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CONNECTORToolStripMenuItem.Click
        Dim sql As String = "SELECT * FROM Table1 WHERE CATEGORY like '%" & "WIRE" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try

        Dim sql1 As String = "SELECT * FROM Table1 WHERE SUBCATEGORY like '%" & "CONNECTOR" & "%'"
        ListView1.Items.Clear()
        cmd = New OleDbCommand(sql1, con)

        Try
            con.Open()
            adapter = New OleDbDataAdapter(cmd)
            adapter.Fill(dt)

            For Each row In dt.Rows
                Populate(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9), row(10), row(11), row(12), row(13), row(14), row(15), row(16))
            Next
            dt.Rows.Clear()
            con.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
            con.Close()

        End Try
    End Sub
End Class
