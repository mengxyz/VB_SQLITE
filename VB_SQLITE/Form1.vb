Imports System.Data.SQLite
Public Class Form1

    Dim conn As SQLiteConnection = New SQLiteConnection("Data Source=swimmer.db;Version=3;FailIfMissing=True;")
    Dim save_sta As String = ""
    Dim a, b, c, d, f As String

    Sub showdata()
        conn.Open()
        Dim sql As String
        sql = "select (Swimer.Fname ||' '|| Swimer.Lname) as 'Swimmer',cast(Distance.Length as text) as 'Length',Stroke.Stroke_ as 'Stroke',Gender.Gender_ as 'Gender',Stage.Stage_ as 'Stage',Result.Time,Result.Place "
        sql += "FROM Swimer,Gender,Result,Stage,Stroke,Distance "
        sql += "where Result.Swimer = Swimer.Swimer_ID and Result.Gender = Gender.Gender_ID and Result.Distance = Distance.Distance_ID "
        sql += "and Result.Stage = Stage.Stage_ID and Result.Stroke = Stroke.Stroke_ID order by Swimer.Swimer_ID"
        Dim da As SQLiteDataAdapter = New SQLiteDataAdapter(sql, conn)
        Dim ds As DataSet = New DataSet()
        da.Fill(ds, "result")
        dataGridView1.DataSource = ds
        dataGridView1.DataMember = "result"
        dataGridView1.ReadOnly = True
        conn.Close()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        showdata()
        cmb_exe()
        button2.Enabled = False
        button3.Enabled = False
        button4.Enabled = False
    End Sub

    Sub sql_exe(ByVal sql As String)
        conn.Open()
        Dim sqlcdm As SQLiteCommand = New SQLiteCommand(sql, conn)
        sqlcdm.ExecuteNonQuery()
        conn.Close()
    End Sub

    Sub cmb_exe()
        conn.Open()
        Dim da1 As SQLiteDataAdapter = New SQLiteDataAdapter("select Swimer_ID,(Fname ||' '|| Lname) as 'name' from Swimer", conn)
        Dim da2 As SQLiteDataAdapter = New SQLiteDataAdapter("select Distance_ID,cast(Length as text) as Length from Distance", conn)
        Dim da3 As SQLiteDataAdapter = New SQLiteDataAdapter("select * from Gender", conn)
        Dim da4 As SQLiteDataAdapter = New SQLiteDataAdapter("select * from Stage", conn)
        Dim da5 As SQLiteDataAdapter = New SQLiteDataAdapter("select * from Stroke", conn)
        Dim ds1 As DataSet = New DataSet()
        da1.Fill(ds1, "d1")
        da2.Fill(ds1, "d2")
        da3.Fill(ds1, "d3")
        da4.Fill(ds1, "d4")
        da5.Fill(ds1, "d5")
        conn.Close()
        cmbDistance.DataSource = ds1.Tables("d2")
        cmbSwimmer.DataSource = ds1.Tables("d1")
        cmbGender.DataSource = ds1.Tables("d3")
        cmbStage.DataSource = ds1.Tables("d4")
        cmbStroke.DataSource = ds1.Tables("d5")

        cmbDistance.ValueMember = "Distance_ID"
        cmbSwimmer.ValueMember = "Swimer_ID"
        cmbGender.ValueMember = "Gender_ID"
        cmbStage.ValueMember = "Stage_ID"
        cmbStroke.ValueMember = "Stroke_ID"

        cmbDistance.DisplayMember = "Length"
        cmbSwimmer.DisplayMember = "name"
        cmbGender.DisplayMember = "Gender_"
        cmbStage.DisplayMember = "Stage_"
        cmbStroke.DisplayMember = "Stroke_"

        cmbDistance.SelectedIndex = -1
        cmbGender.SelectedIndex = -1
        cmbStage.SelectedIndex = -1
        cmbStroke.SelectedIndex = -1
        cmbSwimmer.SelectedIndex = -1

        cmbDistance.Enabled = False
        cmbGender.Enabled = False
        cmbStage.Enabled = False
        cmbStroke.Enabled = False
        cmbSwimmer.Enabled = False

        txtPlace.Text = ""
        txtTime.Text = ""
        txtPlace.Enabled = False
        txtTime.Enabled = False
    End Sub

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        button1.Enabled = False
        button4.Enabled = True

        cmbDistance.SelectedIndex = -1
        cmbGender.SelectedIndex = -1
        cmbStage.SelectedIndex = -1
        cmbStroke.SelectedIndex = -1
        cmbSwimmer.SelectedIndex = -1

        txtPlace.Text = ""
        txtTime.Text = ""

        cmbDistance.Enabled = True
        cmbGender.Enabled = True
        cmbStage.Enabled = True
        cmbStroke.Enabled = True
        cmbSwimmer.Enabled = True
        txtPlace.Enabled = True
        txtTime.Enabled = True

        save_sta = "add"
    End Sub

    Private Sub button2_Click(sender As Object, e As EventArgs) Handles button2.Click
        save_sta = "edit"
        button4.Enabled = True
        button2.Enabled = False
        button3.Enabled = False
        cmbDistance.Enabled = True
        cmbGender.Enabled = True
        cmbStage.Enabled = True
        cmbStroke.Enabled = True
        cmbSwimmer.Enabled = True
        txtPlace.Enabled = True
        txtTime.Enabled = True
    End Sub

    Private Sub button4_Click(sender As Object, e As EventArgs) Handles button4.Click
        If (cmbDistance.SelectedIndex = -1 Or cmbGender.SelectedIndex = -1 Or cmbStage.SelectedIndex = -1 Or cmbStroke.SelectedIndex = -1 Or cmbSwimmer.SelectedIndex = -1 Or txtPlace.Text = Nothing Or txtTime.Text = Nothing) Then
            Exit Sub
        End If
        If (save_sta = "add") Then
            sql_exe("insert into Result values ('" + CStr(cmbSwimmer.SelectedValue) + "','" + CStr(cmbDistance.SelectedValue) + "','" + CStr(cmbStroke.SelectedValue) + "','" + CStr(cmbGender.SelectedValue) + "','" + CStr(cmbStage.SelectedValue) + "','" + CStr(txtTime.Text) + "','" + CStr(txtPlace.Text) + "')")
        End If


        If (save_sta = "edit") Then
            Dim sql As String
            sql = "update Result set Swimer = '" + cmbSwimmer.SelectedValue + "',Distance = '" + cmbDistance.SelectedValue + "',Stroke = '" + cmbStroke.SelectedValue + "',Gender = '" + cmbGender.SelectedValue + "',Stage = '" + cmbStage.SelectedValue + "',"
            sql += "Time = '" + txtTime.Text + "',Place = '" + txtPlace.Text + "' where "
            sql += "Swimer = '" + a + "' and Distance = '" + b + "' and Stroke = '" + c + "' and Gender = '" + d + "' and Stage = '" + f + "'"
            sql_exe(sql)
        End If
        showdata()
        button1.Enabled = True
        button2.Enabled = False
        button3.Enabled = False
        button4.Enabled = False

        cmbDistance.SelectedIndex = -1
        cmbGender.SelectedIndex = -1
        cmbStage.SelectedIndex = -1
        cmbStroke.SelectedIndex = -1
        cmbSwimmer.SelectedIndex = -1

        cmbDistance.Enabled = False
        cmbGender.Enabled = False
        cmbStage.Enabled = False
        cmbStroke.Enabled = False
        cmbSwimmer.Enabled = False

        txtPlace.Text = ""
        txtTime.Text = ""
        txtPlace.Enabled = False
        txtTime.Enabled = False
    End Sub

    Private Sub button3_Click(sender As Object, e As EventArgs) Handles button3.Click
        sql_exe("delete from Result where Swimer = '" + a + "' and Distance = '" + b + "' and Stroke = '" + c + "' and Gender = '" + d + "' and Stage = '" + f + "'")
        showdata()
        button1.Enabled = True
        button2.Enabled = False
        button3.Enabled = False
        button4.Enabled = False

        cmbDistance.SelectedIndex = -1
        cmbGender.SelectedIndex = -1
        cmbStage.SelectedIndex = -1
        cmbStroke.SelectedIndex = -1
        cmbSwimmer.SelectedIndex = -1

        cmbDistance.Enabled = False
        cmbGender.Enabled = False
        cmbStage.Enabled = False
        cmbStroke.Enabled = False
        cmbSwimmer.Enabled = False

        txtPlace.Text = ""
        txtTime.Text = ""
        txtPlace.Enabled = False
        txtTime.Enabled = False
    End Sub

    Private Sub dataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dataGridView1.CellContentDoubleClick
        cmbSwimmer.Text = dataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString()
        cmbDistance.Text = dataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString()
        cmbStroke.Text = dataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString()
        cmbGender.Text = dataGridView1.Rows(e.RowIndex).Cells(3).Value.ToString()
        cmbStage.Text = dataGridView1.Rows(e.RowIndex).Cells(4).Value.ToString()
        txtTime.Text = dataGridView1.Rows(e.RowIndex).Cells(5).Value.ToString()
        txtPlace.Text = dataGridView1.Rows(e.RowIndex).Cells(6).Value.ToString()
        a = cmbSwimmer.SelectedValue.ToString()
        b = cmbDistance.SelectedValue.ToString()
        c = cmbStroke.SelectedValue.ToString()
        d = cmbGender.SelectedValue.ToString()
        f = cmbStage.SelectedValue.ToString()

        button1.Enabled = False
        button2.Enabled = True
        button3.Enabled = True
        button4.Enabled = False
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        button1.Enabled = True
        button2.Enabled = False
        button3.Enabled = False
        button4.Enabled = False

        cmbDistance.SelectedIndex = -1
        cmbGender.SelectedIndex = -1
        cmbStage.SelectedIndex = -1
        cmbStroke.SelectedIndex = -1
        cmbSwimmer.SelectedIndex = -1

        cmbDistance.Enabled = False
        cmbGender.Enabled = False
        cmbStage.Enabled = False
        cmbStroke.Enabled = False
        cmbSwimmer.Enabled = False

        txtPlace.Text = ""
        txtTime.Text = ""
        txtPlace.Enabled = False
        txtTime.Enabled = False
    End Sub
End Class
