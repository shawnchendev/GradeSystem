Imports System.Data.OleDb
Public Class Form1
    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Shawn\Documents\Visual Studio 2012\Projects\FinalProject\FinalProject\bin\Debug\grade.accdb")
    Dim WithEvents student As New grade()
    Dim dt As New DataTable()
    Dim dt1 As New DataTable()
    Dim fmstr As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        setscrollBar()
    End Sub
    'get scroll value as gpa
    Private Sub HScrollBar1_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBar1.Scroll
        gpaTxt.Text = CStr(FormatNumber((HScrollBar1.Value / 10), 1))
        student.GPA_index = HScrollBar1.Value
        scoreTxt.Text = CStr(student.getGPA)
    End Sub
    'set scroll bar
    Sub setscrollBar()
        HScrollBar1.Maximum = 49
        HScrollBar1.Minimum = 20
        HScrollBar1.LargeChange = 10
        HScrollBar1.SmallChange = 5
        HScrollBar1.Value = 20
    End Sub
    'cal calculate the data 
    Private Sub calData()
        fmstr = "{0,-12}{1,10}{2,15}{3,15}{4,15}{5,15}{6,15}{7,15}{8,15}{9,15}{10,15}"
        displayLb1.Items.Clear()
        NameIDTxt()
        SAT_SelectedIndex()
        HSQ_SelectedIndex()
        DOC_SelectedIndex()
        GEO_SelectedIndex()
        Alunmi_SelectedIndex()
        Essay_SelectedIndex()
        LS_SelectedIndex()
        MiS_SelectedIndex()
        displayLb1.Items.Add(String.Format(fmstr, "ID", "Name", "GPA", "SAT", "School", "Curriclum", "Geography", "Alunmi", "Essay", "Leadership", "Miscellaneous"))
        fmstr = "{0,-10}{1,10}{2,18}{3,15}{4,20}{5,20}{6,20}{7,19}{8,19}{9,20}{10,20}"
        displayLb1.Items.Add(String.Format(fmstr, student.ID1, student.Name1, student.getGPA(), student.getSAT(), student.getHSQ(), student.getDOC(), student.getGeo(), student.getAlunmi(), student.getEssay(), student.getLS(), student.getMIS()))
        displayLb1.Items.Add("Total: " & student.getTotal())

        If student.getTotal > 100 Then
            displayLb1.Items.Add("Admited")
        End If
    End Sub
    'set name
    Private Sub NameIDTxt()
        student.Name1 = nametxt.Text
        student.ID1 = idtxt.Text

    End Sub
    ' the Following sub procedure  will represent each group box

    'set SAT index
    Private Sub SAT_SelectedIndex()
        If SATrb1.Checked Then
            student.SAT_index = SATrb1.TabIndex


        ElseIf SATrb2.Checked Then
            student.SAT_index = SATrb2.TabIndex


        ElseIf SATrb3.Checked Then
            student.SAT_index = SATrb3.TabIndex


        ElseIf SATrb4.Checked Then
            student.SAT_index = SATrb4.TabIndex

        ElseIf SATrb5.Checked Then
            student.SAT_index = SATrb5.TabIndex
        End If
    End Sub
    'set High school quality index 
    Private Sub HSQ_SelectedIndex()
        If HSQrb0.Checked Then
            student.HSQ_index = HSQrb0.TabIndex

        ElseIf HSQrb1.Checked Then
            student.HSQ_index = HSQrb1.TabIndex

        ElseIf HSQrb2.Checked Then
            student.HSQ_index = HSQrb2.TabIndex

        ElseIf HSQrb3.Checked Then
            student.HSQ_index = HSQrb3.TabIndex

        ElseIf HSQrb4.Checked Then
            student.HSQ_index = HSQrb4.TabIndex

        ElseIf HSQrb5.Checked Then
            student.HSQ_index = HSQrb5.TabIndex
        End If
    End Sub
    'set difficulty of curriculum
    Private Sub DOC_SelectedIndex()
        If DCrb0.Checked Then
            student.DOC_index = DCrb0.TabIndex
        ElseIf DCrb1.Checked Then
            student.DOC_index = DCrb1.TabIndex

        ElseIf DCrb2.Checked Then
            student.DOC_index = DCrb2.TabIndex

        ElseIf DCrb3.Checked Then
            student.DOC_index = DCrb3.TabIndex

        ElseIf DCrb4.Checked Then
            student.DOC_index = DCrb4.TabIndex

        ElseIf DCrb5.Checked Then
            student.DOC_index = DCrb5.TabIndex

        ElseIf DCrb6.Checked Then
            student.DOC_index = DCrb6.TabIndex
        End If
    End Sub
    'set Geography index
    Private Sub GEO_SelectedIndex()

        If GEOcb1.Checked Then
            student.Geo_geo1 = True
        End If
        If GEOcb2.Checked Then
            student.Geo_geo2 = True
        End If
        If GEOcb3.Checked Then
            student.Geo_geo3 = True
        End If
    End Sub
    'set Alunmi index
    Private Sub Alunmi_SelectedIndex()
        If ALUcb1.Checked Then
            student.Alunmi_Alun1 = True
        End If
        If ALUcb2.Checked Then
            student.Alunmi_Alun2 = True
        End If
    End Sub
    'set essay index
    Private Sub Essay_SelectedIndex()
        If ESSAYrb1.Checked Then
            student.Essay_index = ESSAYrb1.TabIndex
        ElseIf ESSAYrb2.Checked Then
            student.Essay_index = ESSAYrb2.TabIndex
        ElseIf ESSAYrb3.Checked Then
            student.Essay_index = ESSAYrb3.TabIndex
        End If
    End Sub
    'set leader ship and serive index
    Private Sub LS_SelectedIndex()
        If LScb1.Checked Then
            student.Leadership_LS1 = True
        End If
        If LScb2.Checked Then
            student.Leadership_LS2 = True
        End If
        If LScb3.Checked Then
            student.Leadership_LS3 = True
        End If
    End Sub
    'set miscellaneouse
    Private Sub MiS_SelectedIndex()
        If MISrb1.Checked Then
            student.mis_index = MISrb1.TabIndex
        ElseIf MISrb2.Checked Then
            student.mis_index = MISrb2.TabIndex
        ElseIf MISrb3.Checked Then
            student.mis_index = MISrb3.TabIndex
        ElseIf MISrb4.Checked Then
            student.mis_index = MISrb4.TabIndex
        End If
    End Sub
    'write data sqlquery process
    Sub writedata()
        'insert record to index table
        Dim sqlinsert As String = ""
        sqlinsert = "INSERT INTO student_index VALUES (@ID,@Name,@GPA,@SAT,@School,@Difficulty,@Geo1, @Geo2, @Geo3, @Alu1,@Alu2,@Essay,@leader1, @leader2,@leader3, @Misscellaneous)"
        Dim cmd As New OleDbCommand(sqlinsert, con)

        cmd.Parameters.Add(New OleDbParameter("@ID", student.ID1))
        cmd.Parameters.Add(New OleDbParameter("@Name", student.Name1))
        cmd.Parameters.Add(New OleDbParameter("@GPA", student.GPA_index))
        cmd.Parameters.Add(New OleDbParameter("@SAT", student.SAT_index))
        cmd.Parameters.Add(New OleDbParameter("@School", student.HSQ_index))
        cmd.Parameters.Add(New OleDbParameter("@Difficulty", student.DOC_index))
        cmd.Parameters.Add(New OleDbParameter("@Geo1", student.Geo_geo1))
        cmd.Parameters.Add(New OleDbParameter("@Geo2", student.Geo_geo2))
        cmd.Parameters.Add(New OleDbParameter("@Geo3", student.Geo_geo3))
        cmd.Parameters.Add(New OleDbParameter("@Alu1", student.Alunmi_Alun1))
        cmd.Parameters.Add(New OleDbParameter("@Alu2", student.Alunmi_Alun2))
        cmd.Parameters.Add(New OleDbParameter("@Essay", student.Essay_index))
        cmd.Parameters.Add(New OleDbParameter("@leader1", student.Leadership_LS1))
        cmd.Parameters.Add(New OleDbParameter("@leader2", student.Leadership_LS2))
        cmd.Parameters.Add(New OleDbParameter("@leader3", student.Leadership_LS3))
        cmd.Parameters.Add(New OleDbParameter("@Misscellaneous", student.mis_index))

        'insert record to grade table
        Dim sqlinsert1 As String
        sqlinsert1 = "INSERT INTO Grade VALUES(@ID, @Name, @GPA, @SAT, @school, @diff, @geo, @alu, @essay, @leadership,@mis, @total)"
        Dim cmd1 As New OleDbCommand(sqlinsert1, con)
        cmd1.Parameters.Add(New OleDbParameter("@ID", student.ID1))
        cmd1.Parameters.Add(New OleDbParameter("@Name", student.Name1))
        cmd1.Parameters.Add(New OleDbParameter("@GPA", student.getGPA()))
        cmd1.Parameters.Add(New OleDbParameter("@SAT", student.getSAT()))
        cmd1.Parameters.Add(New OleDbParameter("@school", student.getHSQ()))
        cmd1.Parameters.Add(New OleDbParameter("@diff", student.getDOC()))
        cmd1.Parameters.Add(New OleDbParameter("@geo", student.getGeo()))
        cmd1.Parameters.Add(New OleDbParameter("@alu", student.getAlunmi()))
        cmd1.Parameters.Add(New OleDbParameter("@essay", student.getEssay()))
        cmd1.Parameters.Add(New OleDbParameter("@leadership", student.getLS()))
        cmd1.Parameters.Add(New OleDbParameter("@mis", student.getMIS()))
        cmd1.Parameters.Add(New OleDbParameter("@total", student.getTotal()))
        'execute the query statement
        Try
            con.Open()
            cmd.ExecuteNonQuery()
            cmd1.ExecuteNonQuery()
        Catch ex As OleDbException
            MessageBox.Show(ex.Message)
        Finally
            Try
                con.Close()
            Catch ex As Exception

            End Try
        End Try


        MessageBox.Show("Successful Adding")
    End Sub
    'calculate button and write the data to the database
    Private Sub CALbtn1_Click(sender As Object, e As EventArgs) Handles CALbtn1.Click
        calData()
        writedata()
    End Sub
    ' read database query process
    Private Sub readTable()
        Dim sql1 As String
        Dim found As Boolean = False
        Dim row As Integer
        sql1 = "SELECT * FROM student_index"
        Try
            con.Open()
            Dim adapter1 As New OleDbDataAdapter(sql1, con)
            adapter1.Fill(dt)
            adapter1.Dispose()

            For i As Integer = 0 To (dt.Rows.Count - 1)
                If nametxt.Text = dt.Rows(i)("Name") Or idtxt.Text = dt.Rows(i)("ID") Then
                    found = True
                    row = i
                End If
            Next
            If found Then
                student.Name1 = CStr(dt.Rows(row)("Name"))
                student.ID1 = CInt(dt.Rows(row)("ID"))
                student.GPA_index = CInt(dt.Rows(row)("GPA_Index"))
                student.SAT_index = CInt(dt.Rows(row)("SAT_Index"))
                student.HSQ_index = CInt(dt.Rows(row)("School_Index"))
                student.DOC_index = CInt(dt.Rows(row)("Difficulty_Index"))
                student.Geo_geo1 = CStr(dt.Rows(row)("Geo1"))
                student.Geo_geo2 = CStr(dt.Rows(row)("Geo2"))
                student.Geo_geo3 = CStr(dt.Rows(row)("Geo3"))
                student.Alunmi_Alun1 = CStr(dt.Rows(row)("Alunmi1"))
                student.Alunmi_Alun2 = CStr(dt.Rows(row)("Alunmi2"))
                student.Essay_index = CInt(dt.Rows(row)("Essay_Index"))
                student.Leadership_LS1 = CStr(dt.Rows(row)("Leadership1"))
                student.Leadership_LS2 = CStr(dt.Rows(row)("Leadership2"))
                student.Leadership_LS3 = CStr(dt.Rows(row)("Leadership3"))
                student.mis_index = CInt(dt.Rows(row)("Miscellaneous_Index"))
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        con.Close()

    End Sub
    ' group all sub procedure for read database and display into form
    Private Sub setForm()
        rname()
        rGPA()
        rSAT()
        rschool()
        rcurriclum()
        rgeo()
        ralunmi()
        ressay()
        rleadership()
        rmiscellaneous()
    End Sub
    'the following sub procedures are use the data from database to the class variable 

    Sub rname()
        nametxt.Text = student.Name1
        idtxt.Text = student.ID1
    End Sub

    Sub rGPA()
        HScrollBar1.Value = student.GPA_index
        gpaTxt.Text = CStr(HScrollBar1.Value / 10)
        scoreTxt.Text = CStr(HScrollBar1.Value * 2)
    End Sub

    Sub rSAT()
        Select Case student.SAT_index
            Case SATrb1.TabIndex
                SATrb1.Checked = True
            Case SATrb2.TabIndex
                SATrb2.Checked = True
            Case SATrb3.TabIndex
                SATrb3.Checked = True
            Case SATrb4.TabIndex
                SATrb4.Checked = True
            Case SATrb5.TabIndex
                SATrb5.Checked = True
        End Select
    End Sub

    Sub rschool()
        Select Case student.HSQ_index
            Case HSQrb0.TabIndex
                HSQrb0.Checked = True
            Case HSQrb1.TabIndex
                HSQrb1.Checked = True
            Case HSQrb2.TabIndex
                HSQrb2.Checked = True
            Case HSQrb3.TabIndex
                HSQrb3.Checked = True
            Case HSQrb4.TabIndex
                HSQrb4.Checked = True
            Case HSQrb5.TabIndex
                HSQrb5.Checked = True
        End Select
    End Sub

    Sub rcurriclum()
        Select Case student.DOC_index
            Case DCrb0.TabIndex
                DCrb0.Checked = True
            Case DCrb1.TabIndex
                DCrb1.Checked = True
            Case DCrb2.TabIndex
                DCrb2.Checked = True
            Case DCrb3.TabIndex
                DCrb3.Checked = True
            Case DCrb4.TabIndex
                DCrb4.Checked = True
            Case DCrb5.TabIndex
                DCrb5.Checked = True
            Case DCrb6.TabIndex
                DCrb6.Checked = True
        End Select
    End Sub

    Sub rgeo()
        If student.Geo_geo1 = True Then
            GEOcb1.Checked = True
        End If
        If student.Geo_geo2 = True Then
            GEOcb2.Checked = True
        End If
        If student.Geo_geo3 = True Then
            GEOcb3.Checked = True
        End If
    End Sub

    Sub ralunmi()
        If student.Alunmi_Alun1 = True Then
            ALUcb1.Checked = True
        End If
        If student.Alunmi_Alun2 = True Then
            ALUcb2.Checked = True
        End If
    End Sub

    Sub ressay()
        Select Case student.Essay_index
            Case ESSAYrb1.TabIndex
                ESSAYrb1.Checked = True
            Case ESSAYrb2.TabIndex
                ESSAYrb2.Checked = True
            Case ESSAYrb3.TabIndex
                ESSAYrb3.Checked = True
        End Select
    End Sub

    Sub rleadership()
        If student.Leadership_LS1 = True Then
            LScb1.Checked = True
        End If
        If student.Leadership_LS2 = True Then
            LScb2.Checked = True
        End If
        If student.Leadership_LS3 = True Then
            LScb3.Checked = True
        End If
    End Sub

    Sub rmiscellaneous()
        Select Case student.mis_index
            Case MISrb1.TabIndex
                MISrb1.Checked = True
            Case MISrb2.TabIndex
                MISrb2.Checked = True
            Case MISrb3.TabIndex
                MISrb3.Checked = True
            Case MISrb4.TabIndex
                MISrb4.Checked = True
        End Select
    End Sub
    'delete record query process
    Sub deleterecord()
        Dim sqldel As String
        sqldel = "DELETE FROM student_index WHERE ID= '" & idtxt.Text & "' AND Name= '" & nametxt.Text & "'"
        Dim cmd As New OleDbCommand(sqldel, con)

        Dim sqldel1 As String
        sqldel1 = "DELETE FROM Grade WHERE ID= '" & idtxt.Text & "' AND Name= '" & nametxt.Text & "'"
        Dim cmd1 As New OleDbCommand(sqldel1, con)
        Try
            con.Open()
            cmd.ExecuteNonQuery()
            cmd1.ExecuteNonQuery()
        Catch ex As OleDbException
            MessageBox.Show(ex.Message)
        Finally
            Try
                con.Close()
            Catch ex As Exception

            End Try
        End Try
    End Sub

    'update record query process
    Sub updatedata()

        Dim sqlupdate As String
        sqlupdate = "UPDATE Student_index SET ID='@ID', Name='@name', GPA_Index='@gpa',SAT_Index='@sat', School_Index='@scl', Difficulty_Index='@diff', Geo1='@geo1', Geo2='@geo2',Geo3='@geo3', Alunmi1='@alu1',Alunmi2='@alu2',Essay_Index='@essay',Leadership1='@leader1',Leadership2='@leader2',Leadership3='@leader3',Miscellaneous_Index='@mis' WHERE ID='@i'"
        Dim cmd As New OleDbCommand(sqlupdate, con)

        cmd.Parameters.Add(New OleDbParameter("@ID", student.ID1))
        cmd.Parameters.Add(New OleDbParameter("@name", student.Name1))
        cmd.Parameters.Add(New OleDbParameter("@gpa", student.GPA_index))
        cmd.Parameters.Add(New OleDbParameter("@sat", student.SAT_index))
        cmd.Parameters.Add(New OleDbParameter("@scl", student.HSQ_index))
        cmd.Parameters.Add(New OleDbParameter("@diff", student.DOC_index))
        cmd.Parameters.Add(New OleDbParameter("@geo1", student.Geo_geo1))
        cmd.Parameters.Add(New OleDbParameter("@geo2", student.Geo_geo2))
        cmd.Parameters.Add(New OleDbParameter("@geo3", student.Geo_geo3))
        cmd.Parameters.Add(New OleDbParameter("@alu1", student.Alunmi_Alun1))
        cmd.Parameters.Add(New OleDbParameter("@alu2", student.Alunmi_Alun2))
        cmd.Parameters.Add(New OleDbParameter("@essay", student.Essay_index))
        cmd.Parameters.Add(New OleDbParameter("@leader1", student.Leadership_LS1))
        cmd.Parameters.Add(New OleDbParameter("@leader2", student.Leadership_LS2))
        cmd.Parameters.Add(New OleDbParameter("@leader3", student.Leadership_LS3))
        cmd.Parameters.Add(New OleDbParameter("@mis", student.mis_index))
        cmd.Parameters.Add(New OleDbParameter("@i", idtxt.Text))

        Dim sqlupdate1 As String
        sqlupdate1 = "UPDATE Grade SET ID ='@ID', Name='@Name', GPA='@GPA',SAT='@SAT',HSQ='@school',DOC='@diff',GEO='@geo',Alunmi='@alu',Essay='@essay',Leadership='@leadership',Miscellaneous='@mis',Total='@total' WHERE ID='@i'"
        Dim cmd1 As New OleDbCommand(sqlupdate1, con)

        cmd1.Parameters.Add(New OleDbParameter("@ID", student.ID1))
        cmd1.Parameters.Add(New OleDbParameter("@Name", student.Name1))
        cmd1.Parameters.Add(New OleDbParameter("@GPA", student.getGPA))
        cmd1.Parameters.Add(New OleDbParameter("@SAT", student.getSAT))
        cmd1.Parameters.Add(New OleDbParameter("@school", student.getHSQ))
        cmd1.Parameters.Add(New OleDbParameter("@diff", student.getDOC))
        cmd1.Parameters.Add(New OleDbParameter("@geo", student.getGeo))
        cmd1.Parameters.Add(New OleDbParameter("@alu", student.getAlunmi))
        cmd1.Parameters.Add(New OleDbParameter("@essay", student.getEssay))
        cmd1.Parameters.Add(New OleDbParameter("@leadership", student.getLS))
        cmd1.Parameters.Add(New OleDbParameter("@mis", student.getMIS))
        cmd1.Parameters.Add(New OleDbParameter("@total", student.getTotal))
        cmd1.Parameters.Add(New OleDbParameter("@i", idtxt.Text))

        Try
            con.Open()
            cmd.ExecuteNonQuery()
            cmd1.ExecuteNonQuery()
        Catch ex As OleDbException
            MessageBox.Show(ex.Message)
        Finally
            Try
                con.Close()
            Catch ex As Exception

            End Try
        End Try

    End Sub
    'search student and display the grade
    Private Sub readbtn_Click(sender As Object, e As EventArgs) Handles readbtn.Click
        readTable()
        setForm()
        calData()

    End Sub
    'clear the form 
    Private Sub clearbtn_Click(sender As Object, e As EventArgs) Handles clearbtn.Click
        HScrollBar1.Value = 20
        displayLb1.Items.Clear()
        nametxt.Clear()
        idtxt.Clear()
        nametxt.Focus()

        SATrb1.Checked = False
        SATrb3.Checked = False
        SATrb4.Checked = False
        SATrb5.Checked = False

        HSQrb0.Checked = False
        HSQrb1.Checked = False
        HSQrb2.Checked = False
        HSQrb3.Checked = False
        HSQrb4.Checked = False
        HSQrb5.Checked = False

        DCrb0.Checked = False
        DCrb1.Checked = False
        DCrb2.Checked = False
        DCrb3.Checked = False
        DCrb4.Checked = False
        DCrb5.Checked = False
        DCrb6.Checked = False

        GEOcb1.Checked = False
        GEOcb2.Checked = False
        GEOcb3.Checked = False

        ALUcb1.Checked = False
        ALUcb2.Checked = False

        ESSAYrb1.Checked = False
        ESSAYrb2.Checked = False
        ESSAYrb3.Checked = False

        LScb1.Checked = False
        LScb2.Checked = False
        LScb3.Checked = False

        MISrb1.Checked = False
        MISrb2.Checked = False
        MISrb3.Checked = False
        MISrb4.Checked = False

    End Sub
    'update record to the databse
    Private Sub updatebtn_Click(sender As Object, e As EventArgs) Handles updatebtn.Click
        NameIDTxt()
        SAT_SelectedIndex()
        HSQ_SelectedIndex()
        DOC_SelectedIndex()
        GEO_SelectedIndex()
        Alunmi_SelectedIndex()
        Essay_SelectedIndex()
        LS_SelectedIndex()
        MiS_SelectedIndex()
        updatedata()
        calData()
    End Sub
    'delete record form database
    Private Sub delete_Click(sender As Object, e As EventArgs) Handles delete.Click
        Dim confirm As String = ""
        If nametxt.Text = "" And idtxt.Text = "" Then
            MsgBox("You must fill Name and ID to process Delete Function")
        Else
            confirm = InputBox("Do you want delete this Record (Y/N)", "Warning", "N")
            If confirm = "y" Or confirm = "Y" Then
                deleterecord()
                MsgBox("Delete Successfully")
            End If
        End If
    End Sub
    'close application
    Private Sub closebtn_Click(sender As Object, e As EventArgs) Handles closebtn.Click
        Me.Close()
    End Sub
    ' show all student grade in the database
    Private Sub Showbtn_Click(sender As Object, e As EventArgs) Handles Showbtn.Click
        Form2.Show()
        Form2.ListBox1.Items.Clear()
        Dim dt1 As New DataTable()
        Dim sql As String = "SELECT * FROM Grade"
        Dim adapter As New OleDbDataAdapter(sql, con)
        Try
            adapter.Fill(dt1)
            adapter.Dispose()
            fmstr = "{0,-12}{1,10}{2,15}{3,15}{4,15}{5,15}{6,15}{7,15}{8,15}{9,15}{10,15}"
            Form2.ListBox1.Items.Add(String.Format(fmstr, "ID", "Name", "GPA", "SAT", "School", "Curriclum", "Geography", "Alunmi", "Essay", "Leadership", "Miscellaneous"))
            For i As Integer = 0 To (dt1.Rows.Count - 1)
                Form2.ListBox1.Items.Add(Environment.NewLine)
                fmstr = "{0,-10}{1,10}{2,18}{3,15}{4,20}{5,20}{6,20}{7,19}{8,19}{9,20}{10,20}"
                Form2.ListBox1.Items.Add(String.Format(fmstr, CInt(dt1.Rows(i)("ID")), CStr(dt1.Rows(i)("Name")), CInt(dt1.Rows(i)("GPA")), CInt(dt1.Rows(i)("SAT")), CInt(dt1.Rows(i)("HSQ")), CInt(dt1.Rows(i)("DOC")), CInt(dt1.Rows(i)("GEO")), CInt(dt1.Rows(i)("Alunmi")), CInt(dt1.Rows(i)("Essay")), CInt(dt1.Rows(i)("Leadership")), CInt(dt1.Rows(i)("Miscellaneous"))))
                Form2.ListBox1.Items.Add("Total: " & CInt(dt1.Rows(i)("Total")))

                If CInt(dt1.Rows(i)("Total")) > 100 Then
                    Form2.ListBox1.Items.Add("Admited")
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
End Class
