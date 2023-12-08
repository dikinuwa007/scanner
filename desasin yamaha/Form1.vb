Imports MySql.Data.MySqlClient
Imports System.Media

Public Class Form1

    Dim con As MySqlConnection = New MySqlConnection("server=localhost;user id=root;database=yourDB;password=''")
    Dim cmd As MySqlCommand
    Dim da As MySqlDataAdapter
    Dim Ds As New DataSet
    Dim dt As DataTable
    Dim hasil, a, b, itemid_metindo, itemid_cust, dn, jobnoMetindo, jobnoCUST As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        TextScanItemCard.Select()
        TableLayoutPanel1.BackColor = Color.White
        'dataitemfromdn()
    End Sub

    Private Sub TextScanItemCard_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextScanItemCard.KeyPress
        Try


            If Asc(e.KeyChar) = 13 Then
                dn = Mid(TextScanItemCard.Text, 1, 16)
                jobnoCUST = Mid(TextScanItemCard.Text, 17, 7)
                TextScanItemCard.Text = jobnoCUST
                Try

                    cmd = New MySqlCommand
                        With cmd
                            .Connection = con
                        '.CommandText = "SELECT part_id,part_name,qty,imageData FROM tb_item WHERE part_id='" + TextScanItemCard.Text + "'"
                        .CommandText = "SELECT substr(part_id from 7),part_name,id,qty FROM tb_item WHERE jobno = '" + jobnoCUST + "'"
                    End With
                        da = New MySqlDataAdapter
                        dt = New DataTable
                        'Dim arrImage() As Byte
                        da.SelectCommand = cmd
                        da.Fill(dt)
                    itemid_cust = dt.Rows(0).Item(2)
                    Label_ItemPartName.Text = dt.Rows(0).Item(0)
                    TextScanItemCard.Text = jobnoCUST
                    Label_ItemQTY.Text = dt.Rows(0).Item(1)
                        Label_ItemJob.Text = dt.Rows(0).Item(3)
                        'End Try
                        Label_ItemJudgement.Text = "OK"
                    'Label_ItemQTY.Text = "qty"

                Catch ex As Exception
                    itemid_cust = ""
                    Label3.Text = "JOB NO"
                    Label_ItemJudgement.Text = "NG"
                    Label_ItemPartName.ResetText()
                    Label_ItemQTY.ResetText()
                    Label_ItemJob.ResetText()
                    PictureBox1.Image = PictureBox1.InitialImage

                Finally
                    da.Dispose()
                    con.Close()
                End Try
                TextScanPartCard.Select()
                e.Handled = True
            End If
        Catch ex As Exception
            Call CreateObject("WScript.Shell").Popup("Silahkan Scan Customer Card", 1, "WARNING", vbObjectError)
        End Try
    End Sub

    Private Sub btn_clear_Click(sender As Object, e As EventArgs) Handles btn_clear.Click
        txtDN.Clear()
    End Sub

    Private Sub TextScanPartCard_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextScanPartCard.KeyPress
        Try

            If Asc(e.KeyChar) = 13 Then
                Dim tes As String
                'Try
                'tes = Split(TextScanPartCard.Text, "#")
                'Catch ex As Exception
                tes = Microsoft.VisualBasic.Right(TextScanPartCard.Text, 7)
                'TextScanPartCard.Text, 5)
                '       End Try
                jobnoMetindo = tes
                TextScanPartCard.Text = jobnoMetindo
                Try

                    cmd = New MySqlCommand
                        With cmd
                            .Connection = con
                            '    .CommandText = "SELECT part_id,part_name,qty,imageData FROM tb_item WHERE part_id='" + TextScanItemCard.Text + "'"
                            .CommandText = "SELECT substr(part_id from 7),part_name,id,qty FROM tb_item WHERE jobno = '" + jobnoMetindo + "'"

                        End With
                        da = New MySqlDataAdapter
                        dt = New DataTable
                        da.SelectCommand = cmd
                        da.Fill(dt)
                        itemid_metindo = dt.Rows(0).Item(2)
                        TextScanPartCard.Text = jobnoMetindo
                        Label_PartPartName.Text = dt.Rows(0).Item(0)
                        Label_PartQTY.Text = dt.Rows(0).Item(1)
                        Label_PartJob.Text = dt.Rows(0).Item(3)
                    'End Try
                    Label_PartJudgement.Text = "OK"
                    'Label_PartPartName.Text = dt.Rows(0).Item(1)
                    'Label_PartJob.Text = jobno
                Catch ex As Exception
                    Label6.Text = "JOB NO"
                    Label_PartJudgement.Text = "NG"
                    Label_PartPartName.ResetText()
                    Label_PartQTY.ResetText()
                    Label_PartJob.ResetText()
                    PictureBox2.Image = PictureBox2.InitialImage
                    itemid_metindo = ""
                Finally
                    da.Dispose()
                    con.Close()
                End Try

                If Label_ItemJudgement.Text.Equals("OK") And Label_PartJudgement.Text.Equals("OK") And TextScanItemCard.Text.Equals(TextScanPartCard.Text) And Label_PartJob.Text.Equals(Label_ItemJob.Text) And Label_ItemQTY.Text.Equals(Label_PartQTY.Text) Then
                    'My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Asterisk)
                    TableLayoutPanel1.BackColor = Color.LimeGreen
                    hasil = "OK"
                    My.Computer.Audio.Play(My.Resources.OK, AudioPlayMode.WaitToComplete)
                    con.Open()
                    cmd = New MySqlCommand("insert into tb_log(itemid_cust,itemid_metindo,partid_cust,partid_metindo,qty_cust,qty_metindo,status_cust,status_metindo,hasil,jobno_cust,jobno_metindo,dn) values('" + itemid_cust + "','" + itemid_metindo + "','" + Label_ItemPartName.Text + "','" + Label_PartPartName.Text + "','" + Label_ItemJob.Text + "','" + Label_PartJob.Text + "','" + Label_ItemJudgement.Text + "','" + Label_PartJudgement.Text + "','" + hasil + "','" + jobnoCUST + "','" + jobnoMetindo + "','" + dn + "')", con)
                    cmd.ExecuteNonQuery()
                    con.Close()
                    'txtDN.Text = dn 'fitur baru
                    'dataitemfromdn(dn) 'fitur bary
                    'dataitemfromlog(dn) 'fitur baru
                    'RowsColor() 'fitur baru
                Else
                    TableLayoutPanel1.BackColor = Color.Red
                    hasil = "NG"
                    My.Computer.Audio.Play(My.Resources.ENG, AudioPlayMode.WaitToComplete)
                    con.Open()
                    cmd = New MySqlCommand("insert into tb_log(itemid_cust,itemid_metindo,partid_cust,partid_metindo,qty_cust,qty_metindo,status_cust,status_metindo,hasil,jobno_cust,jobno_metindo,dn) values('" + itemid_cust + "','" + itemid_metindo + "','" + Label_ItemPartName.Text + "','" + Label_PartPartName.Text + "','" + Label_ItemJob.Text + "','" + Label_PartJob.Text + "','" + Label_ItemJudgement.Text + "','" + Label_PartJudgement.Text + "','" + hasil + "','" + jobnoCUST + "','" + jobnoMetindo + "','" + dn + "')", con)
                    cmd.ExecuteNonQuery()
                    con.Close()

                End If
                TextScanItemCard.Select()
                e.Handled = True

            End If

        Catch ex As Exception
            'MsgBox()
            'Call CreateObject("WScript.Shell").Popup("Silahkan Scan Part Card", 1, "WARNING", vbObjectError)
        End Try
    End Sub

    Sub dataitemfromdn(tesDN)
        'Dim con As MySqlConnection = New MySqlConnection("server=192.168.2.12;user id=user;database=db_delivery_yamaha")
        Try
            con.Open()
            cmd = New MySqlCommand
            With cmd
                .Connection = con
                .CommandText = "SELECT jobno,totalkanban FROM tb_dn where dn='" + tesDN + "';"
            End With
            da = New MySqlDataAdapter
            da.SelectCommand = cmd
            dt = New DataTable
            da.Fill(dt)
            With dtglist
                .DataSource = dt
            End With

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            da.Dispose()
            con.Close()
        End Try

        With dtglist
            ' --- set incremental/sequential number in HeaderCell
            For rowNum As Integer = 0 To .Rows.Count - 1
                .Rows(rowNum).HeaderCell.Value = (rowNum + 1).ToString
            Next
            ' --- optional: auto adjust width of RowHeader and make string centering ---
            .AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)
            .RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
        Me.dtglist.Font = New Font("Arial", 12, FontStyle.Bold)
    End Sub
    Sub dataitemfromlog(tesDN)
        Try
            con.Open()
            cmd = New MySqlCommand
            With cmd
                .Connection = con
                .CommandText = "SELECT jobno_cust,count(*) as totalkanban FROM tb_log where dn='" + tesDN + "' and hasil='" + "OK" + "' group by jobno_cust;"
            End With
            da = New MySqlDataAdapter
            da.SelectCommand = cmd
            dt = New DataTable
            da.Fill(dt)
            With dtglistscan
                .DataSource = dt
            End With

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            da.Dispose()
            con.Close()
        End Try
        With dtglistscan
            ' --- set incremental/sequential number in HeaderCell
            For rowNum As Integer = 0 To .Rows.Count - 1
                .Rows(rowNum).HeaderCell.Value = (rowNum + 1).ToString
            Next
            ' --- optional: auto adjust width of RowHeader and make string centering ---
            .AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)
            .RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
        Me.dtglistscan.Font = New Font("Arial", 12, FontStyle.Bold)
    End Sub
    Sub modifycell()
        dtglistscan.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        For Each rowscan As DataGridViewRow In dtglistscan.Rows
            For Each row As DataGridViewRow In dtglist.Rows
                If rowscan.Cells("jobno_cust").Equals(row.Cells("jobno")) Then
                    rowscan.DefaultCellStyle.BackColor = Color.Red
                    MsgBox("sama")
                Else
                    rowscan.DefaultCellStyle.BackColor = Color.Green
                    MsgBox("BEDA")
                End If
            Next
        Next

        'dtglist.Item(ColumnIndex, RowIndex).Style.ForeColor = Color.Red
    End Sub
    '=============================================
    Public Sub RowsColor()
        Dim val As Integer = 0
        Dim jobno As String
        Try
            For i As Integer = 0 To dtglistscan.Rows.Count() - 1
                val = 0
                val = dtglistscan.Rows(i).Cells(1).Value
                jobno = dtglistscan.Rows(i).Cells(0).Value
                For a As Integer = 0 To dtglist.Rows.Count() - 1
                    Dim vala As Integer = 0
                    Dim jobnoa As String
                    vala = dtglist.Rows(a).Cells(1).Value
                    jobnoa = dtglist.Rows(a).Cells(0).Value
                    If jobnoa.Equals(jobno) Then
                        If val >= vala Then
                            dtglistscan.Rows(i).DefaultCellStyle.BackColor = Color.Green
                            dtglist.Rows(a).DefaultCellStyle.BackColor = Color.Green
                        Else
                            dtglistscan.Rows(i).DefaultCellStyle.BackColor = Color.Red
                        End If
                        Exit For
                    Else
                        dtglistscan.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    End If

                Next
            Next
        Catch
            'MsgBox("Tidak ada data") UPDATE FITUR
        End Try
    End Sub

    Private Sub txtDN_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDN.KeyPress
        Dim tesDN As String
        If Asc(e.KeyChar) = 13 Then
            tesDN = Mid(txtDN.Text, 1, 16)
            dataitemfromdn(tesDN)
            dataitemfromlog(tesDN)

            RowsColor()
            txtDN.Text = tesDN
        End If
    End Sub

    Private Sub TabPage2_Click(sender As Object, e As MouseEventArgs) Handles TabPage2.Click
        txtDN.Select()
    End Sub


End Class