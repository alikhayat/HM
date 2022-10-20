Imports System.Data.OleDb
Public Class frmGuest

    Private Sub bttnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bttnSave.Click
        con.Open()
        Dim fname As String = Trim(txtFName.Text)
        Dim mname As String = Trim(txtMName.Text)
        Dim lname As String = Trim(txtLName.Text)
        Dim add As String = Trim(txtAddress.Text)
        Dim num As String = Trim(txtNumber.Text)
        Dim email As String = Trim(txtEmail.Text)
        Dim stat As String = "Active"
        Dim remark As String = "Available"

        If fname = Nothing Or mname = Nothing Or lname = Nothing Or add = Nothing Or num = Nothing Then
            MsgBox("Please Fill All Fields", vbInformation, "Note")
        Else
            Dim add_guest As New OleDbCommand("INSERT INTO tblGuest(GuestFName,GuestMName,GuestLName,GuestAddress,GuestContactNumber,GuestGender,GuestEmail,Status,Remarks) values ('" &
                                              fname & "','" &
                                              mname & "','" &
                                              lname & "','" &
                                              add & "','" &
                                              num & "','" &
                                              cboGender.Text & "','" &
                                              email & "','" &
                                              stat & "','" &
                                              remark & "')", con)
            add_guest.ExecuteNonQuery()
            add_guest.Dispose()
            MsgBox("Guest Added!", vbInformation, "Note")
            txtFName.Clear()
            txtMName.Clear()
            txtLName.Clear()
            txtAddress.Clear()
            txtNumber.Clear()
            txtEmail.Clear()
        End If
        con.Close()

    End Sub

    Private Sub frmGuest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'DatabaseDataSet.tblGuest' table. You can move, or remove it, as needed.
        Me.TblGuestTableAdapter.Fill(Me.DatabaseDataSet.tblGuest)
        

        TabControl1.SelectTab(0)
    End Sub

   

    Private Sub bttnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bttnCancel.Click
        txtFName.Clear()
        txtMName.Clear()
        txtLName.Clear()
        txtAddress.Clear()
        txtNumber.Clear()
        txtEmail.Clear()
    End Sub

    
   

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        


        Try
            Try
                Dim id As Integer = Convert.ToInt32(DataGridView1.SelectedRows(0).Cells(0).Value.ToString)
            
            Dim query As String = "DELETE * FROM [tblGuest] WHERE [ID]=? "

            Using conn As New OleDbConnection("Provider=MICROSOFT.ACE.OLEDB.12.0; Data Source=|DataDirectory|/database.accdb")


                Using cmd = New OleDbCommand(query, conn)
                    conn.Open()
                    cmd.Parameters.AddWithValue("@p1", id)

                    cmd.ExecuteNonQuery()
                    conn.Close()
                End Using
            End Using
            Me.TblGuestTableAdapter.Fill(Me.DatabaseDataSet.tblGuest)
            MsgBox("Deleted")
        Catch ex As Exception
            MsgBox(Convert.ToString(ex))

            End Try
        Catch ex As Exception
            MsgBox("nothing selected")
        End Try
    End Sub
End Class