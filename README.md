# Vu-Th29.34
Imports System.Data.SqlClient

Module Module1
    Public conn As SqlConnection

    Public Sub KetNoi()
        Dim str As String = "Data Source=.\SQLEXPRESS;Initial Catalog=QLSV;Integrated Security=True"
        conn = New SqlConnection(str)
        Try
            conn.Open()
            MessageBox.Show("Kết nối thành công!")
        Catch ex As Exception
            MessageBox.Show("Kết nối thất bại: " & ex.Message)
        End Try
    End Sub
End Module

Private Sub btnThem_Click(sender As Object, e As EventArgs) Handles btnThem.Click
    Dim sql As String = "INSERT INTO SinhVien (MaSV, HoTen, Lop, NgaySinh, GioiTinh) 
                         VALUES (@MaSV, @HoTen, @Lop, @NgaySinh, @GioiTinh)"
    Dim cmd As New SqlCommand(sql, conn)
    cmd.Parameters.AddWithValue("@MaSV", txtMaSV.Text)
    cmd.Parameters.AddWithValue("@HoTen", txtHoTen.Text)
    cmd.Parameters.AddWithValue("@Lop", txtLop.Text)
    cmd.Parameters.AddWithValue("@NgaySinh", dtpNgaySinh.Value)
    cmd.Parameters.AddWithValue("@GioiTinh", cboGioiTinh.Text)

    cmd.ExecuteNonQuery()
    MessageBox.Show("Thêm thành công!")
    HienThi() ' Hàm load lại dữ liệu vào DataGridView
End Sub

Private Sub btnXoa_Click(sender As Object, e As EventArgs) Handles btnXoa.Click
    Dim sql As String = "DELETE FROM SinhVien WHERE MaSV=@MaSV"
    Dim cmd As New SqlCommand(sql, conn)
    cmd.Parameters.AddWithValue("@MaSV", txtMaSV.Text)

    cmd.ExecuteNonQuery()
    MessageBox.Show("Xóa thành công!")
    HienThi()
End Sub

Private Sub btnSua_Click(sender As Object, e As EventArgs) Handles btnSua.Click
    Dim sql As String = "UPDATE SinhVien 
                         SET HoTen=@HoTen, Lop=@Lop, NgaySinh=@NgaySinh, GioiTinh=@GioiTinh 
                         WHERE MaSV=@MaSV"
    Dim cmd As New SqlCommand(sql, conn)
    cmd.Parameters.AddWithValue("@MaSV", txtMaSV.Text)
    cmd.Parameters.AddWithValue("@HoTen", txtHoTen.Text)
    cmd.Parameters.AddWithValue("@Lop", txtLop.Text)
    cmd.Parameters.AddWithValue("@NgaySinh", dtpNgaySinh.Value)
    cmd.Parameters.AddWithValue("@GioiTinh", cboGioiTinh.Text)

    cmd.ExecuteNonQuery()
    MessageBox.Show("Sửa thành công!")
    HienThi()
End Sub
