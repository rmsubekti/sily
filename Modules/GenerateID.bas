Attribute VB_Name = "GenerateID"
Public Function generateIDPelanggan() As String
    Dim a As Integer
    sql = "select max(right(id_pelanggan,6)) from pelanggan"
    Set rs = conn.Execute(sql)
    a = IIf(rs(0) <> "NULL", rs(0) + 1, 1)
    If a <= 9 Then
        generateIDPelanggan = "P00000" & a
    ElseIf a > 9 And a < 100 Then
        generateIDPelanggan = "P0000" & a
    ElseIf a > 99 And a < 1000 Then
        generateIDPelanggan = "P000" & a
    ElseIf a > 999 And a < 10000 Then
        generateIDPelanggan = "P00" & a
    ElseIf a > 9999 And a < 100000 Then
        generateIDPelanggan = "P0" & a
    Else
        generateIDPelanggan = "P" & a
    End If
End Function
Public Function generateIDPaket() As String
    Dim a As Integer
    sql = "select max(right(id_paket,3)) from paket"
    Set rs = conn.Execute(sql)
    a = IIf(rs(0) <> "NULL", rs(0) + 1, 1)
    If a <= 9 Then
        generateIDPaket = "L00" & a
    ElseIf a > 9 And a < 100 Then
        generateIDPaket = "L0" & a
    Else
        generateIDPaket = "L" & a
    End If
End Function
Public Function generateIDKaryawan() As String
    Dim a As Integer
    sql = "select max(right(nik,6)) from karyawan"
    Set rs = conn.Execute(sql)
    a = IIf(rs(0) <> "NULL", rs(0) + 1, 1)
    If a <= 9 Then
        generateIDKaryawan = "K00000" & a
    ElseIf a > 9 And a < 100 Then
        generateIDKaryawan = "K0000" & a
    ElseIf a > 99 And a < 1000 Then
        generateIDKaryawan = "K000" & a
    ElseIf a > 999 And a < 10000 Then
        generateIDKaryawan = "K00" & a
    ElseIf a > 9999 And a < 100000 Then
        generateIDKaryawan = "K0" & a
    Else
        generateIDKaryawan = "K" & a
    End If
End Function
Public Function generateIDTransaksi() As String
    Dim a As Integer
    sql = "select max(right(id_transaksi,7)) from transaksi"
    Set rs = conn.Execute(sql)
    a = IIf(rs(0) <> "NULL", rs(0) + 1, 1)
    If a <= 9 Then
        generateIDTransaksi = "T000000" & a
    ElseIf a > 9 And a < 100 Then
        generateIDTransaksi = "T00000" & a
    ElseIf a > 99 And a < 1000 Then
        generateIDTransaksi = "T0000" & a
    ElseIf a > 999 And a < 10000 Then
        generateIDTransaksi = "T000" & a
    ElseIf a > 9999 And a < 100000 Then
        generateIDTransaksi = "T00" & a
    ElseIf a > 99999 And a < 1000000 Then
        generateIDTransaksi = "T0" & a
    Else
        generateIDTransaksi = "T" & a
    End If
End Function
