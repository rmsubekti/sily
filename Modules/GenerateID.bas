Attribute VB_Name = "GenerateID"
Public Function generateIDPelanggan() As String
    Dim a As Integer
    sql = "select max(right(id_pelanggan,6)) from pelanggan"
    Set rs = conn.Execute(sql)
    a = IIf(rs(0) <> "NULL", rs(0) + 1, 1)
    If Val(a) < 10 Then
        generateIDPelanggan = "P00000" & a
    ElseIf Val(a) > 10 And Val(a) < 100 Then
        generateIDPelanggan = "P0000" & a
    ElseIf Val(a) > 100 And Val(a) < 1000 Then
        generateIDPelanggan = "P000" & a
    ElseIf Val(a) > 1000 And Val(a) < 10000 Then
        generateIDPelanggan = "P00" & a
    ElseIf Val(a) > 10000 And Val(a) < 100000 Then
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
    If Val(a) < 10 Then
        generateIDPaket = "P00" & a
    ElseIf Val(a) > 10 And Val(a) < 100 Then
        generateIDPaket = "P0" & a
    Else
        generateIDPaket = "P" & a
    End If
End Function
Public Function generateIDKaryawan() As String
    Dim a As Integer
    sql = "select max(right(nik,6)) from karyawan"
    Set rs = conn.Execute(sql)
    a = IIf(rs(0) <> "NULL", rs(0) + 1, 1)
    If Val(a) < 10 Then
        generateIDKaryawan = "K00000" & a
    ElseIf Val(a) > 10 And Val(a) < 100 Then
        generateIDKaryawan = "K0000" & a
    ElseIf Val(a) > 100 And Val(a) < 1000 Then
        generateIDKaryawan = "K000" & a
    ElseIf Val(a) > 1000 And Val(a) < 10000 Then
        generateIDKaryawan = "K00" & a
    ElseIf Val(a) > 10000 And Val(a) < 100000 Then
        generateIDKaryawan = "K0" & a
    Else
        generateIDKaryawan = "K" & a
    End If
End Function
Public Function generateIDTransaksi() As String
    Dim a As Integer
    sql = "select max(right(id_transaksi,6)) from transaksi"
    Set rs = conn.Execute(sql)
    a = IIf(rs(0) <> "NULL", rs(0) + 1, 1)
    If Val(a) < 10 Then
        generateIDTransaksi = "T00000" & a
    ElseIf Val(a) > 10 And Val(a) < 100 Then
        generateIDTransaksi = "T0000" & a
    ElseIf Val(a) > 100 And Val(a) < 1000 Then
        generateIDTransaksi = "T000" & a
    ElseIf Val(a) > 1000 And Val(a) < 10000 Then
        generateIDTransaksi = "T00" & a
    ElseIf Val(a) > 10000 And Val(a) < 100000 Then
        generateIDTransaksi = "T0" & a
    Else
        generateIDTransaksi = "T" & a
    End If
End Function
