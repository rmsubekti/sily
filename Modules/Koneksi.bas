Attribute VB_Name = "Koneksi"
Public conn As New Connection
Public rs As New Recordset
Public konek As String
Public sql As String
Public needSave As Boolean

Public Sub getKoneksi()
    On Error GoTo koneksiError
    konek = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=db_laundry;Data Source=BEKTI\SQLEXPRESS"
    If conn.State = adStateOpen Then
        conn.Close
        Set conn = New Connection
        conn.Open konek
    Else
        conn.Open konek
    End If
Exit Sub
koneksiError:
    MsgBox "Koneksi ke server gagal!", vbCritical, "Kesalahan Koneksi"
End Sub
