

--drop table karyawan

begin transaction
	declare @id int;
	insert into karyawan values('K000001','Rahmat Subekti','404','Not Found','Software Enginer');
	select @id = scope_identity();
	insert into login_karyawan values (@id, 'secret', 'ADMIN')
commit

begin transaction
	insert into pelanggan(nama, telp,alamat) values('Rahmat Subekti','404','Not Found');
	select @@identity as id;
commit

select * from karyawan
select * from ulogin

--insert into karyawan(nama, no_telpon,jabatan)
--OUTPUT INSERTED.'link','121','pe'
--INTO login (password,akses)
--VALUES  ('secret','kasir');
select cast(getdate() + 3 as datetime)
select * from pelanggan
select max(right(id_pelanggan,6)) from pelanggan
select * from transaksi inner join pelanggan ON transaksi.id_pelanggan=pelanggan.id_pelanggan

delete from pelanggan where id_pelanggan='P000001'

select nik as NIK ,nama as Nama,alamat as Alamat,telp as Telepon,jabatan as Jabatan,password as Password,akses as Akses 
        from karyawan inner join login_karyawan on karyawan.nik = login_karyawan.username
        where nama like '%be%'