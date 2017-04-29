use master
go
create database db_laundry on (
	name = 'Lang Laundry dat',
	filename = 'C:\db_laundry.mdf',
	size=10,
	maxsize=50,
	filegrowth =5
) log on (
	name = 'Lang Laundry log',
	filename ='C:\db_laundry.ldf',
	size=5mb,
	maxsize=25mb,
	filegrowth = 5mb
) 
-- for attach;
-- Above query generate this error message : 
-- CREATE FILE encountered operating system error 5(Access is denied.)
-- while attempting to open or create the physical file 'D:\db_laundry.mdf'.
-- and current windows doesn't supported to install mssql 2000 server components 


create database db_laundry
go
use db_laundry
go
create table paket(
	id_paket varchar(4) not null primary key,
	--id_paket as 'PK' + RIGHT('00000' + cast(id as varchar(5)), 5) persisted primary key,
	nama varchar(20) not null,
	tarif int not null,
	satuan varchar(11) not null
)

--drop table pelanggan

create table pelanggan(
	id_pelanggan varchar(7) not null primary key,
	--id_pelanggan as 'PL' + RIGHT('00000' + cast(id as varchar(5)), 5) persisted primary key,
	nama varchar(50) not null,
	telp varchar(14) not null,
	alamat varchar(100)
)

create table transaksi(
	id_transaksi varchar(8) not null primary key,
	--id_transaksi as 'TR' + RIGHT('00000' + cast(id as varchar(5)), 5) persisted primary key,
	id_pelanggan varchar(7) foreign key references pelanggan(id_pelanggan),
	biaya int not null,
	tgl_terima datetime not null,
	tgl_ambil datetime not null
)

create table det_transaksi(
	id_transaksi varchar(8) foreign key references transaksi(id_transaksi),
	id_paket varchar(4) foreign key references paket(id_paket),
	jumlah int not null,
	total int not null
)

create table karyawan(
	nik varchar(7) not null primary key,
	--uid as 'UID' + RIGHT('00000' + cast(id as varchar(5)), 5) persisted primary key,
	nama varchar(30) not null,
	alamat varchar(50) not null,
	telp varchar(14) not null,
	jabatan varchar(20) not null
)

create table ulogin(
	nik varchar(7) foreign key references karyawan(nik),
	password varchar(20) not null,
	akses varchar(5) not null
)

--drop table karyawan

begin transaction
	declare @id int;
	insert into karyawan(nama, telp,jabatan) values('Rahmat Subekti','404','Software Enginer');
	select @id = scope_identity();
	insert into ulogin(uid,password,akses) values (@id, 'secret', 'ADMIN')
commit
begin transaction
	declare @id int;
	insert into karyawan(nama, telp,jabatan) values('Riski','300','Documentation');
	select @id = scope_identity();
	insert into ulogin(uid,password,akses) values (@id, 'secret', 'KASIR')
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