--use master
--go
--create database db_laundry on (
--	name = 'Lang Laundry dat',
--	filename = 'C:\db_laundry.mdf',
--	size=10,
--	maxsize=50,
--	filegrowth =5
--) log on (
--	name = 'Lang Laundry log',
--	filename ='C:\db_laundry.ldf',
--	size=5mb,
--	maxsize=25mb,
--	filegrowth = 5mb
--) 
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
	id_paket varchar(4) constraint pk_paket primary key,
	--id_paket as 'PK' + RIGHT('00000' + cast(id as varchar(5)), 5) persisted primary key,
	nama varchar(20) not null,
	tarif int not null,
	satuan varchar(11) not null
)

--drop table pelanggan

create table pelanggan(
	id_pelanggan varchar(7) constraint pk_pelanggan primary key,
	--id_pelanggan as 'PL' + RIGHT('00000' + cast(id as varchar(5)), 5) persisted primary key,
	nama varchar(50) not null,
	telp varchar(14) not null,
	alamat varchar(100)
)

create table transaksi(
	id_transaksi varchar(8) constraint pk_transaksi primary key,
	--id_transaksi as 'TR' + RIGHT('00000' + cast(id as varchar(5)), 5) persisted primary key,
	id_pelanggan varchar(7) constraint fk_pelanggan foreign key references pelanggan(id_pelanggan) 
		on update cascade on delete no action,
	biaya int not null,
	tgl_terima datetime default getdate(),
	tgl_ambil datetime default (getdate()+3)
)

create table det_transaksi(
	id_transaksi varchar(8) constraint fk_transaksi foreign key references transaksi(id_transaksi)
		on update cascade on delete no action,
	id_paket varchar(4) constraint fk_paket foreign key references paket(id_paket)
		on update cascade on delete no action,
	jumlah int not null,
	total int not null
)

create table karyawan(
	nik varchar(7) constraint pk_karyawan primary key,
	--uid as 'UID' + RIGHT('00000' + cast(id as varchar(5)), 5) persisted primary key,
	nama varchar(30) not null,
	telp varchar(14) not null,
	alamat varchar(50),
	jabatan varchar(20)
)

create table login_karyawan(
	username varchar(7) constraint fk_karyawan foreign key references karyawan(nik)
		on update cascade on delete no action,
	password varchar(20) not null,
	akses varchar(5) not null
)