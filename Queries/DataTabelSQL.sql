use db_laundry
go
--Dummy Paket
insert into paket values('P001','Setrika','3000','Kg')
insert into paket values('P002','Cuci Kering','6000','Kg')
insert into paket values('P003','Cuci Wangi','7000','Kg')
insert into paket values('P004','Baju Hem','2000','Buah')
insert into paket values('P005','Celana Panjang','2000','Buah')
insert into paket values('P006','Kaos','1500','Buah')
insert into paket values('P007','Jaket','4000','Buah')
insert into paket values('P008','Jemper','3000','Buah')
insert into paket values('P009','Celana Pendek','1500','Buah')
insert into paket values('P010','Jeans Pendek','3500','Buah')

--Dummy Pelanggan
insert into pelanggan values('P000001','Anca','301','Rt.1 Rw.1')
insert into pelanggan values('P000002','Bekti','200','Rt.1 Rw.2')
insert into pelanggan values('P000003','Candil','401','Rt.2 Rw.3')
insert into pelanggan values('P000004','Dona','302','Rt.1 Rw.4')
insert into pelanggan values('P000005','Eka','203','Rt.2 Rw.5')

--Dummy Karyawan
insert into karyawan values('K000001','Liu Kang','302','Not Modified','Kasir')
insert into login_karyawan values('K000001','secret','KASIR')
insert into karyawan values('K000002','Kratos','404','Not Found','Pekerja Tukang')
insert into login_karyawan values('K000002','secret','ADMIN')
