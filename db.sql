create database kart;

CREATE TABLE kart
(
 ID int primary key identity,
 NamePrint nvarchar(200),
 NameKart nvarchar(200),
 countkart int
)
DELETE FROM kart WHERE NamePrint = 'HP T520'
SELECT * FROM kart
INSERT INTO kart VALUES('МФУ монохромное Ricoh Africo SP 3610SF 906386 (407306)','Принт-картридж тип SP4500E (Black) черный',1,'МЦ№ 00-00004464','Склад')