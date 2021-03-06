﻿use NPSF

drop Table GBCVisaDetailAbateDetail

Create Table GBCVisaDetailAbateDetail(
 基金代碼 varchar(3) not null,
 PK_會計年度 varchar(3) not null,
 PK_動支編號 varchar(20) not null,
 PK_種類 varchar(10) not null,
 PK_次別 varchar(10) not null,
 PK_明細號 varchar(20) not null,
 F_核定金額 money null,
 F_傳票年度 varchar(3),
 F_傳票號1 varchar(20) null,
 F_傳票明細號1 int null,
 F_製票日期1 char(10) null,
 F_傳票號2 varchar(20) null,
 F_傳票明細號2 int null,
 F_製票日期2 char(10) null,
 F_受款人 varchar(255) null,
 F_受款人編號 varchar(20) null

 CONSTRAINT GBCVisaDetailAbateDetail_PK PRIMARY KEY (基金代碼,PK_會計年度,PK_動支編號,PK_種類,PK_次別,PK_明細號,F_傳票年度)
)
