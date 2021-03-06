﻿use NPSF

drop Table GBCJSONRecord

Create Table GBCJSONRecord(
 基金代碼 varchar(3) not null,
 PFK_會計年度 varchar(3) not null,
 PFK_動支編號 varchar(20) not null,
 PFK_種類 varchar(10) not null,
 PFK_次別 varchar(10) not null,
 傳票JSON1 nvarchar(MAX) null,
 傳票JSON2 nvarchar(MAX) null,
 是否結案  varchar(1) null

 CONSTRAINT GBCJSONRecord_PK PRIMARY KEY (基金代碼,PFK_會計年度,PFK_動支編號,PFK_種類,PFK_次別)
)