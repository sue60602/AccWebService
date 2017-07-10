use NPSF
drop Table GBCJSONRecordLog
Create Table GBCJSONRecordLog(
 基金代碼 varchar(3) not null,
 條碼 varchar(20) not null,
 JSON紀錄 nvarchar(max) not null,
 接收時間 DateTime not null
)