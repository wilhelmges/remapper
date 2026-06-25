-- fetch not filled on sheet
CREATE TABLE "wasteaccountbook" (
	"id"	INTEGER,
	"date"	TEXT NOT NULL,
	"ordernum"	INTEGER,
	"title"	TEXT,
	"totalsum"	NUMERIC,
	"eventyear"	INTEGER,
	"status"	TEXT,
	PRIMARY KEY("id" AUTOINCREMENT)
)


SELECT * FROM wasted WHERE root_order_id IS NULL AND NOT filled AND sheet_name="БпЛА"