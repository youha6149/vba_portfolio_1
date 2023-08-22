USE MyDB

DROP TABLE market_table;
DROP TABLE nj_buy_table;
DROP TABLE nj_sell_table;


-- market_table�̍쐬
CREATE TABLE market_table (
    date DATE,
    price_date NVARCHAR(50),
    price_hhmm NVARCHAR(10),
    au_ny_end DECIMAL(10,2),
    pt_ny_end DECIMAL(10,2),
    ag_ny_end DECIMAL(10,2),
    ny_exchange_rate DECIMAL(10,2),
    ny_end INT,
    tokyo_exchange_rate DECIMAL(10,2),
    tokyo_start INT,
    au_buy DECIMAL(10,2),
    au_buy_diff INT,
    pt_buy DECIMAL(10,2),
    pt_buy_diff INT,
    au_sell DECIMAL(10,2),
    au_sell_diff INT,
    pt_sell DECIMAL(10,2),
    pt_sell_diff INT
);


-- nj_buy_table�̍쐬
CREATE TABLE nj_buy_table (
    date DATE PRIMARY KEY,
    au DECIMAL(10,2),
    ag DECIMAL(10,2),
    pt DECIMAL(10,2),
    pd DECIMAL(10,2)
);


-- nj_sell_table�̍쐬
CREATE TABLE nj_sell_table (
    date DATE PRIMARY KEY,
    au DECIMAL(10,2),
	ag DECIMAL(10,2),
	pt DECIMAL(10,2),
	pd DECIMAL(10,2)
);
