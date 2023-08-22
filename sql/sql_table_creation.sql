use MyDB

-- market_tableの作成
CREATE TABLE market_table (
    date DATE PRIMARY KEY,
    price_date NVARCHAR(50),
    price_hhmm NVARCHAR(10),
    au_ny_end NVARCHAR(50),
    pt_ny_end NVARCHAR(50),
    ag_ny_end NVARCHAR(50),
    ny_exchange_rate NVARCHAR(50),
    ny_end INT,
    tokyo_exchange_rate NVARCHAR(50),
    tokyo_start INT,
    au_buy NVARCHAR(50),
    au_buy_diff INT,
    pt_buy NVARCHAR(50),
    pt_buy_diff INT,
    au_sell NVARCHAR(50),
    au_sell_diff INT,
    pt_sell NVARCHAR(50),
    pt_sell_diff INT
);

-- nj_buy_tableの作成
CREATE TABLE nj_buy_table (
    date DATE PRIMARY KEY,
    au INT,
    ag INT,
    pt INT,
    pd INT
);


-- nj_sell_tableの作成
CREATE TABLE nj_sell_table (
    date DATE PRIMARY KEY,
    au INT,
	ag INT,
	pt INT,
	pd INT
);
