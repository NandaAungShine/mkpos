import sqlite3

class DatabaseTotal:
    def __init__(self, db):
        self.con = sqlite3.connect(db)
        self.cur = self.con.cursor()
        sql = """
        CREATE TABLE IF NOT EXISTS AddTotal(
            id Integer Primary Key,
            ItemCode text,
            Name text,
            Amount text   
        )
        """
        
        self.cur.execute(sql)
        self.con.commit()

        sql1 = """
        CREATE TABLE IF NOT EXISTS AddTotalBillNum(
            id Integer Primary Key,
            date date,
            BillNum text,
            ItemCodeBill ,
            NameBill text,
            QtyBill int,
            PriceBill int,
            AmountBill text,
            Seller text,
            CashBill text    
        )
        """
        
        self.cur.execute(sql1)
        self.con.commit()
    
    # Insert Function
    def insert(self,ItemCode,Name,Amount):
        self.cur.execute("insert into AddTotal values (NULL,?,?,?)",(ItemCode,Name,Amount))
        self.con.commit()

    
    
    # Fetch All Data from DB
    def fetch(self):
        self.cur.execute("SELECT sum(Amount) from AddTotal")
        rows = self.cur.fetchone()
        # print(rows)
        return rows

    def fetchAll(self):
        self.cur.execute("SELECT * from AddTotalBillNum order by date desc")
        rows = self.cur.fetchall()
        # print(rows)
        return rows
    
    # Delete a Record in DB
    def remove(self):
        self.cur.execute("delete from AddTotal")
        self.con.commit()

    # ---------------- db for BillForm table ----------------

    def insertBill(self,Date,Bill,ItemBill,NameBill,QtyBill,PriceBill,AmountBill,Seller,CashBill):
        self.cur.execute("insert into AddTotalBillNum values (NULL,?,?,?,?,?,?,?,?,?)",(Date,Bill,ItemBill,NameBill,QtyBill,PriceBill,AmountBill,Seller,CashBill))
        self.con.commit()

    # dbTotal Warehouse control item decrease
    def DecItem(self,ItemCode):
        self.cur.execute("select count(*) from AddTotal where ItemCode =?",(ItemCode,))
        rows = self.cur.fetchone()
        return rows
    
    def BestSellerMonth(self,month):
        self.cur.execute("select NameBill from AddTotalBillNum WHERE strftime('%m',date)=? order by ItemCodeBill desc limit 1 ",(month,)) 
        rows = self.cur.fetchone()
        # print(rows)
        return rows
    
    def BestSellerAll(self,year):
        self.cur.execute("select NameBill from AddTotalBillNum WHERE strftime('%y',date)=? order by ItemCodeBill desc limit 1 ",(year,)) 
        rows = self.cur.fetchone()
        # print(rows)
        return rows