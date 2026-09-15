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
            BillNum text,
            ItemCodeBill text,
            NameBill text,
            QtyBill int,
            PriceBill int,
            AmountBill text,
            Seller    
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
        self.cur.execute("SELECT * from AddTotal")
        rows = self.cur.fetchall()
        # print(rows)
        return rows

    # Delete a Record in DB
    def remove(self):
        self.cur.execute("delete from AddTotal")
        self.con.commit()

    # ---------------- db for BillForm table ----------------

    def insertBill(self,Bill,ItemBill,NameBill,QtyBill,PriceBill,AmountBill,Seller):
        self.cur.execute("insert into AddTotalBillNum values (NULL,?,?,?,?,?,?,?)",(Bill,ItemBill,NameBill,QtyBill,PriceBill,AmountBill,Seller))
        self.con.commit()

    # dbTotal Warehouse control item decrease
    def DecItem(self,ItemCode):
        self.cur.execute("select count(*) from AddTotal where ItemCode =?",(ItemCode,))
        rows = self.cur.fetchone()
        return rows