import sqlite3

class DatabaseSales:
    def __init__(self, db):
        self.con = sqlite3.connect(db)
        self.cur = self.con.cursor()
        sql = """
        CREATE TABLE IF NOT EXISTS SalesReport(
            id Integer Primary Key,
            date date,
            sales int,
            billNumber,
            seller text,
            quantity int,
            itemcode
        )
        """
        self.cur.execute(sql)
        self.con.commit()

    # Insert Function
    def insert(self,date,sales,billNumber,seller,quantity,itemcode):
        self.cur.execute("insert into SalesReport values (NULL,?,?,?,?,?,?)",
                         (date,sales,billNumber,seller,quantity,itemcode))
        self.con.commit()

    # Fetch All Data from DB
    def Total(self,date):
        self.cur.execute("SELECT sum(sales) from SalesReport where date=?",(date,))
        rows = self.cur.fetchone()
        # print(rows)
        return rows

    # Delete a Record in DB
    def remove(self, id):
        self.cur.execute("delete from SalesReport where id=?", (id,))
        self.con.commit()
    
    def CountQuan(self,date):
        self.cur.execute("SELECT COUNT(sales) from SalesReport where date=?",(date,))
        rows = self.cur.fetchone()
        # print(rows)
        return rows

    # ----------------- Monthly report ---------------
    #SELECT YEAR(date) AS Year,MONTH(date) AS Month,SUM(sales) AS Total_Sales FROM SalesReport GROUP BY YEAR(date), MONTH(date) ;
    def Monthly(self,month):
        MONTH_DICT={ "Jan" : '01', "Feb" : '02', "Mar" : '03', "Apr" : '04', "May" : '05', "Jun" : '06', "Jul" : '07', "Aug" : '08', "Sep" : '09', "Oct" : 10, "Nov" : 11, "Dec" : 12 }

        self.cur.execute("SELECT sum(sales) FROM SalesReport WHERE strftime('%m',date)=?",(month,)) 
        rows = self.cur.fetchone()
        # print(rows)
        return rows
    
    def MonthlyQuan(self,month):
       
        self.cur.execute("SELECT sum(quantity) FROM SalesReport WHERE strftime('%m',date)=?",(month,)) 
        rows = self.cur.fetchone()
        # print(rows)
        return rows

    def BestSellerMonth(self,month):
        self.cur.execute("select itemcode from SalesReport WHERE strftime('%m',date)=? order by itemcode desc limit 1 ",(month,)) 
        rows = self.cur.fetchone()
        # print(rows)
        return rows
    
    def BestSellerAll(self,year):
        self.cur.execute("select itemcode from SalesReport WHERE strftime('%y',date)=? order by itemcode desc limit 1 ",(year,)) 
        rows = self.cur.fetchone()
        # print(rows)
        return rows