import sqlite3

class warehouse:
    def __init__(self, db):
        self.con = sqlite3.connect(db)
        self.cur = self.con.cursor()
        sql = """
        CREATE TABLE IF NOT EXISTS Order2(
            id Integer Primary Key,
            Item text,
            Name text,
            Category text,
            Quantity int,
            Cost text,
            Price int,
            Exp_Date text,
            Alert_Date datetime
        )
        """
        self.cur.execute(sql)
        self.con.commit()

    # Insert Function
    def insert(self, Item, Name,Category, Quantity,Cost,Price,Exp_Date,Alert_Date):
        self.cur.execute("insert into Order2 values (NULL,?,?,?,?,?,?,?,?)",
                         ( Item, Name,Category, Quantity,Cost,Price,Exp_Date,Alert_Date))
        
        self.con.commit()

    # Fetch All Data from DB
    def fetch(self):
        self.cur.execute("SELECT * from Order2")
        rows = self.cur.fetchall()
        # print(rows)
        return rows

    # Delete a Record in DB
    def remove(self, id):
        self.cur.execute("delete from Order2 where id=?", (id,))
        self.con.commit()

    # Update a Record in DB
    def update(self, id, Item, Name,Category, Quantity,Cost,Price,Exp_Date,Alert_Date):
        self.cur.execute(
            "update Order2 set Item=?, Name=?,Category=?,Quantity=?,Cost=?,Price=?,Exp_Date=?,Alert_Date=? where id=?",
            ( Item, Name,Category, Quantity,Cost,Price,Exp_Date,Alert_Date,id))
        self.con.commit()
    
    # Update Quantity in Db
    def updateQuantity(self, Item,Quantity):
        self.cur.execute(
            "update Order2 set Quantity=? where Item=?",
            (Quantity,Item))
        self.con.commit()
    
    def updateName(self, Item,Name):
        self.cur.execute(
            "update Order2 set Name=? where Item=?",
            (Name,Item))
        self.con.commit()

    # get items code in DB
    def getItemsCodePrice(self,Item):
        self.cur.execute(
            "select Price from  Order2 where Item=?", (Item,)  
        )
        rows = self.cur.fetchone()
        return rows
    
    def Search(self,Item):
        self.cur.execute("SELECT * from Order2 where Item=?",(Item,))
        rows = self.cur.fetchall()
        # print(rows)
        return rows
    # get items code to his Name in DB
    def getQuntity(self,Item):
        self.cur.execute(
            "select Quantity from  Order2 where Item=?", (Item,)  
        )
        rows = self.cur.fetchone()
        return rows
    
    def Total(self,date):
        
        self.cur.execute("SELECT * from Order2 where Alert_Date=?",(date,))
        rows = self.cur.fetchall()
        # print(rows)
        return rows
    
    # ---------------------- Count for all -------------
    def TotalQuantity(self):
        self.cur.execute("SELECT sum(Quantity) from Order2")
        rows = self.cur.fetchone()
        # print(rows)
        return rows
    
    def TotalPrice(self):
        self.cur.execute("SELECT sum(Price) from Order2")
        rows = self.cur.fetchone()
        # print(rows)
        return rows