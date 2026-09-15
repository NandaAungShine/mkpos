import sqlite3

class DatabaseOrder:
    def __init__(self, db):
        self.con = sqlite3.connect(db)
        self.cur = self.con.cursor()
        sql = """
        CREATE TABLE IF NOT EXISTS Order1(
            id Integer Primary Key,
            Item text,
            Name text,
            Quantity text,
            Order_Date text,
            Coming_Date text
        )
        """
        self.cur.execute(sql)
        self.con.commit()

    # Insert Function
    def insert(self, Item, Name, Quantity,Order_Date,Coming_Date):
        self.cur.execute("insert into Order1 values (NULL,?,?,?,?,?)",
                         (Item, Name,Quantity,Order_Date,Coming_Date))
        self.con.commit()

    # Fetch All Data from DB
    def fetch(self):
        self.cur.execute("SELECT * from Order1")
        rows = self.cur.fetchall()
        # print(rows)
        return rows

    # Delete a Record in DB
    def remove(self, id):
        self.cur.execute("delete from Order1 where id=?", (id,))
        self.con.commit()

    # Update a Record in DB
    def update(self, id,Item, Name,Quantity,Order_date,Coming_Date):
        self.cur.execute(
            "update Order1 set Item=?, Name=?,Quantity=?,Order_Date=?,Coming_Date=? where id=?",
            (Item, Name,Quantity,Order_date,Coming_Date,id))
        self.con.commit()
    
    # get items code in DB
    def getItemsCodePrice(self,Item):
        self.cur.execute(
            "select Quantity from  Order1 where Item=?", (Item,)  
        )
        rows = self.cur.fetchone()
        return rows

    # get items code to his Name in DB
    def getItemsCodeName(self,Item):
        self.cur.execute(
            "select Name from  Order1 where Item=?", (Item,)  
        )
        rows = self.cur.fetchone()
        return rows