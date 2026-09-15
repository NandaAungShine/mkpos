import sqlite3

class Database:
    def __init__(self, db):
        self.con = sqlite3.connect(db)
        self.cur = self.con.cursor()
        sql = """
        CREATE TABLE IF NOT EXISTS AddItems1(
            id Integer Primary Key,
            Item text,
            Name text,
            Price text
            
        )
        """
        self.cur.execute(sql)
        self.con.commit()

    # Insert Function
    def insert(self,Item,Name,Price):
        self.cur.execute("insert into AddItems1 values (NULL,?,?,?)",
                         (Item,Name,Price))
        self.con.commit()

    # Fetch All Data from DB
    def fetch(self):
        self.cur.execute("SELECT * from AddItems1")
        rows = self.cur.fetchall()
        # print(rows)
        return rows
    # search 
    def Search(self,Item):
        self.cur.execute("SELECT * from AddItems1 where Item=?",(Item,))
        rows = self.cur.fetchall()
        # print(rows)
        return rows

    # Delete a Record in DB
    def remove(self, id):
        self.cur.execute("delete from AddItems1 where id=?", (id,))
        self.con.commit()

    # Update a Record in DB
    def update(self, id,Item, Name,Price):
        self.cur.execute(
            "update AddItems1 set Item=?, Name=?,Price=? where id=?",
            (Item, Name,Price, id))
        self.con.commit()
    
    # get items code in DB
    def getItemsCodePrice(self,Item):
        self.cur.execute(
            "select Price from  AddItems1 where Item=?", (Item,)  
        )
        rows = self.cur.fetchone()
        return rows

    # get items code to his Name in DB
    def getItemsCodeName(self,Item):
        self.cur.execute(
            "select Name from  AddItems1 where Item=?", (Item,)  
        )
        rows = self.cur.fetchone()
        return rows