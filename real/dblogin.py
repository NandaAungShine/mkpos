import sqlite3

class Registerdata:
    def __init__(self, db):
        self.con = sqlite3.connect(db)
        self.cur = self.con.cursor()
        sql = """
        CREATE TABLE IF NOT EXISTS Admindata(
            id Integer Primary Key,
            Username text,
            Password text,
            Confirm text
            
        )
        """
        self.cur.execute(sql)
        self.con.commit()

        sql1 = """
        CREATE TABLE IF NOT EXISTS Userdata(
            id Integer Primary Key,
            Username text,
            Password text,
            Confirm text
            
        )
        """
        self.cur.execute(sql1)
        self.con.commit()
     
     

    # Insert Function
    def insert(self,Username,Password,Confirm):
        self.cur.execute("insert into Admindata values (NULL,?,?,?)",
                         (Username,Password,Confirm))
        self.con.commit()

    # Fetch All Data from DB
    def fetch(self):
        self.cur.execute("SELECT * from Admindata")
        rows = self.cur.fetchall()
        # print(rows)
        return rows
    # search 
    def Search(self,Username):
        self.cur.execute("SELECT * from Admindata where Username=?",(Username,))
        rows = self.cur.fetchall()
        # print(rows)
        return rows
    
    # ============ User Data ============
    def insertUser(self,Username,Password,Confirm):
        self.cur.execute("insert into Userdata values (NULL,?,?,?)",
                         (Username,Password,Confirm))
        self.con.commit()

    # Fetch All Data from DB
    def fetchUser(self):
        self.cur.execute("SELECT * from Userdata")
        rows = self.cur.fetchall()
        # print(rows)
        return rows