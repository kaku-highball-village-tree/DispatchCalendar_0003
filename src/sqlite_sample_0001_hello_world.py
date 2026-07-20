import sqlite3

connection = sqlite3.connect("hello.db")
cursor = connection.cursor()

cursor.execute("CREATE TABLE IF NOT EXISTS messages (text TEXT)")
cursor.execute("INSERT INTO messages (text) VALUES (?)", ("Hello, world!!",))

connection.commit()

cursor.execute("SELECT text FROM messages")

for row in cursor.fetchall():
    print(row[0])

connection.close()
