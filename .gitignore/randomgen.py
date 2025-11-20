import random
import sqlite3

def get_db_connection():
    conn = sqlite3.connect("hawkeyes.db")
    conn.row_factory = sqlite3.Row
    return conn

def main():
    conn = get_db_connection()
    students = conn.execute("SELECT * FROM students").fetchall()
    
    for student in students:
        if student["code"] == 0:
            code = [str(random.randint(0, 9)) for _ in range(12)]
            code = [''.join(code[0:12])]
            code = int((code[0]))
            conn.execute("UPDATE students SET code = ? WHERE id = ?", (code, student["id"]))
            print("I'm doing something!")
            print(code)

    conn.commit()
    conn.close()

if __name__ == '__main__':
    main()


