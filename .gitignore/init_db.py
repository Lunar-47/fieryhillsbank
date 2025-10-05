import sqlite3
from werkzeug.security import generate_password_hash

# Connect to (or create) the database
conn = sqlite3.connect("hawkeyes.db")
c = conn.cursor()

# Drop existing tables if you want a fresh start
c.execute("DROP TABLE IF EXISTS transactions")
c.execute("DROP TABLE IF EXISTS students")

# Create Students table
c.execute("""
CREATE TABLE students (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    grade TEXT NOT NULL,
    password TEXT NOT NULL
)
""")

# Create Transactions table
c.execute("""
CREATE TABLE transactions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER NOT NULL,
    date TEXT NOT NULL,
    description TEXT,
    debit INTEGER DEFAULT 0,
    credit INTEGER DEFAULT 0,
    FOREIGN KEY(student_id) REFERENCES students(id)
)
""")

# Optionally, insert test students
test_students = [
    ("Amelia", "11", generate_password_hash("Amelia")),
    ("Elise", "11", generate_password_hash("Elise")),
    ("Riti", "11", generate_password_hash("Riti")),
    ("Cairo", "12", generate_password_hash("Cairo")),
    ("Harry", "12", generate_password_hash("Harry")),
    ("Ronald", "12", generate_password_hash("Ronald")),
    ("KGrader", "K", generate_password_hash("KGrader")),
    ("1grader", "1", generate_password_hash("1grader")),
    ("kthekgrader", "K", generate_password_hash("kthekgrader"))
]

c.executemany("INSERT INTO students (name, grade, password) VALUES (?, ?, ?)", test_students)

conn.commit()
conn.close()
print("Database initialized and test students added!")
