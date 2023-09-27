import mysql.connector
import pandas as pd
from text import TextColor as tc

def connect_mysql(host, user, password, database):
    db = mysql.connector.connect(
        host=host,
        user=user,
        password=password,
        database=database
    )
    return db

def insert_participant(db, student_id, name, major, qualified):
    cursor = db.cursor()
    sql = "INSERT INTO participants (student_id, name, study_program, qualified, created_at, updated_at) VALUES (%s, %s, %s, %s, NOW(), NOW())"
    val = (student_id, name, major, qualified)
    cursor.execute(sql, val)
    db.commit()

HOST = "localhost"
USER = "root"
PASSWORD = ""
DATABASE = "pengumuman_fosti"
EXCEL_FILE = "data.xlsx"

if __name__ == "__main__":
    try:
        db = connect_mysql(
            host=HOST,
            user=USER,
            password=PASSWORD,
            database=DATABASE
        )
        print(f"{tc.GREEN}Success: Connected to database{tc.END}")
    except:
        print(f"{tc.RED}Error: Can't connect to database{tc.END}")
        exit(1)

    df = pd.read_excel(EXCEL_FILE)
    participants = tuple(df[['student_id', 'name', 'major', 'qualified']].values)
    for student_id, name, major, qualified in participants:
        try:
            student_id = student_id.upper()
            name = name.title()
            major = major.title()
            print(f"{tc.YELLOW}Action: Inserting {student_id} {name}{tc.END}")
            insert_participant(db, student_id, name, major, qualified)
            print(f"{tc.GREEN}Success: Inserting {student_id} {name}{tc.END}")
        except:
            print(f"{tc.RED}Error: Inserting {student_id} {name}{tc.END}")
            continue
        