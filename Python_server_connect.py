import psycopg2
conn = psycopg2.connect(
    dbname="your_actual_db_name",
    user="your_actual_username",
    password="your_actual_password",
    host="localhost",
    port="5432"
)
print("Connected!")