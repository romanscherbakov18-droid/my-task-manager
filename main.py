from flask import Flask, render_template, request, redirect, url_for, flash, make_response
import sqlite3
import pandas as pd
import plotly.express as px
from io import BytesIO
# --- Импорты для Excel ---
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

app = Flask(__name__, static_url_path='/static')
app.secret_key = 'ваш_секретный_ключ_здесь' # Замените на сложный ключ

def create_db():
    conn = sqlite3.connect('tasks.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            description TEXT,
            assignee TEXT,
            repeat_type TEXT DEFAULT '',
            deadline TEXT DEFAULT '',
            status TEXT DEFAULT 'New'
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS archived_tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            description TEXT,
            assignee TEXT,
            repeat_type TEXT DEFAULT '',
            deadline TEXT DEFAULT '',
            status TEXT DEFAULT 'Archived'
        )
    ''')
    conn.commit()
    conn.close()

create_db()

@app.route('/', methods=['GET'])
def index():
    selected_status = request.args.get('status', '')
    sort_order = request.args.get('sort', '')
    
    sql_query = "SELECT * FROM tasks"
    params = []
    
    if selected_status:
        sql_query += " WHERE status=?"
        params.append(selected_status)
    
    if sort_order == 'asc':
        sql_query += " ORDER BY title ASC"
    elif sort_order == 'desc':
        sql_query += " ORDER BY title DESC"
    
    conn = sqlite3.connect('tasks.db')
    df = pd.read_sql_query(sql_query, conn, params=params)
    tasks = df.values.tolist()
    
    total_tasks = len(df)
    completed_tasks = len(df.query("status == 'Completed'"))
    ongoing_tasks = total_tasks - completed_tasks
    
    if total_tasks > 0:
        fig = px.pie(
            values=[completed_tasks, ongoing_tasks],
            names=['Завершено', 'В процессе'],
            title='Соотношение завершенных и незавершенных задач',
            color_discrete_sequence=px.colors.qualitative.Pastel,
        )
        fig.update_layout(
            template="plotly_dark",
            font_color='white',
            title_x=0.5
        )
        pie_chart_html = fig.to_html(full_html=False)
    else:
        pie_chart_html = ""
    
    conn.close()
    
    return render_template('index.html', tasks=tasks, selected_status=selected_status, 
                           sort_order=sort_order, total_tasks=total_tasks, 
                           completed_tasks=completed_tasks, ongoing_tasks=ongoing_tasks, 
                           pie_chart=pie_chart_html)

@app.route('/add_task', methods=['POST'])
def add_task():
    title = request.form['title']
    description = request.form.get('description', '')
    assignee = request.form.get('assignee', '')
    repeat_type = request.form.get('repeat_type', '')
    deadline = request.form.get('deadline', '')

    conn = sqlite3.connect('tasks.db')
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO tasks (title, description, assignee, repeat_type, deadline) VALUES (?, ?, ?, ?, ?)",
        (title, description, assignee, repeat_type, deadline)
    )
    conn.commit()
    conn.close()
    
    return redirect(url_for('index'))

@app.route('/update_status/<int:id>', methods=['GET', 'POST'])
def update_status(id):
    if request.method == 'POST':
        new_status = request.form['new_status']
        
        conn = sqlite3.connect('tasks.db')
        cursor = conn.cursor()
        cursor.execute("UPDATE tasks SET status=? WHERE id=?", (new_status, id))
        conn.commit()
        
        if new_status == 'Completed':
            flash('Задача успешно завершена!', 'success')
            
        conn.close()
        return redirect(url_for('index'))
    else:
        conn = sqlite3.connect('tasks.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM tasks WHERE id=?", (id,))
        task = cursor.fetchone()
        conn.close()
        
        if task is None:
            return f"Задача с ID {id} не найдена.", 404
            
        return render_template('update_status.html', task=task)

# --- ИСПРАВЛЕННЫЙ ЭКСПОРТ В XLSX ---
@app.route('/export')
def export_csv():
    conn = sqlite3.connect('tasks.db')
    df = pd.read_sql_query("SELECT * FROM tasks", conn)
    conn.close()

    # Создаем объект Excel в памяти
    wb = Workbook()
    ws = wb.active
    ws.title = "Задачи"

    # Добавляем данные из DataFrame в лист Excel
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # --- ИСПРАВЛЕННАЯ ЛОГИКА СОХРАНЕНИЯ ---
    # Используем BytesIO для сохранения файла в памяти
    virtual_io = BytesIO()
    wb.save(virtual_io)
    
    # Получаем байты из потока
    virtual_io.seek(0)
    output_data = virtual_io.read()

    # Формируем ответ для скачивания .xlsx файла
    output = make_response(output_data)
    output.headers["Content-Disposition"] = "attachment; filename=tasks.xlsx"
    output.headers["Content-type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    
    return output

# Остальные функции (edit_task, delete_task, archive_task) остаются без изменений.

@app.route('/edit_task/<int:id>', methods=['GET', 'POST'])
def edit_task(id):
    if request.method == 'POST':
        title = request.form['title']
        description = request.form.get('description', '')
        assignee = request.form.get('assignee', '')
        repeat_type = request.form.get('repeat_type', '')
        deadline = request.form.get('deadline', '')

        conn = sqlite3.connect('tasks.db')
        cursor = conn.cursor()
        
        cursor.execute(
            "UPDATE tasks SET title=?, description=?, assignee=?, repeat_type=?, deadline=? WHERE id=?",
            (title, description, assignee, repeat_type, deadline, id)
        )
        
        conn.commit()
        conn.close()
        
        return redirect(url_for('index'))
        
    else:
        conn = sqlite3.connect('tasks.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM tasks WHERE id=?", (id,))
        task = cursor.fetchone()
        conn.close()
        
        if task is None:
            return f"Задача с ID {id} не найдена.", 404
            
        return render_template('edit_task.html', task=task)


@app.route('/delete_task/<int:id>', methods=['POST'])
def delete_task(id):
    conn = sqlite3.connect('tasks.db')
    cursor = conn.cursor()
    
    cursor.execute("DELETE FROM tasks WHERE id=?", (id,))
    
    conn.commit()
    conn.close()
    
    return redirect(url_for('index'))


@app.route('/archive_task/<int:id>', methods=['POST'])
def archive_task(id):
    conn = sqlite3.connect('tasks.db')
    cursor = conn.cursor()
    
    cursor.execute("SELECT * FROM tasks WHERE id=?", (id,))
    task = cursor.fetchone()
    
    if task:
         cursor.execute(
            "INSERT INTO archived_tasks (title, description, assignee, repeat_type, deadline, status) VALUES (?, ?, ?, ?, ?, ?)",
            (task[1], task[2], task[3], task[4], task[5], 'Archived') 
         )
         cursor.execute("DELETE FROM tasks WHERE id=?", (id,))
         conn.commit()
         flash('Задача успешно архивирована.', 'info')
    else:
         flash('Задача не найдена.', 'danger')
    
    conn.close()
    return redirect(url_for('index'))


if __name__ == '__main__':
   app.run(host='0.0.0.0', port=8080, debug=True)
