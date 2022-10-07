from flask import Flask, render_template, request, flash, redirect, url_for, jsonify
from werkzeug.utils import secure_filename
from data import db_session
from data.users import User
from data.projects import Projects
from DataBase import DData
import os
from gevent.pywsgi import WSGIServer

app = Flask(__name__, static_folder=os.path.abspath(os.getcwd()) + '\\static', template_folder=os.path.abspath(os.getcwd()) + '\\templates')
app.config['UPLOAD_FOLDER'] = os.getcwd()
app.config['SECRET_KEY'] = '********'

CR = 0

def all_users_lenght():
    import sqlite3
    con = sqlite3.connect('db/users.db')
    cur = con.cursor()
    num = 0 
    try:
        num = len([job[0] for job in cur.execute("SELECT id FROM users")])
    except:
        num = 0
    return num


def save_into_db(args):
    user = User()
    user.name = args[0]
    user.surname = args[1]
    user.father = args[2]
    user.tablenum = args[3]
    user.email = args[4]
    user.mobile = args[5]
    user.obr = args[6]
    user.prof = args[7]
    user.company = args[8]
    db_sess = db_session.create_session()
    db_sess.add(user)
    db_sess.commit()


def save_file_into_user_file(files, params, flag):
    path = DData().create_dir(params, flag=flag)
    app.config['UPLOAD_FOLDER'] += '/' + path
    for file in files:
        if file:
            filename = file.filename
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
    app.config['UPLOAD_FOLDER'] = os.getcwd()


@app.route('/form', methods=['POST', 'GET'])
def index():
    error = None
    if request.method == 'GET':
        return render_template('form.html', error=error)
    elif request.method == 'POST':
        name = request.form['name']
        surname = request.form['surname']
        father = request.form['father']
        tablenum = request.form['tablenum']
        mail = request.form['email']
        mobile = request.form['mobile']
        obr = request.form['obr']
        prof = request.form['prof']
        company = request.form['company']
        files = request.files.getlist('file')
        user_id = str(all_users_lenght() + 1)

        params_for_db = [name, surname, father, tablenum, mail, mobile, obr, prof, company]
        params_for_file_path = [name, surname, father, user_id]

        save_file_into_user_file(files, params_for_file_path, 1)
        save_into_db(params_for_db)

        pr_ids = request.form.getlist('r_id')
        pr_urls = request.form.getlist('pr_url')
        project_name = request.form.getlist('pr_name')
        ur_rol_in_project = request.form.getlist('pr_rol')
        about_project = request.form.getlist('pr_content')
        ur_attaches_in_project = request.form.getlist('pr_merit')

        if all([project_name, ur_rol_in_project, about_project, ur_attaches_in_project]):
            print(user_id)

            for i in range(len(project_name)):
                args = [pr_ids[i], pr_urls[i], project_name[i], ur_rol_in_project[i], about_project[i], ur_attaches_in_project[i]]
                DData().insert_into_project(args, user_id)
        return render_template('form_success.html')


if __name__ == '__main__':
    db_session.global_init("db/users.db")
    app.run(port=8080, host='0.0.0.0', debug=True)
    http_server = WSGIServer(('', 5000), app)
    http_server.serve_forever()
