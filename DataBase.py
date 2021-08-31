# -*- coding: utf-8 -*-

from data.users import User
from data import db_session
import sqlite3 as sql
import os
import xlsxwriter


class DData():

    def __init__(self):
        # self.sess = db_session.create_session()
        self.sql = sql.connect("db/users.db")

    def find_data(self, name):
        p = self.sess.query(User).filter(User.name == name)
        self.sess.commit()
        return p

    def needed_column(self, text):
        sp = ['']
        cursor = self.sql.cursor()
        SQL = f"SELECT DISTINCT {text} FROM users"
        result = cursor.execute(SQL).fetchall()
        for i in result:
            sp.append(i[0])
        cursor.close()
        return sp

    def create_dir(self, arg, flag=0):
        eng_word = '_'.join(list(arg))
        if flag:
            os.makedirs(f"users_data/{eng_word}/files")
        return f"users_data/{eng_word}/files"

    # def check_init_dir(self):
    #     previous_path = self.create_dir(params):

    def write_in_file(self, name, data):
        workbook = xlsxwriter.Workbook('teets.xlsx')
        worksheet = workbook.add_worksheet()

        headers = ['Project_id', 'Project_Name', 'Project_Role', 'Project_Content', 'Project_Attachments',
                   'for_User_id', 'User_id', 'Name', 'Surname', 'Father',
                   'Table_number', 'Email', 'Mobile_number', 'Education', 'Profession', 'Company', 'File_Path']

        row, col = 0, 0

        for i in headers:
            worksheet.set_column(row, col, 20)
            worksheet.write(row, col, i)
            col += 1

        for row, (id_pr, pr_name, pr_rol, pr_cont, pr_att, user_id, real_user_id, name, surname, father, tablenum, email, mobile, obr, prof, company) in enumerate(l):
            path = self.dirname((name, surname, father))
            worksheet.write(row, 0, str(id_pr))
            worksheet.write(row, 1, pr_name)
            worksheet.write(row, 2, pr_rol)
            worksheet.write(row, 3, pr_cont)
            worksheet.write(row, 4, pr_att)
            worksheet.write(row, 5, str(user_id))
            worksheet.write(row, 6, str(real_user_id))
            worksheet.write(row, 7, name)
            worksheet.write(row, 8, surname)
            worksheet.write(row, 9, father)
            worksheet.write(row, 10, str(tablenum))
            worksheet.write(row, 11, email)
            worksheet.write(row, 12, mobile)
            worksheet.write(row, 13, obr)
            worksheet.write(row, 14, prof)
            worksheet.write(row, 15, company)

            worksheet.write_url(row, 16, f'external:{path}')
        workbook.close()

    def dirname(self, arg):
        part_name = list(map(str, arg))
        dir_name = '_'.join(part_name)
        path = os.path.abspath(os.getcwd()) + \
            '\\users_data\\' + dir_name + '\\files'
        return path

    def lenght_projects(self):
        con = sql.connect('db/projects.db')
        cur = con.cursor()
        ids = 1
        try:
            ids = len([job[0]
                       for job in cur.execute("SELECT id FROM projects")]) + 1
        except:
            cur.execute("""CREATE TABLE IF NOT EXISTS projects (
                           id INTEGER PRIMARY KEY,
                           pr_id TEXT,
                           url TEXT,
                           name TEXT NOT NULL,
                           rol TEXT NOT NULL,
                           content TEXT NOT NULL,
                           attachments TEXT,
                           user_id INTEGER);""")
            con.commit()
        return ids

    def insert_into_project(self, args, idd):
        ids = self.lenght_projects()
        con = sql.connect('db/projects.db')
        cur = con.cursor()
        cur.execute(f"INSERT INTO projects (id, pr_id, url, name, rol, content, attachments, user_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    (ids, ) + tuple(args) + (idd, ))
        con.commit()
        cur.close()
