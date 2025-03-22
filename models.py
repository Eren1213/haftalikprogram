from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin

db = SQLAlchemy()

class User(UserMixin, db.Model):
    __tablename__ = 'users'

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password = db.Column(db.String(120), nullable=False)
    name = db.Column(db.String(120), nullable=True)
    role = db.Column(db.String(20), nullable=False)
    department_id = db.Column(db.Integer, db.ForeignKey('departments.id'))


class Department(db.Model):
    __tablename__ = 'departments'

    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(10), unique=True, nullable=False)
    name = db.Column(db.String(100), nullable=False)

    users = db.relationship('User', backref='department')
    courses = db.relationship('Course', backref='department')


class Course(db.Model):
    __tablename__ = 'courses'

    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(10), nullable=False)
    name = db.Column(db.String(100), nullable=False)
    theory = db.Column(db.Integer, nullable=True, default=0)
    practice = db.Column(db.Integer, nullable=True, default=0)
    credits = db.Column(db.Integer, nullable=True, default=0)
    semester = db.Column(db.Integer, nullable=True, default=1)
    is_elective = db.Column(db.Boolean, default=False)
    has_fixed_time = db.Column(db.Boolean, default=False)
    department_id = db.Column(db.Integer, db.ForeignKey('departments.id'))
    instructor_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)

    instructor = db.relationship('User', backref='courses')


class Classroom(db.Model):
    __tablename__ = 'classrooms'

    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(20), unique=True, nullable=False)
    capacity = db.Column(db.Integer, nullable=False)
    type = db.Column(db.String(10), nullable=True, default='SINIF')


class Schedule(db.Model):
    __tablename__ = 'schedule_items'

    id = db.Column(db.Integer, primary_key=True)
    course_id = db.Column(db.Integer, db.ForeignKey('courses.id'))
    classroom_id = db.Column(db.Integer, db.ForeignKey('classrooms.id'))
    day = db.Column(db.String(20), nullable=False)
    start_time = db.Column(db.String(5), nullable=False)
    end_time = db.Column(db.String(5), nullable=False)

    course = db.relationship('Course', backref='schedule_items')
    classroom = db.relationship('Classroom', backref='schedule_items')