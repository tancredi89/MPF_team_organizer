from flask import Blueprint, render_template, redirect, url_for, flash, request, send_file
from flask_login import login_required, current_user
from werkzeug.security import generate_password_hash
from collections import defaultdict
from io import BytesIO
import calendar
from datetime import datetime, date

from .models import User, Mission, Assignment, OnCallAssignment
from .forms import UserForm, EditUserForm, AssignmentForm, OnCallAssignmentForm
from . import db

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

main = Blueprint('main', __name__)

def get_month_dates(year, month):
    _, num_days = calendar.monthrange(year, month)
    return [date(year, month, day) for day in range(1, num_days+1)]

@main.route('/')
@login_required
def dashboard():
    now = datetime.now()
    year = request.args.get('year', now.year, type=int)
    month = request.args.get('month', now.month, type=int)
    dates = get_month_dates(year, month)
    users = User.query.all()
    missions = Mission.query.all()
    first_day = date(year, month, 1)
    last_day = date(year, month, calendar.monthrange(year, month)[1])

    assignments = Assignment.query.filter(
        Assignment.date >= first_day,
        Assignment.date <= last_day
    ).all()

    on_calls = OnCallAssignment.query.filter(
        OnCallAssignment.date >= first_day,
        OnCallAssignment.date <= last_day
    ).all()

    assignment_map = {d: {m.id: [] for m in missions} for d in dates}
    on_call_map = {d: {m.id: [] for m in missions} for d in dates}

    for assign in assignments:
        assignment_map[assign.date][assign.mission_id].append(assign.user.username)

    for oc in on_calls:
        on_call_map[oc.date][oc.mission_id].append(oc.user.username)

    user_summary = defaultdict(lambda: defaultdict(int))
    for assign in assignments:
        user_summary[assign.user.username][assign.mission.name] += 1

    return render_template('dashboard.html', year=year, month=month, dates=dates,
                           users=users, missions=missions, assignment_map=assignment_map,
                           on_call_map=on_call_map, user_summary=user_summary)

@main.route('/users', methods=['GET', 'POST'])
@login_required
def users():
    if current_user.role != 'admin':
        flash("Access denied.")
        return redirect(url_for('main.dashboard'))

    form = UserForm()
    if form.validate_on_submit():
        if User.query.filter_by(username=form.username.data).first():
            flash("User already exists.")
        else:
            new_user = User(
                username=form.username.data,
                role=form.role.data
            )
            new_user.set_password(form.password.data)
            db.session.add(new_user)
            db.session.commit()
            flash("User added successfully.")
            return redirect(url_for('main.users'))

    user_list = User.query.all()
    return render_template('users.html', form=form, users=user_list)

@main.route('/users/delete/<int:user_id>', methods=['POST'])
@login_required
def delete_user(user_id):
    if current_user.role != 'admin':
        flash("Access denied.")
        return redirect(url_for('main.dashboard'))
    user = User.query.get(user_id)
    if user and user.username != 'admin':
        db.session.delete(user)
        db.session.commit()
        flash('User deleted.')
    else:
        flash("Cannot delete default admin.")
    return redirect(url_for('main.users'))

@main.route('/users/edit/<int:user_id>', methods=['GET', 'POST'])
@login_required
def edit_user(user_id):
    if current_user.role != 'admin':
        flash("Access denied.")
        return redirect(url_for('main.dashboard'))
    user = User.query.get_or_404(user_id)
    form = EditUserForm(obj=user)
    if form.validate_on_submit():
        user.username = form.username.data
        user.role = form.role.data
        if form.password.data:
            user.set_password(form.password.data)
        db.session.commit()
        flash("User updated.")
        return redirect(url_for('main.users'))
    return render_template('edit_user.html', form=form, user=user)

@main.route('/missions', methods=['GET', 'POST'])
@login_required
def missions():
    if current_user.role != 'admin':
        flash("Access denied.")
        return redirect(url_for('main.dashboard'))

    if request.method == 'POST':
        name = request.form.get('name')
        if name:
            if Mission.query.filter_by(name=name).first():
                flash("Mission name already exists.")
            else:
                db.session.add(Mission(name=name))
                db.session.commit()
                flash("Mission added.")
        else:
            flash("Mission name cannot be empty.")
    missions = Mission.query.all()
    return render_template('missions.html', missions=missions)

@main.route('/assign', methods=['GET', 'POST'])
@login_required
def assign():
    if current_user.role != 'admin':
        flash("Access denied.")
        return redirect(url_for('main.dashboard'))

    form = AssignmentForm()
    form.user_id.choices = [(u.id, u.username) for u in User.query.order_by('username').all()]
    form.mission_id.choices = [(m.id, m.name) for m in Mission.query.order_by('name').all()]

    if form.validate_on_submit():
        # Avoid duplicates
        exists = Assignment.query.filter_by(user_id=form.user_id.data,
                                            mission_id=form.mission_id.data,
                                            date=form.date.data).first()
        if exists:
            flash("Assignment already exists for this user, mission, and date.")
        else:
            a = Assignment(user_id=form.user_id.data, mission_id=form.mission_id.data, date=form.date.data)
            db.session.add(a)
            db.session.commit()
            flash("Assignment created.")
        return redirect(url_for('main.dashboard'))

    return render_template('assign.html', form=form, title="Assign Mission")

@main.route('/oncall_assign', methods=['GET', 'POST'])
@login_required
def oncall_assign():
    if current_user.role != 'admin':
        flash("Access denied.")
        return redirect(url_for('main.dashboard'))

    form = OnCallAssignmentForm()
    form.user_id.choices = [(u.id, u.username) for u in User.query.order_by('username').all()]
    form.mission_id.choices = [(m.id, m.name) for m in Mission.query.order_by('name').all()]

    if form.validate_on_submit():
        exists = OnCallAssignment.query.filter_by(user_id=form.user_id.data,
                                                  mission_id=form.mission_id.data,
                                                  date=form.date.data).first()
        if exists:
            flash("On-call assignment already exists for this user, mission, and date.")
        else:
            oc = OnCallAssignment(user_id=form.user_id.data, mission_id=form.mission_id.data, date=form.date.data)
            db.session.add(oc)
            db.session.commit()
            flash("On-call assignment created.")
        return redirect(url_for('main.dashboard'))

    return render_template('assign.html', form=form, title="Assign On-Call")

@main.route('/export_excel')
@login_required
def export_excel():
    now = datetime.now()
    year = request.args.get('year', now.year, type=int)
    month = request.args.get('month', now.month, type=int)
    dates = get_month_dates(year, month)
    users = User.query.all()
    missions = Mission.query.all()
    first_day = date(year, month, 1)
    last_day = date(year, month, calendar.monthrange(year, month)[1])

    assignments = Assignment.query.filter(
        Assignment.date >= first_day,
        Assignment.date <= last_day
    ).all()

    on_calls = OnCallAssignment.query.filter(
        OnCallAssignment.date >= first_day,
        OnCallAssignment.date <= last_day
    ).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{year}-{month:02d} Assignments"

    header = ['User', 'Mission', 'Date', 'Assignment Type']
    ws.append(header)

    for a in assignments:
        ws.append([a.user.username, a.mission.name, a.date.strftime("%Y-%m-%d"), "Assigned"])

    for oc in on_calls:
        ws.append([oc.user.username, oc.mission.name, oc.date.strftime("%Y-%m-%d"), "On-Call"])

    for col in ws.columns:
        max_length = 0
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)

    filename = f"assignments_{year}_{month:02d}.xlsx"

    return send_file(stream,
                     as_attachment=True,
                     download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
