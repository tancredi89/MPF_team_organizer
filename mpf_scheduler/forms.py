from flask_wtf import FlaskForm
from wtforms import StringField, PasswordField, SelectField, SubmitField, DateField
from wtforms.validators import InputRequired, Length
from .models import User, Mission

class LoginForm(FlaskForm):
    username = StringField('Username', validators=[InputRequired(), Length(min=3, max=150)])
    password = PasswordField('Password', validators=[InputRequired(), Length(min=6)])

class UserForm(FlaskForm):
    username = StringField('Username', validators=[InputRequired(), Length(min=3, max=150)])
    password = PasswordField('Password', validators=[InputRequired(), Length(min=6)])
    role = SelectField('Role', choices=[('admin', 'Admin'), ('user', 'User')])
    submit = SubmitField('Add User')

class EditUserForm(FlaskForm):
    username = StringField('Username', validators=[InputRequired(), Length(min=3, max=150)])
    password = PasswordField('New Password (leave blank to keep)')
    role = SelectField('Role', choices=[('admin', 'Admin'), ('user', 'User')])
    submit = SubmitField('Update User')

class MissionForm(FlaskForm):
    name = StringField('Mission Name', validators=[InputRequired(), Length(min=2, max=150)])
    submit = SubmitField('Add Mission')

class AssignmentForm(FlaskForm):
    user_id = SelectField('User', coerce=int, validators=[InputRequired()])
    mission_id = SelectField('Mission', coerce=int, validators=[InputRequired()])
    date = DateField('Date', format='%Y-%m-%d', validators=[InputRequired()])
    submit = SubmitField('Assign')

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.user_id.choices = [(u.id, u.username) for u in User.query.order_by(User.username).all()]
        self.mission_id.choices = [(m.id, m.name) for m in Mission.query.order_by(Mission.name).all()]

class OnCallAssignmentForm(FlaskForm):
    user_id = SelectField('User', coerce=int, validators=[InputRequired()])
    mission_id = SelectField('Mission', coerce=int, validators=[InputRequired()])
    date = DateField('Date', format='%Y-%m-%d', validators=[InputRequired()])
    submit = SubmitField('Assign On-Call')

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.user_id.choices = [(u.id, u.username) for u in User.query.order_by(User.username).all()]
        self.mission_id.choices = [(m.id, m.name) for m in Mission.query.order_by(Mission.name).all()]
