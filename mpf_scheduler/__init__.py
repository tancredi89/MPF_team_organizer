from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager
from flask_bcrypt import Bcrypt
from flask_wtf.csrf import CSRFProtect

db = SQLAlchemy()
bcrypt = Bcrypt()
csrf = CSRFProtect()
login_manager = LoginManager()
login_manager.login_view = 'auth.login'

def create_app():
    app = Flask(__name__, instance_relative_config=True)
    # Defaults; overridden by instance/config.py if present
    app.config.from_mapping(
        SECRET_KEY='dev',
        SQLALCHEMY_DATABASE_URI='sqlite:///mpf_scheduler.db',
        SQLALCHEMY_TRACK_MODIFICATIONS=False,
        WTF_CSRF_TIME_LIMIT=None
    )
    app.config.from_pyfile('config.py', silent=True)

    db.init_app(app)
    bcrypt.init_app(app)
    csrf.init_app(app)
    login_manager.init_app(app)

    from .models import User  # noqa

    @login_manager.user_loader
    def load_user(user_id):
        from .models import User
        return User.query.get(int(user_id))

    from .auth import auth as auth_blueprint
    from .views import main as main_blueprint
    app.register_blueprint(auth_blueprint)
    app.register_blueprint(main_blueprint)

    return app

def init_db(app):
    from .models import User, Mission
    from werkzeug.security import generate_password_hash
    with app.app_context():
        db.create_all()
        # Ensure default admin, sample missions
        if not User.query.filter_by(username='admin').first():
            from .models import User
            u = User(username='admin', role='admin')
            from . import bcrypt
            u.password = bcrypt.generate_password_hash('admin123').decode('utf-8')
            db.session.add(u)
            db.session.commit()
        # Seed missions if none
        if not Mission.query.first():
            db.session.add_all([Mission(name=n) for n in ["Mission A", "Mission B", "Mission C", "Mission D", "Mission E"]])
            db.session.commit()
