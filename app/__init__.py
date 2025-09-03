# -*- coding: utf-8 -*-
import os
from flask import Flask, render_template
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS
from flask_jwt_extended import JWTManager
from flask_socketio import SocketIO
from .config import Config

db = SQLAlchemy()
jwt = JWTManager()

# SocketIO initialization: use Python's threading mode by default.  Eventlet
# isn't compatible with Python 3.13 at this time, and using threading avoids
# dependency on a green‑thread implementation.  This mode provides
# acceptable performance for moderate loads and full WebSocket support.
socketio = SocketIO(cors_allowed_origins="*", async_mode='threading')


def create_app() -> Flask:
    app = Flask(__name__, 
                static_folder="../static", 
                static_url_path="/static",
                template_folder="../templates")
    
    app.config.from_object(Config())

    CORS(app)  # allow cross-origin for MVP
    db.init_app(app)
    jwt.init_app(app)
    # Inicialize o SocketIO para esta instância de app.  Não especifique
    # explicitamente 'eventlet' aqui; deixe o ``async_mode`` herdado do
    # objeto global ``socketio`` (que foi configurado para 'threading').
    socketio.init_app(app, cors_allowed_origins="*")

    with app.app_context():
        from . import models  # noqa: F401
        db.create_all()

        # Register blueprints
        from .blueprints.auth import bp as auth_bp
        from .blueprints.procurements import bp as proc_bp
        from .blueprints.tr import bp as tr_bp
        from .blueprints.proposals import bp as proposals_bp

        app.register_blueprint(auth_bp, url_prefix="/api/auth")
        app.register_blueprint(proc_bp, url_prefix="/api")
        app.register_blueprint(tr_bp, url_prefix="/api")
        app.register_blueprint(proposals_bp, url_prefix="/api")

        # Rota principal para servir o HTML
        @app.route('/')
        def index():
            return render_template('index.html')

        # Simple healthcheck
        @app.get("/healthz")
        def healthz():
            return {"status": "ok"}

    return app
