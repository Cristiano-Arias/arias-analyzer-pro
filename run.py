# -*- coding: utf-8 -*-
"""
Arquivo principal da aplicação
"""
# Para ambiente de produção com Render e Python 3.13 não utilizamos
# eventlet.  O SocketIO está configurado para usar o modo 'threading'
# (ver app/__init__.py), então não é necessário aplicar monkey_patch.

from app import create_app, socketio
from flask_socketio import join_room

# Criar a aplicação Flask
application = create_app()
app = application  # Alias para compatibilidade com Gunicorn

# Socket.IO event handlers
@socketio.on("join_procurement")
def on_join_proc(data):
    proc_id = data.get("procurement_id")
    if not proc_id:
        return
    room = f"proc:{proc_id}"
    join_room(room)


@socketio.on("join_user")
def on_join_user(data):
    user_id = data.get("user_id")
    if not user_id:
        return
    room = f"user:{user_id}"
    join_room(room)


@socketio.on("join_role") 
def on_join_role(data):
    role = data.get("role")
    if not role:
        return
    room = f"role:{role}"
    join_room(room)


if __name__ == "__main__":
    # Para execução local ou em produção.  Use a porta definida no
    # ambiente (por ex. Render) se disponível; caso contrário, utilize 5000.
    import os
    port = int(os.environ.get("PORT", 5000))
    socketio.run(
        application,
        host="0.0.0.0",
        port=port,
        debug=False,
        # O servidor integrado Werkzeug não é recomendado para produção, mas
        # ao utilizar o modo 'threading' esta é a opção suportada.  Passamos
        # explicitamente ``allow_unsafe_werkzeug=True`` para suprimir a
        # exceção lançada pelo Flask-SocketIO em ambientes de produção.
        allow_unsafe_werkzeug=True,
    )
