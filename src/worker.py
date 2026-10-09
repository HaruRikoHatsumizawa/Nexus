import os

from workers import WorkerEntrypoint, wsgi

from app import app
import src.app as backend

app.template_folder = os.path.join(os.path.dirname(__file__), "templates")


class Default(WorkerEntrypoint):
    async def fetch(self, request):
        secret_key = getattr(self.env, "SECRET_KEY", None)
        admin_user = getattr(self.env, "ADMIN_USER", None)
        admin_hash = getattr(self.env, "ADMIN_PASS_HASH", None)
        openai_key = getattr(self.env, "OPENAI_API_KEY", None)

        if secret_key:
            app.secret_key = secret_key
        if admin_user is not None:
            backend.ADMIN_USER = str(admin_user).strip()
        if admin_hash is not None:
            backend.ADMIN_PASS_HASH = str(admin_hash)
        if openai_key is not None:
            os.environ["OPENAI_API_KEY"] = str(openai_key)

        app.config["SESSION_COOKIE_SECURE"] = True
        app.config["SESSION_COOKIE_HTTPONLY"] = True
        app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

        return await wsgi.fetch(app, request, self.env)