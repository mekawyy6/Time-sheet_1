
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from starlette.middleware.sessions import SessionMiddleware

from app.database import Base, SessionLocal, engine
from app.models import DailyEntry, Site, SiteEvaluation, User  # noqa: F401
from app.routes.web import router as web_router

app = FastAPI(title="Field Team PWA")
app.add_middleware(SessionMiddleware, secret_key="local-single-user-secret")
app.mount("/static", StaticFiles(directory="app/static"), name="static")


def seed_profile() -> None:
    db = SessionLocal()
    try:
        user = db.query(User).filter(User.id == 1).first()
        if not user:
            user = User(
                id=1,
                employee_id="",
                name="",
                email="local@device",
                password_hash="local-only",
                role="local_user",
                team_leader="",
                active=True,
            )
            db.add(user)
            db.commit()
    finally:
        db.close()


@app.on_event("startup")
def on_startup() -> None:
    Base.metadata.create_all(bind=engine)
    seed_profile()


app.include_router(web_router)
