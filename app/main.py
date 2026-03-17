from contextlib import asynccontextmanager

from fastapi import FastAPI

from app.api.routes import arm_admin, documents, journal, local_llm, schedules, telegram
from app.core.config import settings
from app.db.init_db import init_db


@asynccontextmanager
async def lifespan(_: FastAPI):
    init_db()
    yield


app = FastAPI(
    title=settings.app_name,
    debug=settings.debug,
    lifespan=lifespan,
)


@app.get("/health", tags=["system"])
def healthcheck() -> dict[str, str]:
    return {
        "status": "ok",
        "environment": settings.app_env,
    }


app.include_router(documents.router)
app.include_router(schedules.router)
app.include_router(journal.router)
app.include_router(telegram.router)
app.include_router(local_llm.router)
app.include_router(arm_admin.router)
