"""Veritabanı olmadan açılan resmi liste ekranı."""

from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from app.liste_api import SAYFA, router
from app.liste_cekici import ROOT

app = FastAPI(title="Resmi Liste Çekici")
app.mount("/static", StaticFiles(directory=ROOT / "app" / "static"), name="static")
app.include_router(router)


@app.get("/")
def index():
    return FileResponse(SAYFA)
