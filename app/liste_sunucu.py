"""Veritabanı olmadan açılan resmi liste ekranı."""

import mimetypes

from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from app.liste_api import SAYFA, router
from app.liste_cekici import ROOT

# Python türü Windows kayıt defterinden okur; bazı bilgisayarlarda .js/.css "text/plain" kayıtlı ve
# tarayıcı dosyayı reddeder.
mimetypes.add_type("text/javascript", ".js")
mimetypes.add_type("text/css", ".css")

app = FastAPI(title="Resmi Liste Çekici")
app.mount("/static", StaticFiles(directory=ROOT / "app" / "static"), name="static")
app.include_router(router)


@app.get("/")
def index():
    return FileResponse(SAYFA)
