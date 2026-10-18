"""Resmi liste çekici ekranının adresleri."""

from pathlib import Path

from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel

from app.liste_cekici import ROOT, ayarlari_oku, ayarlari_yaz, is_durumu, isi_baslat, plani_oku

router = APIRouter()
SAYFA = ROOT / "app" / "static" / "listeler.html"


class Yollar(BaseModel):
    girdi: str
    cikti_klasoru: str


@router.get("/listeler")
def sayfa():
    return FileResponse(SAYFA)


@router.get("/api/liste-cek/ayar")
def ayar():
    return ayarlari_oku()


@router.post("/api/liste-cek/ayar")
def ayar_kaydet(istek: Yollar):
    if not istek.girdi.strip() or not istek.cikti_klasoru.strip():
        raise HTTPException(400, "Girdi dosyası ve çıktı klasörü dolu olmalı.")
    return ayarlari_yaz(istek.girdi, istek.cikti_klasoru)


@router.post("/api/liste-cek/onizle")
def onizle(istek: Yollar):
    path = Path(istek.girdi.strip())
    if not path.is_file():
        raise HTTPException(400, f"Girdi dosyası bulunamadı: {path}")
    try:
        return {"satirlar": plani_oku(path)}
    except Exception as exc:
        raise HTTPException(400, f"Excel okunamadı: {exc}") from exc


@router.post("/api/liste-cek/calistir")
def calistir(istek: Yollar):
    if not istek.girdi.strip() or not istek.cikti_klasoru.strip():
        raise HTTPException(400, "Girdi dosyası ve çıktı klasörü dolu olmalı.")
    ayarlari_yaz(istek.girdi, istek.cikti_klasoru)
    try:
        return isi_baslat(istek.girdi, istek.cikti_klasoru)
    except RuntimeError as exc:
        raise HTTPException(409, str(exc)) from exc
    except Exception as exc:
        raise HTTPException(400, str(exc)) from exc


@router.get("/api/liste-cek/durum")
def durum():
    data = is_durumu()
    data.pop("baslangic", None)
    data.pop("bitis", None)
    return data
