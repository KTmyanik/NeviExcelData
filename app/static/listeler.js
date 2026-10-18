const girdi = document.querySelector("#girdi");
const cikti = document.querySelector("#cikti");
const asama = document.querySelector("#asama");
const ozet = document.querySelector("#ozet");
const icerik = document.querySelector("#icerik");
const calistirBtn = document.querySelector("#calistir");

let izleniyor = false;
let zamanlayici = 0;

function yollar() {
  return { girdi: girdi.value.trim(), cikti_klasoru: cikti.value.trim() };
}

function sure(saniye) {
  if (saniye == null || Number.isNaN(Number(saniye))) return "—";
  const deger = Number(saniye);
  if (deger < 60) return `${deger.toLocaleString("tr-TR", { maximumFractionDigits: 1 })} sn`;
  const dakika = Math.floor(deger / 60);
  const kalan = Math.round(deger % 60);
  return `${dakika} dk ${kalan} sn`;
}

function sayi(deger) {
  return Number(deger || 0).toLocaleString("tr-TR");
}

function kacis(value) {
  return String(value ?? "").replace(/[&<>"']/g, (ch) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[ch]));
}

async function gonder(url, body) {
  const yanit = await fetch(url, {
    method: body ? "POST" : "GET",
    headers: body ? { "Content-Type": "application/json" } : undefined,
    body: body ? JSON.stringify(body) : undefined,
  });
  const data = await yanit.json().catch(() => ({}));
  if (!yanit.ok) throw new Error(data.detail || "İstek tamamlanamadı.");
  return data;
}

function rozet(durum) {
  const ad = durum === "tamam" ? "Tamam" : durum === "uyari" ? "Uyarı" : durum === "hata" ? "Hata" : durum;
  return `<span class="badge ${kacis(durum)}">${kacis(ad)}</span>`;
}

function sonucCiz(data) {
  const kartlar = [
    ["Süre", sure(data.sure)],
    ["Dosya", sayi(data.dosya)],
    ["Toplam satır", sayi(data.toplam_satir)],
    ["Hata", sayi((data.satirlar || []).filter((satir) => satir.durum === "hata").length)],
  ];
  ozet.innerHTML = kartlar.map(([ad, deger]) => `<div class="stat"><strong>${deger}</strong><span>${ad}</span></div>`).join("");
  if (!data.satirlar || !data.satirlar.length) {
    icerik.innerHTML = "";
    return;
  }
  const govde = data.satirlar.map((satir) => {
    const kaynaklar = (satir.kaynaklar || []).map((kaynak) => `
      <li>${kacis(kaynak.etiket)} · ${sayi(kaynak.satir)} satır · ${sure(kaynak.sure)}
        ${kaynak.hata ? `<div class="warn">${kacis(kaynak.hata)}</div>` : ""}
      </li>`).join("");
    const uyarilar = (satir.uyarilar || []).map((uyari) => `<p class="warn">${kacis(uyari)}</p>`).join("");
    return `<tr>
      <td>${kacis(satir.ad)}<div class="sub">${kacis(satir.yol || satir.dosya || "dosya yazılmadı")}</div></td>
      <td>${sayi(satir.satir_sayisi)}</td>
      <td>${sure(satir.sure)}</td>
      <td>${rozet(satir.durum)}${uyarilar}</td>
      <td><ul class="links">${kaynaklar}</ul></td>
    </tr>`;
  }).join("");
  icerik.innerHTML = `<section class="panel"><h2>Son çalışma</h2><table>
    <thead><tr><th>Liste</th><th>Satır</th><th>Süre</th><th>Durum</th><th>Kaynaklar</th></tr></thead>
    <tbody>${govde}</tbody></table></section>`;
}

function onizlemeCiz(satirlar) {
  const govde = satirlar.map((satir) => {
    const adresler = (satir.kaynaklar || []).map((kaynak) =>
      `<li><a href="${kacis(kaynak.url)}" target="_blank" rel="noreferrer">${kacis(kaynak.baslik || kaynak.metin || kaynak.url)}</a></li>`
    ).join("");
    const islemler = (satir.islemler || []).map((islem) => `<div class="sub">${kacis(islem)}</div>`).join("");
    return `<tr>
      <td>${kacis(satir.dosya)}<div class="sub">Excel satırı ${kacis(satir.satir)}</div></td>
      <td>${kacis(satir.ad)}${satir.ipucu ? `<div class="sub">${kacis(satir.ipucu)}</div>` : ""}</td>
      <td>${islemler || '<span class="sub">Adres yok</span>'}</td>
      <td>${adresler ? `<ul class="links">${adresler}</ul>` : "—"}</td>
    </tr>`;
  }).join("");
  icerik.innerHTML = `<section class="panel"><h2>Excel’de bulunan satırlar</h2><table>
    <thead><tr><th>Çıktı</th><th>Satır</th><th>Nasıl okunacak</th><th>Adresler</th></tr></thead>
    <tbody>${govde}</tbody></table></section>`;
  ozet.innerHTML = `<div class="stat"><strong>${sayi(satirlar.length)}</strong><span>Satır</span></div>
    <div class="stat"><strong>${sayi(satirlar.reduce((toplam, satir) => toplam + (satir.kaynaklar || []).length, 0))}</strong><span>Adres</span></div>`;
}

async function durumOku() {
  const data = await gonder("/api/liste-cek/durum");
  asama.textContent = data.calisiyor ? data.asama || "Çalışıyor" : (data.hata || "");
  if (data.hata && !data.calisiyor) asama.textContent = data.hata;
  if (data.satirlar && data.satirlar.length) sonucCiz(data);
  else if (!data.calisiyor) asama.textContent = data.hata || asama.textContent;
  calistirBtn.disabled = Boolean(data.calisiyor);
  if (data.calisiyor && !izleniyor) izle();
  if (!data.calisiyor && izleniyor) {
    window.clearInterval(zamanlayici);
    izleniyor = false;
  }
  return data;
}

function izle() {
  if (izleniyor) return;
  izleniyor = true;
  zamanlayici = window.setInterval(() => {
    durumOku().catch((error) => { asama.textContent = error.message; });
  }, 1000);
}

document.querySelector("#kaydet").addEventListener("click", async () => {
  try {
    await gonder("/api/liste-cek/ayar", yollar());
    asama.textContent = "Yollar kaydedildi.";
  } catch (error) {
    asama.textContent = error.message;
  }
});

document.querySelector("#onizle").addEventListener("click", async () => {
  asama.textContent = "Excel okunuyor";
  try {
    const data = await gonder("/api/liste-cek/onizle", yollar());
    onizlemeCiz(data.satirlar || []);
    asama.textContent = "";
  } catch (error) {
    asama.textContent = error.message;
  }
});

calistirBtn.addEventListener("click", async () => {
  calistirBtn.disabled = true;
  asama.textContent = "Başlıyor";
  try {
    await gonder("/api/liste-cek/calistir", yollar());
    izle();
    await durumOku();
  } catch (error) {
    asama.textContent = error.message;
    calistirBtn.disabled = false;
  }
});

fetch("/api/durum").then((yanit) => {
  if (yanit.ok) document.querySelector("#geri").hidden = false;
}).catch(() => {});

gonder("/api/liste-cek/ayar").then((data) => {
  girdi.value = data.girdi || "";
  cikti.value = data.cikti_klasoru || "";
  return durumOku();
}).catch((error) => { asama.textContent = error.message; });
