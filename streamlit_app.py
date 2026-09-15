import streamlit as st
import streamlit.components.v1 as components
import io, os, datetime, math, json, numpy as np, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter

# =====================================================================
# KONFIGURACIJA (secrets) + SUPABASE
# =====================================================================
try:
    from supabase import create_client
except Exception:
    create_client = None

def _cfg(key, default=None):
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return default

APP_PASSWORD   = _cfg("APP_PASSWORD", "vape2024")     # analiticar (pun pristup)
ADMIN_PASSWORD = _cfg("ADMIN_PASSWORD", "aman2024")   # administracija 1 (koleginica 1)
ADMIN_PASSWORD_2 = _cfg("ADMIN_PASSWORD_2", "aman2025")  # administracija 2 (koleginica 2)
DIREKTOR_PASSWORD = _cfg("DIREKTOR_PASSWORD", "2026vape")  # direktori (pregled izveštaja)
KOMERCIJALA_PASSWORD = _cfg("KOMERCIJALA_PASSWORD", "komerc2026")     # komercijala 1
KOMERCIJALA_PASSWORD_2 = _cfg("KOMERCIJALA_PASSWORD_2", "komerc2027")  # komercijala 2

# --- korisnicka imena (prijava: ime + sifra) ---
APP_KORISNIK          = _cfg("APP_KORISNIK", "analitika")
ADMIN_KORISNIK        = _cfg("ADMIN_KORISNIK", "aleksandraapatovic")
ADMIN_KORISNIK_2      = _cfg("ADMIN_KORISNIK_2", "zoranahromis")
DIREKTOR_KORISNIK     = _cfg("DIREKTOR_KORISNIK", "direktor")
KOMERCIJALA_KORISNIK  = _cfg("KOMERCIJALA_KORISNIK", "komercijala1")
KOMERCIJALA_KORISNIK_2 = _cfg("KOMERCIJALA_KORISNIK_2", "komercijala2")

# --- imena koja se prikazuju i upisuju u dnevnik ---
ADMIN_IME        = _cfg("ADMIN_IME", "Aleksandra Apatović")
ADMIN_IME_2      = _cfg("ADMIN_IME_2", "Zorana Hromiš")
KOMERCIJALA_IME  = _cfg("KOMERCIJALA_IME", "Komercijala 1")
KOMERCIJALA_IME_2 = _cfg("KOMERCIJALA_IME_2", "Komercijala 2")


def _norm_korisnik(s):
    """Sve svede na mala slova bez razmaka, tacaka i nasih kvacica."""
    t = str(s or "").strip().lower()
    for _a, _b in (("š", "s"), ("đ", "dj"), ("č", "c"), ("ć", "c"), ("ž", "z")):
        t = t.replace(_a, _b)
    for _z in (" ", ".", "-", "_"):
        t = t.replace(_z, "")
    return t


def _aliasi(*vrednosti):
    """Prihvatljiva korisnicka imena: i skraceno ime, i puno ime, i mejl (ceo ili do @)."""
    out = set()
    for v in vrednosti:
        v = str(v or "").strip()
        if not v:
            continue
        out.add(_norm_korisnik(v))
        if "@" in v:
            out.add(_norm_korisnik(v.split("@")[0]))
    out.discard("")
    return out

SUPABASE_URL   = _cfg("SUPABASE_URL", "")
SUPABASE_KEY   = _cfg("SUPABASE_KEY", "")

MESEC_NAZIVI = {1:'Januar',2:'Februar',3:'Mart',4:'April',5:'Maj',6:'Jun',
                7:'Jul',8:'Avgust',9:'Septembar',10:'Oktobar',11:'Novembar',12:'Decembar'}

def _now():
    """Trenutno vreme po Beogradu (server radi po UTC-u)."""
    try:
        from zoneinfo import ZoneInfo
        return datetime.datetime.now(ZoneInfo("Europe/Belgrade")).replace(tzinfo=None)
    except Exception:
        # rezerva: leto UTC+2 (CEST). Ako zoneinfo/tzdata nedostaje.
        return datetime.datetime.utcnow() + datetime.timedelta(hours=2)

def mesec_label(key):
    try:
        y, m = str(key).split('-')
        return f"{MESEC_NAZIVI.get(int(m), m)} {y}"
    except Exception:
        return str(key)

@st.cache_resource
def _sb():
    if not (create_client and SUPABASE_URL and SUPABASE_KEY):
        return None
    try:
        return create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception:
        return None

def sb_dostupan():
    return _sb() is not None

def sb_objavi(mesec_key, sistem, podaci, xlsx_b64=None):
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen (SUPABASE_URL / SUPABASE_KEY u secrets).")
    payload = {"mesec": mesec_key, "sistem": sistem.strip(), "podaci": podaci}
    if xlsx_b64:
        payload["analitika_xlsx"] = xlsx_b64
    try:
        cli.table("porudzbine").upsert(payload, on_conflict="mesec,sistem").execute()
    except Exception:
        # kolona 'analitika_xlsx' možda ne postoji -> objavi bez nje (ne ruši objavu)
        payload.pop("analitika_xlsx", None)
        cli.table("porudzbine").upsert(payload, on_conflict="mesec,sistem").execute()
    # osvezi kes da koleginice odmah vide
    for fn in (sb_meseci, sb_sisteme, sb_svi_sistemi, sb_ucitaj, sb_pregled, sb_ucitaj_xlsx):
        try: fn.clear()
        except Exception: pass


def sb_obrisi(mesec_key, sistem):
    """Obriši objavljeni izveštaj (mesec + sistem) iz baze. Trajno."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    cli.table("porudzbine").delete().eq("mesec", mesec_key).eq("sistem", sistem).execute()
    for fn in (sb_meseci, sb_sisteme, sb_svi_sistemi, sb_ucitaj, sb_pregled, sb_ucitaj_xlsx):
        try: fn.clear()
        except Exception: pass


@st.cache_data(ttl=60)
def sb_ucitaj_xlsx(mesec_key, sistem):
    """Vrati base64 analitika Excel-a za sistem/mesec (ili None). Odvojen upit da ne
    opterećuje običan sb_ucitaj."""
    cli = _sb()
    if cli is None:
        return None
    try:
        res = cli.table("porudzbine").select("analitika_xlsx").eq("mesec", mesec_key).eq("sistem", sistem).limit(1).execute()
        if not res.data:
            return None
        return res.data[0].get("analitika_xlsx")
    except Exception:
        return None


def sb_objavi_izvestaj_prodaje(html, xlsx_b64, mesec_label, prodaja_json=""):
    """Sačuvaj (samo poslednji) direktorski Izveštaj prodaje u tabelu izvestaj_prodaje."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    payload = {"kljuc": "latest", "html": html, "xlsx_b64": xlsx_b64,
               "mesec_label": mesec_label or "", "prodaja_json": prodaja_json or "",
               "generisano": _now().strftime("%d.%m.%Y %H:%M")}
    try:
        cli.table("izvestaj_prodaje").upsert(payload, on_conflict="kljuc").execute()
    except Exception:
        # kolona prodaja_json možda ne postoji -> sačuvaj bez nje
        payload.pop("prodaja_json", None)
        cli.table("izvestaj_prodaje").upsert(payload, on_conflict="kljuc").execute()
    try:
        sb_ucitaj_izvestaj_prodaje.clear()
    except Exception:
        pass


@st.cache_data(ttl=60)
def sb_ucitaj_izvestaj_prodaje():
    cli = _sb()
    if cli is None:
        return None
    try:
        res = cli.table("izvestaj_prodaje").select("html,xlsx_b64,mesec_label,generisano,prodaja_json").eq("kljuc", "latest").limit(1).execute()
        if res.data:
            return res.data[0]
    except Exception:
        pass
    try:
        res = cli.table("izvestaj_prodaje").select("html,xlsx_b64,mesec_label,generisano").eq("kljuc", "latest").limit(1).execute()
        if not res.data:
            return None
        return res.data[0]
    except Exception:
        return None


@st.cache_data(ttl=30)
def sb_meseci():
    cli = _sb()
    if cli is None: return []
    res = cli.table("porudzbine").select("mesec").execute()
    keys = sorted({r["mesec"] for r in (res.data or [])}, reverse=True)
    return [{"key": k, "label": mesec_label(k)} for k in keys]

@st.cache_data(ttl=30)
def sb_sisteme(mesec_key):
    cli = _sb()
    if cli is None: return []
    res = cli.table("porudzbine").select("sistem").eq("mesec", mesec_key).execute()
    return sorted({r["sistem"] for r in (res.data or [])})

@st.cache_data(ttl=30)
def sb_svi_sistemi():
    cli = _sb()
    if cli is None: return []
    res = cli.table("porudzbine").select("sistem").execute()
    return sorted({r["sistem"] for r in (res.data or [])})

@st.cache_data(ttl=30)
def sb_ucitaj(mesec_key, sistem):
    cli = _sb()
    if cli is None: return None
    res = cli.table("porudzbine").select("podaci").eq("mesec", mesec_key).eq("sistem", sistem).limit(1).execute()
    if not res.data: return None
    return res.data[0]["podaci"]

def sb_predaj(mesec_key, sistem):
    """Označi izveštaj (mesec+sistem) kao PREDAT — zaključava izmene za taj mesec.
    Upisuje meta.predato u JSON reda u tabeli porudzbine."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    res = cli.table("porudzbine").select("podaci").eq("mesec", mesec_key).eq("sistem", sistem.strip()).limit(1).execute()
    if not res.data:
        raise RuntimeError("Nema objavljenog izveštaja za taj mesec/sistem.")
    podaci = res.data[0].get("podaci") or {}
    meta = podaci.get("meta") or {}
    meta["predato"] = True
    meta["predato_at"] = _now().strftime("%d.%m.%Y %H:%M")
    podaci["meta"] = meta
    cli.table("porudzbine").update({"podaci": podaci}).eq("mesec", mesec_key).eq("sistem", sistem.strip()).execute()
    try:
        sb_ucitaj.clear()
    except Exception:
        pass


def sb_napomena_sistem(mesec_key, sistem, tekst):
    """Sačuvaj sistemsku napomenu (za nedeljne/sistemske sisteme) u meta.napomena_sistem."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    res = cli.table("porudzbine").select("podaci").eq("mesec", mesec_key).eq("sistem", sistem.strip()).limit(1).execute()
    if not res.data:
        raise RuntimeError("Nema objavljenog izveštaja za taj mesec/sistem.")
    podaci = res.data[0].get("podaci") or {}
    meta = podaci.get("meta") or {}
    meta["napomena_sistem"] = tekst or ""
    meta["napomena_sistem_at"] = _now().strftime("%d.%m.%Y %H:%M")
    podaci["meta"] = meta
    cli.table("porudzbine").update({"podaci": podaci}).eq("mesec", mesec_key).eq("sistem", sistem.strip()).execute()
    try:
        sb_ucitaj.clear()
    except Exception:
        pass


def sb_nedeljni_prijava_set(mesec_key, sistem, idk, prijavljen, napomena="", ko=""):
    """Za sistemske/nedeljne sisteme: označi (ili skini oznaku) da je za konkretan
    komitent PRIJAVLJEN problem komercijali. Čuva se u meta.nedeljni_prijave =
    {idk: {napomena, at, ko}}. Prisustvo ključa = prijavljeno."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    res = cli.table("porudzbine").select("podaci").eq("mesec", mesec_key).eq("sistem", sistem.strip()).limit(1).execute()
    if not res.data:
        raise RuntimeError("Nema objavljenog izveštaja za taj mesec/sistem.")
    podaci = res.data[0].get("podaci") or {}
    meta = podaci.get("meta") or {}
    prij = dict(meta.get("nedeljni_prijave") or {})
    _k = str(int(idk))
    if prijavljen:
        prij[_k] = {"napomena": napomena or "", "at": _now().strftime("%d.%m.%Y %H:%M"), "ko": ko or ""}
    else:
        prij.pop(_k, None)
    meta["nedeljni_prijave"] = prij
    podaci["meta"] = meta
    cli.table("porudzbine").update({"podaci": podaci}).eq("mesec", mesec_key).eq("sistem", sistem.strip()).execute()
    try:
        sb_ucitaj.clear()
    except Exception:
        pass


def sb_nedeljni_mail_set(mesec_key, sistem, to_email, ko="", poslato=False):
    """Za sistemske sisteme: zapamti mejl nadležnog (glavni kontakt) i podatak o
    poslednjem slanju. Čuva se u meta.nedeljni_mail = {to, at, ko}.
    poslato=False samo pamti adresu; poslato=True beleži i vreme slanja."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    res = cli.table("porudzbine").select("podaci").eq("mesec", mesec_key).eq("sistem", sistem.strip()).limit(1).execute()
    if not res.data:
        raise RuntimeError("Nema objavljenog izveštaja za taj mesec/sistem.")
    podaci = res.data[0].get("podaci") or {}
    meta = podaci.get("meta") or {}
    _m = dict(meta.get("nedeljni_mail") or {})
    _m["to"] = (to_email or "").strip()
    if poslato:
        _m["at"] = _now().strftime("%d.%m.%Y %H:%M")
        _m["ko"] = ko or ""
    meta["nedeljni_mail"] = _m
    podaci["meta"] = meta
    cli.table("porudzbine").update({"podaci": podaci}).eq("mesec", mesec_key).eq("sistem", sistem.strip()).execute()
    try:
        sb_ucitaj.clear()
    except Exception:
        pass


def sb_nedeljni_start_set(mesec_key, sistem, snap):
    """Za sistemske sisteme: zabeleži STARTNO stanje izveštaja (broj objekata sa
    problemom u trenutku pokretanja) u meta.nedeljni_start = {kada, n, problem}.
    Snima se JEDNOM — kasnija ažuriranja iz admina ne diraju ovaj start."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    res = cli.table("porudzbine").select("podaci").eq("mesec", mesec_key).eq("sistem", sistem.strip()).limit(1).execute()
    if not res.data:
        raise RuntimeError("Nema objavljenog izveštaja za taj mesec/sistem.")
    podaci = res.data[0].get("podaci") or {}
    meta = podaci.get("meta") or {}
    meta["nedeljni_start"] = snap
    podaci["meta"] = meta
    cli.table("porudzbine").update({"podaci": podaci}).eq("mesec", mesec_key).eq("sistem", sistem.strip()).execute()
    try:
        sb_ucitaj.clear()
    except Exception:
        pass


# =====================================================================
# KOMERCIJALA — rute i kontrola (tabele: rute_dan, ruta_objekti)
# =====================================================================
RUTA_SQL_PORUKA = ("Tabele za rute još nisu napravljene u bazi. Analitičar treba da pokrene SQL "
                   "(KORAK 1 iz fajla „SQL_komercijala_rute.sql“) u Supabase → SQL Editor.")


def sb_ruta_tabele_ok():
    """Da li postoje tabele rute_dan i ruta_objekti (i da li aplikacija ima pristup)."""
    cli = _sb()
    if cli is None:
        return (False, "Supabase nije podešen.")
    try:
        cli.table("rute_dan").select("id").limit(1).execute()
        cli.table("ruta_objekti").select("id").limit(1).execute()
        return (True, "")
    except Exception as _e:
        _m = str(_e)
        if ("does not exist" in _m) or ("PGRST205" in _m) or ("42P01" in _m) or ("schema cache" in _m):
            return (False, RUTA_SQL_PORUKA)
        return (False, "Baza javlja grešku: " + _m)


def sb_ruta_dan_get(datum, ko):
    """Vrati {start_at, kraj_at, napomena} za dan rute (ili prazno)."""
    cli = _sb()
    if cli is None:
        return {}
    try:
        res = cli.table("rute_dan").select("*").eq("datum", str(datum)).eq("ko", ko).limit(1).execute()
        return (res.data or [{}])[0] or {}
    except Exception:
        return {}


def sb_ruta_dan_set(datum, ko, start_at=None, kraj_at=None, napomena=None):
    """Upiši start/kraj rute. Šalju se samo polja koja nisu None."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    _p = {"datum": str(datum), "ko": ko}
    if start_at is not None:
        _p["start_at"] = start_at
    if kraj_at is not None:
        _p["kraj_at"] = kraj_at
    if napomena is not None:
        _p["napomena"] = napomena
    try:
        cli.table("rute_dan").upsert(_p, on_conflict="datum,ko").execute()
    except Exception as _e:
        _m = str(_e)
        if ("does not exist" in _m) or ("PGRST205" in _m) or ("42P01" in _m) or ("schema cache" in _m):
            raise RuntimeError(RUTA_SQL_PORUKA)
        raise RuntimeError("Nije sačuvano u bazu: " + _m)


def sb_ruta_objekti(datum, ko):
    """Svi objekti u ruti za dati dan i komercijalistu."""
    cli = _sb()
    if cli is None:
        return []
    try:
        res = cli.table("ruta_objekti").select("*").eq("datum", str(datum)).eq("ko", ko).execute()
        return list(res.data or [])
    except Exception:
        return []


def sb_ruta_obj_get(datum, ko, idk):
    cli = _sb()
    if cli is None:
        return {}
    try:
        res = (cli.table("ruta_objekti").select("*").eq("datum", str(datum))
               .eq("ko", ko).eq("idk", int(idk)).limit(1).execute())
        return (res.data or [{}])[0] or {}
    except Exception:
        return {}


def sb_ruta_obj_save(datum, ko, idk, **polja):
    """Upsert jednog objekta u ruti. Šalju se samo prosleđena polja."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    _p = {"datum": str(datum), "ko": ko, "idk": int(idk)}
    for _k, _v in (polja or {}).items():
        if _v is not None:
            _p[_k] = _v
    try:
        cli.table("ruta_objekti").upsert(_p, on_conflict="datum,ko,idk").execute()
    except Exception as _e:
        _m = str(_e)
        if ("does not exist" in _m) or ("PGRST205" in _m) or ("42P01" in _m) or ("schema cache" in _m):
            raise RuntimeError(RUTA_SQL_PORUKA)
        raise RuntimeError("Nije sačuvano u bazu: " + _m)


def sb_ruta_obj_del(datum, ko, idk):
    cli = _sb()
    if cli is None:
        return
    try:
        (cli.table("ruta_objekti").delete().eq("datum", str(datum))
         .eq("ko", ko).eq("idk", int(idk)).execute())
    except Exception:
        pass


def sb_ruta_mesec(mesec_key):
    """Svi obilasci u mesecu (svi komercijalisti) — za karticu Kontrola.
    Vrati {idk: red_sa_najnovijim_obilaskom}."""
    cli = _sb()
    if cli is None:
        return {}
    try:
        res = cli.table("ruta_objekti").select("*").eq("mesec", str(mesec_key)).execute()
    except Exception:
        return {}
    out = {}
    for r in (res.data or []):
        _i = int(r.get("idk") or 0)
        _prev = out.get(_i)
        # zadrži najnoviji obilazak (ili onaj koji je završen)
        if (_prev is None) or (str(r.get("datum") or "") > str(_prev.get("datum") or "")):
            out[_i] = r
    return out


RUTA_BUCKET = "ruta-slike"
RUTA_BUCKET_PORUKA = ("Prostor za slike („bucket“) još ne postoji u Supabase i aplikacija nije uspela sama da ga "
                      "napravi. Uradi to jednom ručno: Supabase → Storage → New bucket → ime "
                      "ruta-slike → čekiraj Public bucket → Save.")


def _sb_napravi_bucket():
    """Pokušaj da napraviš javni bucket za slike (radi ako ključ ima prava)."""
    cli = _sb()
    if cli is None:
        return False
    for _args, _kw in (((RUTA_BUCKET,), {"options": {"public": True}}),
                       ((RUTA_BUCKET,), {"options": {"public": "true"}}),
                       ((RUTA_BUCKET, RUTA_BUCKET), {"options": {"public": True}}),
                       ((RUTA_BUCKET,), {})):
        try:
            cli.storage.create_bucket(*_args, **_kw)
            return True
        except Exception as _e:
            if "already exists" in str(_e).lower() or "duplicate" in str(_e).lower():
                return True
            continue
    return False


def sb_ruta_slika_upload(file_bytes, filename, content_type="image/jpeg"):
    """Otpremi sliku u Supabase Storage bucket 'ruta-slike' i vrati javni URL.
    Ako bucket ne postoji — pokuša sam da ga napravi, pa ponovi slanje."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    _path = _now().strftime("%Y/%m/") + filename

    def _posalji():
        return cli.storage.from_(RUTA_BUCKET).upload(
            _path, file_bytes, {"content-type": content_type, "upsert": "true"})

    try:
        _posalji()
    except Exception as _e:
        _m = str(_e)
        if ("bucket not found" in _m.lower()) or ("404" in _m):
            if _sb_napravi_bucket():
                try:
                    _posalji()
                except Exception as _e2:
                    raise RuntimeError("Slanje slike nije uspelo: " + str(_e2))
            else:
                raise RuntimeError(RUTA_BUCKET_PORUKA)
        else:
            # možda fajl već postoji — probaj update
            try:
                cli.storage.from_(RUTA_BUCKET).update(
                    _path, file_bytes, {"content-type": content_type})
            except Exception:
                raise RuntimeError("Slanje slike nije uspelo: " + _m)
    try:
        return cli.storage.from_(RUTA_BUCKET).get_public_url(_path)
    except Exception:
        return ""


def _smanji_sliku(file_bytes, max_px=1400, kvalitet=78):
    """Smanji sliku pre slanja (da ne troši prostor). Ako PIL nije dostupan — vrati original."""
    try:
        from PIL import Image as _PILImage
        import io as _io
        _im = _PILImage.open(_io.BytesIO(file_bytes))
        if _im.mode in ("RGBA", "P", "LA"):
            _im = _im.convert("RGB")
        _w, _h = _im.size
        if max(_w, _h) > max_px:
            _sc = max_px / float(max(_w, _h))
            _im = _im.resize((int(_w * _sc), int(_h * _sc)), _PILImage.LANCZOS)
        _b = _io.BytesIO()
        _im.save(_b, format="JPEG", quality=kvalitet, optimize=True)
        return _b.getvalue()
    except Exception:
        return file_bytes


def sb_start_snapshot(mesec_key, sistem, snap):
    """Zabeleži „startni rezultat" (zone u trenutku prvog povlačenja porudžbina od 01.)
    u meta.start_zone. Jednom snimljeno — koristi se u PDF izveštaju kao START."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    res = cli.table("porudzbine").select("podaci").eq("mesec", mesec_key).eq("sistem", sistem.strip()).limit(1).execute()
    if not res.data:
        raise RuntimeError("Nema objavljenog izveštaja za taj mesec/sistem.")
    podaci = res.data[0].get("podaci") or {}
    meta = podaci.get("meta") or {}
    meta["start_zone"] = snap
    podaci["meta"] = meta
    cli.table("porudzbine").update({"podaci": podaci}).eq("mesec", mesec_key).eq("sistem", sistem.strip()).execute()
    try:
        sb_ucitaj.clear()
    except Exception:
        pass


def sb_admin_hist_set(mesec_key, sistem, hist_map, kada=None, svi_idk=None):
    """Zapamti (trajno) povučene prethodne porudžbine iz admina za ceo sistem.
    hist_map: {idk: [ {id,datum,status,cena,stavke} ]}. Upisuje se u meta.admin_hist
    (ključevi kao stringovi) + meta.admin_hist_at (vreme poslednjeg ažuriranja).
    svi_idk: SVI objekti koji su obuhvaćeni ovim povlačenjem (i oni bez ijedne
    porudžbine) — upisuje se u meta.admin_hist_idk, da bi aplikacija znala da su
    i ti objekti već provereni i da za njih NE traži ponovo ažuriranje.
    Tako se posle osvežavanja stranice zadržavaju „posle 01." i sortiranje po zonama."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    res = cli.table("porudzbine").select("podaci").eq("mesec", mesec_key).eq("sistem", sistem.strip()).limit(1).execute()
    if not res.data:
        raise RuntimeError("Nema objavljenog izveštaja za taj mesec/sistem.")
    podaci = res.data[0].get("podaci") or {}
    meta = podaci.get("meta") or {}
    meta["admin_hist"] = {str(k): (v or []) for k, v in (hist_map or {}).items() if v}
    meta["admin_hist_at"] = kada or _now().isoformat()
    if svi_idk:
        _prev_cov = meta.get("admin_hist_idk") or []
        _cov = set(str(x) for x in _prev_cov) | set(str(x) for x in svi_idk)
        meta["admin_hist_idk"] = sorted(_cov)
    podaci["meta"] = meta
    cli.table("porudzbine").update({"podaci": podaci}).eq("mesec", mesec_key).eq("sistem", sistem.strip()).execute()
    try:
        sb_ucitaj.clear()
    except Exception:
        pass


@st.cache_data(ttl=30)
def sb_pregled():
    """Lagani pregled svega objavljenog: mesec, sistem, kada (bez povlacenja stavki)."""
    cli = _sb()
    if cli is None: return []
    res = cli.table("porudzbine").select("mesec,sistem,objavljeno").order("mesec", desc=True).execute()
    return res.data or []

REAKCIJE_OPCIJE = ["Pozvala sam", "Poslala sam mejl", "Obavestila direktorku"]

def _reak_short(r):
    return {"Pozvala sam": "\U0001F4DE Pozvala", "Poslala sam mejl": "\u2709\uFE0F Mejl",
            "Obavestila direktorku": "\U0001F454 Komercijala",
            "Ubačena porudžbina": "\U0001F4E6 Ubačena porudžbina"}.get(r, r)

def _ko_kratko(name):
    """Kratka oznaka: 'Aleksandra Apatović' -> 'AA', 'Administracija 1' -> 'A1'."""
    n = str(name or "").strip()
    _m = {"Administracija 1": "A1", "Administracija 2": "A2"}
    if n in _m:
        return _m[n]
    _d = [d for d in n.split() if d]
    if len(_d) >= 2:
        return (_d[0][:1] + _d[1][:1]).upper()
    return n[:10] if n else ""

def _reak_short_ko(r, ko_map):
    """Reakcija + ko ju je uneo, npr. 'Pozvala · A1'."""
    _base = _reak_short(r)
    _ko = (ko_map or {}).get(r, "")
    return _base + ((" · " + _ko_kratko(_ko)) if _ko else "")

def _zona_disp(nivo):
    return {"crveno": ("z-red", "\U0001F534 Hitno pozvati", "\U0001F534", "Hitno pozvati"),
            "zuto":   ("z-org", "\U0001F7E0 Iskontrolisati", "\U0001F7E0", "Iskontrolisati"),
            "zeleno": ("z-grn", "\U0001F7E2 Dobra", "\U0001F7E2", "Dobra")}[nivo]

def sb_load_obrada(mesec_key, sistem):
    cli = _sb()
    if cli is None:
        return {}
    try:
        res = cli.table("obrada").select("idk,reakcije,trebovali,trebovali_tip,njihova,napomena,reakcije_ko,azurirao,dnevnik").eq("mesec", mesec_key).eq("sistem", sistem).execute()
    except Exception:
        # kolone reakcije_ko/azurirao/dnevnik možda još ne postoje -> učitaj bez njih
        try:
            res = cli.table("obrada").select("idk,reakcije,trebovali,trebovali_tip,njihova,napomena").eq("mesec", mesec_key).eq("sistem", sistem).execute()
        except Exception:
            return {}
    out = {}
    for r in (res.data or []):
        out[int(r["idk"])] = {"reakcije": r.get("reakcije") or [], "trebovali_tip": r.get("trebovali_tip") or "",
                              "njihova": r.get("njihova") or {}, "napomena": r.get("napomena") or "",
                              "reakcije_ko": r.get("reakcije_ko") or {}, "azurirao": r.get("azurirao") or "",
                              "dnevnik": r.get("dnevnik") or {}}
    return out


def _dt_kratko(iso):
    """ISO -> 'DD.MM.YYYY HH:MM' (ili original)."""
    if not iso:
        return ""
    try:
        return datetime.datetime.fromisoformat(str(iso)).strftime("%d.%m.%Y %H:%M")
    except Exception:
        return str(iso)


def sb_obrada_ocisti(mesec_key, sistem, idk, sta="mejl"):
    """Obriši dnevnik i reakciju za objekat — za čišćenje testnih zapisa.
    sta: 'mejl' | 'poziv' | 'komercijala' | 'sve'. Ostalo (napomena, njihova,
    trebovanje) se ne dira."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    try:
        res = (cli.table("obrada")
               .select("reakcije,trebovali_tip,njihova,napomena,reakcije_ko,dnevnik")
               .eq("mesec", mesec_key).eq("sistem", sistem).eq("idk", int(idk)).limit(1).execute())
        _r = res.data[0] if res.data else {}
    except Exception:
        _r = {}
    if not _r:
        return 0
    reakcije = list(_r.get("reakcije") or [])
    reakcije_ko = dict(_r.get("reakcije_ko") or {})
    dnevnik = dict(_r.get("dnevnik") or {})
    _mapa = {"mejl": ("mejlovi", "Poslala sam mejl"),
             "poziv": ("pozivi", "Pozvala sam"),
             "komercijala": ("komercijala", "Obavestila direktorku")}
    _kljucevi = list(_mapa.keys()) if sta == "sve" else [sta]
    _obrisano = 0
    for _k in _kljucevi:
        _dk, _reak = _mapa.get(_k, (None, None))
        if not _dk:
            continue
        _obrisano += len(dnevnik.get(_dk) or [])
        dnevnik.pop(_dk, None)
        if _reak in reakcije:
            reakcije.remove(_reak)
            if _obrisano == 0:
                _obrisano = 1     # stari zapis bez dnevnika
        reakcije_ko.pop(_reak, None)
    _row = {"mesec": mesec_key, "sistem": sistem, "idk": int(idk),
            "reakcije": reakcije, "trebovali": bool(_r.get("trebovali_tip")),
            "trebovali_tip": _r.get("trebovali_tip") or "", "njihova": _r.get("njihova") or {},
            "napomena": _r.get("napomena") or "", "reakcije_ko": reakcije_ko, "dnevnik": dnevnik,
            "azurirano": _now().isoformat()}
    try:
        cli.table("obrada").upsert(_row, on_conflict="mesec,sistem,idk").execute()
    except Exception:
        for _c in ("dnevnik", "reakcije_ko"):
            _row.pop(_c, None)
        cli.table("obrada").upsert(_row, on_conflict="mesec,sistem,idk").execute()
    return _obrisano


def sb_obrada_log(mesec_key, sistem, idk, kind, ko="", kopija=None):
    """Zabeleži poziv ili mejl u dnevnik obrade (dnevnik.pozivi / dnevnik.mejlovi = lista {ko, at}).
    Uz to upali odgovarajuću reakciju (Pozvala sam / Poslala sam mejl) da se vidi u pregledu.
    kind: 'poziv' ili 'mejl'. Čuva postojeća polja (trebovanje, njihova, napomena)."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    try:
        res = cli.table("obrada").select("reakcije,trebovali_tip,njihova,napomena,reakcije_ko,dnevnik").eq("mesec", mesec_key).eq("sistem", sistem).eq("idk", int(idk)).limit(1).execute()
        _r = res.data[0] if res.data else {}
    except Exception:
        _r = {}
    reakcije = list(_r.get("reakcije") or [])
    reakcije_ko = dict(_r.get("reakcije_ko") or {})
    dnevnik = dict(_r.get("dnevnik") or {})
    _at = _now().isoformat()
    if kind == "poziv":
        dnevnik.setdefault("pozivi", []).append({"ko": ko or "", "at": _at})
        if "Pozvala sam" not in reakcije:
            reakcije.append("Pozvala sam")
        reakcije_ko["Pozvala sam"] = reakcije_ko.get("Pozvala sam") or (ko or "")
    elif kind == "ubacena":
        # MI smo porudžbinu ubacili kroz aplikaciju (dugme Naše/Njihove → admin).
        # Takav objekat NE ide u automatsko prepoznavanje načina trebovanja —
        # način se i dalje bira ručno (po našem / po njihovom).
        dnevnik.setdefault("ubaceno", []).append({"ko": ko or "", "at": _at,
                                                  "tip": str(kopija or "")[:16]})
        if "Ubačena porudžbina" not in reakcije:
            reakcije.append("Ubačena porudžbina")
        reakcije_ko["Ubačena porudžbina"] = reakcije_ko.get("Ubačena porudžbina") or (ko or "")
    elif kind == "komercijala":
        # prosleđeno komercijali — pamti se vreme da bi se videlo da li je objekat
        # trebovao POSLE prosleđivanja
        dnevnik.setdefault("komercijala", []).append({"ko": ko or "", "at": _at})
        if "Obavestila direktorku" not in reakcije:
            reakcije.append("Obavestila direktorku")
        reakcije_ko["Obavestila direktorku"] = reakcije_ko.get("Obavestila direktorku") or (ko or "")
    else:
        _unos = {"ko": ko or "", "at": _at}
        # da li je kopija upisana u folder Poslato (da se posle zna zašto je nema u sandučetu)
        if kopija is not None:
            try:
                _unos["kopija"] = bool(kopija[0])
                if not kopija[0]:
                    _unos["kopija_zasto"] = str(kopija[1])[:160]
            except Exception:
                pass
        dnevnik.setdefault("mejlovi", []).append(_unos)
        # uspelo je — skloni crvenu oznaku o ranijem neuspehu
        dnevnik.pop("greske", None)
        if "Poslala sam mejl" not in reakcije:
            reakcije.append("Poslala sam mejl")
        reakcije_ko["Poslala sam mejl"] = reakcije_ko.get("Poslala sam mejl") or (ko or "")
    _row = {"mesec": mesec_key, "sistem": sistem, "idk": int(idk),
            "reakcije": reakcije, "trebovali": bool(_r.get("trebovali_tip")),
            "trebovali_tip": _r.get("trebovali_tip") or "", "njihova": _r.get("njihova") or {},
            "napomena": _r.get("napomena") or "", "reakcije_ko": reakcije_ko, "dnevnik": dnevnik,
            "azurirao": ko or "", "azurirano": _at}
    try:
        cli.table("obrada").upsert(_row, on_conflict="mesec,sistem,idk").execute()
    except Exception:
        for _c in ("dnevnik", "reakcije_ko", "azurirao"):
            _row.pop(_c, None)
        cli.table("obrada").upsert(_row, on_conflict="mesec,sistem,idk").execute()


def sb_oznaci_trebovao(mesec_key, sistem, idk, kada="", n_por=0,
                       tip=None, njihova=None, info=None):
    """Trajno zabeleži da je objekat POSLAO PORUDŽBINU posle starta.

    Zove se pri ažuriranju iz admina. Piše u bazu, pa status ostaje i kad se
    izađe iz aplikacije i uđe ponovo bez novog ažuriranja.

    tip / njihova / info: automatski prepoznat način trebovanja (po našem ili po
    njihovom, na osnovu količina) i stvarno poručene količine po artiklu. Kada su
    prosleđeni, upisuju se u bazu i NE menjaju se ručno — objekat je završen."""
    cli = _sb()
    if cli is None:
        return False
    try:
        res = (cli.table("obrada")
               .select("reakcije,trebovali_tip,njihova,napomena,reakcije_ko,dnevnik")
               .eq("mesec", mesec_key).eq("sistem", sistem).eq("idk", int(idk)).limit(1).execute())
        _r = res.data[0] if res.data else {}
        _tip = str(_r.get("trebovali_tip") or "")
        _dn = dict(_r.get("dnevnik") or {})
        _vec = _dn.get("trebovao_posle_starta") or {}
        _stara_njih = dict(_r.get("njihova") or {})
        _nov_tip = str(tip or "") or (_tip or "nas")
        _nova_njih = ({str(k): int(v) for k, v in (njihova or {}).items()}
                      if njihova is not None else _stara_njih)
        # Ništa novo — ne diraj bazu (da se pri svakom ažuriranju ne prepisuje bez potrebe).
        # Ako se promenilo PRAVILO po kom se prepoznaje način trebovanja, upis ide i
        # kad su tip i količine isti (da zapis dobije novu oznaku pravila).
        _p_st = int(((_vec.get("info") or {}).get("pravilo") or 0))
        _p_nov = int(((info or {}).get("pravilo") or 0))
        if _vec and _tip == _nov_tip and _stara_njih == _nova_njih and _p_st >= _p_nov:
            return False
        _dn["trebovao_posle_starta"] = {"at": (_vec.get("at") or _now().isoformat()),
                                        "kada": str(kada or "")[:32],
                                        "porudzbina": int(n_por or 0),
                                        "tip": _nov_tip,
                                        "auto": bool(tip),
                                        "info": dict(info or {})}
        _row = {"mesec": mesec_key, "sistem": sistem, "idk": int(idk),
                "reakcije": list(_r.get("reakcije") or []),
                "trebovali": True,
                "trebovali_tip": _nov_tip,
                "njihova": _nova_njih,
                "napomena": _r.get("napomena") or "",
                "reakcije_ko": dict(_r.get("reakcije_ko") or {}),
                "dnevnik": _dn, "azurirano": _now().isoformat()}
        cli.table("obrada").upsert(_row, on_conflict="mesec,sistem,idk").execute()
        return True
    except Exception:
        return False


def sb_mejl_greska(mesec_key, sistem, idk, poruka, ko=""):
    """Zapiši NEUSPELO slanje mejla u dnevnik (dnevnik.greske = lista {ko, at, sta}).

    Važno: ovo se upisuje odmah, u bazu, pa ostaje zapisano i ako se strana
    „izgubi“ usred grupnog slanja. Ne dira reakcije — objekat NE dobija oznaku
    „Poslala sam mejl“, jer mejl nije ni otišao."""
    cli = _sb()
    if cli is None:
        return False
    try:
        res = cli.table("obrada").select("dnevnik").eq("mesec", mesec_key).eq("sistem", sistem).eq("idk", int(idk)).limit(1).execute()
        _r = res.data[0] if res.data else {}
        dnevnik = dict(_r.get("dnevnik") or {})
        _lst = list(dnevnik.get("greske") or [])
        _lst.append({"ko": ko or "", "at": _now().isoformat(), "sta": str(poruka)[:300]})
        dnevnik["greske"] = _lst[-20:]
        if res.data:
            cli.table("obrada").update({"dnevnik": dnevnik}).eq("mesec", mesec_key).eq("sistem", sistem).eq("idk", int(idk)).execute()
        else:
            cli.table("obrada").insert({"mesec": mesec_key, "sistem": sistem, "idk": int(idk),
                                        "dnevnik": dnevnik}).execute()
        return True
    except Exception:
        return False


def sb_vraceno_upisi(mesec_key, sistem, idk, podaci, postojeci_dnevnik=None):
    """Trajno zapamti da se mejl objektu VRATIO (dnevnik.vraceno).

    Zašto: odbijenice se čitaju iz sandučeta, pa čim se poruka obriše iz Outlooka
    oznaka „VRAĆENO“ nestane iz aplikacije. Ovako se podatak čuva u bazi i ostaje
    i kad se sanduče očisti."""
    cli = _sb()
    if cli is None:
        return False
    try:
        if postojeci_dnevnik is None:
            res = cli.table("obrada").select("dnevnik").eq("mesec", mesec_key).eq("sistem", sistem).eq("idk", int(idk)).limit(1).execute()
            _ima = bool(res.data)
            dnevnik = dict((res.data[0].get("dnevnik") or {}) if res.data else {})
        else:
            _ima = True
            dnevnik = dict(postojeci_dnevnik or {})
        dnevnik["vraceno"] = {"adresa": str(podaci.get("adresa", ""))[:120],
                              "kada": str(podaci.get("kada", ""))[:32],
                              "razlog": str(podaci.get("razlog", ""))[:200],
                              "zabelezeno": _now().isoformat()}
        if _ima:
            cli.table("obrada").update({"dnevnik": dnevnik}).eq("mesec", mesec_key).eq("sistem", sistem).eq("idk", int(idk)).execute()
        else:
            cli.table("obrada").insert({"mesec": mesec_key, "sistem": sistem, "idk": int(idk),
                                        "dnevnik": dnevnik}).execute()
        return True
    except Exception:
        return False


def _vraceno_vazi(vr, mejlovi, trenutna_adresa=""):
    """Da li oznaku „VRAĆENO“ treba i dalje prikazivati.

    Ne treba u dva slučaja:
      1) adresa objekta je u međuvremenu ispravljena (odbijenica se odnosi na staru),
      2) mejl je posle te odbijenice ponovo poslat i ovaj put se NIJE vratio.
    Tolerancija od 6 sati postoji zato što odbijenica nosi vreme servera koji ju je
    poslao, pa sat ne mora biti isti kao naš."""
    if not vr:
        return False
    import datetime as _dt
    _a = _ocisti_mejl(vr.get("adresa", "")).lower()
    _t = _ocisti_mejl(trenutna_adresa or "").lower()
    if _a and _t and _a != _t:
        return False
    try:
        _b = _dt.datetime.strptime(str(vr.get("kada") or ""), "%d.%m.%Y %H:%M")
    except Exception:
        return True
    _last = None
    for _m in (mejlovi or []):
        try:
            _d = _dt.datetime.fromisoformat(str(_m.get("at") or ""))
        except Exception:
            continue
        if _last is None or _d > _last:
            _last = _d
    if _last is None:
        return True
    return not (_last > _b + _dt.timedelta(hours=6))


def _dnevnik_lista_html(dnevnik, kind):
    """Sitni sivi kosi tekst: lista poziva/mejlova (ko + datum + vreme)."""
    _arr = (dnevnik or {}).get("pozivi" if kind == "poziv" else "mejlovi") or []
    if not _arr:
        return ""
    _rows = []
    for _i, _e in enumerate(_arr, 1):
        _ko = _ko_kratko(_e.get("ko", "")) or "?"
        _kd = str(_e.get("ko", "")) or "?"
        _tm = _dt_kratko(_e.get("at", ""))
        _pref = (str(_i) + ". poziv — ") if kind == "poziv" else ""
        _kop = ""
        if kind != "poziv" and _e.get("kopija") is False:
            _kop = (' <span style="color:#b45309;">· nije upisan u Poslato ('
                    + str(_e.get("kopija_zasto", ""))[:70] + ')</span>')
        _rows.append('<div style="font-size:11px;color:#9ca3af;font-style:italic;">' + _pref
                     + _h_escape(_kd) + " · " + _h_escape(_tm) + _kop + "</div>")
    return "".join(_rows)

def sb_save_obrada(mesec_key, sistem, idk, reakcije, trebovali_tip, njihova=None, napomena="", reakcije_ko=None, azurirao=""):
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    _row = {"mesec": mesec_key, "sistem": sistem, "idk": int(idk),
            "reakcije": list(reakcije), "trebovali": bool(trebovali_tip), "trebovali_tip": trebovali_tip or "",
            "njihova": dict(njihova or {}), "napomena": napomena or "",
            "reakcije_ko": dict(reakcije_ko or {}), "azurirao": azurirao or "",
            "azurirano": _now().isoformat()}
    try:
        cli.table("obrada").upsert(_row, on_conflict="mesec,sistem,idk").execute()
    except Exception:
        # fallback ako kolone reakcije_ko/azurirao ne postoje
        _row.pop("reakcije_ko", None)
        _row.pop("azurirao", None)
        cli.table("obrada").upsert(_row, on_conflict="mesec,sistem,idk").execute()

def sb_bulk_ubaci(mesec_key, sistem, ids, ko=""):
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    _at = _now().isoformat()
    _kod = {"Ubačena porudžbina": ko} if ko else {}
    rows = [{"mesec": mesec_key, "sistem": sistem, "idk": int(i),
             "reakcije": ["Ubačena porudžbina"], "trebovali": True, "trebovali_tip": "nas",
             "njihova": {}, "napomena": "", "reakcije_ko": dict(_kod), "azurirao": ko or "",
             "azurirano": _at} for i in ids]
    if rows:
        try:
            cli.table("obrada").upsert(rows, on_conflict="mesec,sistem,idk").execute()
        except Exception:
            for _r in rows:
                _r.pop("reakcije_ko", None); _r.pop("azurirao", None)
            cli.table("obrada").upsert(rows, on_conflict="mesec,sistem,idk").execute()

def sb_bulk_reset(mesec_key, sistem):
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    cli.table("obrada").delete().eq("mesec", mesec_key).eq("sistem", sistem).execute()

def sb_load_plan(mesec_key):
    cli = _sb()
    if cli is None:
        return None
    try:
        res = cli.table("plan_objave").select("datum").eq("mesec", mesec_key).limit(1).execute()
    except Exception:
        return None
    if not res.data:
        return None
    return res.data[0].get("datum")

def sb_save_plan(mesec_key, datum):
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    cli.table("plan_objave").upsert({"mesec": mesec_key, "datum": datum,
        "azurirano": _now().isoformat()}, on_conflict="mesec").execute()

def sb_plan_meseci():
    cli = _sb()
    if cli is None:
        return []
    try:
        res = cli.table("plan_objave").select("mesec").execute()
    except Exception:
        return []
    return [r["mesec"] for r in (res.data or [])]

# ---- Rokovi (postavlja direktor; 3 posebna: administracija / sistemi / prodaja) ----
@st.cache_data(ttl=30)
def sb_rokovi_all():
    cli = _sb()
    if cli is None:
        return {}
    try:
        res = cli.table("rokovi").select("mesec,rok_admin,rok_sistemi,rok_prodaja,rok_syx,rok_potraz,rok_kontrola,napomena").execute()
        return {r["mesec"]: r for r in (res.data or [])}
    except Exception:
        # kolone rok_syx/rok_potraz možda još ne postoje -> učitaj bez njih
        try:
            res = cli.table("rokovi").select("mesec,rok_admin,rok_sistemi,rok_prodaja,napomena").execute()
            return {r["mesec"]: r for r in (res.data or [])}
        except Exception:
            return {}

def sb_rokovi_get(mesec_key):
    try:
        return sb_rokovi_all().get(mesec_key, {}) or {}
    except Exception:
        return {}

def sb_rokovi_set(mesec_key, rok_admin, rok_sistemi, rok_prodaja, napomena, rok_syx=None, rok_potraz=None, rok_kontrola=None):
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    payload = {"mesec": mesec_key,
               "rok_admin": rok_admin or None, "rok_sistemi": rok_sistemi or None,
               "rok_prodaja": rok_prodaja or None, "rok_syx": rok_syx or None,
               "rok_potraz": rok_potraz or None, "rok_kontrola": rok_kontrola or None, "napomena": napomena or "",
               "azurirano": _now().isoformat()}
    try:
        cli.table("rokovi").upsert(payload, on_conflict="mesec").execute()
    except Exception:
        # fallback ako rok_syx/rok_potraz kolone ne postoje
        payload.pop("rok_syx", None)
        payload.pop("rok_potraz", None)
        payload.pop("rok_kontrola", None)
        cli.table("rokovi").upsert(payload, on_conflict="mesec").execute()
    try:
        sb_rokovi_all.clear()
    except Exception:
        pass

def _rok_fmt(s):
    """'YYYY-MM-DD' -> 'DD.MM.YYYY' (ili prazno)."""
    if not s:
        return ""
    try:
        return datetime.date.fromisoformat(str(s)[:10]).strftime("%d.%m.%Y")
    except Exception:
        return str(s)

def _dt_fmt(s):
    """ISO datetime -> 'DD.MM.YYYY u HH:MM' (ili original ako ne uspe)."""
    if not s:
        return ""
    try:
        return datetime.datetime.fromisoformat(str(s)).strftime("%d.%m.%Y u %H:%M")
    except Exception:
        return str(s)

def _rok_je_prosao(s):
    """True ako je rok (YYYY-MM-DD) prošao (danas je posle roka)."""
    if not s:
        return False
    try:
        return datetime.date.today() > datetime.date.fromisoformat(str(s)[:10])
    except Exception:
        return False

# ---- Izveštaj SYX (Word dokument po mesecu; analitičar ubacuje, direktori preuzimaju) ----
@st.cache_data(ttl=30)
def sb_syx_list():
    cli = _sb()
    if cli is None:
        return []
    try:
        res = cli.table("izvestaj_syx").select("mesec,filename,azurirano").execute()
        return sorted(res.data or [], key=lambda r: r.get("mesec", ""), reverse=True)
    except Exception:
        return []

@st.cache_data(ttl=300)
def sb_syx_get(mesec_key):
    cli = _sb()
    if cli is None:
        return None
    try:
        res = cli.table("izvestaj_syx").select("filename,docx_b64").eq("mesec", mesec_key).limit(1).execute()
        if res.data:
            return res.data[0]
    except Exception:
        pass
    return None

def sb_syx_set(mesec_key, filename, b64):
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    cli.table("izvestaj_syx").upsert({"mesec": mesec_key, "filename": filename, "docx_b64": b64,
        "azurirano": _now().isoformat()}, on_conflict="mesec").execute()
    for fn in (sb_syx_list, sb_syx_get):
        try:
            fn.clear()
        except Exception:
            pass

def sb_syx_obrisi(mesec_key):
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    cli.table("izvestaj_syx").delete().eq("mesec", mesec_key).execute()
    for fn in (sb_syx_list, sb_syx_get):
        try:
            fn.clear()
        except Exception:
            pass

# ---- Izveštaj potraživanja (analitičar uploaduje Excel, administracija dopunjava u aplikaciji, direktor vidi + izvozi) ----
def _potraz_txt(v):
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%d.%m.%Y")
    return str(v)

def _potraz_tip(naziv):
    n = " ".join(str(naziv).lower().split())
    if "status" in n and "komunik" in n:
        return "dd_a"
    if "status" in n and ("tužb" in n or "tuzb" in n):
        return "dd_b"
    if any(k in n for k in ["vrednost", "ukupni dug", "za uplatu", "dana", "iznos"]):
        return "num"
    return "text"

def _potraz_num(v):
    if isinstance(v, (int, float)):
        return float(v)
    try:
        s = str(v).strip().replace(" ", "").replace(" ", "")
        if s == "":
            return None
        if s == "—":
            return None
        if "," in s:                    # zarez = decimala; tačke = hiljade
            s = s.replace(".", "").replace(",", ".")
        elif s.count(".") > 1:          # 1.000.000 -> tačke su hiljade
            s = s.replace(".", "")
        return float(s)
    except Exception:
        return None

def _potraz_fmt_num(v):
    """Broj -> '1 879 376,60' (razmak hiljade, zarez decimala); prazno -> ''."""
    if v is None:
        return ""
    try:
        return "{:,.2f}".format(float(v)).replace(",", " ").replace(".", ",")
    except Exception:
        return str(v)

def potraz_parse(xlsx_bytes):
    """Rasčlani Excel potraživanja u strukturu (listovi -> sekcije -> kolone/redovi) sa
    koordinatama ćelija, da bi administracija mogla da dopunjava u aplikaciji, a izvoz
    upisuje nazad u originalni fajl (identično formatiranje)."""
    import openpyxl as _ox
    wb = _ox.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    dd_a, dd_b = [], []
    if "_LISTE" in wb.sheetnames:
        for r in wb["_LISTE"].iter_rows(values_only=True):
            if r and len(r) >= 1 and r[0]:
                dd_a.append(_potraz_txt(r[0]))
            if r and len(r) >= 2 and r[1]:
                dd_b.append(_potraz_txt(r[1]))
    out = {"stanje_na_dan": "", "dd_a": dd_a, "dd_b": dd_b, "listovi": []}
    stanje = ""
    for ws in wb.worksheets:
        if ws.title == "_LISTE":
            continue
        rows = list(ws.iter_rows(values_only=False))
        n = len(rows)
        for r in rows[:6]:
            for c in r:
                if c.value and "Stanje na dan" in str(c.value):
                    stanje = str(c.value).replace("Stanje na dan:", "").strip()
        sekcije = []
        i = 0
        last = ""
        while i < n:
            rc = rows[i]
            vals = [_potraz_txt(c.value).strip() for c in rc]
            for m in ("PO FAKTURI", "PO ODJAVI"):
                if any(v == m for v in vals):
                    last = m
            is_hdr = any(v == "Komitent" for v in vals) or any("Naziv komitenta" in v for v in vals)
            if is_hdr:
                kolone = []
                for c in rc:
                    nz = _potraz_txt(c.value).strip()
                    if nz:
                        kolone.append({"col": c.column, "naziv": " ".join(nz.split()), "tip": _potraz_tip(nz)})
                redovi = []
                ukupno_row = None
                sum_cols = []
                j = i + 1
                while j < n:
                    r2 = rows[j]
                    a = _potraz_txt(r2[0].value).strip()
                    if a.upper().startswith("UKUPNO"):
                        ukupno_row = r2[0].row
                        for c in r2:
                            if isinstance(c.value, (int, float)):
                                sum_cols.append(c.column)
                        break
                    if all(_potraz_txt(c.value).strip() == "" for c in r2):
                        break
                    cells = {}
                    for k in kolone:
                        cells[str(k["col"])] = _potraz_txt(r2[k["col"] - 1].value)
                    redovi.append({"r": r2[0].row, "cells": cells})
                    j += 1
                sekcije.append({"naslov": last or ws.title, "kolone": kolone, "redovi": redovi,
                                "ukupno_row": ukupno_row, "sum_cols": sum_cols})
                i = j
                last = ""
                continue
            i += 1
        out["listovi"].append({"sheet": ws.title, "sekcije": sekcije})
    out["stanje_na_dan"] = stanje
    return out

def potraz_init_popuna(struktura):
    """Početna popuna = vrednosti iz uploadovanog fajla (administracija ih dalje menja)."""
    pop = {}
    for L in struktura.get("listovi", []):
        sh = L["sheet"]
        pop[sh] = {}
        for s in L["sekcije"]:
            for rr in s["redovi"]:
                pop[sh].setdefault(str(rr["r"]), {})
                for col, val in rr["cells"].items():
                    pop[sh][str(rr["r"])][str(col)] = val
    return pop

def potraz_export(original_b64, struktura, popuna):
    """Upiši trenutne vrednosti (popuna) u originalni Excel i vrati bytes (identično formatiranje).
    Tip kolone se određuje PO SEKCIJI (ista kolona može biti različita u „po fakturi" i „po odjavi")."""
    import openpyxl as _ox, base64 as _b64
    wb = _ox.load_workbook(io.BytesIO(_b64.b64decode(original_b64)))
    popuna = popuna or {}
    for L in struktura.get("listovi", []):
        sh = L["sheet"]
        if sh not in wb.sheetnames:
            continue
        ws = wb[sh]
        for s in L["sekcije"]:
            coltip = {str(k["col"]): k["tip"] for k in s["kolone"]}
            for rr in s["redovi"]:
                rv = popuna.get(sh, {}).get(str(rr["r"]), {})
                for col, tip in coltip.items():
                    if col not in rv:
                        continue
                    val = rv.get(col)
                    c = ws.cell(row=int(rr["r"]), column=int(col))
                    if tip == "num":
                        nv = _potraz_num(val)
                        c.value = nv if nv is not None else None
                    else:
                        c.value = val if (val is not None and str(val) != "") else None
            if s.get("ukupno_row") and s.get("sum_cols"):
                for scol in s["sum_cols"]:
                    tot = 0.0
                    for rr in s["redovi"]:
                        nv = _potraz_num(popuna.get(sh, {}).get(str(rr["r"]), {}).get(str(scol)))
                        if nv is not None:
                            tot += nv
                    ws.cell(row=int(s["ukupno_row"]), column=int(scol)).value = tot
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

def _potraz_eff_kolone(sekcija):
    """Kolone za prikaz — bez 'Dana kašnjenja' (izbačeno na zahtev)."""
    out = []
    for k in sekcija["kolone"]:
        nl = " ".join(str(k["naziv"]).lower().split())
        if "dana kašnjenja" in nl or "dana kasnjenja" in nl:
            continue
        out.append(k)
    return out

def _potraz_role_cols(kolone):
    """Vrati col-indekse (str) za: ukupni dug, vrednost lagera, za uplatu (ako postoje)."""
    c_dug = c_lager = c_zau = None
    for k in kolone:
        nl = " ".join(str(k["naziv"]).lower().split())
        if "ukupni dug" in nl:
            c_dug = str(k["col"])
        elif "vrednost lagera" in nl or ("lager" in nl and "vrednost" in nl):
            c_lager = str(k["col"])
        elif "za uplatu" in nl:
            c_zau = str(k["col"])
    return c_dug, c_lager, c_zau

def potraz_section_df(sekcija, sheet, popuna):
    """DataFrame za sekciju + labele + redosled kolona + labele koje su read-only (računate)."""
    kolone = _potraz_eff_kolone(sekcija)
    labels = []
    counts = {}
    col_order = []
    for k in kolone:
        base = k["naziv"]
        c = counts.get(base, 0) + 1
        counts[base] = c
        labels.append(base if c == 1 else base + " (" + str(c) + ")")
        col_order.append(str(k["col"]))
    # prvo izračunaj vrednosti (brojevi) po ćeliji, pa Za uplatu, pa formatiraj za prikaz
    c_dug, c_lager, c_zau = _potraz_role_cols(kolone)
    disabled = []
    data = {lab: [] for lab in labels}
    for rr in sekcija["redovi"]:
        rv = popuna.get(sheet, {}).get(str(rr["r"]), {})
        # izračunaj Za uplatu = Ukupni dug − Vrednost lagera (ako oba postoje)
        _zau_val = None
        if c_zau and c_dug:
            _dug = _potraz_num(rv.get(c_dug, rr["cells"].get(c_dug, "")))
            _lag = _potraz_num(rv.get(c_lager, rr["cells"].get(c_lager, ""))) if c_lager else None
            _zau_val = (_dug - (_lag or 0)) if _dug is not None else None
        for idx, k in enumerate(kolone):
            col = str(k["col"])
            val = rv.get(col, rr["cells"].get(col, ""))
            if c_zau and col == c_zau:
                data[labels[idx]].append(_potraz_fmt_num(_zau_val))   # računato, tekst
            elif k["tip"] == "num":
                data[labels[idx]].append(_potraz_fmt_num(_potraz_num(val)))
            elif k["tip"] in ("dd_a", "dd_b"):
                sv = "" if val is None else str(val).strip()
                data[labels[idx]].append(sv if sv not in ("", "—") else "—")  # „—" = prazno
            else:
                data[labels[idx]].append("" if val is None else str(val))
    df = pd.DataFrame(data, columns=labels)
    if c_zau:
        disabled.append(labels[col_order.index(c_zau)])
    return df, labels, col_order, disabled

def potraz_col_config(sekcija, labels, df, dd_a, dd_b):
    cfg = {}
    kolone = _potraz_eff_kolone(sekcija)
    for idx, k in enumerate(kolone):
        lab = labels[idx]
        tip = k["tip"]
        if tip in ("dd_a", "dd_b"):
            base = list(dd_a if tip == "dd_a" else dd_b)
            try:
                existing = [str(x) for x in df[lab].dropna().unique() if str(x).strip() not in ("", "—")]
            except Exception:
                existing = []
            opts = ["—"] + base + [e for e in existing if e not in base]
            cfg[lab] = st.column_config.SelectboxColumn(lab, options=opts, required=False, width="medium")
        elif tip == "num":
            cfg[lab] = st.column_config.TextColumn(lab, help="Iznos u RSD (npr. 1 000 000 ili 1000000).")
        elif "komentar" in lab.lower():
            cfg[lab] = st.column_config.TextColumn(lab, width="large")
        else:
            cfg[lab] = st.column_config.TextColumn(lab)
    return cfg

@st.cache_data(ttl=30)
def sb_potraz_list():
    cli = _sb()
    if cli is None:
        return []
    try:
        res = cli.table("izvestaj_potrazivanja").select("mesec,naziv,predato,azurirano").execute()
        return sorted(res.data or [], key=lambda r: r.get("mesec", ""), reverse=True)
    except Exception:
        return []

@st.cache_data(ttl=120)
def sb_potraz_get(mesec_key):
    cli = _sb()
    if cli is None:
        return None
    try:
        res = cli.table("izvestaj_potrazivanja").select("mesec,naziv,original_b64,struktura,popuna,predato,predato_at,azurirano").eq("mesec", mesec_key).limit(1).execute()
        if res.data:
            return res.data[0]
    except Exception:
        pass
    return None

def sb_potraz_set(mesec_key, naziv, original_b64, struktura_json, popuna_json):
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    cli.table("izvestaj_potrazivanja").upsert({"mesec": mesec_key, "naziv": naziv,
        "original_b64": original_b64, "struktura": struktura_json, "popuna": popuna_json,
        "predato": False, "predato_at": None,
        "azurirano": _now().isoformat()}, on_conflict="mesec").execute()
    for fn in (sb_potraz_list, sb_potraz_get):
        try:
            fn.clear()
        except Exception:
            pass

def sb_potraz_popuni(mesec_key, popuna_json, predato=False):
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    _upd = {"popuna": popuna_json, "azurirano": _now().isoformat()}
    if predato:
        _upd["predato"] = True
        _upd["predato_at"] = _now().strftime("%d.%m.%Y %H:%M")
    cli.table("izvestaj_potrazivanja").update(_upd).eq("mesec", mesec_key).execute()
    for fn in (sb_potraz_list, sb_potraz_get):
        try:
            fn.clear()
        except Exception:
            pass

def sb_potraz_obrisi(mesec_key):
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    cli.table("izvestaj_potrazivanja").delete().eq("mesec", mesec_key).execute()
    for fn in (sb_potraz_list, sb_potraz_get):
        try:
            fn.clear()
        except Exception:
            pass

def sb_potraz_reopen(mesec_key):
    """Vrati predati izveštaj potraživanja na dopunu (otključaj)."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    cli.table("izvestaj_potrazivanja").update({"predato": False, "predato_at": None,
        "azurirano": _now().isoformat()}).eq("mesec", mesec_key).execute()
    for fn in (sb_potraz_list, sb_potraz_get):
        try:
            fn.clear()
        except Exception:
            pass

def _potraz_collect(edited, base=None):
    """Iz izmenjenih data_editor tabela sklopi popunu {sheet: {r: {col: val}}}, čuvajući
    sve ostale ćelije iz `base`. Brojevi se parsiraju, „—" -> prazno, Za uplatu = dug − lager."""
    newpop = json.loads(json.dumps(base)) if base else {}
    for (sh, sidx), (ed, col_order, sek) in edited.items():
        newpop.setdefault(sh, {})
        eff = _potraz_eff_kolone(sek)
        tip_of = {str(k["col"]): k["tip"] for k in eff}
        c_dug, c_lager, c_zau = _potraz_role_cols(eff)
        for ri, rr in enumerate(sek["redovi"]):
            newpop[sh].setdefault(str(rr["r"]), {})
            for ci, col in enumerate(col_order):
                if c_zau and col == c_zau:
                    continue  # računa se ispod
                try:
                    val = ed.iloc[ri, ci]
                except Exception:
                    val = ""
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    val = ""
                if tip_of.get(col) == "num":
                    nv = _potraz_num(val)
                    newpop[sh][str(rr["r"])][col] = (round(nv, 2) if nv is not None else "")
                elif tip_of.get(col) in ("dd_a", "dd_b"):
                    sv = str(val).strip()
                    newpop[sh][str(rr["r"])][col] = ("" if sv in ("", "—") else sv)
                else:
                    newpop[sh][str(rr["r"])][col] = val
            # Za uplatu = Ukupni dug − Vrednost lagera (računato)
            if c_zau and c_dug:
                _dug = _potraz_num(newpop[sh][str(rr["r"])].get(c_dug))
                _lag = _potraz_num(newpop[sh][str(rr["r"])].get(c_lager)) if c_lager else None
                newpop[sh][str(rr["r"])][c_zau] = (round(_dug - (_lag or 0), 2) if _dug is not None else "")
    return newpop

def _potraz_norm_work(struct, pop):
    """Normalizuj popunu za radnu kopiju: brojevi zaokruženi na 2 decimale, „—" -> '',
    Za uplatu preračunata. Time round-trip prikaz<->čuvanje ostaje stabilan (bez petlje)."""
    w = json.loads(json.dumps(pop or {}))
    for L in struct.get("listovi", []):
        sh = L["sheet"]
        w.setdefault(sh, {})
        for s in L["sekcije"]:
            eff = _potraz_eff_kolone(s)
            tip_of = {str(k["col"]): k["tip"] for k in eff}
            c_dug, c_lager, c_zau = _potraz_role_cols(eff)
            for rr in s["redovi"]:
                row = w[sh].setdefault(str(rr["r"]), {})
                for col, tp in tip_of.items():
                    cur = row.get(col, rr["cells"].get(col, ""))
                    if c_zau and col == c_zau:
                        continue
                    if tp == "num":
                        nv = _potraz_num(cur)
                        row[col] = (round(nv, 2) if nv is not None else "")
                    elif tp in ("dd_a", "dd_b"):
                        sv = str(cur or "").strip()
                        row[col] = "" if sv in ("", "—") else sv
                    else:
                        row[col] = "" if cur is None else str(cur)
                if c_zau and c_dug:
                    _d = _potraz_num(row.get(c_dug))
                    _l = _potraz_num(row.get(c_lager)) if c_lager else None
                    row[c_zau] = (round(_d - (_l or 0), 2) if _d is not None else "")
    return w

# ============================ IZVEŠTAJ KNEZ PETROL ============================
# Sve pumpe iz šifarnika komitenata; grupno slanje mejla kojim se traži stanje
# zaliha na kraju izabranog meseca i prodaja za taj mesec. Nema zona ni
# ograničenja koliko puta se šalje — samo se broji koliko je puta poslato.

KNEZ_SIS = "KNEZ PETROL — IZVEŠTAJ"     # ključ pod kojim se pamti dnevnik slanja


def _knez_meseci(n=18):
    """Poslednjih n meseci (ključ 'YYYY-MM'), najnoviji prvi."""
    _d = datetime.date.today().replace(day=1)
    _out = []
    for _ in range(int(n)):
        _out.append(_d.strftime("%Y-%m"))
        _d = (_d - datetime.timedelta(days=1)).replace(day=1)
    return _out


def _knez_period(mesec_key):
    """Vrati (prvi_dan, poslednji_dan) izabranog meseca kao 'DD.MM.YYYY.'"""
    _y = int(str(mesec_key).split("-")[0]); _m = int(str(mesec_key).split("-")[1])
    _prvi = datetime.date(_y, _m, 1)
    if _m == 12:
        _posl = datetime.date(_y, 12, 31)
    else:
        _posl = datetime.date(_y, _m + 1, 1) - datetime.timedelta(days=1)
    return (_prvi.strftime("%d.%m.%Y."), _posl.strftime("%d.%m.%Y."))


KNEZ_TEKST_DEFAULT = (
    "Poštovani,\n\n"
    "molimo Vas da nam pošaljete:\n\n"
    "•  stanje zaliha na dan {do}\n"
    "•  prodaju u periodu od {od} do {do}\n\n"
    "za objekat {objekat}.\n\n"
    "Podatke možete poslati kao odgovor na ovaj mejl (Excel ili tabela u poruci).\n\n"
    "Hvala unapred.\n\n"
    "Srdačan pozdrav,")


def _knez_tekst(sablon, mesec_key, naziv=""):
    _od, _do = _knez_period(mesec_key)
    _t = str(sablon or "")
    return (_t.replace("{od}", _od).replace("{do}", _do)
            .replace("{mesec}", mesec_label(mesec_key))
            .replace("{objekat}", str(naziv or "")))


def knez_admin_ui():
    st.markdown("<div style='font-size:18px;font-weight:800;margin:4px 0 10px;'>"
                "⛽ Izveštaj Knez Petrol</div>", unsafe_allow_html=True)
    if not sb_dostupan():
        st.error("Veza sa bazom trenutno nije podešena. Javi se analitičaru.")
        return

    _mk_opts = _knez_meseci(18)
    _mk_lbls = [mesec_label(k) for k in _mk_opts]
    _kc1, _kc2, _kc3 = st.columns([1.2, 2, 2])
    with _kc1:
        _sel_lbl_k = st.selectbox("Mesec izveštaja", _mk_lbls,
                                  index=(1 if len(_mk_lbls) > 1 else 0), key="knez_mes")
    mesec_key = _mk_opts[_mk_lbls.index(_sel_lbl_k)]
    _od, _do = _knez_period(mesec_key)
    with _kc2:
        _filt = st.text_input("Naziv komitenta POČINJE sa",
                              value=str(_cfg("KNEZ_FILTER", "KNEZ PETROL")), key="knez_filt",
                              help="Uzimaju se SAMO komitenti čiji naziv počinje ovim tekstom "
                                   "(ne oni koji ga imaju negde u nazivu, npr. u adresi).")
    with _kc3:
        st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
        st.caption("Traži se stanje zaliha na **" + _do + "** i prodaja **"
                   + _od.rstrip(".") + " – " + _do + "**")

    # --- Pumpe iz šifarnika ---
    if st.session_state.get("_komfull") is None:
        st.session_state["_komfull"] = sb_komitenti_full()
    _kom = st.session_state.get("_komfull") or {}
    import re as _rek

    def _norm_nz(_s):
        """mala slova, bez viška razmaka i vodećih znakova — za poređenje početka naziva"""
        return _rek.sub(r"\s+", " ", str(_s or "")).strip().lstrip("-–—·.,").strip().lower()

    _q = _norm_nz(_filt)
    _pumpe = []
    for _idk, _inf in (_kom or {}).items():
        _nz = str((_inf or {}).get("naziv", "") or "")
        # SAMO oni čiji naziv POČINJE zadatim tekstom
        if _q and not _norm_nz(_nz).startswith(_q):
            continue
        _pumpe.append({"idk": int(_idk), "naziv": _nz or ("ID " + str(_idk)),
                       "email": str((_inf or {}).get("email", "") or "").strip(),
                       "mesto": str((_inf or {}).get("mesto", "") or ""),
                       "telefon": str((_inf or {}).get("telefon", "") or "")})
    _pumpe.sort(key=lambda r: r["naziv"])
    if not _pumpe:
        st.warning("U šifarniku nema komitenata čiji naziv POČINJE sa „" + str(_filt) + "“. "
                   "Proveri tekst ili učitaj šifarnik komitenata.")
        return

    # --- Dnevnik slanja (pamti se u bazi, po mesecu) ---
    _obr = sb_load_obrada(mesec_key, KNEZ_SIS)
    for r in _pumpe:
        _v = _obr.get(int(r["idk"])) or {}
        _mj = (_v.get("dnevnik") or {}).get("mejlovi") or []
        r["mail_n"] = len(_mj)
        r["mail_at"] = (_mj[-1].get("at") if _mj else None)
        r["mail_ko"] = (_mj[-1].get("ko", "") if _mj else "")
        r["_email_ok"] = _mejl_ok(r["email"])
    # --- Odgovori iz sandučeta (čita se na klik) ---
    _odg_k = "_knez_odg_" + str(mesec_key)
    _odg_ses = st.session_state.get(_odg_k) or {}
    _odg_live = (_odg_ses.get("po") or {})
    _nepr = (_odg_ses.get("nep") or [])
    for r in _pumpe:
        _v = _obr.get(int(r["idk"])) or {}
        _sac = ((_v.get("dnevnik") or {}).get("odgovori") or [])
        _liv = _odg_live.get((r["email"] or "").lower(), [])
        r["odg_sac"] = _sac
        r["odg_live"] = _liv
        r["odg_n"] = max(len(_sac), len(_liv))
    _n_uk = len(_pumpe)
    _n_mail = sum(1 for r in _pumpe if r["_email_ok"])
    _n_pos = sum(1 for r in _pumpe if r["mail_n"] > 0)
    _n_ost = _n_uk - _n_pos
    _n_odg = sum(1 for r in _pumpe if r["odg_n"] > 0)

    st.markdown('<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin:6px 0 10px;">'
                '<div style="background:#faf7ff;border:1px solid #e9d5ff;border-radius:12px;padding:15px 18px;">'
                '<div style="font-size:22px;font-weight:800;color:#7c3aed;">' + str(_n_uk) + '</div>'
                '<div style="font-size:12px;color:#8b7fa8;margin-top:3px;">Pumpi u šifarniku</div></div>'
                '<div style="background:#f6fdf9;border:1px solid #bbf7d0;border-radius:12px;padding:15px 18px;">'
                '<div style="font-size:22px;font-weight:800;color:#158a3f;">' + str(_n_mail) + '</div>'
                '<div style="font-size:12px;color:#6f9a80;margin-top:3px;">Sa ispravnim mejlom</div></div>'
                '<div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:12px;padding:15px 18px;">'
                '<div style="font-size:22px;font-weight:800;color:#1e40af;">' + str(_n_pos) + '</div>'
                '<div style="font-size:12px;color:#6b82ad;margin-top:3px;">Poslat mejl (' + _sel_lbl_k + ')</div></div>'
                '<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:12px;padding:15px 18px;">'
                '<div style="font-size:22px;font-weight:800;color:#b45309;">' + str(_n_ost) + '</div>'
                '<div style="font-size:12px;color:#9a7b3a;margin-top:3px;">Još nije poslato</div></div>'
                '<div style="background:#dcfce7;border:1px solid #86efac;border-radius:12px;padding:15px 18px;">'
                '<div style="font-size:22px;font-weight:800;color:#14532d;">' + str(_n_odg) + '</div>'
                '<div style="font-size:12px;color:#166534;margin-top:3px;font-weight:700;">📥 Odgovorili</div></div></div>',
                unsafe_allow_html=True)

    # --- Provera odgovora u sandučetu ---
    # Čita se od 1. u TEKUĆEM mesecu (ili od prvog poslatog mejla, ako je raniji) —
    # odgovor ne može da stigne pre nego što smo poslali zahtev.
    _od_def = _now().date().replace(day=1)
    try:
        _prvi_mejl = None
        for r in _pumpe:
            for _m0 in (((_obr.get(int(r["idk"])) or {}).get("dnevnik") or {}).get("mejlovi") or []):
                _d0 = str(_m0.get("at", ""))[:10]
                if len(_d0) == 10 and (_prvi_mejl is None or _d0 < _prvi_mejl):
                    _prvi_mejl = _d0
        if _prvi_mejl:
            _dp = datetime.date.fromisoformat(_prvi_mejl)
            if _dp < _od_def:
                _od_def = _dp
    except Exception:
        pass
    _oc1, _ocd, _oc2 = st.columns([1.5, 1.3, 3])
    with _oc1:
        st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
        _chk_odg = st.button("📥 Proveri odgovore", key="knez_scan", use_container_width=True,
                             help="Čita sanduče i vezuje odgovore za pumpe po adresi pošiljaoca.")
    with _ocd:
        _od_dat = st.date_input("Čitaj poruke od", value=_od_def, key="knez_od_dat",
                                format="DD.MM.YYYY",
                                help="Podrazumevano 1. u tekućem mesecu (ili od prvog poslatog "
                                     "mejla). Što kraći period — to brža provera.")
    with _oc2:
        if _odg_ses.get("kada"):
            st.caption("📥 Poslednja provera: " + str(_odg_ses.get("kada"))
                       + " · pronađeno odgovora: " + str(sum(len(v) for v in _odg_live.values()))
                       + " · neprepoznatih: " + str(len(_nepr))
                       + ". Prilozi se preuzimaju iz sandučeta — dostupni su do sledeće provere.")
        else:
            st.caption("Klikni „Proveri odgovore“ da se iz sandučeta povuku odgovori pumpi "
                       "(tekst i prilozi) i prikažu ispod svake pumpe.")
    if _chk_odg:
        _adr = set((r["email"] or "").lower() for r in _pumpe if r["_email_ok"])
        import time as _tm0
        _t0 = _tm0.time()
        with st.spinner("📥 Čitam sanduče od " + _od_dat.strftime("%d.%m.%Y") + "…"):
            _po, _nep, _err = knez_odgovori(_adr, od_datum=_od_dat)
        _trajalo = round(_tm0.time() - _t0, 1)
        if _err and not _po:
            st.error("Čitanje sandučeta nije uspelo: " + str(_err))
        _n_up = 0
        for _a, _lst in (_po or {}).items():
            _ik = next((r["idk"] for r in _pumpe if (r["email"] or "").lower() == _a), None)
            if _ik is None:
                continue
            for _z in _lst:
                try:
                    if sb_knez_odgovor_set(mesec_key, _ik, _z):
                        _n_up += 1
                except Exception:
                    pass
        st.session_state[_odg_k] = {"po": _po, "nep": _nep,
                                    "kada": _now().strftime("%d.%m.%Y. %H:%M")}
        st.session_state["_knez_scan_flash"] = (
            "📥 Pronađeno " + str(sum(len(v) for v in (_po or {}).values())) + " odgovora od "
            + str(len(_po or {})) + " pumpi" + ((" · novih zabeleženo: " + str(_n_up)) if _n_up else "")
            + ((" · " + str(len(_nep)) + " poruka nije prepoznato (vidi dole)") if _nep else "")
            + "  ·  provera je trajala " + str(_trajalo) + " s")
        st.rerun()
    _sf = st.session_state.pop("_knez_scan_flash", None)
    if _sf:
        st.success(_sf)

    # --- Naslov i tekst mejla (isti za pojedinačno i za grupno slanje) ---
    if not smtp_dostupan():
        _n = _mail_nalog()
        _suf = ("_" + _n) if _n else ""
        st.warning("✉️ Slanje mejlova nije podešeno za tvoj nalog — dodaj u Secrets: "
                   "SMTP_HOST" + _suf + " / SMTP_USER" + _suf + " / SMTP_PASSWORD" + _suf + ".")
    with st.expander("✉️ Naslov i tekst mejla" + (("  ·  šalje se sa " + str(_smtp_cfg().get("from_email", "")))
                                                  if smtp_dostupan() else ""), expanded=False):
        _subj_k = st.text_input(
            "Naslov mejla", key="knez_subj_" + str(mesec_key),
            value=("VAPE SHOP - Zahtev za izveštaj o prodaji i zalihama - " + _sel_lbl_k))
        _tekst_k = st.text_area(
            "Tekst mejla", key="knez_body_" + str(mesec_key), height=210,
            value=str(_cfg("KNEZ_MEJL_TEKST", KNEZ_TEKST_DEFAULT)),
            help="{objekat} = naziv pumpe · {od} = prvi dan meseca · {do} = poslednji dan meseca "
                 "· {mesec} = naziv meseca.")
        _prva = next((r for r in _pumpe if r["_email_ok"]), _pumpe[0])
        st.caption("Ovako izgleda za „" + str(_prva["naziv"])[:44] + "“:")
        st.code("Za: " + (_prva["email"] or "—") + "\nNaslov: " + str(_subj_k) + "\n\n"
                + _knez_tekst(_tekst_k, mesec_key, _prva["naziv"]))

    def _knez_posalji(_r):
        """Pošalji mejl jednoj pumpi i zabeleži u dnevnik (ko + kada)."""
        _ko1 = st.session_state.get("admin_user", "Administracija")
        _sk1 = "knez_sent_" + str(mesec_key) + "_" + str(_r["idk"])
        try:
            with st.spinner("✉️ Šaljem mejl…"):
                posalji_mejl_sa_prilogom(_r["email"], str(_subj_k),
                                         _knez_tekst(_tekst_k, mesec_key, _r["naziv"]))
            try:
                sb_obrada_log(mesec_key, KNEZ_SIS, _r["idk"], "mejl", _ko1,
                              kopija=st.session_state.get("_zadnja_kopija"))
            except Exception:
                pass
            st.session_state[_sk1] = {"ok": True, "msg": "Poslato na " + _r["email"]}
        except Exception as _e1:
            st.session_state[_sk1] = {"ok": False, "msg": str(_e1)}
            try:
                sb_mejl_greska(mesec_key, KNEZ_SIS, _r["idk"], str(_e1), _ko1)
            except Exception:
                pass

    def _knez_poziv(_idk, _tel):
        """Klik na slušalicu = zabeležen poziv (ko + kada) + pokreće pozivanje."""
        try:
            sb_obrada_log(mesec_key, KNEZ_SIS, int(_idk), "poziv",
                          st.session_state.get("admin_user", "Administracija"))
        except Exception:
            pass
        st.session_state["_knez_dial_" + str(_idk)] = _tel

    _KT = ["Pregled", "📧 Grupno slanje mejlova"]
    try:
        _kt1, _kt2 = st.tabs(_KT, key="knez_tabs", on_change="rerun")
    except TypeError:
        _kt1, _kt2 = st.tabs(_KT)

    with _kt1:
        _pc1, _pc2 = st.columns([2.2, 1.4])
        with _pc1:
            _pq = st.text_input("Pretraga (naziv, mesto ili mejl)", key="knez_pq",
                                placeholder="npr. Ćuprija")
        with _pc2:
            _pf = st.selectbox("Prikaži", ["Sve", "Nije poslato", "Već poslato",
                                           "📥 Odgovorili", "Nisu odgovorili", "Bez mejla"],
                               key="knez_pf")

        def _pok(r):
            if _pf == "Nije poslato" and r["mail_n"] > 0:
                return False
            if _pf == "Već poslato" and r["mail_n"] == 0:
                return False
            if _pf == "📥 Odgovorili" and not r["odg_n"]:
                return False
            if _pf == "Nisu odgovorili" and (r["odg_n"] or r["mail_n"] == 0):
                return False
            if _pf == "Bez mejla" and r["_email_ok"]:
                return False
            if _pq.strip():
                _qq = _pq.strip().lower()
                if (_qq not in r["naziv"].lower() and _qq not in r["email"].lower()
                        and _qq not in str(r["mesto"]).lower()):
                    return False
            return True
        _vidljive = [r for r in _pumpe if _pok(r)]
        st.caption("Prikazano: " + str(len(_vidljive)) + " od " + str(len(_pumpe))
                   + " pumpi.  ·  👆 Klikni na pumpu da se ispod otvori slanje mejla i poziv.")

        for r in _vidljive:
            _idk = int(r["idk"])
            _v = _obr.get(_idk) or {}
            _dnv = _v.get("dnevnik") or {}
            _pozivi = _dnv.get("pozivi") or []
            _mejlovi = _dnv.get("mejlovi") or []
            _greske = _dnv.get("greske") or []
            _hdr = (("📥 " if r["odg_n"] else ("✅ " if _mejlovi else
                                              ("⚠️ " if not r["_email_ok"] else "✉️ ")))
                    + str(r["naziv"]))
            if r["mesto"]:
                _hdr += "   ·   " + str(r["mesto"])
            if _mejlovi:
                _hdr += "   ·   poslato " + str(len(_mejlovi)) + "×"
            if _pozivi:
                _hdr += "   ·   📞 " + str(len(_pozivi)) + " poziv" + ("a" if len(_pozivi) > 1 else "")
            if r["odg_n"]:
                _hdr += "   ·   📥 ODGOVORILI"
            with st.expander(_hdr, expanded=False):
                _kb = []
                if r["email"]:
                    _kb.append("✉️ " + _h_escape(r["email"]))
                if r["telefon"]:
                    _kb.append("📞 " + _h_escape(r["telefon"]))
                _kb.append("ID " + str(_idk))
                st.markdown('<div style="color:#6b7280;font-size:12.5px;margin:0 0 8px;">'
                            + "&nbsp;&nbsp;·&nbsp;&nbsp;".join(_kb) + '</div>',
                            unsafe_allow_html=True)

                # --- Slušalica: klik = zabeležen poziv ---
                _tel_raw = str(r["telefon"] or "").strip()
                import re as _ret
                _tel_c = _ret.sub(r"[^\d+]", "", _tel_raw)
                if _tel_c.count("+") > 1:
                    _tel_c = "+" + _tel_c.replace("+", "")
                if _tel_c.startswith("00"):
                    _tel_c = "+" + _tel_c[2:]
                elif _tel_c.startswith("0"):
                    _tel_c = "+381" + _tel_c[1:]
                _b1, _b2, _b3 = st.columns([1.5, 1.5, 2])
                with _b1:
                    st.button("📞 Pozovi " + (_tel_raw or "—"),
                              key="knez_call_" + str(mesec_key) + "_" + str(_idk),
                              use_container_width=True, disabled=(not _tel_c),
                              on_click=_knez_poziv, args=(_idk, _tel_c),
                              help="Klik se automatski beleži kao poziv (ko i kada).")
                with _b2:
                    st.button("✉️ Pošalji mejl",
                              key="knez_send1_" + str(mesec_key) + "_" + str(_idk),
                              type="primary", use_container_width=True,
                              disabled=(not r["_email_ok"] or not smtp_dostupan()),
                              on_click=_knez_posalji, args=(r,))
                with _b3:
                    if not r["_email_ok"]:
                        st.caption("Nema ispravan mejl u šifarniku.")
                    elif not smtp_dostupan():
                        st.caption("Slanje mejlova nije podešeno.")

                if st.session_state.get("_knez_dial_" + str(_idk)):
                    _dn = st.session_state.pop("_knez_dial_" + str(_idk))
                    components.html("<script>window.location.href='tel:" + _dn + "';</script>", height=0)
                    st.markdown('<a href="tel:' + _h_escape(_dn) + '" style="display:inline-block;'
                                'background:#16a34a;color:#fff;text-decoration:none;font-weight:700;'
                                'font-size:13px;padding:8px 16px;border-radius:9px;">'
                                '📞 Ako se poziv ne otvori sam — klikni</a>', unsafe_allow_html=True)

                _res1 = st.session_state.get("knez_sent_" + str(mesec_key) + "_" + str(_idk))
                if _res1:
                    (st.success if _res1["ok"] else st.error)(
                        ("✅ " if _res1["ok"] else "❌ ") + str(_res1["msg"]))

                # --- Trajna oznaka: mejl je poslat + ko i kada ---
                if _mejlovi:
                    _z = _mejlovi[-1]
                    st.markdown('<div style="background:#dcfce7;border:1px solid #86efac;'
                                'border-radius:9px;padding:8px 13px;margin:6px 0 2px;'
                                'font-size:13px;color:#14532d;font-weight:700;">'
                                '✅ Mejl je poslat ' + str(len(_mejlovi)) + '× · poslednji put '
                                + _h_escape(_dt_kratko(_z.get("at", "")))
                                + (" · " + _h_escape(_ko_kratko(_z.get("ko", ""))) if _z.get("ko") else "")
                                + '</div>' + _dnevnik_lista_html(_dnv, "mejl"),
                                unsafe_allow_html=True)
                else:
                    st.caption("✉️ Mejl još nije poslat za " + _sel_lbl_k + ".")

                if _pozivi:
                    st.markdown('<div style="font-size:12px;color:#6b7280;font-weight:600;'
                                'margin:8px 0 1px;">📞 Pozvano ' + str(len(_pozivi)) + '× :</div>'
                                + _dnevnik_lista_html(_dnv, "poziv"), unsafe_allow_html=True)

                if _greske:
                    _g = _greske[-1]
                    st.error("❌ Poslednji pokušaj slanja NIJE uspeo — "
                             + _dt_kratko(_g.get("at", "")) + " · " + str(_g.get("sta", ""))[:200])

                # --- ODGOVOR PUMPE (iz sandučeta) ---
                _odg_prikaz = r["odg_live"] or r["odg_sac"]
                if _odg_prikaz:
                    st.markdown('<div style="background:#eef2ff;border:2px solid #818cf8;'
                                'border-radius:10px;padding:10px 14px;margin:10px 0 4px;">'
                                '<div style="font-size:14px;font-weight:900;color:#312e81;">'
                                '📥 ODGOVORILI — ' + str(len(_odg_prikaz)) + ' poruka</div></div>',
                                unsafe_allow_html=True)
                    for _oi, _o in enumerate(_odg_prikaz):
                        _ko_o = (str(_o.get("ime", "")) or str(_o.get("od", "")))
                        st.markdown('<div style="font-size:13px;font-weight:700;color:#3730a3;'
                                    'margin:6px 0 2px;">✉️ ' + _h_escape(_ko_o) + ' &lt;'
                                    + _h_escape(str(_o.get("od", ""))) + '&gt;  ·  '
                                    + _h_escape(str(_o.get("at", ""))) + '</div>'
                                    '<div style="font-size:12.5px;color:#4b5563;font-weight:600;">'
                                    + _h_escape(str(_o.get("naslov", ""))) + '</div>',
                                    unsafe_allow_html=True)
                        if _o.get("tekst"):
                            st.markdown('<div style="background:#f8fafc;border:1px solid #e5e7eb;'
                                        'border-radius:8px;padding:8px 12px;margin:4px 0 6px;'
                                        'font-size:12.5px;color:#374151;white-space:pre-wrap;">'
                                        + _h_escape(str(_o.get("tekst", ""))) + '</div>',
                                        unsafe_allow_html=True)
                        for _pi, _pz in enumerate(_o.get("prilozi") or []):
                            _vel = int(_pz.get("vel", 0) or 0)
                            _vt = (str(round(_vel / 1024.0)) + " KB" if _vel < 1024 * 1024
                                   else str(round(_vel / 1048576.0, 1)) + " MB")
                            if _pz.get("data"):
                                st.download_button(
                                    "⬇️ " + str(_pz.get("ime", "prilog")) + "  (" + _vt + ")",
                                    _pz["data"], file_name=str(_pz.get("ime", "prilog")),
                                    key=("knez_dl_" + str(mesec_key) + "_" + str(_idk) + "_"
                                         + str(_oi) + "_" + str(_pi)))
                            else:
                                st.caption("📎 " + str(_pz.get("ime", "prilog")) + " (" + _vt
                                           + ") — klikni „📥 Proveri odgovore“ gore da se prilog "
                                             "povuče iz sandučeta.")
                    if not r["odg_live"] and r["odg_sac"]:
                        st.caption("ℹ️ Prikazano iz zabeleženog. Za preuzimanje priloga klikni "
                                   "„📥 Proveri odgovore“ gore.")
                elif _mejlovi:
                    st.caption("📥 Još nema odgovora od ove pumpe.")

        # --- Odgovori koje aplikacija nije prepoznala (drugi mejl pošiljaoca) ---
        if _nepr:
            st.markdown("<hr style='margin:16px 0 8px;border:none;border-top:1px solid #e5e7eb;'>",
                        unsafe_allow_html=True)
            with st.expander("❓ Neprepoznati odgovori (" + str(len(_nepr))
                             + ") — pošiljalac nije u šifarniku, dodeli ih ručno", expanded=False):
                st.caption("Ovo su poruke iz sandučeta koje nisu stigle sa adrese neke pumpe "
                           "(npr. odgovorili su sa lične adrese). Izaberi pumpu i klikni Dodeli.")
                _opts = ["—"] + [str(r["idk"]) + " · " + r["naziv"] for r in _pumpe]
                for _ni, _z in enumerate(_nepr[:30]):
                    st.markdown('<div style="font-size:13px;font-weight:700;color:#374151;'
                                'margin:8px 0 2px;">✉️ ' + _h_escape(str(_z.get("ime", "")) or "")
                                + ' &lt;' + _h_escape(str(_z.get("od", ""))) + '&gt;  ·  '
                                + _h_escape(str(_z.get("at", ""))) + '</div>'
                                '<div style="font-size:12.5px;color:#6b7280;">'
                                + _h_escape(str(_z.get("naslov", ""))) + '</div>',
                                unsafe_allow_html=True)
                    if _z.get("tekst"):
                        st.caption(str(_z.get("tekst", ""))[:220])
                    for _pi2, _pz2 in enumerate(_z.get("prilozi") or []):
                        if _pz2.get("data"):
                            st.download_button("⬇️ " + str(_pz2.get("ime", "prilog")),
                                               _pz2["data"], file_name=str(_pz2.get("ime", "prilog")),
                                               key="knez_ndl_" + str(_ni) + "_" + str(_pi2))
                    _dc1, _dc2 = st.columns([3, 1])
                    with _dc1:
                        _pick = st.selectbox("Dodeli pumpi", _opts, key="knez_nas_" + str(_ni),
                                             label_visibility="collapsed")
                    with _dc2:
                        if st.button("✓ Dodeli", key="knez_nasb_" + str(_ni),
                                     use_container_width=True, disabled=(_pick == "—")):
                            try:
                                _tid = int(str(_pick).split("·")[0].strip())
                                if sb_knez_odgovor_set(mesec_key, _tid, _z):
                                    st.success("Dodeljeno pumpi " + str(_tid) + ".")
                                # prikaži ga odmah i pod tom pumpom
                                _em = next((p["email"] for p in _pumpe if p["idk"] == _tid), "")
                                if _em:
                                    _cur = st.session_state.get(_odg_k) or {"po": {}, "nep": []}
                                    _cur.setdefault("po", {}).setdefault(_em.lower(), []).append(_z)
                                    _cur["nep"] = [x for x in (_cur.get("nep") or []) if x is not _z]
                                    st.session_state[_odg_k] = _cur
                                st.rerun()
                            except Exception as _de:
                                st.error("Nije uspelo: " + str(_de))
                    st.markdown("<hr style='margin:6px 0;border:none;border-top:1px dashed #e5e7eb;'>",
                                unsafe_allow_html=True)

        _bez = [r for r in _pumpe if not r["_email_ok"]]
        if _bez:
            st.warning("✉️ " + str(len(_bez)) + " pumpi nema ispravan mejl u šifarniku — "
                       "njima se ne može poslati (ispravi u šifarniku komitenata): "
                       + ", ".join(str(r["naziv"])[:38] for r in _bez[:10])
                       + ("…" if len(_bez) > 10 else ""))
        try:
            import io as _io
            _pdf = pd.DataFrame([{
                "ID": r["idk"], "Naziv": r["naziv"], "Mesto": r["mesto"],
                "Email": r["email"] or "—", "Poslato ×": int(r["mail_n"]),
                "Poslednji put": (_dt_kratko(r["mail_at"]) if r["mail_at"] else ""),
                "Ko": _ko_kratko(r["mail_ko"]) if r["mail_ko"] else "",
                "Poziva": len(((_obr.get(int(r["idk"])) or {}).get("dnevnik") or {}).get("pozivi") or []),
                "Odgovorili": ("DA" if r["odg_n"] else "—"),
                "Odgovor stigao": ((r["odg_live"] or r["odg_sac"])[-1].get("at", "")
                                   if r["odg_n"] else ""),
            } for r in _pumpe])
            _buf = _io.BytesIO()
            with pd.ExcelWriter(_buf, engine="openpyxl") as _w:
                _pdf.to_excel(_w, index=False, sheet_name="Knez Petrol")
            st.download_button("⬇️ Izvezi pregled u Excel", _buf.getvalue(),
                               file_name="Knez_Petrol_" + str(mesec_key) + ".xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="knez_xlsx")
        except Exception as _ke:
            st.caption("Izvoz trenutno nije moguć: " + str(_ke))

    with _kt2:
        if smtp_dostupan():
            st.caption("✉️ Mejlovi se šalju sa: " + str(_smtp_cfg().get("from_email", "")))
        st.caption("Naslov i tekst menjaš gore, u „✉️ Naslov i tekst mejla“.")

        _selk = "knez_sel_" + str(mesec_key)
        _verk = "knez_ver_" + str(mesec_key)
        if _selk not in st.session_state:
            st.session_state[_selk] = set()
        if _verk not in st.session_state:
            st.session_state[_verk] = 0

        _fc1, _fc2 = st.columns([1.4, 2])
        with _fc1:
            _f_st = st.selectbox("Status slanja", ["Sve", "Nije poslato", "Već poslato"],
                                 key="knez_f_st")
        with _fc2:
            _f_q = st.text_input("Pretraga (naziv ili mejl)", key="knez_f_q", placeholder="npr. Novi Sad")

        def _ok(r):
            if _f_st == "Nije poslato" and r["mail_n"] > 0:
                return False
            if _f_st == "Već poslato" and r["mail_n"] == 0:
                return False
            if _f_q.strip():
                _qq = _f_q.strip().lower()
                if _qq not in r["naziv"].lower() and _qq not in r["email"].lower():
                    return False
            return True
        _view = [r for r in _pumpe if _ok(r)]

        _sig = str(_f_st) + "|" + _f_q.strip().lower() + "|" + str(mesec_key)
        _sigk = "knez_sig"
        if st.session_state.get(_sigk) != _sig:
            st.session_state[_sigk] = _sig
            st.session_state[_selk] = set()
            st.session_state[_verk] += 1

        _a1, _a2, _a3 = st.columns([1.3, 1.3, 3])
        with _a1:
            if st.button("☑️ Izaberi sve (filtrirane)", key="knez_all", use_container_width=True):
                for r in _view:
                    if r["_email_ok"]:
                        st.session_state[_selk].add(r["idk"])
                st.session_state[_verk] += 1
                st.rerun()
        with _a2:
            if st.button("✖️ Poništi izbor", key="knez_none", use_container_width=True):
                st.session_state[_selk] = set()
                st.session_state[_verk] += 1
                st.rerun()

        if not _view:
            st.info("Nema pumpi za izabrane filtere.")
            return
        _edf = pd.DataFrame([{
            "Izabrano": r["idk"] in st.session_state[_selk],
            "Naziv": r["naziv"],
            "Email": r["email"] or "—",
            "Poslato ×": int(r["mail_n"]),
            "Poslednji put": r["mail_at"],
        } for r in _view])
        _edf.index = [int(r["idk"]) for r in _view]
        _edf["Poslednji put"] = pd.to_datetime(_edf["Poslednji put"], errors="coerce")
        _ed = st.data_editor(
            _edf, hide_index=True, use_container_width=True,
            height=min(60 + 36 * len(_view), 700),
            key="knez_editor_" + str(mesec_key) + "_" + str(st.session_state[_verk]),
            disabled=["Naziv", "Email", "Poslato ×", "Poslednji put"],
            column_config={
                "Izabrano": st.column_config.CheckboxColumn("Izabrano", width="small"),
                "Poslato ×": st.column_config.NumberColumn("Poslato ×", width="small"),
                "Poslednji put": st.column_config.DatetimeColumn(
                    "Poslednji put", format="DD.MM.YYYY. HH:mm")})
        _new = set()
        for _ix, _rw in _ed.iterrows():
            if bool(_rw["Izabrano"]):
                try:
                    _new.add(int(_ix))
                except Exception:
                    pass
        st.session_state[_selk] = _new
        _sel_ids = st.session_state[_selk] & set(r["idk"] for r in _view)
        _spremni = [r for r in _view if r["idk"] in _sel_ids and r["_email_ok"]]
        _bez_m = sum(1 for r in _view if r["idk"] in _sel_ids and not r["_email_ok"])
        st.caption("Izabrano: " + str(len(_sel_ids)) + " · spremno za slanje: " + str(len(_spremni))
                   + ((" · bez ispravnog mejla: " + str(_bez_m)) if _bez_m else ""))

        _fl = st.session_state.pop("_knez_flash", None)
        if _fl:
            (st.success if _fl[0] == "ok" else st.warning)(_fl[1])

        if st.button("📧 Pošalji izabranima (" + str(len(_spremni)) + ")", key="knez_send",
                     type="primary", use_container_width=True,
                     disabled=(len(_spremni) == 0 or not smtp_dostupan())):
            try:
                _bmax = int(str(_cfg("GRUPNO_BATCH", "8")).strip() or 0)
            except Exception:
                _bmax = 8
            _sada = _spremni[:_bmax] if _bmax > 0 else _spremni
            try:
                _pauza = float(str(_cfg("GRUPNO_PAUZA", "2")).strip() or 0)
            except Exception:
                _pauza = 2.0
            _ko = st.session_state.get("admin_user", "Administracija")
            _prog = st.progress(0, "✉️ Pripremam slanje…")
            _ok_n = _fail_n = 0
            _stop = ""
            _ses = MejlSesija()
            for _i, r in enumerate(_sada):
                _prog.progress(int(_i / len(_sada) * 100),
                               "✉️ Šaljem " + str(_i + 1) + "/" + str(len(_sada))
                               + " — " + str(r["naziv"])[:40] + " …")
                try:
                    posalji_mejl_sa_prilogom(
                        r["email"], str(_subj_k),
                        _knez_tekst(_tekst_k, mesec_key, r["naziv"]), sesija=_ses)
                    _ok_n += 1
                    try:
                        st.session_state[_selk].discard(r["idk"])
                    except Exception:
                        pass
                    try:
                        sb_obrada_log(mesec_key, KNEZ_SIS, r["idk"], "mejl", _ko,
                                      kopija=st.session_state.get("_zadnja_kopija"))
                    except Exception:
                        pass
                except MejlOgranicenje as _mo:
                    _fail_n += 1
                    _stop = str(_mo)
                    try:
                        sb_mejl_greska(mesec_key, KNEZ_SIS, r["idk"], str(_mo), _ko)
                    except Exception:
                        pass
                    break
                except Exception as _me:
                    _fail_n += 1
                    try:
                        sb_mejl_greska(mesec_key, KNEZ_SIS, r["idk"], str(_me), _ko)
                    except Exception:
                        pass
                _prog.progress(int((_i + 1) / len(_sada) * 100),
                               "✅ Poslato " + str(_i + 1) + "/" + str(len(_sada)))
                if _pauza > 0 and _i < len(_sada) - 1:
                    import time as _ts
                    _ts.sleep(_pauza)
            try:
                _ses.zatvori()
            except Exception:
                pass
            _prog.empty()
            _por = ["✅ Poslato " + str(_ok_n) + " mejlova."] if _fail_n == 0 else [
                "Poslato " + str(_ok_n) + " · nije uspelo " + str(_fail_n) + "."]
            _preostalo = len(_spremni) - _ok_n - _fail_n
            if _stop:
                _por.append("⏸️ Zaustavljeno jer nas server privremeno koči: " + _stop
                            + "  Sačekaj pa klikni ponovo — poslate pumpe su već odštiklirane.")
            elif _preostalo > 0:
                _por.append("Ostalo je još " + str(_preostalo)
                            + " — i dalje su štiklirane, samo klikni ponovo.")
            st.session_state["_knez_flash"] = (("ok" if _fail_n == 0 and not _stop else "warn"),
                                               "\n\n".join(_por))
            st.session_state[_verk] += 1
            st.rerun()


def potraz_admin_ui():
    st.markdown("<div style='font-size:18px;font-weight:800;margin:4px 0 10px;'>💳 Izveštaj potraživanja</div>", unsafe_allow_html=True)
    _lst = sb_potraz_list()
    if not _lst:
        st.info("Analitičar još nije objavio nijedan izveštaj potraživanja.")
        return
    _labels = [mesec_label(r["mesec"]) for r in _lst]
    _keys = [r["mesec"] for r in _lst]
    _sel = st.selectbox("Mesec", _labels, index=0, key="pz_adm_mes")
    _mk = _keys[_labels.index(_sel)]
    rec = sb_potraz_get(_mk)
    if not rec:
        st.info("Nema podataka za ovaj mesec.")
        return
    try:
        struct = json.loads(rec.get("struktura") or "{}")
        pop = json.loads(rec.get("popuna") or "{}")
    except Exception:
        st.error("Greška u podacima izveštaja.")
        return
    st.caption("📄 Fajl: " + str(rec.get("naziv", "") or "—") + "  ·  poslednje ažuriranje: "
               + (_dt_fmt(rec.get("azurirano")) or "—"))
    _predato = bool(rec.get("predato"))
    _rok_pz = sb_rokovi_get(_mk).get("rok_potraz")
    if _rok_pz and not _predato:
        if _rok_je_prosao(_rok_pz):
            st.warning("⏰ Rok za predaju potraživanja (" + _rok_fmt(_rok_pz) + ") je istekao.")
        else:
            st.info("⏰ Rok za predaju potraživanja: " + _rok_fmt(_rok_pz) + ".")
    if _predato:
        _pc_a, _pc_b = st.columns([3, 1])
        with _pc_a:
            st.success("Ovaj izveštaj je predat direktoru (" + str(rec.get("predato_at") or "") + "). Prikaz je samo za pregled.")
        with _pc_b:
            if st.button("🔓 Vrati na dopunu", key="pz_reopen", use_container_width=True):
                try:
                    sb_potraz_reopen(_mk)
                    st.session_state.pop("pzwork_for", None)  # osveži radnu kopiju
                    st.success("Otključano — možeš ponovo da dopunjavaš.")
                    st.rerun()
                except Exception as _e:
                    st.error("Greška: " + str(_e))
    else:
        st.caption("Stanje na dan: " + str(struct.get("stanje_na_dan", "")) + ". Dopuni iznose, statuse i komentare, pa klikni Prosledi komercijali.")
    dd_a = struct.get("dd_a", [])
    dd_b = struct.get("dd_b", [])
    def _sek_boja(naslov):
        n = str(naslov).upper()
        if "FAKTUR" in n:
            return ("#16a34a", "#ecfdf5", "#bbf7d0")   # zelena — po fakturi
        if "ODJAV" in n:
            return ("#2563eb", "#eff6ff", "#bfdbfe")   # plava — po odjavi
        return ("#7c3aed", "#faf5ff", "#e9d5ff")       # ljubičasta — ostalo
    # Radna kopija (u sesiji) — da se Za uplatu računa odmah dok kucaju, pre čuvanja
    _wkey = "pzwork_" + _mk
    if st.session_state.get("pzwork_for") != _mk:
        st.session_state[_wkey] = _potraz_norm_work(struct, pop)
        st.session_state["pzwork_for"] = _mk
    _work = st.session_state.get(_wkey) or _potraz_norm_work(struct, pop)

    _has_zau = False
    edited = {}
    for L in struct.get("listovi", []):
        sh = L["sheet"]
        with st.container(border=True):
            st.markdown("<div style='display:inline-block;background:linear-gradient(135deg,#a855f7,#ec4899);color:#fff;"
                        "font-weight:800;font-size:14px;padding:5px 14px;border-radius:20px;margin:2px 0 6px;'>👤 "
                        + _h_escape(sh) + "</div>", unsafe_allow_html=True)
            for sidx, s in enumerate(L["sekcije"]):
                _cbc, _cbg, _cbb = _sek_boja(s["naslov"])
                st.markdown("<div style='display:inline-block;background:" + _cbg + ";color:" + _cbc + ";border:1px solid "
                            + _cbb + ";font-weight:700;font-size:12px;padding:3px 11px;border-radius:8px;margin:10px 0 4px;'>"
                            + _h_escape(str(s["naslov"])) + "</div>", unsafe_allow_html=True)
                df, labels, col_order, _disabled = potraz_section_df(s, sh, _work)
                if _disabled and not _predato and not _has_zau:
                    _has_zau = True
                    st.caption("ℹ️ Za uplatu se računa automatski (Ukupan dug − Vrednost lagera) čim upišete Ukupan dug.")
                cfg = potraz_col_config(s, labels, df, dd_a, dd_b)
                _dis = True if _predato else (_disabled or False)
                _ed = st.data_editor(df, column_config=cfg, hide_index=True, use_container_width=True,
                                     num_rows="fixed", key="pz_ed_" + _mk + "_" + sh + "_" + str(sidx),
                                     disabled=_dis)
                edited[(sh, sidx)] = (_ed, col_order, s)
    # Primeni izmene na radnu kopiju i preračunaj Za uplatu; ako ima promene -> osveži prikaz
    if not _predato:
        _newwork = _potraz_collect(edited, base=_work)
        if _newwork != _work:
            st.session_state[_wkey] = _newwork
            st.rerun()
        _b1, _b2 = st.columns(2)
        with _b1:
            if st.button("💾 Sačuvaj (bez prosleđivanja)", key="pz_adm_save", use_container_width=True):
                try:
                    sb_potraz_popuni(_mk, json.dumps(_work, default=str))
                    st.success("Sačuvano.")
                    st.rerun()
                except Exception as _e:
                    st.error("Greška pri čuvanju: " + str(_e))
        with _b2:
            if st.button("📨 Prosledi komercijali", key="pz_adm_send", use_container_width=True, type="primary"):
                try:
                    sb_potraz_popuni(_mk, json.dumps(_work, default=str), predato=True)
                    st.success("Prosleđeno komercijali.")
                    st.rerun()
                except Exception as _e:
                    st.error("Greška pri prosleđivanju: " + str(_e))

def potraz_director_ui():
    st.markdown('<div style="font-size:20px;font-weight:800;margin:6px 0 6px;">💳 Izveštaj potraživanja</div>', unsafe_allow_html=True)
    _lst = [r for r in sb_potraz_list() if r.get("predato")]
    if not _lst:
        st.info("Još nema prosleđenih izveštaja potraživanja. Administracija ih popunjava i prosleđuje.")
        return
    _labels = [mesec_label(r["mesec"]) for r in _lst]
    _keys = [r["mesec"] for r in _lst]
    _sel = st.selectbox("Mesec", _labels, index=0, key="pz_dir_mes")
    _mk = _keys[_labels.index(_sel)]
    rec = sb_potraz_get(_mk)
    if not rec:
        st.info("Nema podataka.")
        return
    try:
        struct = json.loads(rec.get("struktura") or "{}")
        pop = json.loads(rec.get("popuna") or "{}")
    except Exception:
        st.error("Greška u podacima.")
        return
    _tc1, _tc2 = st.columns([3, 1])
    with _tc1:
        st.caption("Stanje na dan: " + str(struct.get("stanje_na_dan", ""))
                   + "  ·  predato " + str(rec.get("predato_at") or "")
                   + "  ·  poslednje ažuriranje: " + (_dt_fmt(rec.get("azurirano")) or "—"))
    with _tc2:
        try:
            _xb = potraz_export(rec.get("original_b64"), struct, pop)
            st.download_button("⬇️ Izvezi u Excel", _xb,
                file_name="Izvestaj_potrazivanja_" + _mk + ".xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="pz_dir_dl", use_container_width=True)
        except Exception as _e:
            st.caption("Izvoz trenutno nije moguć.")
    def _sek_boja2(naslov):
        n = str(naslov).upper()
        if "FAKTUR" in n:
            return ("#16a34a", "#ecfdf5", "#bbf7d0")
        if "ODJAV" in n:
            return ("#2563eb", "#eff6ff", "#bfdbfe")
        return ("#7c3aed", "#faf5ff", "#e9d5ff")
    for L in struct.get("listovi", []):
        sh = L["sheet"]
        with st.container(border=True):
            st.markdown("<div style='display:inline-block;background:linear-gradient(135deg,#a855f7,#ec4899);color:#fff;"
                        "font-weight:800;font-size:14px;padding:5px 14px;border-radius:20px;margin:2px 0 6px;'>👤 "
                        + _h_escape(sh) + "</div>", unsafe_allow_html=True)
            for s in L["sekcije"]:
                _cbc, _cbg, _cbb = _sek_boja2(s["naslov"])
                st.markdown("<div style='display:inline-block;background:" + _cbg + ";color:" + _cbc + ";border:1px solid "
                            + _cbb + ";font-weight:700;font-size:12px;padding:3px 11px;border-radius:8px;margin:10px 0 4px;'>"
                            + _h_escape(str(s["naslov"])) + "</div>", unsafe_allow_html=True)
                df, labels, col_order, _disabled = potraz_section_df(s, sh, pop)
                # numeričke kolone -> lepo formatiranje sa hiljadama; boje zaglavlja
                try:
                    _sty = df.style.set_table_styles([{"selector": "th", "props": [("background-color", _cbg), ("color", _cbc)]}])
                    st.dataframe(_sty, hide_index=True, use_container_width=True)
                except Exception:
                    st.dataframe(df, hide_index=True, use_container_width=True)

# ---- HITNOST po objektu (na osnovu niskog lagera) ----
# Pragovi su namerno apsolutni i lako se menjaju (dole dve brojke).
HIT_CRVENO_KOM   = 15   # >= ovoliko kom/mesec na artiklima bez lagera -> HITNO
HIT_CRVENO_ART   = 8    # ILI >= ovoliko artikala na nuli -> HITNO
HIT_ZUTO_KOM     = 5
HIT_ZUTO_ART     = 3
def hitnost_objekta(stavke_obj):
    """stavke_obj = lista stavki za jedan objekat (mogu biti i artikli sa porudzbinom 0).
    Urgentnost racunamo SAMO na artiklima koje predlazemo za porudzbinu (kol > 0),
    da mrtvi artikli (bez prodaje, bez porudzbine) ne bi lazno dizali hitnost.
    izgubljeno = predvidjena mesecna prodaja na artiklima koji su TRENUTNO na 0 lagera."""
    _ord = [s for s in stavke_obj if int(s.get('kol', 0)) > 0]
    n_nula = sum(1 for s in _ord if int(s.get('lager', 0)) == 0)
    izgubljeno = sum(int(s.get('pred', 0)) for s in _ord if int(s.get('lager', 0)) == 0)
    if izgubljeno >= HIT_CRVENO_KOM or n_nula >= HIT_CRVENO_ART:
        nivo = 'crveno'
    elif izgubljeno >= HIT_ZUTO_KOM or n_nula >= HIT_ZUTO_ART:
        nivo = 'zuto'
    else:
        nivo = 'zeleno'
    return nivo, n_nula, izgubljeno

def hitnost_objekta_dodatna(stavke_obj, poruceno):
    """Kao hitnost_objekta, ali na osnovu DODATNE porudžbine — pošto su objekti već
    sami poručili (posle preseka), računamo hitnost na onome što STVARNO još treba:
    dodatna = preporuka − već poručeno (min 0), realni lager = lager + već poručeno."""
    def _por(s):
        try:
            return int(poruceno.get(int(s.get('ida', -1)), 0))
        except Exception:
            return 0
    _ord = [s for s in stavke_obj if max(int(s.get('kol', 0)) - _por(s), 0) > 0]
    n_nula = sum(1 for s in _ord if int(s.get('lager', 0)) + _por(s) <= 0)
    izgubljeno = sum(int(s.get('pred', 0)) for s in _ord if int(s.get('lager', 0)) + _por(s) <= 0)
    if izgubljeno >= HIT_CRVENO_KOM or n_nula >= HIT_CRVENO_ART:
        nivo = 'crveno'
    elif izgubljeno >= HIT_ZUTO_KOM or n_nula >= HIT_ZUTO_ART:
        nivo = 'zuto'
    else:
        nivo = 'zeleno'
    return nivo, n_nula, izgubljeno

HIT_EMOJI = {'crveno': '🔴', 'zuto': '🟡', 'zeleno': '🟢'}
HIT_RANG  = {'crveno': 0, 'zuto': 1, 'zeleno': 2}
HIT_TEKST = {'crveno': 'Hitno', 'zuto': 'Srednje', 'zeleno': 'Može da čeka'}

def stavke_iz_rezultata(result, engine):
    """Napravi listu stavki za cuvanje u Supabase.
    Cuvamo SVE artikle (i one sa porudzbinom 0) za objekte koji imaju bar
    jedan artikal za porudzbinu — da koleginice u detalju vide ceo asortiman
    objekta, a ne samo ono sto se poruci."""
    reg = engine.region_map if hasattr(engine, 'region_map') else {}
    _ordered_ids = set(result[result['Porudzbina_2'] > 0]['ID KOMITENTA'].tolist())
    order = result[result['ID KOMITENTA'].isin(_ordered_ids)]
    _mo = list(getattr(engine, 'meseci_order', []) or [])
    _pdict = getattr(engine, 'prodaja_dict', {}) or {}
    def _mesecna_prodaja(idk, ida):
        # niz mesečne prodaje artikla u tom objektu, poravnat sa engine.meseci_order
        _ser = []
        for (_g, _m) in _mo:
            _v = _pdict.get((idk, ida, _g, _m))
            try:
                _kol = int(round(float(_v[0]))) if (_v and _v[0] == _v[0]) else 0  # NaN-safe
            except Exception:
                _kol = 0
            _ser.append(max(_kol, 0))
        return _ser
    stavke = []
    for _, r in order.iterrows():
        idk = int(r['ID KOMITENTA'])
        _ida = int(r['id artikla'])
        stavke.append({
            'idk': idk,
            'region': str(reg.get(r['ID KOMITENTA'], '') or ''),
            'ida': _ida,
            'naziv': str(r['Naziv artikla']),
            'grupa': str(r.get('Grupa', '') or ''),
            'pred': int(r.get('Predikcija', 0)),
            'lager': int(r.get('Lager_danas', 0)),
            'kol': int(r['Porudzbina_2']),
            'prodaja_mesecno': _mesecna_prodaja(idk, _ida),
        })
    return stavke


def direktor_blok(engine, res):
    """Iz rezultata analitike spakuj podatke za DIREKTORSKI izveštaj (prodaja, trend,
    poređenja, po grupama, OOS). Bezbedno — nikad ne ruši objavu (sve u try/except)."""
    out = {}
    try:
        ml = list(engine.mesec_labels)
    except Exception:
        ml = []

    def _col(lb, suf):
        c = str(lb) + suf
        return c if c in res.columns else None

    # Trend prodaje po mesecu (kom + rsd)
    trend = []
    for lb in ml:
        cp = _col(lb, "_Prodaja")
        cr = _col(lb, "_Promet")
        if cp:
            try:
                trend.append({"mesec": lb, "kom": int(res[cp].sum()),
                              "rsd": int(res[cr].sum()) if cr else 0})
            except Exception:
                pass
    out["prodaja_trend"] = trend
    _last = ml[-1] if ml else None
    out["prodaja_tekuci"] = trend[-1] if trend else {"mesec": _last, "kom": 0, "rsd": 0}

    # Poređenja: prošli mesec, 6-mesečni prosek, isti mesec lani (ako ima podataka)
    comp = {}
    if len(trend) >= 2:
        comp["prosli_mesec"] = trend[-2]
        _prev6 = trend[-7:-1] if len(trend) >= 7 else trend[:-1]
        if _prev6:
            comp["prosek_6m"] = {"kom": int(round(sum(t["kom"] for t in _prev6) / len(_prev6))), "n": len(_prev6)}
    if _last and " " in str(_last):
        _mnaz, _god = str(_last).rsplit(" ", 1)
        if _god.isdigit():
            _lani = _mnaz + " " + str(int(_god) - 1)
            for t in trend:
                if t["mesec"] == _lani:
                    comp["isti_mesec_lani"] = t
                    break
    out["poredjenja"] = comp

    # Prodaja po grupama (tekući mesec)
    grupe = []
    _lc = _col(_last, "_Prodaja") if _last else None
    if _lc and ("Grupa" in res.columns):
        try:
            g = res.groupby("Grupa")[_lc].sum().sort_values(ascending=False)
            for naz, kom in g.items():
                grupe.append({"grupa": str(naz), "kom": int(kom)})
        except Exception:
            pass
    out["po_grupama"] = grupe

    # Out of stock (na osnovu lagera danas)
    try:
        _oos = res[(res["Lager_danas"] == 0) & (res["Predikcija"] > 0)]
        out["oos"] = {"kombinacija_na_0": int(len(_oos)), "izgubljeno_kom": int(_oos["Predikcija"].sum())}
        _pa = (_oos.groupby("Naziv artikla")
               .agg(objekata=("ID KOMITENTA", "nunique"), izgubljeno=("Predikcija", "sum"))
               .sort_values("izgubljeno", ascending=False).head(10))
        out["oos_po_artiklu"] = [{"artikal": str(i), "objekata": int(r["objekata"]),
                                  "izgubljeno": int(r["izgubljeno"])} for i, r in _pa.iterrows()]
    except Exception:
        pass

    # ---- Prosečna prodaja po objektu (mesečno) ----
    try:
        _ppo = []
        for lb in ml:
            cp = _col(lb, "_Prodaja")
            if cp:
                _tot = float(res[cp].sum())
                _no = int((res[cp] > 0).sum())
                _ppo.append({"mesec": lb, "prosek": round(_tot / _no, 1) if _no else 0.0})
        out["prosek_po_objektu"] = _ppo
    except Exception:
        pass

    # ---- Mesečne grupe (za složeni grafikon) iz analitike (fallback ako nema tabele prodaje) ----
    try:
        if "Grupa" in res.columns:
            _gm = {}
            for lb in ml:
                cp = _col(lb, "_Prodaja")
                if cp:
                    _gs = res.groupby("Grupa")[cp].sum()
                    for _gn, _gv in _gs.items():
                        _gm.setdefault(str(_gn), []).append(int(_gv))
            if _gm:
                out["nazivi"] = list(ml)
                out["grupe_mesecno"] = _gm
    except Exception:
        pass

    # ---- OOS po količinama za poslednji mesec ----
    try:
        _dfoos = getattr(engine, "df_oos", None)
        if _dfoos is not None and len(_dfoos) > 0 and ml:
            _last = ml[-1]
            _colo = "OOS_" + str(_last)
            _ok = {"mesec": _last, "izgubljeno_kom": 0, "objekata_na_0": 0, "po_artiklu": []}
            if _colo in _dfoos.columns:
                _sub = _dfoos[_dfoos[_colo] > 0]
                _ok["izgubljeno_kom"] = int(round(_sub[_colo].sum()))
                _per = (_sub.groupby("Naziv artikla")
                        .agg(objekata=("ID KOMITENTA", "nunique"), izg=(_colo, "sum"))
                        .reset_index().sort_values("izg", ascending=False))
                _ok["po_artiklu"] = [{"artikal": str(r["Naziv artikla"]), "objekata": int(r["objekata"]),
                                      "izgubljeno": int(round(r["izg"]))} for _, r in _per.iterrows()]
            try:
                _ok["objekata_na_0"] = int(res[res["Lager_danas"] == 0]["ID KOMITENTA"].nunique())
            except Exception:
                _ok["objekata_na_0"] = int((_dfoos.get("Lager_danas", 0) == 0).sum())
            out["oos_kom"] = _ok
    except Exception:
        pass

    # ---- Predlog porudžbine za sistem + pokrivenost lagera ----
    try:
        if "Porudzbina_2" in res.columns:
            _exc = getattr(engine, "excluded", None) or set()
            _rv = res[~res["ID KOMITENTA"].isin(_exc)] if _exc else res
            _pr = {"ukupno": int(_rv["Porudzbina_2"].sum()),
                   "objekata": int(_rv[_rv["Porudzbina_2"] > 0]["ID KOMITENTA"].nunique()),
                   "po_grupi": []}
            if "Grupa" in _rv.columns:
                _gp = _rv.groupby("Grupa")["Porudzbina_2"].sum().sort_values(ascending=False)
                _pr["po_grupi"] = [{"grupa": str(g), "kom": int(v)} for g, v in _gp.items() if int(v) > 0]
            # prosečna pokrivenost (dani) — ponderisano prodajom, iz df_promo
            _dp = getattr(engine, "df_promo", None)
            if _dp is not None and len(_dp) > 0 and "Dani_pokrivanja" in _dp.columns:
                _dd = _dp[(_dp["Dani_pokrivanja"] < 900) & (_dp["Prodato_kom"] > 0)]
                if len(_dd) > 0:
                    _wsum = float((_dd["Dani_pokrivanja"] * _dd["Prodato_kom"]).sum())
                    _psum = float(_dd["Prodato_kom"].sum())
                    _pr["dani_avg"] = int(round(_wsum / _psum)) if _psum else 0
            out["porudzbina"] = _pr
    except Exception:
        pass

    # ---- Bestseleri i najslabiji artikli (iz res, po ukupnoj prodaji perioda) ----
    try:
        _art = {}
        _grp_of = {}
        for lb in ml:
            cp = _col(lb, "_Prodaja")
            if cp and "Naziv artikla" in res.columns:
                _gsum = res.groupby("Naziv artikla")[cp].sum()
                for _nz, _vv in _gsum.items():
                    _art[str(_nz)] = _art.get(str(_nz), 0) + int(_vv)
        if "Grupa" in res.columns and "Naziv artikla" in res.columns:
            for _nz, _gg in res.groupby("Naziv artikla")["Grupa"].first().items():
                _grp_of[str(_nz)] = str(_gg)
        if _art:
            _srt = sorted(_art.items(), key=lambda kv: kv[1], reverse=True)
            _best = [{"artikal": k, "grupa": _grp_of.get(k, ""), "prodato": v} for k, v in _srt[:8]]
            _slab = [{"artikal": k, "grupa": _grp_of.get(k, ""), "prodato": v}
                     for k, v in sorted(_srt, key=lambda kv: kv[1])[:8]]
            out["artikli_rang"] = {"best": _best, "slab": _slab}
    except Exception:
        pass

    # ---- Uspešnost akcije (iz df_promo; samo ako ima cena) ----
    try:
        _dp = getattr(engine, "df_promo", None)
        if _dp is not None and len(_dp) > 0:
            _ak = {"ukupno_akcija": int(_dp["Profit_akcija"].sum()),
                   "ukupno_redovna": int(_dp["Profit_da_je_redovna"].sum())}
            _ak["razlika"] = _ak["ukupno_redovna"] - _ak["ukupno_akcija"]
            _tp = _dp.sort_values("Prodato_kom", ascending=False).head(12)
            _ak["artikli"] = [{"naziv": str(r["Naziv"]), "grupa": str(r["Grupa"]),
                               "prodato": int(r["Prodato_kom"]), "obrt": float(r["Obrt_x"]),
                               "popust": float(r["Popust_%"]), "profit_akcija": int(r["Profit_akcija"]),
                               "cena_akcije": int(r["Cena_akcije"]), "dani": int(r["Dani_pokrivanja"]) if r["Dani_pokrivanja"] < 900 else 0}
                              for _, r in _tp.iterrows()]
            out["akcija"] = _ak
    except Exception:
        pass

    # ---- Profitabilnost (identično kao u analitici; puni se samo ako ima cena) ----
    try:
        if getattr(engine, "has_prices", False) and len(getattr(engine, "df_profit_obj", [])) > 0:
            prof = engine.df_profit_obj.copy()
            a_labels = list(engine.analitika_labels) if getattr(engine, "analitika_labels", None) else list(ml)
            n_mes = max(len(a_labels), 1)
            pf = {}
            pf["period"] = ", ".join(a_labels) if a_labels else "svi meseci"
            pf["n_mes"] = n_mes
            pf["n_obj"] = int(getattr(engine, "num_komitenti", len(prof)))
            pf["total_trosak"] = int(prof["Trosak_mkt"].sum())
            pf["total_bruto"] = int(prof["Bruto_profit"].sum())
            pf["total_neto"] = int(prof["Neto_profit"].sum())
            _dfoos = getattr(engine, "df_oos", None)
            _has_oos = _dfoos is not None and len(_dfoos) > 0
            pf["total_oos"] = int(_dfoos["Izgubljeni_profit"].sum()) if _has_oos else 0
            # mesečni trend bruto / neto
            _bm = []; _nm = []
            for lb in a_labels:
                cb = "Bruto_" + str(lb); cn = "Neto_" + str(lb)
                _bm.append([lb, int(prof[cb].sum()) if cb in prof.columns else 0])
                _nm.append([lb, int(prof[cn].sum()) if cn in prof.columns else 0])
            pf["bruto_po_mes"] = _bm
            pf["neto_po_mes"] = _nm
            # profitabilnost po objektima
            ukupno = len(prof)
            _neg = prof[prof["Neto_profit"] <= 0]
            _oos_neg = prof[(prof["Neto_profit"] <= 0) & (prof["Potencijalni_profit"] > 0)]
            _pravi_neg = prof[(prof["Neto_profit"] <= 0) & (prof["Potencijalni_profit"] <= 0)]
            pf["obj_ukupno"] = int(ukupno)
            pf["obj_profit"] = int(ukupno - len(_neg))
            pf["obj_oos_neg"] = int(len(_oos_neg))
            pf["obj_pravi_neg"] = int(len(_pravi_neg))
            _tpo = float(getattr(engine, "trosak_po_objektu", 0) or 0)
            _usteda = len(_pravi_neg) * _tpo + (abs(_pravi_neg["Neto_profit"].sum()) if len(_pravi_neg) > 0 else 0)
            pf["usteda_ukupno"] = int(_usteda)
            pf["objekti"] = [{"id": int(r["ID KOMITENTA"]), "neto": int(r["Neto_profit"]),
                              "potencijal": int(r["Potencijalni_profit"]), "bruto": int(r["Bruto_profit"]),
                              "trosak": int(r["Trosak_mkt"]), "oos": int(r["Izgubljeno_OOS"])}
                             for _, r in prof.iterrows()]
            # OOS u dinarima (identično kao analitika)
            if _has_oos:
                pf["oos_0_danas"] = int((_dfoos.get("Lager_danas", 0) == 0).sum()) if "Lager_danas" in _dfoos.columns else 0
                _om = []
                for lb in a_labels:
                    ci = "Izgub_" + str(lb); co = "OOS_" + str(lb)
                    _om.append([lb, int(_dfoos[ci].sum()) if ci in _dfoos.columns else 0,
                                int((_dfoos[co] > 0).sum()) if co in _dfoos.columns else 0])
                pf["oos_po_mes"] = _om
                _oa = (_dfoos.groupby(["id artikla", "Naziv artikla"])
                       .agg(Objekata=("ID KOMITENTA", "nunique"), OOS_meseci=("OOS_meseci", "sum"),
                            Izgubljeni_profit=("Izgubljeni_profit", "sum"))
                       .reset_index().sort_values("Izgubljeni_profit", ascending=False))
                pf["oos_artikli"] = [{"naziv": str(r["Naziv artikla"]), "objekata": int(r["Objekata"]),
                                      "meseci": int(r["OOS_meseci"]), "rsd": int(r["Izgubljeni_profit"])}
                                     for _, r in _oa.iterrows()]
            out["profit"] = pf
    except Exception:
        pass

    return out

# =====================================================================
# PRIJAVA (dve uloge: analitika / administracija)
# =====================================================================
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
    html, body, .stApp { background: #f4f1fb !important; font-family: 'Poppins', sans-serif; }
    .stApp { background: linear-gradient(135deg, #f7f5fc 0%, #efeaf8 55%, #f7f5fc 100%) !important; }
    header[data-testid="stHeader"] { background: transparent !important; }
    .stDeployButton { display: none; }
    footer { display: none; }
    #MainMenu { display: none; }
    .block-container { max-width: 440px !important; margin: 0 auto !important; padding-top: 90px !important; }
    .login-card-wrap { background:#ffffff; border:1px solid #ece7f6; border-radius:20px; padding:34px 30px;
        box-shadow:0 10px 40px rgba(124,58,237,0.10); }
    .stTextInput > div > div > input {
        background: #ffffff !important; border: 1px solid #e3e0ee !important;
        color: #1f2430 !important; border-radius: 12px !important; padding: 13px 16px !important; font-size: 15px !important; }
    .stTextInput > div > div > input::placeholder { color: #9aa0ad !important; }
    .stTextInput > div > div > input:focus {
        border-color: #a855f7 !important; box-shadow: 0 0 0 3px rgba(168,85,247,0.14) !important; }
    .stButton > button {
        background: linear-gradient(135deg, #a855f7 0%, #ec4899 100%) !important; color: white !important;
        border: none !important; border-radius: 12px !important; padding: 13px 32px !important;
        font-weight: 700 !important; font-size: 15px !important; width: 100% !important;
        box-shadow: 0 6px 20px rgba(168,85,247,0.28) !important; transition: opacity 0.2s !important; }
    .stButton > button:hover { opacity: 0.9 !important; }
    .stAlert { border-radius: 10px !important; background: #fdecec !important;
        border: 1px solid #f6c9c9 !important; color: #b42318 !important; }
    </style>
    """, unsafe_allow_html=True)
    st.markdown("""
    <div style="text-align:center; margin-bottom: 32px;">
        <div style="display:inline-flex; align-items:center; gap:10px; margin-bottom: 22px;">
            <div style="width:38px; height:38px; background:linear-gradient(135deg,#a855f7,#ec4899);
                border-radius:10px; display:inline-flex; align-items:center; justify-content:center;
                box-shadow:0 6px 18px rgba(168,85,247,0.3);">
                <div style="width:13px; height:13px; background:white; border-radius:3px; opacity:0.95;"></div>
            </div>
            <span style="font-size:22px; font-weight:800; color:#1f2430; letter-spacing:0.3px;">Vape Shop</span>
            <span style="font-size:22px; font-weight:300; color:#a99bd1;">Porudžbine</span>
        </div>
        <div style="height:1px; background:linear-gradient(90deg, transparent, #e3ddf2, transparent); margin-bottom:26px;"></div>
        <h2 style="color:#1f2430; font-size:23px; font-weight:700; margin:0 0 8px 0; line-height:1.35;">
            Dobrodošli 👋
        </h2>
        <p style="color:#8b8fa0; font-size:14px; margin:0;">
            Unesite korisničko ime i šifru
        </p>
    </div>
    """, unsafe_allow_html=True)
    usr = st.text_input("Korisničko ime", placeholder="Korisničko ime ili mejl",
                        label_visibility="collapsed")
    pwd = st.text_input("Šifra", type="password", placeholder="Šifra",
                        label_visibility="collapsed")
    btn = st.button("Prijavi se", use_container_width=True)
    if btn:
        _u = _norm_korisnik(usr)
        # (spisak prihvacenih imena, sifra, uloga, ime za prikaz, nalog za mejl)
        _nalozi = [
            (_aliasi(APP_KORISNIK), APP_PASSWORD, "analitika", None, ""),
            (_aliasi(ADMIN_KORISNIK, ADMIN_IME, _cfg("SMTP_USER_1", "")),
             ADMIN_PASSWORD, "administracija", ADMIN_IME, "1"),
            (_aliasi(ADMIN_KORISNIK_2, ADMIN_IME_2, _cfg("SMTP_USER_2", "")),
             ADMIN_PASSWORD_2, "administracija", ADMIN_IME_2, "2"),
            (_aliasi(DIREKTOR_KORISNIK), DIREKTOR_PASSWORD, "direktori", None, ""),
            (_aliasi(KOMERCIJALA_KORISNIK, KOMERCIJALA_IME), KOMERCIJALA_PASSWORD,
             "komercijala", KOMERCIJALA_IME, ""),
            (_aliasi(KOMERCIJALA_KORISNIK_2, KOMERCIJALA_IME_2), KOMERCIJALA_PASSWORD_2,
             "komercijala", KOMERCIJALA_IME_2, ""),
        ]
        _nadjen = None
        for _ku, _kp, _rola, _ime, _nalog in _nalozi:
            # ako je ime uneto -> mora da se poklopi; ako je prazno -> vazi samo sifra
            if pwd == _kp and (not _u or _u in _ku):
                _nadjen = (_rola, _ime, _nalog)
                break
        if _nadjen:
            _rola, _ime, _nalog = _nadjen
            st.session_state.authenticated = True
            st.session_state.role = _rola
            st.session_state.mail_nalog = _nalog
            if _rola == "administracija":
                st.session_state.admin_user = _ime
            elif _rola == "komercijala":
                st.session_state.komerc_user = _ime
            st.rerun()
        else:
            st.error("Pogrešno korisničko ime ili šifra")
    st.markdown("""
    <div style="text-align:center; margin-top:28px;">
        <p style="color:#b9b3c9; font-size:12px; margin:0;">
            Vape Shop · Sistem porudžbina
        </p>
    </div>
    """, unsafe_allow_html=True)
    return False


# =====================================================================
# PREGLED ZA KOLEGINICE (administracija)
# =====================================================================
def _admin_css():
    st.markdown("""<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
    section[data-testid="stSidebar"] { display: none !important; }
    header[data-testid="stHeader"] { display: none !important; }
    #MainMenu { visibility: hidden !important; }
    footer { visibility: hidden !important; }
    .stApp { background: #f5f0ff !important; font-family: 'Poppins', sans-serif; }
    div[data-testid="stMainBlockContainer"], .main .block-container {
        padding: 12px 18px 0 18px !important; max-width: 100% !important; }
    .stButton > button {
        background: linear-gradient(135deg, #a855f7 0%, #ec4899 100%) !important; color: white !important;
        border: none !important; border-radius: 10px !important; font-weight: 600 !important; }
    </style>""", unsafe_allow_html=True)

def _admin_header():
    st.markdown('''<div style="background:#12002a;border-radius:16px;padding:0 28px;height:60px;
        display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;
        border-bottom:3px solid;border-image:linear-gradient(90deg,#a855f7,#ec4899) 1;
        box-shadow:0 4px 20px rgba(18,0,42,0.18);">
        <div style="display:flex;align-items:center;gap:12px;">
            <div style="width:30px;height:30px;background:linear-gradient(135deg,#a855f7,#ec4899);
                border-radius:8px;display:flex;align-items:center;justify-content:center;">
                <div style="width:11px;height:11px;background:white;border-radius:3px;"></div>
            </div>
            <span style="font-size:18px;font-weight:700;color:white;">VAPE</span>
            <span style="font-size:18px;font-weight:300;color:rgba(255,255,255,0.4);">Porudžbine</span>
            <span style="font-size:11px;color:rgba(255,255,255,0.25);margin-left:8px;">·</span>
            <span style="font-size:12px;color:rgba(255,255,255,0.35);">Pregled za administraciju</span>
        </div>
        <div style="display:flex;gap:6px;align-items:center;">
            <div style="width:8px;height:8px;border-radius:50%;background:rgba(168,85,247,0.7);"></div>
            <div style="width:8px;height:8px;border-radius:50%;background:rgba(236,72,153,0.5);"></div>
            <div style="width:8px;height:8px;border-radius:50%;background:rgba(255,255,255,0.15);"></div>
        </div>
    </div>''', unsafe_allow_html=True)

@st.cache_resource
def _pdf_font():
    import os
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    cands = [("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")]
    try:
        import matplotlib
        _d = os.path.join(matplotlib.get_data_path(), "fonts/ttf")
        cands.append((os.path.join(_d, "DejaVuSans.ttf"), os.path.join(_d, "DejaVuSans-Bold.ttf")))
    except Exception:
        pass
    for _r, _b in cands:
        if os.path.exists(_r):
            try:
                pdfmetrics.registerFont(TTFont("DejaVu", _r))
                if os.path.exists(_b):
                    pdfmetrics.registerFont(TTFont("DejaVu-Bold", _b))
                    pdfmetrics.registerFontFamily("DejaVu", normal="DejaVu", bold="DejaVu-Bold",
                                                  italic="DejaVu", boldItalic="DejaVu-Bold")
                    return "DejaVu", "DejaVu-Bold"
                return "DejaVu", "DejaVu"
            except Exception:
                pass
    return "Helvetica", "Helvetica-Bold"


def napravi_pdf_izvestaj(mesec_key, mesec_lbl):
    import io as _io
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as _plt
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, KeepTogether, HRFlowable

    FN, FB = _pdf_font()
    H1 = ParagraphStyle('H1', fontName=FB, fontSize=19, leading=23, textColor=colors.HexColor('#3b0764'), spaceAfter=2)
    Hsub = ParagraphStyle('Hsub', fontName=FN, fontSize=10, leading=13, textColor=colors.HexColor('#9188a5'), spaceAfter=14)
    Hsys = ParagraphStyle('Hsys', fontName=FB, fontSize=13, leading=16, textColor=colors.HexColor('#7c3aed'), spaceBefore=2, spaceAfter=4)
    Nar = ParagraphStyle('Nar', fontName=FN, fontSize=10, leading=14.5, textColor=colors.HexColor('#333333'))

    def _chart(nred, norg, ngrn, n, nnas, nnj, nnije, neob, nrev):
        fig, ax = _plt.subplots(1, 2, figsize=(9.0, 2.15))
        ax[0].pie([nred, norg, ngrn], colors=['#e5484d', '#f2820c', '#17a34a'], startangle=90,
                  counterclock=False, wedgeprops=dict(width=0.44, edgecolor='white', linewidth=1.6))
        ax[0].text(0, 0, str(n), ha='center', va='center', fontsize=12, fontweight='bold', color='#3b0764')
        ax[0].set_title('Zone', fontsize=9.5, fontweight='bold', color='#3b0764', pad=3)
        ax[0].legend(["Hitno (" + str(nred) + ")", "Iskontrolisati (" + str(norg) + ")", "Dobra (" + str(ngrn) + ")"],
                     loc='center left', bbox_to_anchor=(0.94, 0.5), frameon=False, fontsize=7.5, handlelength=1)
        ax[1].pie([max(nnas, 0), max(nnj, 0), max(nnije, 0), max(neob, 0)],
                  colors=['#17a34a', '#f2820c', '#c9b8ec', '#e5e2ee'], startangle=90,
                  counterclock=False, wedgeprops=dict(width=0.44, edgecolor='white', linewidth=1.6))
        ax[1].text(0, 0, str(nrev), ha='center', va='center', fontsize=12, fontweight='bold', color='#3b0764')
        ax[1].set_title('Trebovanje', fontsize=9.5, fontweight='bold', color='#3b0764', pad=3)
        ax[1].legend(["Prema našem predlogu (" + str(nnas) + ")", "Po njihovom (" + str(nnj) + ")",
                      "Bez porudžbine (" + str(nnije) + ")", "Neobrađeno (" + str(neob) + ")"],
                     loc='center left', bbox_to_anchor=(0.94, 0.5), frameon=False, fontsize=7.5, handlelength=1)
        fig.subplots_adjust(left=0.01, right=0.72, top=0.84, bottom=0.04, wspace=1.7)
        b = _io.BytesIO()
        fig.savefig(b, format='png', dpi=200, facecolor='white')
        _plt.close(fig)
        b.seek(0)
        return b

    buf = _io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=16 * mm, bottomMargin=14 * mm,
                            leftMargin=18 * mm, rightMargin=18 * mm)
    el = [Paragraph("Izveštaj administracije", H1),
          Paragraph("Mesec: " + mesec_lbl + "   ·   generisano " + _now().strftime("%d.%m.%Y"), Hsub)]

    sistemi = sb_sisteme(mesec_key)
    if not sistemi:
        el.append(Paragraph("Nema objavljenih sistema za ovaj mesec.", Nar))
        doc.build(el)
        buf.seek(0)
        return buf.getvalue()

    for _sis in sistemi:
        podaci = sb_ucitaj(mesec_key, _sis)
        if not podaci or not podaci.get("stavke"):
            continue
        stavke = podaci["stavke"]
        po = {}
        for s in stavke:
            po.setdefault(int(s["idk"]), []).append(s)
        _meta_s = podaci.get("meta") or {}
        # Sistemski/nedeljni sistem — u izveštaju samo problem + napomena (bez zona/trebovanja)
        if _meta_s.get("nedeljni"):
            _dani_s = int(_meta_s.get("nedeljni_dani", 7) or 7)
            _perl_s = str(_dani_s) + " dana"
            _cut_s = _admin_presek(_meta_s, mesec_key)
            _ahist_s = (_meta_s.get("admin_hist") or {}) if isinstance(_meta_s, dict) else {}
            def _pa_pdf(lst, idk):
                _pm = _treb_posle_preseka(_ahist_s.get(str(int(idk)), []) or [], _cut_s) if _cut_s else {}
                _o = []
                for s in lst:
                    _por = int(_pm.get(int(s.get('ida', -1)), 0) or 0)
                    _lg = int(s.get('lager', 0) or 0) + _por   # realni lager
                    _pr = int(s.get('pred', 0) or 0)
                    _prag = _pr * _dani_s / 30.0
                    if _pr > 0 and _lg < _prag:
                        _o.append({"naziv": str(s.get('naziv', '')), "lager": _lg, "prag": int(round(_prag))})
                return _o
            _prob_pdf = {idk: _pa_pdf(lst, idk) for idk, lst in po.items()}
            _prob_pdf = {k: v for k, v in _prob_pdf.items() if v}
            _npb = len(_prob_pdf)
            _prijave = dict(_meta_s.get("nedeljni_prijave") or {})
            _nst = _meta_s.get("nedeljni_start") or {}
            try:
                _kf_pdf = sb_komitenti_full()
            except Exception:
                _kf_pdf = {}
            _start_line = ("Na startu (" + str(_nst.get("kada", "")) + "): <b>" + str(_nst.get("problem", 0))
                           + "</b> objekata sa problemom od " + str(_nst.get("n", len(po))) + ".  ") if _nst else ""
            _flow = [Paragraph(str(_sis) + "  (sistemski · period " + _perl_s + ")", Hsys),
                     Paragraph(_start_line + "Trenutno sa problemom (realni lager ispod prodaje za " + _perl_s + "): <b>"
                               + str(_npb) + "</b> od " + str(len(po)) + ".  Prijavljeno komercijali: <b>"
                               + str(len(_prijave)) + "</b>.", Nar),
                     Spacer(1, 3)]
            # Prijavljeni komitenti — naziv, kontakt, artikli u problemu, napomena
            _any_prij = False
            for _idk, _info in _prijave.items():
                try:
                    _ik = int(_idk)
                except Exception:
                    continue
                _arts_p = _prob_pdf.get(_ik, [])
                _ki = (_kf_pdf.get(_ik, {}) or {})
                _naz = _ki.get("naziv", "") or ("ID " + str(_ik))
                _kbits = []
                if _ki.get("telefon"): _kbits.append(str(_ki["telefon"]))
                if _ki.get("email"): _kbits.append(str(_ki["email"]))
                _kline = ("  ·  " + "  ·  ".join(_kbits)) if _kbits else ""
                _any_prij = True
                _artline = "; ".join((str(a["naziv"]) + " (realni lager " + str(a["lager"])
                                      + ", za " + _perl_s + " ~" + str(a["prag"]) + ")") for a in _arts_p) or "—"
                _naptxt = (_info.get("napomena", "") or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                _flow.append(Paragraph("• <b>" + _h_escape(str(_naz)) + "</b>" + _h_escape(_kline), Nar))
                _flow.append(Paragraph("&nbsp;&nbsp;&nbsp;Problem: " + _h_escape(_artline), Nar))
                if _naptxt:
                    _flow.append(Paragraph("&nbsp;&nbsp;&nbsp;Napomena: " + _naptxt, Nar))
                _flow.append(Spacer(1, 3))
            if not _any_prij:
                _flow.append(Paragraph("<i>Nema prijavljenih problema za ovaj sistem.</i>", Nar))
            # Backward-compat: stara jedna napomena sistema (ako postoji)
            _napt_old = _meta_s.get("napomena_sistem", "") or ""
            if _napt_old.strip():
                _naph = (_napt_old.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>"))
                _flow.append(Spacer(1, 2))
                _flow.append(Paragraph("<b>Napomena administracije:</b><br/>" + _naph, Nar))
            _flow.append(Spacer(1, 4))
            _flow.append(HRFlowable(width="100%", thickness=0.6, color=colors.HexColor('#eae4f7'), spaceAfter=6))
            el.append(KeepTogether(_flow))
            continue
        objekti = []
        for idk, lst in po.items():
            nivo, n0, izg = hitnost_objekta(lst)
            objekti.append({"idk": idk, "nivo": nivo, "lst": lst})
        obrada = sb_load_obrada(mesec_key, _sis)
        # Količine: koliko je predloženo (dodatna, posle već poručenog) po objektu
        _cut_pdf = _admin_presek(_meta_s, mesec_key)
        _ahist_pdf = (_meta_s.get("admin_hist") or {}) if isinstance(_meta_s, dict) else {}
        def _kom_predlog(o):
            _pm = _treb_posle_preseka(_ahist_pdf.get(str(int(o["idk"])), []) or [], _cut_pdf) if _cut_pdf else {}
            _s = 0
            for a in (o.get("lst") or []):
                _s += max(int(a.get("kol", 0) or 0) - int(_pm.get(int(a.get("ida", -1)), 0) or 0), 0)
            return _s

        def _ima(o, naz):
            return naz in (obrada.get(o["idk"], {}).get("reakcije") or [])

        n = len(objekti)
        # START zone: ako je zabeležen „startni rezultat" (posle porudžbina do objave) — koristi njega;
        # inače sirova preporuka (staro ponašanje).
        _sz = ((podaci.get("meta") or {}).get("start_zone")) if isinstance(podaci.get("meta"), dict) else None
        if _sz:
            nred = int(_sz.get("crveno", 0)); norg = int(_sz.get("zuto", 0)); ngrn = int(_sz.get("zeleno", 0))
        else:
            nred = sum(1 for o in objekti if o["nivo"] == "crveno")
            norg = sum(1 for o in objekti if o["nivo"] == "zuto")
            ngrn = sum(1 for o in objekti if o["nivo"] == "zeleno")
        rev = [o for o in objekti if obrada.get(o["idk"], {}).get("reakcije")]
        nrev = len(rev)
        nPoz = sum(1 for o in objekti if _ima(o, "Pozvala sam"))
        nMej = sum(1 for o in objekti if _ima(o, "Poslala sam mejl"))
        nDir = sum(1 for o in objekti if _ima(o, "Obavestila direktorku"))
        # Ko je šta uradio (po osobi)
        _ko_stat = {}
        for o in objekti:
            _ov = obrada.get(o["idk"], {})
            _rk = _ov.get("reakcije_ko") or {}
            for _rr in (_ov.get("reakcije") or []):
                _u = _rk.get(_rr)
                if _u:
                    _ko_stat.setdefault(_u, {}).setdefault(_rr, 0)
                    _ko_stat[_u][_rr] += 1
        def _osoba_linija(stat):
            _pp = []
            for _u in sorted(stat.keys()):
                _d = stat[_u]
                _pp.append("<b>" + _h_escape(str(_u)) + "</b>: pozvano " + str(_d.get("Pozvala sam", 0))
                           + ", mejlova " + str(_d.get("Poslala sam mejl", 0))
                           + ", komercijali " + str(_d.get("Obavestila direktorku", 0)))
            return "  ·  ".join(_pp)
        _osoba_txt = _osoba_linija(_ko_stat)
        _nas_o = [o for o in rev if obrada.get(o["idk"], {}).get("trebovali_tip") == "nas"]
        _nj_o = [o for o in rev if obrada.get(o["idk"], {}).get("trebovali_tip") == "njihov"]
        nnas = len(_nas_o)
        nnj = len(_nj_o)
        nnije = sum(1 for o in rev if not obrada.get(o["idk"], {}).get("trebovali_tip"))
        neob = n - nrev
        pct = round(nrev / max(n, 1) * 100)
        # KOLIČINE trebovanja (kom)
        _kom_nas = sum(_kom_predlog(o) for o in _nas_o)
        _kom_nj = sum(int(x or 0) for o in _nj_o
                      for x in (obrada.get(o["idk"], {}).get("njihova") or {}).values())
        _kom_uk = _kom_nas + _kom_nj
        _kom_predlozeno = sum(_kom_predlog(o) for o in objekti)

        narr = ("Od <b>" + str(n) + "</b> objekata za porudžbinu, <b><font color='#d33'>" + str(nred) +
                "</font></b> je u hitnoj zoni, " + str(norg) + " za kontrolu i " + str(ngrn) + " u dobroj. "
                "Obrađeno je <b>" + str(nrev) + "</b> (" + str(pct) + "%) — pozvano " + str(nPoz) +
                ", mejlova " + str(nMej) + ", prosleđeno komercijali " + str(nDir) + ". "
                "Trebovanje: <b><font color='#158a3f'>" + str(nnas) + " prema našem predlogu ("
                + format(_kom_nas, ",").replace(",", ".") + " kom)</font></b>, "
                "<font color='#c66a00'>" + str(nnj) + " po njihovom ("
                + format(_kom_nj, ",").replace(",", ".") + " kom)</font>, " + str(nnije) +
                " bez porudžbine; <b>" + str(neob) + "</b> neobrađeno. "
                "Ukupno trebovano: <b>" + format(_kom_uk, ",").replace(",", ".") + " kom</b> "
                "(predloženo " + format(_kom_predlozeno, ",").replace(",", ".") + " kom).")
        img = Image(_chart(nred, norg, ngrn, n, nnas, nnj, nnije, neob, nrev), width=150 * mm, height=150 * mm * 2.15 / 9.0)
        img.hAlign = 'CENTER'
        _blk = [Paragraph(str(_sis), Hsys), Paragraph(narr, Nar)]
        if _osoba_txt:
            _blk.append(Spacer(1, 2))
            _blk.append(Paragraph("Ko je šta uradio — " + _osoba_txt, Nar))
        _blk += [Spacer(1, 2), img, Spacer(1, 4),
                 HRFlowable(width="100%", thickness=0.6, color=colors.HexColor('#eae4f7'), spaceAfter=6)]
        el.append(KeepTogether(_blk))

    # ===== ZBIRNO: rad administracije po osobi + efektivnost kontakta (ceo mesec) =====
    try:
        _adm = {}
        def _A(name):
            return _adm.setdefault(name or "Bez oznake", {"poz": 0, "mej": 0, "dir": 0, "obj": set()})
        _em_obj = 0; _em_reag = 0
        _call_b = {1: [0, 0], 2: [0, 0], 3: [0, 0]}   # broj poziva -> [ukupno objekata, reagovalo]
        _eff = {}   # {osoba: {mej:[uk,reag], p1, p2, p3}} — efektivnost po osobi
        def _E(name):
            return _eff.setdefault(name or "Bez oznake",
                                   {"mej": [0, 0], "p1": [0, 0], "p2": [0, 0], "p3": [0, 0]})
        for _sis in sistemi:
            _obr = sb_load_obrada(mesec_key, _sis)
            for _idk, _v in _obr.items():
                _dn = _v.get("dnevnik") or {}
                _reak = _v.get("reakcije") or []
                _rko = _v.get("reakcije_ko") or {}
                _pozivi = _dn.get("pozivi") or []
                _mejlovi = _dn.get("mejlovi") or []
                _reacted = (_v.get("trebovali_tip") == "njihov")
                # rad po osobi
                if _pozivi:
                    for _e in _pozivi:
                        _d = _A(_e.get("ko", "")); _d["poz"] += 1; _d["obj"].add((_sis, _idk))
                elif "Pozvala sam" in _reak:
                    _d = _A(_rko.get("Pozvala sam", "")); _d["poz"] += 1; _d["obj"].add((_sis, _idk))
                if _mejlovi:
                    for _e in _mejlovi:
                        _d = _A(_e.get("ko", "")); _d["mej"] += 1; _d["obj"].add((_sis, _idk))
                elif "Poslala sam mejl" in _reak:
                    _d = _A(_rko.get("Poslala sam mejl", "")); _d["mej"] += 1; _d["obj"].add((_sis, _idk))
                if "Obavestila direktorku" in _reak:
                    _d = _A(_rko.get("Obavestila direktorku", "")); _d["dir"] += 1; _d["obj"].add((_sis, _idk))
                # efektivnost kontakta
                _nmej = len(_mejlovi) if _mejlovi else (1 if "Poslala sam mejl" in _reak else 0)
                if _nmej > 0:
                    _em_obj += 1
                    if _reacted:
                        _em_reag += 1
                _ncall = len(_pozivi) if _pozivi else (1 if "Pozvala sam" in _reak else 0)
                if _ncall > 0:
                    _bk = 3 if _ncall >= 3 else _ncall
                    _call_b[_bk][0] += 1
                    if _reacted:
                        _call_b[_bk][1] += 1
                # efektivnost PO OSOBI
                _mailers = (set(_e.get("ko", "") for _e in _mejlovi) if _mejlovi
                            else ({_rko.get("Poslala sam mejl", "")} if "Poslala sam mejl" in _reak else set()))
                for _m in _mailers:
                    _de = _E(_m); _de["mej"][0] += 1
                    if _reacted: _de["mej"][1] += 1
                if _pozivi:
                    for _n in (1, 2, 3):
                        if len(_pozivi) >= _n:
                            _de = _E(_pozivi[_n - 1].get("ko", "")); _de["p" + str(_n)][0] += 1
                            if _reacted: _de["p" + str(_n)][1] += 1
                elif "Pozvala sam" in _reak:
                    _de = _E(_rko.get("Pozvala sam", "")); _de["p1"][0] += 1
                    if _reacted: _de["p1"][1] += 1
        _adm = {k: v for k, v in _adm.items() if (v["poz"] or v["mej"] or v["dir"])}
        if _adm or _em_obj or any(_call_b[b][0] for b in _call_b):
            Hbig = ParagraphStyle('Hbig', fontName=FB, fontSize=14, leading=18,
                                  textColor=colors.HexColor('#3b0764'), spaceBefore=8, spaceAfter=6)
            _zb = [HRFlowable(width="100%", thickness=1, color=colors.HexColor('#c9b8ec'), spaceAfter=8),
                   Paragraph("Rad administracije — ko je šta uradio (" + str(mesec_lbl) + ")", Hbig)]
            # grafik po osobi
            if _adm:
                def _adm_chart(adm):
                    _names = sorted(adm.keys(), key=lambda u: (u == "Bez oznake", u))
                    _xx = np.arange(len(_names)); _w = 0.26
                    fig, ax = _plt.subplots(figsize=(8.2, 2.7))
                    ax.bar(_xx - _w, [adm[n]["poz"] for n in _names], _w, label="Pozivi", color="#16a34a")
                    ax.bar(_xx, [adm[n]["mej"] for n in _names], _w, label="Mejlovi", color="#7c3aed")
                    ax.bar(_xx + _w, [adm[n]["dir"] for n in _names], _w, label="Direktoru", color="#ec4899")
                    ax.set_xticks(_xx); ax.set_xticklabels(_names, fontsize=9)
                    ax.legend(fontsize=8, frameon=False, ncol=3, loc='upper right')
                    for _sp in ['top', 'right']:
                        ax.spines[_sp].set_visible(False)
                    ax.tick_params(axis='y', labelsize=8)
                    ax.grid(axis='y', color='#eee', linewidth=0.7)
                    fig.tight_layout()
                    _b = _io.BytesIO(); fig.savefig(_b, format='png', dpi=200, facecolor='white'); _plt.close(fig); _b.seek(0)
                    return _b
                _cimg = Image(_adm_chart(_adm), width=170 * mm, height=170 * mm * 2.7 / 8.2)
                _cimg.hAlign = 'CENTER'
                _zb.append(_cimg)
                _zb.append(Spacer(1, 3))
                for _u in sorted(_adm.keys(), key=lambda u: (u == "Bez oznake", u)):
                    _d = _adm[_u]
                    _zb.append(Paragraph("<b>" + _h_escape(str(_u)) + "</b>: pozvano " + str(_d["poz"])
                                         + ", mejlova " + str(_d["mej"]) + ", prosleđeno komercijali " + str(_d["dir"])
                                         + " · obradila " + str(len(_d["obj"])) + " objekata.", Nar))
            # efektivnost kontakta
            _zb.append(Spacer(1, 6))
            _zb.append(Paragraph("Efektivnost kontakta <font color='#9aa0ad'>(reagovao = objekat poručio po svom "
                                 "sistemu posle kontakta)</font>", Hsys))
            def _fun(lbl, uk, re):
                _p = (str(round(re / uk * 100)) + "%") if uk else "—"
                return Paragraph("• <b>" + lbl + "</b>: " + str(uk) + " objekata → reagovalo <b><font color='#158a3f'>"
                                 + str(re) + "</font></b> (" + _p + ")", Nar)
            _zb.append(_fun("Poslat mejl", _em_obj, _em_reag))
            _zb.append(_fun("1 poziv", _call_b[1][0], _call_b[1][1]))
            _zb.append(_fun("2 poziva", _call_b[2][0], _call_b[2][1]))
            _zb.append(_fun("3+ poziva", _call_b[3][0], _call_b[3][1]))
            # efektivnost PO OSOBI
            _eff = {k: v for k, v in _eff.items()
                    if any(v[_kk][0] for _kk in ("mej", "p1", "p2", "p3"))}
            if _eff:
                def _fp(arr):
                    _p = (str(round(arr[1] / arr[0] * 100)) + "%") if arr[0] else "—"
                    return str(arr[0]) + "→<font color='#158a3f'>" + str(arr[1]) + "</font> (" + _p + ")"
                _zb.append(Spacer(1, 6))
                _zb.append(Paragraph("Efektivnost po osobi <font color='#9aa0ad'>(kontaktirano → reagovalo)</font>", Hsys))
                for _u in sorted(_eff.keys(), key=lambda u: (u == "Bez oznake", u)):
                    _e = _eff[_u]
                    _zb.append(Paragraph("<b>" + _h_escape(str(_u)) + "</b> — mejl: " + _fp(_e["mej"])
                                         + " · 1. poziv: " + _fp(_e["p1"]) + " · 2. poziv: " + _fp(_e["p2"])
                                         + " · 3. poziv: " + _fp(_e["p3"]), Nar))
            el.append(KeepTogether(_zb))
    except Exception:
        pass

    doc.build(el)
    buf.seek(0)
    return buf.getvalue()


def _admin_order_xlsx(rows):
    """rows = lista tuplova (id_kupca, id_artikla, kolicina).
    Vraca bajtove .xlsx u formatu koji admin (Nova porudzbina iz Excel-a) ocekuje:
    kolone 'Id kupca', 'Id artikla', 'Kolicina' (tacno kao u sablonu)."""
    import io as _io
    from openpyxl import Workbook as _WB
    _wb = _WB()
    _ws = _wb.active
    _ws.title = "Sheet1"
    _ws.append(["Id kupca", "Id artikla", "Količina"])
    for _k, _a, _q in rows:
        _ws.append([int(_k), int(_a), int(_q)])
    _buf = _io.BytesIO()
    _wb.save(_buf)
    return _buf.getvalue()


def _zadaci_xlsx(rows, sistem, mesec_lbl):
    """Excel liste zadataka administracije: ID, naziv, mejl + 3 poziva (Da/Ne),
    4 kolone ko+kada, i Završeno. rows = lista dict-ova."""
    import io as _io
    from openpyxl import Workbook as _WB
    from openpyxl.styles import Font as _F, PatternFill as _PF, Alignment as _AL, Border as _BD, Side as _SD
    _wb = _WB(); _ws = _wb.active; _ws.title = "Zadaci"
    _thin = _SD(style="thin", color="E5E0F0")
    _bord = _BD(left=_thin, right=_thin, top=_thin, bottom=_thin)
    _ws.merge_cells("A1:K1")
    _t = _ws["A1"]; _t.value = "Lista zadataka · " + str(sistem) + " · " + str(mesec_lbl)
    _t.font = _F(bold=True, size=13, color="3730A3"); _t.alignment = _AL(horizontal="left", vertical="center")
    _ws.row_dimensions[1].height = 22
    _hdr = ["ID komitenta", "Naziv objekta", "Mejl", "Pozvala 1. put", "Pozvala 2. put", "Pozvala 3. put",
            "Mejl — ko i kada", "1. poziv — ko i kada", "2. poziv — ko i kada", "3. poziv — ko i kada", "Završeno"]
    _ws.append([])           # red 2 prazan
    _ws.append(_hdr)         # red 3 zaglavlje
    _hf = _PF("solid", fgColor="EDE9FE")
    for _ci in range(1, len(_hdr) + 1):
        _c = _ws.cell(row=3, column=_ci)
        _c.font = _F(bold=True, size=10, color="4C1D95"); _c.fill = _hf
        _c.alignment = _AL(horizontal="center", vertical="center", wrap_text=True); _c.border = _bord
    _green = _PF("solid", fgColor="DCFCE7")
    for _r in rows:
        _ws.append([
            _r.get("idk"), _r.get("naziv", ""),
            "Da" if _r.get("mejl") else "Ne",
            "Da" if _r.get("p1") else "Ne",
            "Da" if _r.get("p2") else "Ne",
            "Da" if _r.get("p3") else "Ne",
            _r.get("mejl_ko", ""), _r.get("p1_ko", ""), _r.get("p2_ko", ""), _r.get("p3_ko", ""),
            "ZAVRŠENO" if _r.get("zavrseno") else "",
        ])
        _rr = _ws.max_row
        for _ci in range(1, len(_hdr) + 1):
            _cc = _ws.cell(row=_rr, column=_ci)
            _cc.border = _bord
            _cc.alignment = _AL(horizontal=("left" if _ci == 2 else "center"), vertical="center")
        if _r.get("zavrseno"):
            _zc = _ws.cell(row=_rr, column=11); _zc.fill = _green
            _zc.font = _F(bold=True, color="14532D")
    _ws.column_dimensions["A"].width = 12
    _ws.column_dimensions["B"].width = 42
    for _cl in ("C", "D", "E", "F"):
        _ws.column_dimensions[_cl].width = 13
    for _cl in ("G", "H", "I", "J"):
        _ws.column_dimensions[_cl].width = 26
    _ws.column_dimensions["K"].width = 12
    _ws.freeze_panes = "A4"
    _buf = _io.BytesIO(); _wb.save(_buf); return _buf.getvalue()


def _objekat_order_xlsx(naziv, idk, mesec_lbl, rows, meseci=None):
    """Excel za slanje objektu (mejlom): status (kružić), naziv artikla, Lager
    (realni: lager + poručeno posle 01.), Predikcija (za zadati broj meseci) i
    Porudžbina (dodatna). rows = lista dict-ova sa ključevima
    'kruzic','naziv','lager','predikcija','dodatna'. Vraća bajtove .xlsx."""
    import io as _io
    from openpyxl import Workbook as _WB
    from openpyxl.styles import Font as _F, PatternFill as _PF, Alignment as _AL, Border as _BD, Side as _SD
    _wb = _WB()
    _ws = _wb.active
    _ws.title = "Porudžbina"
    _thin = _SD(style="thin", color="E5E0F0")
    _bord = _BD(left=_thin, right=_thin, top=_thin, bottom=_thin)
    _pred_lbl = "Predikcija"
    _ncols = 5
    _last = "E"
    # Naslov (koji objekat / mesec)
    _ws.merge_cells("A1:" + _last + "1")
    _t = _ws["A1"]
    _t.value = ("Porudžbina · " + str(naziv or ("ID " + str(idk))) + " · "
                + _now().strftime("%d.%m.%Y."))
    _t.font = _F(bold=True, size=13, color="3730A3")
    _t.alignment = _AL(horizontal="left", vertical="center")
    _ws.row_dimensions[1].height = 22
    # Zaglavlje tabele
    _hdr = ["Status", "Naziv artikla", "Lager", _pred_lbl, "Porudžbina"]
    _ws.append([])  # red 2 prazan
    _ws.append(_hdr)
    _hr = 3
    _hfill = _PF("solid", fgColor="EDE9FE")
    for _ci, _hv in enumerate(_hdr, start=1):
        _c = _ws.cell(row=_hr, column=_ci)
        _c.font = _F(bold=True, size=11, color="4C1D95")
        _c.fill = _hfill
        _c.alignment = _AL(horizontal=("left" if _ci == 2 else "center"), vertical="center", wrap_text=True)
        _c.border = _bord
    for _r in rows:
        _ws.append([str(_r.get("kruzic", "")), str(_r.get("naziv", "")),
                    int(_r.get("lager", 0) or 0), int(_r.get("predikcija", 0) or 0),
                    int(_r.get("dodatna", 0) or 0)])
        _rr = _ws.max_row
        for _ci in range(1, _ncols + 1):
            _ws.cell(row=_rr, column=_ci).alignment = _AL(
                horizontal=("left" if _ci == 2 else "center"), vertical="center")
            _ws.cell(row=_rr, column=_ci).border = _bord
    _ws.column_dimensions["A"].width = 9
    _ws.column_dimensions["B"].width = 48
    _ws.column_dimensions["C"].width = 10
    _ws.column_dimensions["D"].width = 14
    _ws.column_dimensions["E"].width = 13
    _ws.freeze_panes = "A4"
    _buf = _io.BytesIO()
    _wb.save(_buf)
    return _buf.getvalue()


@st.cache_data(ttl=300, show_spinner=False)
def _cene_iz_analitike(mesec_key, sistem):
    """Prodajna cena po komadu za svaki artikal, iz analitika Excel-a sačuvanog pri objavi
    (list „Analiza akcije“: prihod / prodato kom). Vraća {id_artikla: cena_po_komadu}."""
    import base64 as _b64, io as _io
    try:
        _b = sb_ucitaj_xlsx(mesec_key, sistem)
        if not _b:
            return {}
        from openpyxl import load_workbook as _lw
        _wb = _lw(_io.BytesIO(_b64.b64decode(_b)), data_only=True, read_only=True)
        _ws = None
        for _n in _wb.sheetnames:
            if "analiza akcije" in str(_n).strip().lower():
                _ws = _wb[_n]
                break
        if _ws is None:
            return {}
        _rows = _ws.iter_rows(values_only=True)
        _h = next(_rows, None)
        if not _h:
            return {}
        _ix = {}
        for _i, _v in enumerate(_h):
            _t = str(_v or "").replace("\n", " ").strip().lower()
            if _t.startswith("id artikla"):
                _ix["ida"] = _i
            elif _t.startswith("prodato"):
                _ix["kom"] = _i
            elif _t.startswith("prihod akcija"):
                _ix["prihod"] = _i
        if not all(_k in _ix for _k in ("ida", "kom", "prihod")):
            return {}
        _out = {}
        for _r in _rows:
            if not _r:
                continue
            try:
                _ida = int(_r[_ix["ida"]])
                _kom = float(_r[_ix["kom"]] or 0)
                _pri = float(_r[_ix["prihod"]] or 0)
            except Exception:
                continue
            if _kom > 0 and _pri > 0:
                _out[_ida] = _pri / _kom
        return _out
    except Exception:
        return {}


@st.cache_data(ttl=300, show_spinner=False)
def _serije_prodaje_iz_analitike(mesec_key, sistem):
    """Rekonstruiši mesečnu prodaju po (objekat, artikal) iz analitika Excel-a koji je
    sačuvan pri objavi sistema. Koristi se kad stavke nemaju 'prodaja_mesecno'
    (sistemi objavljeni starijom verzijom aplikacije).
    Vraća (mapa {(idk, ida): [kom po mesecima]}, [nazivi meseci])."""
    import base64 as _b64, io as _io
    try:
        _b = sb_ucitaj_xlsx(mesec_key, sistem)
        if not _b:
            return {}, []
        from openpyxl import load_workbook as _lw
        _wb = _lw(_io.BytesIO(_b64.b64decode(_b)), data_only=True, read_only=True)
        _ws = None
        for _n in _wb.sheetnames:
            if "pregled po objektima" in str(_n).strip().lower():
                _ws = _wb[_n]
                break
        if _ws is None:
            return {}, []
        _rows = _ws.iter_rows(values_only=True)
        _r1 = next(_rows, None)
        _r2 = next(_rows, None)
        if not _r1 or not _r2:
            return {}, []
        _cur = ""
        _kol = []            # [(indeks kolone, naziv meseca)]
        _labels = []
        for _i in range(len(_r2)):
            _v1 = _r1[_i] if _i < len(_r1) else None
            if _v1 not in (None, ""):
                _cur = str(_v1).strip()
            if str(_r2[_i] or "").strip().lower() == "prodaja" and _cur:
                _kol.append((_i, _cur))
                if _cur not in _labels:
                    _labels.append(_cur)
        if not _kol:
            return {}, []
        _mapa = {}
        for _row in _rows:
            if not _row or _row[0] in (None, ""):
                continue
            try:
                _idk = int(_row[0]); _ida = int(_row[1])
            except Exception:
                continue
            _ser = []
            for _i, _lbl in _kol:
                _v = _row[_i] if _i < len(_row) else 0
                try:
                    _ser.append(max(int(round(float(_v or 0))), 0))
                except Exception:
                    _ser.append(0)
            _mapa[(_idk, _ida)] = _ser
        return _mapa, _labels
    except Exception:
        return {}, []


def _nedeljni_predlog_xlsx(sistem, dani, datum, grupe, payload=None):
    """Jedan Excel koji se šalje sistemu — zamenjuje raniji PDF + Excel.

      List 1 „Izveštaj“              — dijagnoza: tri kružna grafika, primer sa grafikom, zahtev
      List 2 „Predlog po objektima“  — grupisano po objektu, sa zbirovima
      List 3 „Tabela“                — ravan spisak sa filterom, za pivot
      List 4 „Podaci za grafikone“   — skriven, hrani grafikone sa lista 1

    grupe = [{'objekat': naziv, 'arts': [{'naziv','lager','pred','predlog'}, ...]}, ...]
    payload = isti rečnik koji je pravio PDF (dokazi, brojevi objekata, izgubljena prodaja).
              Ako ga nema ili u njemu nema dokaza, list „Izveštaj“ se izostavlja.
    """
    import io as _io
    from openpyxl import Workbook as _WB
    from openpyxl.styles import Font as _F, PatternFill as _PF, Alignment as _AL, Border as _BD, Side as _SD
    from openpyxl.chart import DoughnutChart as _DC, BarChart as _BC, Reference as _R
    from openpyxl.chart.series import DataPoint as _DP
    from openpyxl.chart.label import DataLabelList as _DLL
    from openpyxl.drawing.line import LineProperties as _LP
    from openpyxl.chart.shapes import GraphicalProperties as _GP

    _p = payload or {}
    _dok = _p.get("dokazi") or {}
    _FN = "Arial"
    INK, INK2, MUTC = "3730A3", "44403C", "8B8FA3"
    CRV, ZUT, ZEL = "E5484D", "F2820C", "17A34A"
    LJ = ["4C1D95", "5F21BD", "8B5CF6", "B7A3FB"]
    PLAVA, TAMNA = "2A78D6", "2E1065"
    _thin = _SD(style="thin", color="E5E0F0")
    _bord = _BD(left=_thin, right=_thin, top=_thin, bottom=_thin)
    _f_hdr = _PF("solid", fgColor="EDE9FE")
    _f_obj = _PF("solid", fgColor="F5F3FF")
    _f_sum = _PF("solid", fgColor="FFF7ED")
    _f_crv = _PF("solid", fgColor="FEF2F2")
    _f_zel = _PF("solid", fgColor="F0FDF4")
    _f_tam = _PF("solid", fgColor=TAMNA)

    def _per(n):
        n = int(n)
        if n % 7 == 0 and n >= 7:
            _w = n // 7
            return {1: "nedelju dana", 2: "dve nedelje", 3: "tri nedelje",
                    4: "četiri nedelje"}.get(_w, str(_w) + " nedelja")
        return str(n) + " dana"
    _perl = _per(dani)

    def _rs(n):
        return "{:,.0f}".format(int(n)).replace(",", ".")

    def _sir(ws, mapa):
        for k, v in mapa.items():
            ws.column_dimensions[k].width = v

    def _c(ws, adr, val, bold=False, size=10, boja=INK2, fill=None,
           wrap=False, ha="left", va="center"):
        _cl = ws[adr]
        _cl.value = val
        _cl.font = _F(name=_FN, bold=bold, size=size, color=boja)
        _cl.alignment = _AL(horizontal=ha, vertical=va, wrap_text=wrap)
        if fill:
            _cl.fill = fill
        return _cl

    _wb = _WB()

    # ============================================================ LIST 1: Izveštaj
    _zone = _dok.get("zone_objekti")          # [("Roba nedostaje", n), ("Na granici", n), ("Zalihe u redu", n)]
    _ucest = _dok.get("ucestalost")           # [("nijednom", n), ("1 put", n), ...]
    _art = _dok.get("stanje_artikala")        # [("Lager 0", n), ("Ispod minimuma", n), ("Dovoljno", n)]
    _pr = _dok.get("primer_neredovno")
    _ima_izv = bool(_zone or _ucest or _art or _pr)

    if _ima_izv:
        _ws = _wb.active
        _ws_izv = _ws
        _ws.title = "Izveštaj"
        _ws.sheet_view.showGridLines = False
        _sir(_ws, dict([("A", 2.5)] + [(chr(66 + i), 15) for i in range(10)] + [("L", 2.5)]))

        _n_uk = int(_p.get("n_obj_ukupno", 0) or 0)
        _n_pr = int(_p.get("n_obj_problem", 0) or 0)
        _izg = int(_p.get("izgub_rsd", 0) or 0)
        _pk = int(_p.get("predlog_kom", 0) or 0)

        _ws.merge_cells("B2:K2")
        _c(_ws, "B2", "Stanje zaliha i predlog dopune", True, 18, "FFFFFF", _f_tam)
        _ws.row_dimensions[2].height = 30
        _ws.merge_cells("B3:K3")
        _c(_ws, "B3", str(sistem) + "   ·   presek na dan " + str(datum)
           + (("   ·   " + str(_n_uk) + " objekata u sistemu") if _n_uk else ""),
           False, 10, "D7C9F5", _f_tam)
        _ws.row_dimensions[3].height = 18
        _ws.merge_cells("B5:K5")
        _c(_ws, "B5", "Poštovani, u nastavku je pregled stanja zaliha u vašim objektima i predlog "
                      "dopune. Brojke dolaze iz evidencije prodaje po objektu i iz porudžbina "
                      "zavedenih kod nas.", False, 10, INK2, wrap=True)
        _ws.row_dimensions[5].height = 28

        # --- skriveni list sa podacima za grafikone (pravi se sada, premešta se na kraj)
        _wp = _wb.create_sheet("Podaci za grafikone")
        _c(_wp, "A1", "Ovaj list hrani grafikone na listu „Izveštaj“. Ne briši ga.", True, 10, MUTC)
        _sir(_wp, {"A": 26, "B": 12, "C": 22, "D": 12})
        _r = 3
        _blok = {}
        for _kljuc, _naz, _lst in (("zone", "Stanje objekata", _zone),
                                   ("ucest", "Koliko puta je poručeno", _ucest),
                                   ("art", "Stanje artikala", _art)):
            if not _lst:
                continue
            _c(_wp, "A" + str(_r), _naz, True)
            _c(_wp, "B" + str(_r), "broj", True)
            _blok[_kljuc] = (_r + 1, len(_lst))
            for _i, _x in enumerate(_lst):
                _c(_wp, "A" + str(_r + 1 + _i), str(_x[0]))
                _c(_wp, "B" + str(_r + 1 + _i), int(_x[1]))
            _r += len(_lst) + 2
        _prim_r = None
        if _pr and _pr.get("meseci"):
            _mes = [str(m).split()[0] for m in _pr["meseci"]]
            _prim_r = _r
            _c(_wp, "A" + str(_r), "Mesec", True)
            _c(_wp, "B" + str(_r), "Prodato krajnjim kupcima", True)
            _c(_wp, "C" + str(_r), "Poručeno od nas", True)
            for _i, _m in enumerate(_mes):
                _c(_wp, "A" + str(_r + 1 + _i), _m)
                _c(_wp, "B" + str(_r + 1 + _i), int((_pr.get("prodaja") or [0] * 99)[_i]))
                _c(_wp, "C" + str(_r + 1 + _i), int((_pr.get("poruceno") or [0] * 99)[_i]))
            _r += len(_mes) + 2

        def _krug(naslov, prvi_red, n, boje):
            _ch = _DC(holeSize=58)
            _ch.title = naslov
            _ch.add_data(_R(_wp, min_col=2, min_row=prvi_red - 1, max_row=prvi_red + n - 1),
                         titles_from_data=True)
            _ch.set_categories(_R(_wp, min_col=1, min_row=prvi_red, max_row=prvi_red + n - 1))
            try:
                from openpyxl.chart.data_source import AxDataSource as _ADS2, StrRef as _SR2
                _ch.series[0].cat = _ADS2(strRef=_SR2(
                    f="'" + _wp.title + "'!$A$" + str(prvi_red) + ":$A$" + str(prvi_red + n - 1)))
            except Exception:
                pass
            _s = _ch.series[0]
            _s.data_points = []
            for _b in boje[:n]:
                _dp = _DP(idx=len(_s.data_points))
                _dp.graphicalProperties.solidFill = _b
                _dp.graphicalProperties.line.solidFill = "FFFFFF"
                _dp.graphicalProperties.line.width = 22000
                _s.data_points.append(_dp)
            _ch.dataLabels = _DLL()
            _ch.dataLabels.showVal = True
            _ch.dataLabels.showPercent = False
            _ch.dataLabels.showSerName = False
            _ch.dataLabels.showCatName = False
            _ch.dataLabels.showLegendKey = False
            _ch.width, _ch.height = 8.2, 7.4
            _ch.legend.position = "b"
            _ch.legend.overlay = False
            return _ch

        _c(_ws, "B7", "DIJAGNOZA — ŠTA POKAZUJU BROJKE", True, 11, INK)
        _koloni = ["B9", "E9", "H9"]
        _i_k = 0
        if "zone" in _blok:
            _ws.add_chart(_krug("1. Gde je roba nestala", _blok["zone"][0], _blok["zone"][1],
                                [CRV, ZUT, ZEL]), _koloni[_i_k]); _i_k += 1
        if "ucest" in _blok:
            _ws.add_chart(_krug("2. Kako se poručuje", _blok["ucest"][0], _blok["ucest"][1],
                                LJ), _koloni[_i_k]); _i_k += 1
        if "art" in _blok:
            _ws.add_chart(_krug("3. Koliko artikala fali", _blok["art"][0], _blok["art"][1],
                                [CRV, ZUT, ZEL]), _koloni[_i_k]); _i_k += 1
        for _rr in range(9, 24):
            _ws.row_dimensions[_rr].height = 15

        _txt = []
        if _zone:
            _txt.append(("B25:D26", "U " + str(_n_pr) + " objekata artikli koji se prodaju trenutno "
                         "su na nuli ili ispod nivoa koji pokriva prodaju za " + _perl + "."))
        if _ucest:
            _u = dict((str(a), int(b)) for a, b in _ucest)
            _retko = _u.get("nijednom", 0) + _u.get("1 put", 0)
            _svu = sum(_u.values()) or 1
            _txt.append(("E25:G26", str(_retko) + " od " + str(_svu) + " objekata poručilo je jednom "
                         "ili nijednom u posmatranom periodu, dok prodaja teče svakog meseca."))
        if _art:
            _a = [(str(x[0]), int(x[1])) for x in _art]
            _uka = sum(v for _n, v in _a) or 1
            _txt.append(("H25:K26", "Od " + _rs(_uka) + " artikala koji se prodaju, " + str(_a[0][1])
                         + " je na nuli, a još " + str(_a[1][1] if len(_a) > 1 else 0)
                         + " neće izdržati " + _perl + "."))
        for _rng, _t in _txt:
            _ws.merge_cells(_rng)
            _c(_ws, _rng.split(":")[0], _t, False, 9, INK2, wrap=True, va="top")
        _ws.row_dimensions[25].height = 26
        _ws.row_dimensions[26].height = 14

        _red = 28
        if _prim_r:
            _c(_ws, "B" + str(_red), "PRIMER — KAD ZALIHE NESTANU, PRODAJA STANE", True, 11, INK)
            _ws.merge_cells("B" + str(_red + 1) + ":K" + str(_red + 1))
            _c(_ws, "B" + str(_red + 1), str(_pr.get("obj", "")) + "   ·   " + str(_pr.get("art", "")),
               True, 10, INK2)
            _nm = len(_pr["meseci"])
            _bar = _BC(); _bar.type = "col"; _bar.style = 2
            _bar.add_data(_R(_wp, min_col=2, min_row=_prim_r, max_row=_prim_r + _nm), titles_from_data=True)
            _bar.add_data(_R(_wp, min_col=3, min_row=_prim_r, max_row=_prim_r + _nm), titles_from_data=True)
            _bar.set_categories(_R(_wp, min_col=1, min_row=_prim_r + 1, max_row=_prim_r + _nm))
            try:
                from openpyxl.chart.data_source import AxDataSource as _ADS, StrRef as _SR
                _cat_f = ("'" + _wp.title + "'!$A$" + str(_prim_r + 1) + ":$A$" + str(_prim_r + _nm))
                for _sr in _bar.series:
                    _sr.cat = _ADS(strRef=_SR(f=_cat_f))
            except Exception:
                pass
            for _i, _b in enumerate((PLAVA, "4A3AA7")):
                _bar.series[_i].graphicalProperties.solidFill = _b
                _bar.series[_i].graphicalProperties.line.noFill = True
            _bar.gapWidth = 60
            _bar.y_axis.title = "komada"
            _bar.y_axis.majorGridlines.spPr = _GP(ln=_LP(solidFill="ECE9F4"))
            # Bez ovoga Excel sakrije obe ose — nema ni meseci ni brojeva.
            # (openpyxl ne upiše <c:delete>, pa Excel podrazumeva da su obrisane.)
            _bar.x_axis.delete = False
            _bar.y_axis.delete = False
            _bar.x_axis.tickLblPos = "low"
            _bar.y_axis.tickLblPos = "nextTo"
            _bar.x_axis.majorTickMark = "none"
            _bar.y_axis.majorTickMark = "out"
            # broj iznad svakog stubića prodaje — da se vrednost vidi i bez ose
            _bar.series[0].dLbls = _DLL()
            _bar.series[0].dLbls.showVal = True
            _bar.series[0].dLbls.showSerName = False
            _bar.series[0].dLbls.showCatName = False
            _bar.series[0].dLbls.showLegendKey = False
            _bar.width, _bar.height = 23.5, 8.4
            _bar.legend.position = "t"
            _bar.legend.overlay = False
            _ws.add_chart(_bar, "B" + str(_red + 3))
            for _rr in range(_red + 3, _red + 20):
                _ws.row_dimensions[_rr].height = 15
            _red += 21
            _pros = _pr.get("mes_prosek") or 0
            _op = ("Ovaj artikal je danas na lageru NULA, a u posmatranom periodu prodato je "
                   + _rs(_pr.get("uk_prodato", 0)) + " komada — prosečno " + str(_pros).replace(".", ",")
                   + " komada mesečno.")
            if int(_pr.get("n_porudzbina", 0) or 0) > 0:
                _op += (" Objekat ga je poručio " + str(_pr.get("n_porudzbina"))
                        + (" put" if int(_pr.get("n_porudzbina")) == 1 else " puta") + ", ukupno "
                        + _rs(_pr.get("uk_poruceno", 0)) + " komada — manje nego što je prodao.")
            else:
                _op += " Objekat ga u tom periodu nije poručio nijednom."
            _op += (" Potražnja postoji i dalje, ali robe nema, pa je svaki dan od sada prodaja "
                    "koja se ne ostvaruje. Isti obrazac ponavlja se u većini objekata sa lista "
                    "„Predlog po objektima“.")
            _ws.merge_cells("B" + str(_red) + ":K" + str(_red + 2))
            _c(_ws, "B" + str(_red), _op, False, 10, "991B1B", _f_crv, wrap=True, va="top")
            for _rr in range(_red, _red + 3):
                _ws.row_dimensions[_rr].height = 16
            _red += 4

        _zt = "Molimo vas da porudžbinu pošaljete u najkraćem roku."
        if _pk:
            _zt += (" Predlažemo dopunu od " + _rs(_pk) + " komada, raspoređenu po objektima — "
                    "spisak je na listu „Predlog po objektima“.")
        _zt += " Svaki dan sa praznom policom je prodaja koja se ne nadoknađuje."
        if _izg:
            _zt += (" Po dosadašnjoj prodaji, odlaganje od " + _perl + " znači oko " + _rs(_izg)
                    + " RSD neostvarenog prometa.")
        _ws.merge_cells("B" + str(_red) + ":K" + str(_red + 2))
        _c(_ws, "B" + str(_red), _zt, True, 10, "14532D", _f_zel, wrap=True, va="top")
        for _rr in range(_red, _red + 3):
            _ws.row_dimensions[_rr].height = 16
        _ws.print_area = "A1:L" + str(_red + 3)
        _ws.page_setup.orientation = "portrait"
        _ws.page_setup.fitToWidth = 1
        _ws.page_setup.fitToHeight = 1
        _ws.sheet_properties.pageSetUpPr.fitToPage = True
        _w2 = _wb.create_sheet("Predlog po objektima")
    else:
        _wp = None
        _ws_izv = None
        _w2 = _wb.active
        _w2.title = "Predlog po objektima"

    # ============================================================ LIST 2: po objektima
    _w2.sheet_view.showGridLines = False
    _sir(_w2, {"A": 46, "B": 12, "C": 12, "D": 13})
    _w2.merge_cells("A1:D1")
    _c(_w2, "A1", "Predlog porudžbine · " + str(sistem) + " · " + str(datum), True, 13, INK)
    _w2.row_dimensions[1].height = 22
    _w2.merge_cells("A2:D2")
    _c(_w2, "A2", "Objekti kod kojih lager ne pokriva prodaju ni za " + _perl
       + ". Predlog = koliko komada nedostaje do pokrivenosti za " + _perl + ".", False, 9, MUTC)
    _r = 4
    for _h, _t in zip("ABCD", ["Objekat / artikal", "Lager", "Predikcija", "Predlog (kom)"]):
        _c(_w2, _h + str(_r), _t, True, 10, INK, _f_hdr, ha=("left" if _h == "A" else "center"))
        _w2[_h + str(_r)].border = _bord
    _r += 1
    _prvi = _r
    for _g in grupe:
        _c(_w2, "A" + str(_r), str(_g.get("objekat", "")), True, 10, INK, _f_obj)
        for _h in "BCD":
            _c(_w2, _h + str(_r), None, fill=_f_obj)
        _od = _r + 1
        _r += 1
        for _a in (_g.get("arts") or []):
            _c(_w2, "A" + str(_r), "    " + str(_a.get("naziv", "")))
            _c(_w2, "B" + str(_r), int(_a.get("lager", 0) or 0), ha="center")
            _c(_w2, "C" + str(_r), int(_a.get("pred", 0) or 0), ha="center")
            _c(_w2, "D" + str(_r), int(_a.get("predlog", 0) or 0), True, ha="center")
            for _h in "ABCD":
                _w2[_h + str(_r)].border = _bord
            _r += 1
        _c(_w2, "A" + str(_r), "    Ukupno — " + str(_g.get("objekat", "")), True, 9, INK2, _f_sum)
        for _h in "BC":
            _c(_w2, _h + str(_r), None, fill=_f_sum)
        # SUBTOTAL i ovde, da ga ukupan zbir ispod ne bi brojao dvaput
        _c(_w2, "D" + str(_r), "=SUBTOTAL(9,D" + str(_od) + ":D" + str(_r - 1) + ")",
           True, 10, INK, _f_sum, ha="center")
        for _h in "ABCD":
            _w2[_h + str(_r)].border = _bord
        _r += 2
    _c(_w2, "A" + str(_r), "UKUPAN PREDLOG", True, 11, "FFFFFF", _f_tam)
    for _h in "BC":
        _c(_w2, _h + str(_r), None, fill=_f_tam)
    _c(_w2, "D" + str(_r), "=SUBTOTAL(9,D" + str(_prvi) + ":D" + str(_r - 1) + ")",
       True, 12, "FFFFFF", _f_tam, ha="center")
    _w2.freeze_panes = "A5"

    # ============================================================ LIST 3: ravna tabela
    _w3 = _wb.create_sheet("Tabela")
    _sir(_w3, {"A": 40, "B": 38, "C": 10, "D": 12, "E": 14})
    for _h, _t in zip("ABCDE", ["Objekat", "Artikal", "Lager", "Predikcija", "Predlog (kom)"]):
        _c(_w3, _h + "1", _t, True, 10, "FFFFFF", _f_tam)
    _r = 2
    for _g in grupe:
        for _a in (_g.get("arts") or []):
            _c(_w3, "A" + str(_r), str(_g.get("objekat", "")))
            _c(_w3, "B" + str(_r), str(_a.get("naziv", "")))
            _c(_w3, "C" + str(_r), int(_a.get("lager", 0) or 0), ha="center")
            _c(_w3, "D" + str(_r), int(_a.get("pred", 0) or 0), ha="center")
            _c(_w3, "E" + str(_r), int(_a.get("predlog", 0) or 0), True, ha="center")
            _r += 1
    _c(_w3, "A" + str(_r), "UKUPNO", True, 10, INK, _f_sum)
    for _h in "BCD":
        _c(_w3, _h + str(_r), None, fill=_f_sum)
    _c(_w3, "E" + str(_r), "=SUM(E2:E" + str(_r - 1) + ")", True, 11, INK, _f_sum, ha="center")
    if _r > 2:
        _w3.auto_filter.ref = "A1:E" + str(_r - 1)
    _w3.freeze_panes = "A2"

    # Redosled kartica: prvo ono što se koristi (predlog i tabela), pa izveštaj,
    # a skriveni list sa podacima za grafikone sasvim na kraju.
    for _sh in [x for x in (_w2, _w3, _ws_izv, _wp) if x is not None]:
        try:
            _wb.move_sheet(_sh, offset=len(_wb.worksheets) - _wb.worksheets.index(_sh) - 1)
        except Exception:
            pass
    if _wp is not None:
        _wp.sheet_state = "hidden"
    try:
        _wb.active = 0          # fajl se otvara na „Predlog po objektima“
    except Exception:
        pass

    _buf = _io.BytesIO()
    _wb.save(_buf)
    return _buf.getvalue()

def _admin_secret(k, d=""):
    try:
        import streamlit as _s
        if hasattr(_s, "secrets") and k in _s.secrets:
            return str(_s.secrets[k])
    except Exception:
        pass
    import os
    return os.environ.get(k, d)


def _sb_select_all(table, columns, step=1000):
    """Pročitaj SVE redove iz tabele kroz paginaciju (Supabase/PostgREST vraća
    max 1000 po upitu). Vraća listu dict-ova."""
    cli = _sb()
    if cli is None:
        return []
    out = []
    start = 0
    while True:
        res = cli.table(table).select(columns).range(start, start + step - 1).execute()
        batch = res.data or []
        out.extend(batch)
        if len(batch) < step:
            break
        start += step
    return out


def sb_komitenti_map():
    if _sb() is None:
        return {}
    try:
        rows = _sb_select_all("komitenti", "idk,naziv")
        return {int(r["idk"]): (r.get("naziv") or "") for r in rows}
    except Exception:
        return {}


def sb_komitenti_save(mapping):
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    rows = [{"idk": int(k), "naziv": v} for k, v in mapping.items()]
    for _i in range(0, len(rows), 500):
        cli.table("komitenti").upsert(rows[_i:_i + 500], on_conflict="idk").execute()


def sb_komitenti_full():
    """Vrati {idk: {naziv,email,telefon,mesto,adresa}} za SVE komitente (paginacija).
    Radi i ako kolone kontakata još ne postoje (tada su prazne)."""
    if _sb() is None:
        return {}
    try:
        rows = _sb_select_all("komitenti", "idk,naziv,email,telefon,mesto,adresa")
        out = {}
        for r in rows:
            out[int(r["idk"])] = {"naziv": r.get("naziv") or "", "email": r.get("email") or "",
                                  "telefon": r.get("telefon") or "", "mesto": r.get("mesto") or "",
                                  "adresa": r.get("adresa") or ""}
        return out
    except Exception:
        try:
            rows = _sb_select_all("komitenti", "idk,naziv")
            return {int(r["idk"]): {"naziv": r.get("naziv") or "", "email": "", "telefon": "",
                                    "mesto": "", "adresa": ""} for r in rows}
        except Exception:
            return {}


def sb_komitenti_upsert_rows(rows):
    """rows: lista dict sa poljima idk,naziv,email,telefon,mesto,adresa. Upsert po idk.
    Vraca broj komitenata koji SU ZAISTA u bazi posle upisa (verifikovano čitanjem).
    Ako baza vrati 0 a poslali smo >0 -> tabela nije dobro podešena (baca grešku)."""
    cli = _sb()
    if cli is None:
        raise RuntimeError("Supabase nije podešen.")
    clean = []
    for r in rows:
        try:
            _id = int(r.get("idk"))
        except Exception:
            continue
        clean.append({"idk": _id, "naziv": r.get("naziv") or "", "email": r.get("email") or "",
                      "telefon": r.get("telefon") or "", "mesto": r.get("mesto") or "",
                      "adresa": r.get("adresa") or ""})
    if not clean:
        return 0

    def _do_upsert(_rows):
        for _i in range(0, len(_rows), 500):
            cli.table("komitenti").upsert(_rows[_i:_i + 500], on_conflict="idk").execute()

    _err = None
    try:
        _do_upsert(clean)
    except Exception as _e1:
        _err = _e1
        # Fallback: možda kolone email/telefon/mesto/adresa još ne postoje -> probaj samo idk+naziv
        slim = [{"idk": r["idk"], "naziv": r["naziv"]} for r in clean]
        try:
            _do_upsert(slim)
            _err = None
        except Exception as _e2:
            _err = _e2

    # Verifikacija: pročitaj nazad koliko redova zaista ima u tabeli posle upisa
    _u_bazi = None
    try:
        _chk = cli.table("komitenti").select("idk", count="exact").limit(1).execute()
        _u_bazi = _chk.count
    except Exception:
        _u_bazi = None

    if _u_bazi is not None:
        if _u_bazi == 0 and len(clean) > 0:
            raise RuntimeError(
                "Upis nije sačuvan — u tabeli 'komitenti' ima 0 redova posle upisa. "
                "Tabela najverovatnije nije podešena kako treba (pokreni SQL setup). "
                + (("Detalj: " + str(_err)) if _err else ""))
        return _u_bazi
    if _err is not None:
        raise RuntimeError("Upis nije uspeo: " + str(_err))
    return len(clean)


def posalji_u_admin(id_kupca, items):
    """Prijavi se u admin i kreira porudzbinu direktno (bez rucnog uvoza).
    items = lista dict {'idArticle': int, 'quantity': int}.
    Vraca (ok: bool, poruka: str)."""
    import re, requests
    base = (_admin_secret("ADMIN_BASE_URL", "https://admin.vapeshop.rs") or "").rstrip("/")
    email = _admin_secret("ADMIN_LOGIN_EMAIL", "")
    pwd = _admin_secret("ADMIN_LOGIN_PASSWORD", "")
    if not email or not pwd:
        return (False, "Nije podešena admin prijava. Analitičar treba da doda ADMIN_LOGIN_EMAIL i ADMIN_LOGIN_PASSWORD u Secrets.")
    items = [{"idArticle": int(i["idArticle"]), "quantity": int(i["quantity"])}
             for i in items if int(i.get("quantity", 0)) > 0]
    if not items:
        return (False, "Nema stavki za slanje (sve količine su 0).")
    _tok = re.compile(r'name="__RequestVerificationToken"[^>]*value="([^"]+)"')
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0", "Accept-Language": "sr,en;q=0.8"})
    try:
        r = s.get(base + "/login", timeout=30)
        m = _tok.search(r.text)
        if not m:
            return (False, "Ne mogu da otvorim login stranicu admina (proveri ADMIN_BASE_URL).")
        s.post(base + "/login",
               data={"Email": email, "Password": pwd,
                     "__RequestVerificationToken": m.group(1)},
               headers={"Referer": base + "/login"}, timeout=30, allow_redirects=True)
        chk = s.get(base + "/orders-processing/new-order-from-excel", timeout=30)
        if ("/login" in chk.url) or ('name="Password"' in chk.text):
            return (False, "Prijava na admin nije uspela. Proveri email/lozinku u Secrets — ili admin blokira pristup sa servera aplikacije.")
        m2 = _tok.search(chk.text)
        if not m2:
            return (False, "Ne mogu da nađem sigurnosni token na stranici za uvoz.")
        xlsx = _admin_order_xlsx([(id_kupca, it["idArticle"], it["quantity"]) for it in items])
        # Jedan korak: uvoz Excel-a = admin ODMAH kreira porudžbinu (kao ručno).
        # Namerno NE zovemo "create-order-from-in-memory-cart" da se ne naprave dve.
        up = s.post(base + "/orders-processing/load-order-from-excel",
                    files={"fileArticles": ("porudzbina.xlsx", xlsx,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
                    data={"__RequestVerificationToken": m2.group(1)},
                    headers={"Referer": base + "/orders-processing/new-order-from-excel"},
                    timeout=90, allow_redirects=True)
        if up.status_code >= 400:
            return (False, "Slanje u admin nije prošlo (status " + str(up.status_code) + ").")
        _low = (up.text or "").lower()
        if ("greška" in _low) or ("greska" in _low) or ("error" in _low) or ("nije prona" in _low):
            return (False, "Admin je prijavio problem pri uvozu (proveri da su šifre artikala ispravne).")
        _mid = re.search(r'/order/details/(\d+)', up.text or "")
        if _mid:
            return (True, "Porudžbina #" + _mid.group(1) + " je kreirana u adminu.")
        return (True, "Porudžbina je poslata u admin (vidi „Pregled porudžbina“).")
    except requests.exceptions.RequestException as _e:
        return (False, "Greška u komunikaciji sa adminom: " + str(_e))


# ---- SMTP: slanje mejla objektu (sa prilogom) ----
def _mail_nalog():
    """Koji SMTP nalog koristi trenutno prijavljeni korisnik ('1', '2' ili '')."""
    try:
        return str(st.session_state.get("mail_nalog", "") or "")
    except Exception:
        return ""


def _smtp_kljuc(ime, nalog, default=None):
    """Prvo trazi npr. SMTP_USER_1, pa tek onda zajednicki SMTP_USER."""
    if nalog:
        v = _cfg(ime + "_" + str(nalog), None)
        if v not in (None, ""):
            return v
    return _cfg(ime, default)


def _smtp_cfg(nalog=None):
    n = _mail_nalog() if nalog is None else str(nalog or "")
    _ime = ADMIN_IME if n == "1" else (ADMIN_IME_2 if n == "2" else "Vape Shop")
    # Ako za ovaj nalog postoji svoj SMTP_USER, onda i lozinka mora da bude njegova —
    # ne sme da se povuce zajednicka (to bi bila tudja lozinka i prijava ne bi uspela).
    _svoj = _cfg("SMTP_USER_" + n, "") if n else ""
    if _svoj:
        _user = _svoj
        _pwd = _cfg("SMTP_PASSWORD_" + n, "") or ""
        _from = _cfg("SMTP_FROM_" + n, "") or _user
    else:
        _user = _cfg("SMTP_USER", "")
        _pwd = _cfg("SMTP_PASSWORD", "")
        _from = _cfg("SMTP_FROM", "") or _user
    return {
        "nalog": n,
        "host": _smtp_kljuc("SMTP_HOST", n, ""),
        "port": int(_smtp_kljuc("SMTP_PORT", n, 587) or 587),
        "user": _user,
        "password": _pwd,
        "from_email": _from,
        "from_name": _smtp_kljuc("SMTP_FROM_NAME", n, _ime),
        "use_ssl": bool(_smtp_kljuc("SMTP_USE_SSL", n, False)),
    }


def smtp_dostupan(nalog=None):
    c = _smtp_cfg(nalog)
    return bool(c["host"] and c["user"] and c["password"])


def _imap_cfg(nalog=None):
    """Podaci za IMAP (da se poslati mejl upiše u folder Poslato).
    Podrazumevano isti server i nalog kao SMTP; može se pregaziti sa IMAP_HOST_n / IMAP_PORT_n."""
    c = _smtp_cfg(nalog)
    n = c.get("nalog") or ""
    _h = _smtp_kljuc("IMAP_HOST", n, "") or ""
    if not _h:
        _h = str(c.get("host") or "")
        if _h.lower().startswith("smtp."):
            _h = "imap." + _h[5:]
    try:
        _p = int(_smtp_kljuc("IMAP_PORT", n, 993) or 993)
    except Exception:
        _p = 993
    return {"host": _h, "port": _p, "user": c.get("user", ""), "password": c.get("password", ""),
            "nalog": n}


def _imap_folderi(imap):
    """Vrati listu (ime_foldera, zastavice) iz IMAP LIST odgovora, ispravno isparsirano."""
    import re as _re
    _out = []
    try:
        _ok, _lst = imap.list()
        if _ok != "OK" or not _lst:
            return _out
    except Exception:
        return _out
    _rx = _re.compile(r'^\((?P<flags>[^)]*)\)\s+(?:"(?P<delim>[^"]*)"|NIL)\s+(?P<name>.+)$')
    for _raw in _lst:
        try:
            _lin = _raw.decode("utf-8", "ignore") if isinstance(_raw, bytes) else str(_raw)
        except Exception:
            continue
        _lin = _lin.strip()
        _m = _rx.match(_lin)
        if _m:
            _ime = _m.group("name").strip()
            _fl = _m.group("flags") or ""
        else:
            # rezerva: uzmi sve posle poslednjeg navodnika-para
            _d = _lin.split()
            _ime = _d[-1] if _d else ""
            _fl = ""
        if len(_ime) >= 2 and _ime[0] == '"' and _ime[-1] == '"':
            _ime = _ime[1:-1]
        _ime = _ime.strip()
        if _ime:
            _out.append((_ime, _fl))
    return _out


def _nadji_poslato_folder(imap):
    """Pronađi folder za poslatu poštu (razlikuje se po serveru i jeziku)."""
    _f = _imap_folderi(imap)
    if not _f:
        return None
    for _ime, _fl in _f:
        if "\\Sent" in _fl:
            return _ime
    for _ime, _fl in _f:
        _low = _ime.lower()
        if _low.endswith("sent") or _low.endswith("sent items") or _low.endswith("sent mail"):
            return _ime
    for _ime, _fl in _f:
        _low = _ime.lower()
        if "poslat" in _low or "poslano" in _low or "sent" in _low:
            return _ime
    return None


def _upisi_u_poslato(msg, nalog=None):
    """Upiši kopiju poslatog mejla u folder Poslato, da se vidi u Outlooku.
    Vraća (True, ime_foldera) ili (False, razlog). Nikad ne baca izuzetak."""
    if str(_cfg("SMTP_KOPIJA", "1")).strip().lower() in ("0", "false", "ne", "off"):
        return (False, "isključeno u podešavanjima")
    import imaplib, time as _t
    c = _imap_cfg(nalog)
    if not (c["host"] and c["user"] and c["password"]):
        return (False, "IMAP nije podešen")
    try:
        _im = imaplib.IMAP4_SSL(c["host"], c["port"], timeout=20)
        _im.login(c["user"], c["password"])
        _f = _smtp_kljuc("IMAP_SENT", c.get("nalog") or _mail_nalog(), "") or ""
        if not _f:
            _f = _nadji_poslato_folder(_im)
        if not _f:
            _im.logout()
            return (False, "ne mogu da nađem folder za poslatu poštu")
        _im.append(_f, "\\Seen", imaplib.Time2Internaldate(_t.time()),
                   msg.as_string().encode("utf-8"))
        _im.logout()
        return (True, _f)
    except Exception as _e:
        try:
            _im.logout()
        except Exception:
            pass
        return (False, str(_e)[:160])


def imap_test_upis(nalog=None):
    """Stvarno upiše probnu poruku u folder Poslato i vrati (True, folder) ili (False, razlog).
    Služi da se vidi da li kopija zaista stiže u sanduče."""
    from email.mime.text import MIMEText
    from email.utils import formataddr, formatdate
    c = _smtp_cfg(nalog)
    _m = MIMEText("Probna poruka iz aplikacije VAPE — Porudžbine.\n"
                  "Sluzi samo da se proveri da li se poslati mejlovi upisuju u folder Poslato.\n"
                  "Ovu poruku slobodno obriši.", "plain", "utf-8")
    _m["Subject"] = "PROBA — upis u Poslato"
    _m["From"] = formataddr((str(c.get("from_name", "")), str(c.get("from_email", ""))))
    _m["To"] = str(c.get("from_email", ""))
    _m["Date"] = formatdate(localtime=True)
    return _upisi_u_poslato(_m, nalog)


def _mail_telo_tekst(msg, maks=800):
    """Izvuci čitljiv tekst poruke (bez HTML-a, bez citiranog dela)."""
    import re as _re
    _txt = ""
    try:
        if msg.is_multipart():
            for _p in msg.walk():
                if _p.get_content_maintype() == "multipart":
                    continue
                if (_p.get("Content-Disposition") or "").lower().startswith("attachment"):
                    continue
                if _p.get_content_type() == "text/plain":
                    _txt = _p.get_payload(decode=True).decode(
                        _p.get_content_charset() or "utf-8", "ignore")
                    break
            if not _txt:
                for _p in msg.walk():
                    if _p.get_content_type() == "text/html":
                        _h = _p.get_payload(decode=True).decode(
                            _p.get_content_charset() or "utf-8", "ignore")
                        _txt = _re.sub(r"<[^>]+>", " ", _h)
                        break
        else:
            _txt = msg.get_payload(decode=True).decode(
                msg.get_content_charset() or "utf-8", "ignore")
            if msg.get_content_type() == "text/html":
                _txt = _re.sub(r"<[^>]+>", " ", _txt)
    except Exception:
        _txt = ""
    # HTML entiteti (&nbsp; &amp; …) u čitljiv tekst
    try:
        import html as _hh
        _txt = _hh.unescape(str(_txt or ""))
    except Exception:
        pass
    _lin = []
    for _l in str(_txt or "").splitlines():
        _s = _l.strip()
        if _s.startswith(">"):
            continue                      # citirani deo naše poruke
        if _re.match(r"^(-{2,}\s*Original|On .+ wrote:|Od:|From:|Poslato:|Sent:)", _s):
            break
        _lin.append(_s)
    _out = _re.sub(r"\n{3,}", "\n\n", "\n".join(_lin)).strip()
    return _out[:int(maks)]


def _mail_prilozi(msg, maks_po_fajlu=6_000_000, maks_fajlova=5):
    """Vrati [{ime, vel, data}] za priloge poruke (preskače inline slike i potpise)."""
    _out = []
    try:
        for _p in msg.walk():
            if _p.get_content_maintype() == "multipart":
                continue
            _disp = (_p.get("Content-Disposition") or "")
            _ime = _p.get_filename()
            if not _ime:
                continue
            # Slika iz potpisa (logo) nije prilog — prepoznaje se po Content-ID / inline
            if _p.get_content_maintype() == "image" and (_p.get("Content-ID")
                                                         or "inline" in _disp.lower()):
                continue
            try:
                from email.header import decode_header as _dh
                _dec = _dh(_ime)
                _ime = "".join((_b.decode(_c or "utf-8", "ignore") if isinstance(_b, bytes) else _b)
                               for _b, _c in _dec)
            except Exception:
                pass
            try:
                _dat = _p.get_payload(decode=True) or b""
            except Exception:
                _dat = b""
            if not _dat or len(_dat) > int(maks_po_fajlu):
                _out.append({"ime": str(_ime)[:120], "vel": len(_dat), "data": None})
            else:
                _out.append({"ime": str(_ime)[:120], "vel": len(_dat), "data": _dat})
            if len(_out) >= int(maks_fajlova):
                break
    except Exception:
        pass
    return _out


def knez_odgovori(adrese, nalog=None, od_datum=None, _v=0):
    """Pro\u010ditaj sandu\u010de i na\u0111i ODGOVORE pumpi. `adrese` = skup mejlova iz \u0161ifarnika.

    BRZINA: prvo se povla\u010de SAMO zaglavlja (From/Subject/Date) svih poruka od
    `od_datum` \u2014 to je par kilobajta. Cela poruka (sa prilozima) se skida tek za
    one \u010diji se po\u0161iljalac poklapa sa pumpom. Zato traje sekundama, a ne minutima.

    Vrati (po_adresi, neprepoznati, greska)."""
    import imaplib, email, re as _re
    from email.utils import parseaddr, parsedate_to_datetime
    from email.header import decode_header as _dh
    import datetime as _dt
    _po = {}
    _nep = []
    c = _imap_cfg(nalog)
    if not (c["host"] and c["user"] and c["password"]):
        return (_po, _nep, "IMAP nije pode\u0161en (IMAP_HOST / SMTP_USER / SMTP_PASSWORD u Secrets).")
    _skup = set(str(a or "").strip().lower() for a in (adrese or set()) if a)
    _nas = str(c.get("user") or "").strip().lower()
    _domeni = set()
    for _a in _skup:
        if "@" in _a:
            _domeni.add(_a.split("@")[-1])

    def _zaglavlje(_txt, _polje):
        _m = _re.search(r"^" + _polje + r":\s*(.*(?:\n[ \t].*)*)", _txt, _re.I | _re.M)
        return _re.sub(r"\s+", " ", _m.group(1)).strip() if _m else ""

    try:
        _im = imaplib.IMAP4_SSL(c["host"], c["port"], timeout=30)
        _im.login(c["user"], c["password"])
        _im.select("INBOX", readonly=True)
        if isinstance(od_datum, _dt.date):
            _od = od_datum.strftime("%d-%b-%Y")
        else:
            _od = _now().replace(day=1).strftime("%d-%b-%Y")
        _ok, _dat = _im.search(None, '(SINCE "' + _od + '")')
        _ids = (_dat[0].split() if (_ok == "OK" and _dat and _dat[0]) else [])
        if not _ids:
            try:
                _im.logout()
            except Exception:
                pass
            return (_po, _nep, "")
        # --- 1) samo ZAGLAVLJA, u paketima (brzo) ---
        _kand = []          # [(broj_poruke, adresa, ime, naslov, datum)]
        _KOR = 200
        for _i in range(0, len(_ids), _KOR):
            _grupa = _ids[_i:_i + _KOR]
            _set = b",".join(_grupa)
            try:
                _ok2, _raw = _im.fetch(_set, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
            except Exception:
                continue
            if _ok2 != "OK" or not _raw:
                continue
            for _st in _raw:
                if not isinstance(_st, tuple) or len(_st) < 2:
                    continue
                try:
                    _pref = _st[0].decode("utf-8", "ignore")
                    _br = _re.match(r"\s*(\d+)", _pref).group(1)
                    _htxt = _st[1].decode("utf-8", "ignore")
                except Exception:
                    continue
                _ime, _adr = parseaddr(_zaglavlje(_htxt, "From"))
                _adr = str(_adr or "").strip().lower()
                if not _adr or _adr == _nas:
                    continue
                if ("mailer-daemon" in _adr or "postmaster" in _adr
                        or "no-reply" in _adr or "noreply" in _adr):
                    continue
                _dom = _adr.split("@")[-1] if "@" in _adr else ""
                # cela poruka se skida samo za pumpe ili za isti domen (npr. knezpetrol.com)
                if _adr not in _skup and _dom not in _domeni:
                    continue
                _kand.append((_br, _adr, str(_ime or ""), _zaglavlje(_htxt, "Subject"),
                              _zaglavlje(_htxt, "Date")))
        # --- 2) cela poruka SAMO za kandidate ---
        for _br, _adr, _ime, _nas_raw, _dat_raw in _kand[:300]:
            try:
                _ok3, _rawm = _im.fetch(_br, "(RFC822)")
                if _ok3 != "OK" or not _rawm or not _rawm[0]:
                    continue
                _msg = email.message_from_bytes(_rawm[0][1])
            except Exception:
                continue
            try:
                _nas_txt = "".join((_b.decode(_ch or "utf-8", "ignore") if isinstance(_b, bytes) else _b)
                                   for _b, _ch in _dh(_msg.get("Subject", "") or _nas_raw))
            except Exception:
                _nas_txt = str(_nas_raw or "")
            try:
                _kad = parsedate_to_datetime(_msg.get("Date", "") or _dat_raw).strftime("%d.%m.%Y. %H:%M")
            except Exception:
                _kad = ""
            try:
                _ime2 = "".join((_b.decode(_ch or "utf-8", "ignore") if isinstance(_b, bytes) else _b)
                                for _b, _ch in _dh(_ime)) if _ime else ""
            except Exception:
                _ime2 = str(_ime or "")
            _z = {"od": _adr, "ime": str(_ime2 or "")[:80], "at": _kad,
                  "naslov": str(_nas_txt or "")[:160],
                  "tekst": _mail_telo_tekst(_msg),
                  "prilozi": _mail_prilozi(_msg)}
            if _adr in _skup:
                _po.setdefault(_adr, []).append(_z)
            elif _z["prilozi"] or _z["tekst"]:
                _nep.append(_z)
        try:
            _im.logout()
        except Exception:
            pass
        return (_po, _nep[:60], "")
    except Exception as _e:
        try:
            _im.logout()
        except Exception:
            pass
        return (_po, _nep, str(_e)[:200])


def sb_knez_odgovor_set(mesec_key, idk, zapis):
    """Trajno zabeleži da je pumpa odgovorila (ko, kada, naslov, tekst, imena priloga).
    Sadržaj priloga se NE čuva u bazi — preuzima se iz sandučeta pri proveri."""
    cli = _sb()
    if cli is None:
        return False
    try:
        res = (cli.table("obrada")
               .select("reakcije,trebovali_tip,njihova,napomena,reakcije_ko,dnevnik")
               .eq("mesec", mesec_key).eq("sistem", KNEZ_SIS).eq("idk", int(idk)).limit(1).execute())
        _r = res.data[0] if res.data else {}
        _dn = dict(_r.get("dnevnik") or {})
        _lst = list(_dn.get("odgovori") or [])
        _kljuc = (str(zapis.get("od", "")) + "|" + str(zapis.get("at", ""))
                  + "|" + str(zapis.get("naslov", ""))[:60])
        for _p in _lst:
            if (str(_p.get("od", "")) + "|" + str(_p.get("at", ""))
                    + "|" + str(_p.get("naslov", ""))[:60]) == _kljuc:
                return False                      # već zabeleženo
        _lst.append({"od": zapis.get("od", ""), "ime": zapis.get("ime", ""),
                     "at": zapis.get("at", ""), "naslov": zapis.get("naslov", ""),
                     "tekst": str(zapis.get("tekst", ""))[:800],
                     "prilozi": [{"ime": _p.get("ime", ""), "vel": int(_p.get("vel", 0) or 0)}
                                 for _p in (zapis.get("prilozi") or [])],
                     "upisano": _now().isoformat()})
        _dn["odgovori"] = _lst[-20:]
        _row = {"mesec": mesec_key, "sistem": KNEZ_SIS, "idk": int(idk),
                "reakcije": list(_r.get("reakcije") or []),
                "trebovali": bool(_r.get("trebovali_tip")),
                "trebovali_tip": _r.get("trebovali_tip") or "",
                "njihova": _r.get("njihova") or {}, "napomena": _r.get("napomena") or "",
                "reakcije_ko": dict(_r.get("reakcije_ko") or {}),
                "dnevnik": _dn, "azurirano": _now().isoformat()}
        cli.table("obrada").upsert(_row, on_conflict="mesec,sistem,idk").execute()
        return True
    except Exception:
        return False


@st.cache_data(ttl=300, show_spinner=False)
def vraceni_mejlovi(nalog=None, dana=30, _v=0):
    """Pročitaj sanduče i nađi poruke koje su se vratile (Mail Delivery System).
    Vraća {adresa_malim_slovima: {"kada": tekst, "razlog": tekst}}.
    Nikad ne baca izuzetak — ako ne uspe, vraća prazno."""
    import imaplib, email, re as _re
    from email.utils import parsedate_to_datetime
    import datetime as _dt
    _out = {}
    c = _imap_cfg(nalog)
    if not (c["host"] and c["user"] and c["password"]):
        return _out
    try:
        _im = imaplib.IMAP4_SSL(c["host"], c["port"], timeout=25)
        _im.login(c["user"], c["password"])
        _im.select("INBOX", readonly=True)
        _od = (_now() - _dt.timedelta(days=int(dana or 30))).strftime("%d-%b-%Y")
        _ids = []
        for _kr in ('(SINCE "' + _od + '" FROM "MAILER-DAEMON")',
                    '(SINCE "' + _od + '" FROM "postmaster")',
                    '(SINCE "' + _od + '" SUBJECT "Undelivered")',
                    '(SINCE "' + _od + '" SUBJECT "Delivery has failed")',
                    '(SINCE "' + _od + '" SUBJECT "Undeliverable")'):
            try:
                _ok, _d = _im.search(None, _kr)
                if _ok == "OK" and _d and _d[0]:
                    _ids += _d[0].split()
            except Exception:
                pass
        _ids = list(dict.fromkeys(_ids))[-400:]
        _rx_adr = _re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
        for _i in _ids:
            try:
                _ok, _dat = _im.fetch(_i, "(RFC822)")
                if _ok != "OK" or not _dat or not _dat[0]:
                    continue
                _msg = email.message_from_bytes(_dat[0][1])
            except Exception:
                continue
            try:
                _kada = parsedate_to_datetime(_msg.get("Date")).strftime("%d.%m.%Y %H:%M")
            except Exception:
                _kada = ""
            _adrese, _razlog = [], ""
            for _p in _msg.walk():
                _ct = (_p.get_content_type() or "").lower()
                if _ct == "message/delivery-status":
                    try:
                        _txt = _p.as_string()
                    except Exception:
                        continue
                    for _m in _re.finditer(r"(?i)final-recipient:\s*[^;]*;\s*(\S+)", _txt):
                        _adrese.append(_m.group(1).strip("<> "))
                    _dc = _re.search(r"(?i)diagnostic-code:\s*(.+)", _txt)
                    if _dc and not _razlog:
                        _razlog = _dc.group(1).strip()[:160]
            if not _adrese:
                # rezerva: nađi adresu i razlog u tekstu poruke
                _tel = ""
                for _p in _msg.walk():
                    if (_p.get_content_type() or "").lower() in ("text/plain", "text/html"):
                        try:
                            _tel += _p.get_payload(decode=True).decode("utf-8", "ignore")[:8000]
                        except Exception:
                            pass
                _svi = [a2 for a2 in _rx_adr.findall(_tel)
                        if not a2.lower().endswith(("loopia.se", "vapeshop.rs"))]
                _adrese = _svi[:3]
                _dc2 = _re.search(r"(?i)(5\d\d[ \-]5\.\d\.\d[^\n\r]{0,140}|550[^\n\r]{0,140})", _tel)
                if _dc2 and not _razlog:
                    _razlog = _dc2.group(1).strip()[:160]
            for _a2 in _adrese:
                _a2 = _ocisti_mejl(_a2).lower()
                if _a2 and _mejl_ok(_a2) and not _a2.endswith("vapeshop.rs"):
                    _st = _out.get(_a2)
                    if not _st or (_kada and _kada > _st.get("kada", "")):
                        _out[_a2] = {"kada": _kada, "razlog": _razlog or "poruka se vratila"}
        _im.logout()
    except Exception:
        try:
            _im.logout()
        except Exception:
            pass
    return _out


def smtp_test(nalog=None):
    """Proveri SMTP nalog bez slanja mejla: poveži se i prijavi.
    Vraća (True, poruka) ili (False, razlog)."""
    import smtplib
    cfg = _smtp_cfg(nalog)
    if not (cfg["host"] and cfg["user"] and cfg["password"]):
        return (False, "Nije podešeno — fale SMTP_HOST / SMTP_USER / SMTP_PASSWORD za ovaj nalog.")
    try:
        if cfg["use_ssl"]:
            srv = smtplib.SMTP_SSL(cfg["host"], cfg["port"], timeout=20)
        else:
            srv = smtplib.SMTP(cfg["host"], cfg["port"], timeout=20)
            srv.ehlo()
            srv.starttls()
            srv.ehlo()
        srv.login(cfg["user"], cfg["password"])
        srv.quit()
        _por = ("Veza radi — mejlovi će ići sa " + str(cfg["from_email"])
                + " (" + str(cfg["host"]) + ":" + str(cfg["port"]) + ").")
        # provera da li kopija može da se upiše u Poslato (da se vidi u Outlooku)
        try:
            import imaplib as _il
            _ic = _imap_cfg(nalog)
            if _ic["host"] and _ic["user"] and _ic["password"]:
                _im = _il.IMAP4_SSL(_ic["host"], _ic["port"], timeout=20)
                _im.login(_ic["user"], _ic["password"])
                _fold = _nadji_poslato_folder(_im)
                _im.logout()
                if _fold:
                    _por += " Kopija se upisuje u folder „" + str(_fold) + "“."
                else:
                    _por += " (Folder za poslatu poštu nije pronađen — kopije neće biti.)"
            else:
                _por += " (IMAP nije podešen — kopije u Poslato neće biti.)"
        except Exception as _ie:
            _por += " (Kopija u Poslato ne radi: " + str(_ie)[:90] + ")"
        return (True, _por)
    except smtplib.SMTPAuthenticationError as _e:
        return (False, "Server odbija prijavu — pogrešno korisničko ime ili lozinka za "
                + str(cfg["user"]) + ". (" + str(_e).split("\n")[0][:160] + ")")
    except smtplib.SMTPNotSupportedError as _e:
        return (False, "Server ne podržava traženi način prijave. Probaj port 465 uz "
                       "SMTP_USE_SSL. (" + str(_e)[:160] + ")")
    except Exception as _e:
        return (False, "Ne mogu da se povežem na " + str(cfg["host"]) + ":" + str(cfg["port"])
                + " — " + str(_e)[:200])


# Tekst mejla koji ide objektu uz Excel predlog porudžbine.
# Može da se promeni iz Secrets (MEJL_TEKST), bez diranja koda.
# {objekat} se zamenjuje nazivom objekta, ako se upotrebi.
MEJL_TEKST_DEFAULT = _cfg("MEJL_TEKST", (
    "Poštovani,\n\n"
    "u prilogu Vam šaljemo predlog porudžbine za Vaš objekat, sačinjen na osnovu "
    "trenutnog stanja zaliha i prodaje u prethodnom periodu.\n\n"
    "Molimo Vas da nam trebovanje pošaljete što pre.\n\n"
    "Za sva pitanja stojimo Vam na raspolaganju.\n\n"
    "Srdačan pozdrav,"))


def _mejl_tekst(naziv_objekta=""):
    """Tekst mejla za objekat; {objekat} se zamenjuje nazivom ako je upotrebljen."""
    _t = str(MEJL_TEKST_DEFAULT or "")
    try:
        return _t.replace("{objekat}", str(naziv_objekta or ""))
    except Exception:
        return _t


VAPE_LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAAoAAAACUCAIAAACMdw/QAABJOElEQVR42u19aXvcPK4swZ1a2pn7/3/jeeNubdzA+wGSWnYcx27biZ2wzjwzOYndi0QRBFCoglIKq6ioqKioqPi94PUSVFRUVFRU1ABcUVFRUVFRA3BFRUVFRUVFDcAVFRUVFRU1AFdUVFRUVFTUAFxRUVFRUVEDcEVFRUVFRUUNwBUVFRUVFTUAV1RUVFRUVNQAXFFRUVFRUQNwRUVFRUVFDcAVFRUVFRUVNQBXVFRUVFTUAFxRUVFRUVFRA3BFRUVFRUUNwBUVFRUVFRU1AFdUVFRUVNQAXFFRUVFRUVEDcEVFRUVFRQ3AFRUVFRUVNQBXVFRUVFRU1ABcUVFRUVFRA3BFRUVFRUVFDcAVFRUVFRU1AFdUVFRUVFTUAFxRUVFRUVEDcEVFRUVFRUUNwBUVFRUVFTUAV1RUVFRU1ABcUVFRUVFRUQNwRUVFRUVFDcAVFRUVFRUVNQBXVFRUVFTUAFxRUVFRUVFRA3BFRUVFRUUNwBUVFRUVFRU1AFdUVFRUVNQAXFFRUVFRUQNwRUVFRUVFRQ3AFRUVFRUVNQBXVFRUVFRU1ABcUVFRUVFRA3BFRUVFRUVFDcAVFRUVFRU1AFdUVFRUVFTUAFxRUVFRUVEDcEVFRUVFRQ3AFRUVFRUVFTUAV1RUVFRU1ABcUVFRUVFRUQNwRUVFRUVFDcAVFRUVFRUVNQBXVFRUVFTUAFxRUVFRUVFRA3BFRUVFRUUNwBUVFRUVFTUAV1RUVFRUVPwuyJ/9QymFMQYA9Rq9EH/witFbvwqf6s7e8Pm/Ov6yJ6sUxlipN6Wi4nVL7h/c+yoqKioqKj5dBlwT35cf+o/nmB+u4cdewv3Y9Ka3ORy+6h3/JAWAz3wjHqW5wIDBX/pY16eh4g9mwKUURGSMCcHrYnxJAEbEUgoAcP672+qlFLpfpRTEwljJOT/aRQA4AOMcCJxz4Pwz3NdSWClIf/inzgAAe7SFL/4YFFa+bBOhMAZ0sGAADACgAIO651X83gz4UeIbgp/nBQDappFKbQu14vEuSlcu5xxCiDEKwdu2pcv4Pknqr+oTiBhjDCHGmFIKKaWcEfGJAMw5CMGFkEorpZRSWil1DMJ/pPiBmEOIKeW/oIP44tALAIxzzjkAcCGEEOKXN/qPZ+c/fhgsmDMiYkFELAVLYeXrtrQAQEqplfr9Z+iKfz0A7zEj5+y9H4ZhmiYppRCikZLX+uSzMZiu2LwsUghEdM4JIT8iw9w3wZxzzphyDCH6xVMMTjnmlHPOWB7X0YDRjs+lFFJJpbTWRhtttRJCCrFGg99/+VLK4zgus8dS/pF9DwAYMA50Q7iSUipJzxrfwvLnKUQ/PO0VxEz/nXNOOcWYEde/K1jY1wzAWBhjTHDeNE4KwUUNwBW/LQAfan/e+2mahmG8XAbvvTFGKSOkMkbvEbo2C7fMYL1sGXGcpvvzeZpnDuC977u+7TrrjHinkPIoH4opTdM8T7MPS/AhxogZcykFCsOfUooBkbESImczWyOxVNZoa13jnHOGb3nY7+lh0xuklIdhHC4DFuQcGMDfnwkDY4xxBgyAcxBcCCWVkloZa421Rim9n97+VCr84/vmnBcfvPchhBRjDDFjTjkjUg+Eblv5WrcPALAgFmCsKCkASts0j5ZoRcUHBmCqG1ERdRzHYRimaVkWn3NiDC7DQMdzJUW9WE/uU3RqmaY5xMgBEMecMebcJmetlUJwId5YQnhQnwhhnpdxnOZ5DiHknBCRFWDAQHBOm/qTH5WxQm1ixFJKCMC594tcfAg+hOCss1orIcTv3HdKwZSS9x4Lcs7/nT4wbG0dAM4FF4JrpZfFWGutNcYYKaWU4k/FgIe1lhRD9CHMS/B+CSHmlFJKpRQs5UuPUQAAImKhx0umlAqrUyEVvzEAl1JCjNM4Xi7DPE8xxZwZF5xzXViZpkkqYa1VoqlN4EdH41KKD2EYpmmacs5SSM55KTjNcwhxmqambdqmaRqnlbrtvRBx70gF74dxGsZhnpcQ4kb74oIDMVJ3WurTu3YpwHkpRQixb5op5TSOfvHjODaN6/qu71sh5J4FAcAHE0MBAIQQgGxNwf+RGFzWMgpjDLEUTDll7/04jlpr51zTNG3bai0frrqPJek+5oLEOE3jPM3LvPgYY0yIiIgAQD/KOV+p0V80WwQquhTGmBACave34ndnwAxDCOfL5ft/3zMiF5yDkEIwgJRSjHEa59GNWkq9FaIraJPKOU/TPAxDjJGiiJQi55xSWtIcYggxsFK0VjcHYIq+OWfvwzhczpdhmqaYUilrT5dzvkbfct3UnyQ0cX4ldwIwOvZjKZizzz6EEFMkBlfbNkopzn9H6x+ACSGEkhz5etT4gNb5HvE+0+swuv4rhbggFswxhxh9CCHEEGJKqWmcMVoIeXj0PpAQuZ8sY8re+2VZxnGYpyUET3SrjUYPnAu29bMfTiP92QzytZ8EkPNSkDGo3d+KPxCAGSuIKYQQYhRSCi4fPY0hxPP5LIS4E3dUiK7dEYIP8XIZxnHMmfLUsk0iAWMcEVOKKUW2pZuvum6IhWrJiDgM4/l8HofRh1Awc+CMw6Erv241wAAPm045ULHolM+3ButeNYQ1xkMpGGJM54v3fvH93d2dc44D7WLlqd3tI1PDH3fR5/fV53+evfh3n9+3n//JV0TncvyYhRUAVhgA4yAYKwVL8cGHGBc/u8ndnU5t22qtDmWJ94/B+3pjjC3LchmmYRyXZY4xFsRSCh3ItsGpI3t+bfw+MVFVnr16t9QMfnVn4U3vDrX8XPGbAzDnXEpJgykMHlBh6Z8QyzjOUiprrRSuhl56mlPO8zzP80zpL+ecqCjAGAdeoEjJrTHWmB+HTF6W+64cpWmavn+/P18uIQRWmBScC7HXacsBFI7prwGgPOBCF8YYwp4BU/xe+8UAgAgZMcaYUsyItCO3jV0zbIYfFW1LoZpBKYX/k4KUOw5JGdLpLedI8S/n3PedVgo4h4+JEtt6S37x52EYLuO0zDFGxpjgnHPxaDiHVhepBXzp0SNq5TDGMjDMyCoqfmsABmGM7ftTSnma5pQSjUPQKAvnvJScUpqmaRhHKYQ2Gl6fz/0doK8MwHLOwzBehsF7f9hG1/QES0Eszpq7093d6aS2+vMLL9d+YXPOl8vl+/395TKsW6EUj9IfCrTbGGYBRvVBiqzUUmUFsRSGK9uEEp0HJy3agDgAE5xoZf/3f/+llAD+1zYOgNP3/qCriqXknEthpTyVR8ILsplfZKKHg8hL8t2f5lg/JFkvz7yfeoftkj64tLucCwCUgimm8/kSU8w5n/rONY5++x2fvv2lYozn8+VyvozzFGMqpQguGDD+8wmIQme1a43nV5WJt557f1WfOH6CB5f1p8cOWnuMlSw4IlYOVsVvDcDAQCnd913OOca0LAsiAvB9f+ecI5YY4zAMSkqhhBLyH79qdDXGcUREzsVWuwNSoUJEIYRzru9769wNp3KKvtM03Z/Pl8sQQqAkRHDxKOulri8wJqTgXEghBBdciq09zLbQm3PGnJFY04iklgUAbGdNb5s+yzkti2cMhJDAoGk+sOyxqR9oLEVw/iFF7l3o6F1e52eB/LWvhTQzWzJizgyAAWcc+KE+wQoTOaWYMo4TKwxY4UJo/c5UDDrDhRiHy/D9+/0wjjElDkAH8WOcvmqurfzttXwi1tnlm67LW8rUT5SdfxKAf/JeRN2gp0hJeVuxqqLi9gBMB1ujdde20zinlEjbBgA4AP2rEKIUnMZJCmGtkY38d8vQpeScl3mep9EvC2MghFh3SwZUUAXOnXNt22qt2WuqBfuPIeI4jvf3Z2J4cc6FEGyrKtOOuSW+TAiulbLOGGO00kpKIQWJTa77JhbEnDLGEEMMy+JpmrMUBACxSvBdX5xzwRiGEO6/f2elcM6ds9sG9j53fr8aQoimaeiQJzjbO9bPRM8f653P/zxjL/3dJ/bwH3d7eG63f74WC9ekH1PGTByBmBImhkwwzsR2ZSgmwBoGMedxnInH3ve9NWb7qTflwfuvLz7c33+/v78sy5JzFtfTW9nD6l5ryYisMM6Bc66UEkJIKYFzsdZd4PG9uC24vvIc9fw6eeYilb0CxLl1budh1V5bxe8IwPtqM8b0fZtzGqcRU+JCsO08KARPCWOM0zSP4yilNPpfZERTx2helmEcvfc00sNXFiWjwFkQtVJd1zZtwwW/YYukIvD5PJwvlxAi5bJCCOq30bvgyosRSkprjXOGpkdpQ3xS2hARU0w+RmO89wuNEceYKP/aBZh2fnVKaVkWzrkxWkqplPwIEpaUsmtba00phQNjBX4S954tZn4xEhYrDFNCmjIIIfjgNw1RpHXGYY1kaw5aGKY0LYs4n4FzIYWS6l3Wcyll8eF8Pn//fj9N876kd04DiXXT/zBgAEwrKYRQSimptFZCSCklCZ0CPJRR/iIkrI1EBkrJqkNZ8VsD8B4elJR93yHmGMOcF7Y25a6FJsZYCOF8vnAhxN2dFP8QI/raJ0vpMoyXYYwxw8pZI8YT0mFaSOGc6bu2cZZ2sZecpo+zxSHEYZguw7AsnlIgdhjQ3NUPpBCuabq27brGaMM3HcMnN5stleJSK2dNzu28+GkaL+czNR0KMACx71tUUUQsi/fn80UIcTr14gPuuBDcOVOK/vcqKaWUkjOmFL330zxP47z4kHMCgCIk7Lx66thLgYjDOAHnSuuulYK/Q+7rffjvv+/39/fzMhPVec996We2z5loi9BaO2eNsdZapaSUaidA/AX7wKcSAa34hzJgetisNSm1RMVKGdlBBYJ2cCzEiJbOWuH+OUY0Ii7ej+O4LL4wJoTcj9mlMLJAsMZ0XeusveEoTbPF8zwPw3DMsHeiJuW+jDGjddM0Xdd1XeN+aDOXpyqt1LQTAIJzpZTSWikpOL9cLvO8pJwZYzTcSaCcGAuO4yilMEa7D7jj//iWpxRjzFhrtNZa6XGa5nlJMSIiK2WNaqUArLeGVEjNMCqpnDU3XzoqdcQUL8Nwf3+epomqEVsVZGUX5Jzpz6RX7Zx1zjrnjDFam5osVlS8TwDeH0xjTN/3OeMwjiR5A1snGAAwZ5LmGNwkuDRGPTpQ/8W5LxWfx2FcFp9zFkJwDuT/QpkuYtFadV3XdZ2UD1wuXo4Y19lipOj70C+S1BqMMXd3d6fTyVqjlHxye/116sl545xWyhjzf//33zCOOWdWGFUS9ztecgkxjNPkJieE3KZR6yz4e0II2baNMbZpmsvlchmGeV4youAcgB/nsIFBjGkYRiWlWvsCr7sXRxXu8/35v+/3yzIzVjg/yl7CSh5ALIUpJdum6bq2bRutNR0K692vqHjnAEwpV9u2KWUfwsaIhi0lAl542RjRUggp+z3S/N0AgJTSNI7jMKSU90rdnhlvXDbTda21jsEt3V9Kf6dpCpu01q6DsXv9aqO7rru767vuse8he1mt+5jjaq2JAMUABvpqALsrw3awwBCIAy+kPL07U3SX7/qXwbnQWmgt1z5CYbP3hW4MXMOvlBIR/TyPSrZNc7NYdEac5/n+/jKNIyW41E+5dn0RCyscQBvdNE3f913XGGN+UWj5+o95DQkVfyAAH6IsN8Z0XTfPU04pI5JOG+3FNJWUEcdxEIJbp6UU/4ZTcCGzinleiSp795eIV4wxY0zTNM45OpS8cHfac82c8ziONHSEpUgaClptl0gyHrVSd3enu9Nd0zRwUNV4yxYjpby7OwHnOeM4TgVLAdxkfgsD4IJjzsMwSMGtseL1g1W/+EiMsbrxbRfDOSsEBxBwPk/zjJh3StR23mIxpWVZ5nlWWiqpXjVfDsAQcZ7n8+U8z3PKecuzWbkS7BELcgDnzN3dt77vjdE/HrxquKqoeOvJ+6nDOBWiO+r5pZSvptwkWQ4sxjjP0zhOIYS/+1Hc2WcUfWNKbOemUvpbSkbknLdt03WdulX2OaY0DOM4jjnnRyU+av1yzmm2uG0bMlR4YwpC2S3F4K7rTv3JWgvAqB+8g4aPY4jjOI/TnFKum+/7n+/Wem/hnFtrT6f+1HdGqf3v99VIay+lNM+L9+H1awBSSsMwDsMQc+QbAZ5tZ0ZabGSOe3d3d3d3atuG2sPH4fOKiop3y4AfHWmVEn3fp4whxpzzUUB4mxqFENLlMsiVEP0XMqKv3bKcyCM5hMQOtPBNh7kwYMbovu/6rhVCHGUKXoiMZVn8OE3L4tl6yiH2DWBBCsnO2r5rnXN7LgLv4XJI0Ere3Z0Qc85x8WH1uCkFNp5UKcz7eL4MQoi+74SoquDvf87br6W1BrEPIeScY0oZkfQu9gcwZ5znxVqz0/1+eS82KkNZFj+M47z4gnsthwHAGuoLAoBz7v/9v//1fb/PstdT17tvL89f10cHnTde/H2w/G0f9Z3XwPOvXA720h/zvvCM/zj80Bz7iPUvfxZ4jDFtuzKiSUppl8fiHBjbNaKFtdb9vYzoUor38XIZx3HOWDi/Dm4VxjBnxphSyjnnnBPkmlzKy2uqFNuWxY/j7H3A62xxOaa/Sqmua7quVe/adD84OoC1pu/aeZljSjv/bv0xzjkAljKSHKk1Toi6Gb93DL5uDXQ7uq6LMaVxJNLfoUbFqSqzLD6l9HIeRiklhDBN8zL7nLPg4qphvkqoIufgnCOWH0VfxPJ3TBl9wiPX44hDO/8HbPc36pR98MHriaB7nH195LP1Me/783f4Hc2x5x5da/Sp7zCncZwwIxewBQbgnKWEIYRxnJ2duBDmlapPnzvoXoWNQkzTNI/THGIE4EJw0nymDDhnlIK3jeu71mj1yne58lGHYRiGIaVEA5V7RkI/KQS3Rrdt55zb24EfMBHEjDVd16YU53lnevNjlzqEMI6Ta1op5E7ALqxaRb/PHTg+8ESHDDEt3qeUCmLhnB04+SmnEENMyf5qJzmqi4/jOI5jSolKG2yTl2GMUaXbGPO/b3d3d3d7M4XzenvfOfd6pMSOOeNW2ycJQnhI83yfLJPKdex1W8dRouC9QvGT6ggk4ktLEYALAZzLRzLp73Qa+Il0+QsC8HvJ2v46AG+M6CalFEJcGdH8ejojn4YQ4mUYaEzwL1NS3Qp98zCMMUYAttFhVgcYqtgppbquIz3Fm3WvDvSuQ3eZsZIzY4yqEc7ZD528JDZs17YppRBSShngMAjOOS+YEvoYhmFQSgjRijoJ+mH3AgCUVs45rTV5fjDE7XDGqTKSUkopIebjAPczSDFN00QrTT70e6aNT2vdtm3f9UR4ri2GD8q9SLOWtNlJjQ4LHrRgSfxulXff/+aGe4GI3vsUE65uaUUqZbR+CVWllJJSiinllHLOrBS5Gee9cavfv0jOOeccY0ZMKWXEvAdgLrgUUkohpSQpwHdZijljTDGGiJhfkglzDpwLIbngclMd/uAM+EdG9DSvZcmCKyN6WygCMZNQg3OGFCFeU3/9vBsg3Q9SKpimkfY4enY2CzMGjCmlrHNt0xqtAW7RvfIhjPO0LEvKSawBfp0tpudHSd023ZHe9RFqGPsdd86ljPPkiXyHWNYcaK18ECP6IgU3Wglr2d9yyz8hOIDWyhi9zCrntEsWUxJDjaEUY04J1HO789799cEvi48xkvgo2wa+KR5Ipbqu608nfR03qrf1nQ9V9Ge/+HGa52UJMeSUaPJrT87IzkxwLriQSiqlrDVam5crZR5rHsMwjsMQUyaloLZt//ftjlh1v9ydQgiXYZjG0XsPjFlr+75vus5sZMCbj2iFsRTiOE3LvPjgU045IcmeMsaAA2NcCi4k11o7a5vm8RTcbbcgpTQM0/l8H0LgXAjBD7vXAzlZ2GxXaFzTWts2jTHqfR+KX3SP1hjcd2RMmxIeGEBMCJFTijHO4zQ2Ix2P/pqC1TqVO07ekyQkP9QuCs2HOOe6rjPGwK26V9M0D8MQU9w9elfvXkRqxGqjuq49Fp8/dtMnJ4mujSkuy5Jz4pyvzoaMCSFyziH4cRzb1tFZuG7SHwfBudZaaZ2XjKXwh3cfsVASLNUv9oVSSoxxWUKMkdy6gAPV03Yfa021nN+10v5NUB/nfL4MwzgvS8wZU9qtvHdVauLccc6lFFpL713btl3XvbYGRkF0HGcfQ06Jynh93708XC3LMgzDPM/AWM5ZaW3ePIiYcl6WZZ7mcZyWxYcYck6YsTyoMwMXXHCutPTLEmNs2tZaK/ib0tCcMYQwDOOyLFJK0tgv5QHhaw9wsJb+uNLKLj7F2LaNdVbwd6v1yl9mRUrJU9/nlGMIOfldF4LtPril+FUjmp9OJylerc7z2U6pACwjjtN8uYwheETcFQ+271VKKUqqrmv7vr1Z98qHQLpXOWXKSNjWk6PUUynVNK5p7G8QPNk/vJSy7/ucU0opBH80oN1HUb3358vIuey6lsrmdcv+oPOQUlIp6f3VOn5jbq4shJQLYhGC/XgXjjpui/dkdrTuLmydCabXVEo5ZxtntZIA79/rqrkvPdfjNJ3P5/P54r2nhixJn9CWWtjebWWsYEbSHoQYEwB3zlEZ7LZnbeNq0But7/T8PV79Ttf/AJZ3sEtOKQ3j9P3+fhwnsjlfDcsfaP6tUSYjlhBiiPOyuGn+9u2ubdqDHt/rlyiwg7rR+h/Meb0BD82+AYBzhliCxxjCssyLX/73v/+1TfNe3UD5kqVDPch5mlJKWFaG7L5BsJUfOwkpjHHCia+/EcPu+LtN5T7SvWJCCGNN2zbGmJ2W9dpj4DzP8zzHENchk/IgwxZCtI3rulZr9duuKH0RZ03q2nmeU4ql4ENBNA4AGZEY0cYYIXTdZD8O1AnjnOeH89nsavaMv3wR6gX6EBDLo72DRsyNMU3jtFL1FPVxiDFehuH+fF7mhW6EFEIKwfnqgIUFGGOYsRQkZ+4YE+er7Pxtt4Zz4EIIFGQ4Bq/pJVP7WQgphGRsZUa9LfrmYZjuz/eXy2VZfCmFC84ZvctV4nT1UMVSsKSEBXOMMabMGCtY+v52uQUAxjmIFUDWXyAlnUjpuFGgUJhGXL1v6CmLMSAWKZXgwjn7Lk/Ki/IqYMzZ3axwPvJjaeHknEOM07SM4yyENFrefkL5DGUimsodp2VZGFtnJbduGeaMnHNrTdu2zhr+ekoh0buGYboMow9+ryUce3KMFaNV13dt2whxY4Z94xol3rV1bdumlLx/ghGdUl6WRQjpmkYqIUVlY31QBkyFOEGHPLYNoNNWsmpWvUAbAxFDCDFGLOW4BVPaIaRwTXOdo6sjvx9QW6LR7XGY/BIY40ICUIvBGGsUeSpjYQUxpZxSyjkvy5JT4gA06nlb1lUYK7DlfAxKgUf54POfvRQoDLaNHNYa+U1ALNO8/Pf9++VyIQk2pFfnXElpjFFSkQ4umWX7EGKIJdOhgeeE5/MFM3LOu669rSgIrABjwMjcugjBpJRGK60VFwLLWn4AxgpiCMn7hdJ0qorHGC+XixBUl7q9GvG6AEzmd5tGdPQ+ID7ixxYqFQzDIKWUspVfkxFNg9ne+3GcKPAQB+/Hqdy2bddFcNMUToxpGIZ1vusHZemtJOjaptFK/5ECr1Ky6zraC1IKuwYTJcGcl5yT934YRqVk27jqjfNh5yHgz/rJv2QLTtRRSIke2EcTGFJK66wxBiqj7oMKS4yllLwPIcScUQihpBKCW2ParnVWCyHZxpXLa/zN86wXrQTnXdcdFXi+5BUoZfHLMAzDOPkQBOdCcMFACmGssYbszCXx+RFzTIkm3ZdloXVL7dsJwBgtpWjeXAemNp9WsmmbtiF5wWsAxoIhpGXW0zwH71POjEEpOM9eqblr22eIbO8WgI9SDNa6nHGaFuLHEkGWb7OJXPCc0zAMQghrlVwb9V9mRvTo+DsMwzhuU7ls78uup0jOQWvddW3jttGjV47W5YzLskzT5JeFXXWvGIN1PkFw4axr205r/Zvj7l4CEpy3jaNjOMlPEltny4MZ53THL1JwrZQxf+ag8E+E4LfdTUQMMcWY1hkPtrYead1zzpVSxmi16U2+6h3/rDLlJ19sV0YxYkwppkQXLGeUshC/9bQVVMvKvL0qfnZd65ellGKMOW73b1bFWoeOocCPXd1yaIRti+T2jfygdpCGYRiGS0pxe5eilOra9nQ6OWfJj2vbf6grjH7x9+fLOI5kzwoAKefLMEgptNYkFHPLngNr9KXt1xjtrKXyzz7uXApzDtum0ZfL9//+yzkVBmtfIIQQorX49uGoV2TxnHNrXd/31LxMOfND4VFwnhBjDNM0jqOVUhGb44sd0xC998T6o7nY44GJRDmNMW3bkGj+DftFKcUvfhzH4FctwB8y7KK1aLuu2TSf/9QuI4RwzrVtm3Lyi885HztAxIimIeamcVKKv2wQ/NNs4ohvC3LbxHA+rqV9e5VCKKnUrcf5et56YQZMQhPbFbvu/s+M1WqtrNHbKOC7VZg45+sE/1NaU0e2x3v5Tm5k7HGeZ/rWrKCUouvaU3/apW2Pn3G9Akox2IfukMYxvF+mSXZdq96DsgAAggPtw483QMaVlKXkZZ5CDBmvu3TOiaYJ3novXvWMET/21PdSSuJus4eMaMZK8P5yuYzDkHPaCbSf/JS6X1bv/TCO87yQ6cJuh8w25Q0hRNu1D1gAL/aioT+klC7DZRiHnPLejTu6oK/0rq5x9up5/kfyYKrPnHpyOBaIeWcn7mfVnHHxy+VymaZp52pVvf73AmJBzHnzu3yYoBBHBvhT9Zf9BmAp1EZAzJuM3eY8yBgAKCm1UuL18y3XT8LYH7zfX2ipwZpfMcH5Nhi2+MU/8yWEEFJKGpi5eROA611jx17Sr3cBzt+oubePI/sQQggkbAwM6HD/7e7uqej74Nfbprm7OzWNoxoAFtyq00tK6dX5ydXRgK0Hoc2D86dJqpTGaCpGloKkkEQHgrdvdPJVj9wqUdu20zzFlAoWhJ/wY6Uwxjj3ZZJgKm6M0657dd2n2NXxt2j9pqlcRFyl8Oe1+MxgC/+slIyMMWtM2zbWmD/eVS2lcA7O2ZjSMk8xxo0Cf73jpbCU8zCOUgqtDRWiK94RKWNKTzOdd3W2X6y6hJif2iwQOV/3dw6vXmyISNzUvbL9g7Q9vHuIvL5iKQxACqGVluqzu5JzxpQQQnDGkJXMuGSMxRjHcWKM+Ri01lKQNxXfAh//mLN3yRlzRnZo+B/fYR83z4g557fHGJonpuY3IuO8cMGNUU3TNE1D0ZeiMoMf9x8uBG+cnZuGBBlTpn5woYj+9iLwNkmAj6RF9j9wzqVSQkjGNvMxWNf/22/GqxcuACPV/pzyNM05IzG5r7NupYQYx2m2bhZC6q/DiPYhXoZpmmbqgO7M5+2kA1rpxrnGOSUlPJJO/9XBZY2+PgzjtMxLzklwyTlcda9KyYhaqa5rT12n5G9lPj9zw6WUjbNd06SUgg8ZdxF/BgBS8IR5XhYhZdO0UkpRGdHveQbClOJ+0r+K1BO5FUCSYiH/oTu4PW+UQdNOSi9QtrIN/Y2UQioBL9bPuTL5ES+XYRjHlNKejn907eMwppKBQePct7s7KQXNkHy2DebIodFaa62k4AHo3q0ytDHFYZi0llpJKZXWUimtlHpH/gdNu+6zyMuyADBgvFxnyh/8MOW9JJcRY6SDOGPsVWPAx0gWY4oxkuI1IhOCO+ucbY4Vvme+LOe8cTZ4H2MovhQGGVmMOcZkbXn0di85wZVffJOy/ffer/moFSJffy+LlKJr25QwhriEB4xoCsUpofdhGAYlpZTtJ28NrtyruE3lxkiDYuXBVC4KIVzTdl1n1C1TuVuGPQ3DEFMC4Mctj+oZAGC0XovPW/j/DJeI9K5TzvcppxiBPWBEQyk5Z7/4YRillE3jal/wvZARQ4ghBMbwx6vKOaf89fnQ89ysMGdc8NvOTAVxWfzlMsZIpuBcCI4fHIFhVVIoOSc6TXRd9yVuJRnHNY1NKcaEKceCAMAwFFa8lEIrIaVUWmmllVLGGKWUUvLJ9uRtOx3nnMhQi/dbhv2TAMwAMYcQaQiHc35zwoeIKeWc918vNHeufjXfcey5aKXNVhTkW/Pr6FV/+1Uh5e1DN3BPNfdnMMa4NyX3cvW73BT5mtt3YEQ71+Uyz3PMqRSaVmZ7Hiy4wJyHYZBCWKPcZ2VEP3CJGYbhMgTvGSu0NPf686p7pVTf913XiVfnpuuPhRDGjd519GGlgXPGmNG6WccxJTWMPsPRZLXl6LqMOC8+5sSOjOh1TgZSWifktFbvMiFXgaV4H0m9mQQEDoyEtZknlRQre5n9zFYWy8N+FflY0wP5BqF/duigbSulfPSi3b8D/V9Zx1s3gaNPudzWBwHAWns63RUG5/MQQqA9k2oapWCIJabsfeBiBgAlldaqaWzbts41j+qirz22lG2SNWccx+nRWik/PPjb/SywRe5SbtTBKoztK5DedtUGORwsfqlkzoUQQnIuOHAGhQrpb+rCXguYK5PiZz8YfFy8DyHQjpdzIR8EIfiHjyE9UxNwzlBWNE1zSpnvGtGMCcFTKjGEcRzbxiqlpPzMzeBCulfTOFI2f7yspHslpXTONY0jFbRbdK9iPGTY/MfZYiFk2zab6cKnu1hXRvTWztnLnhwYFyLnPC+zHETTuHc7s/+rWKsyIc7zHEIoJZM28MpIKHhMf19wqcszZ0MO/IYG8CEFF4gCgHZp/jv5d9QC/0KGEVKqtu0KA1ZgmufjdB9iwYKYcdOjBMm9EML7lWpk7fvMAeNmPPTc+ihrR58Eqt7+vqvAxXHlwOs6kqtLIxz7xPguS62UklOOMdEc8NGNFrF4Hy7D6H0gHbqtLU2uB783A350VJFS9f0pZQwheu85f6wRXQrzIZwvFyFEf+oE/5Qa0aWkTI6/0+IDB+Cb6tDuEsO5aBrXde3uefyqky/A6kkyXC7Bh/0a0iVa8wbGjNFd33VdK4VYRwI/x4W6MqK1PvV9TnifzzEG4HK1kqUviZhzmpeF8mDnHBUSah782jXDNqOOcRy3eXT+sF5CvVuptVbXSblXJYCwbrKboPt+i15zsyivewdr6teStbbxfMZYeYm1+id5ghhjVEjTSs/zMi8LyZOllBjL+2QqkFjKugamFGNK+XTKXbcShm94pg763qsRFvzsbAb7LCwcx0DeqXTxdA3yRU/Gwcf3rZ9o7YgzxlhKaZ6XUpgQchXkZmv49Skt07L4kFJim0G7Utpaq7V+ox/UmzLgVSPamrZt53kiw/BH/FhgDDOO4yiE0MY4++mSYNrO5nkehsEHX0qBdSoXD+kvKqm7jiQhb3SJiTENwziOE5byMMPGg+6VbZzbXO4/Y7mejJJiTMs8pxRIIfqYDJXCU87DMEpJJJI6FvyqK0yjEasQgd9caJ6syojVIs38BpeO5zeyw25IhMJy08s8eEnGnphP/aGqWQDZl5t3o7kv2UqtjbVm8SHGQCJlOWNOKSMiZrqUmDMizosv7EJ90zcmo6WglMpaS/XTNd97uNesWw+HUlja8BbGLzDG9yPkuoBLfsxLKL/kMeSVrYxUF6fOyc2fimSxaXOecFqWQNv7vhgRMaTkl5AzcmAk+aqUatum7dr3qlO+6ekVnDm3akTP04I586tiMAPOMSXvwzhOzs1SyH129o8zog+6V3kYxmEYU8q7EjjNe5H0lRDCOtNtc0EvdPw9vhNpwE7z7Ml04XCGXeldnG+mC/qGXOZ3QinZtnaebMohxEyHrmvvX4iccVq8HEfnGqV0TX1fviC3uVxgjPlluVwu40j8Jn44a3MyzqGOgLP2BV002DPGH3JgauLuLNnXJyWkErKq98ErB8DLNfY+cHB6QTq7hfqyGcgyxuDrqGgCMKWkEI21lmrPFILTNuG6+LCyn4QsiD7EaVq6LmhtOH9NyaFc/4OlcFakFHd3J2ct56IU/CkJCzj5sQ7DiDSHjnijVAsDIfiVqw9QCsYYYwxKq62MWx6a8u5LkTLfnUedqWrIOXB+w2x0YYUVLKTeSH3flFKKsfzkkcRSWEFcHzreOPftru+79mDL+7uUsJ78PkqKrmtTyilE7xM+UAwGznnKxft4uQxSylMvPtWMCqW/JPvMaCr3cEajBM9Z07XtbVO5a4a9LMM4eurhC3HkE1I5USnZdW3bNOIr+LAqpfq+Szmn+0vK+TiHyrmgseB59uM4aq2sNbX+/MKlss2qFR/85Xw+Xy4+BirKHLhXiIjAmJTSWWvMiy4v/FwtZifI3Jy173bChbGCcNvUb/khD3s+t929Bb6u4gsAEw/55zljzinGaIxR8zyOYwiRMZYZ5JxjTCGGnDPnt27aZLwhhLW273sAVrD8SD4qBal5RFND3od5XhDLzY8x51wruZdqOI1C+cV6bVfC6TMnNwBgmMvi/bx4OvGvCm5SqhcRIH7+wRhnjGEpOWf8wWqMrSJZHAQH4IJz6+zpdGrb9mYvpncLwLsvLOfcWZc7XOY5pYSbjOfKj+UgGMdVmkNao52zfzbHO07lzvM8DsO8+Iy4S0KWUjhwyk2N0V3X9V0ntysOL9a92pWlL8NwGUaKVbD9E67ygiCFsNa2bWtuy7B/Y4RYyx5Cdl2fMs5LSNO0KyLtTUogRvQwSCmEkErd2LX6uErvJ1RP2i/OPM/3l/P5/n5ZFkR2dOJCpHylSCmtNc5Zuak3P0siZRvNGQ486Kux0hs0fUAIoZUEhpQBA/BXdnL3ucxjAP51R7gAI+kudZQQ+eJHPSE4gBRCaK21MaUwxCHGtEk3l5wxY5LlVumJPbflq4Xbk9Tf3Xp147htlWq4cUkDgNZKa82FoHdHLMu0jFK17VXw4KCr9ngXTSmN00Rae5TAcC6MfjAq/bprAtc/cDoKkew/HP+FsZV9zbXWjXVt1+yWzH84AD864DjnyDlnmpeUkAs4nPJESjlsjGitlJDijz8pAOT4Ow7DkB+mcVtiyoBzMkI21sJNU7mkbTmOo99sDY+rDHPmAMZco++XqJQCgJCyaZq2XWJMKREj+urKKDnHUuZ5llI41whhP9VX+0H551NcVcRM0/OXy+U8XObZl1I4f6ATXgrLOXPgzrq2bV8u1LAtb3jqrRFTwpxuffAtsJLWX4ejyvGrAvDDTPYFlCxgVHo2xr7vhvjRJbecM2OFA+dPlU/pSRFCAOfjIH8Qdi3vwoe69lGfSoSOKcq7iD0xYHSqUEpxwQttsBmX2Q/DyDlX6ulElg6dNLg8T3OMgR2m77TRb/EjQlYAkTOupJRKXacJtqMGDbYB50IKrbWz1jn7LsSrdwvAjzSiU8YQk/eeLKX2nha1VEkjWgrR9x0XfzgrIg2aYRjnecGrH1HZVL8LABhjmqYlwsJtGfayzOMwLMuSc96HxqiTR/0FqXTXtV3Xyc+ie/VSaK27vkspnc9DygkAAPiqBQ3AciZu4TBchIC9Uvqb73g57Oyf+cLmnOZ5Gad5mtZBtcIY5xIeN1ULK0Vp1a5r5rku1IEOSfbjkrTdHhW9SaWIWDYvXOeHp170fdu27tao8Ij685oAvEcs4J/hQP/SeliMwzCklIwxzjlj9M/OgjnRrFA5TkNw4GQZ9Jb9+ocH4qeL51103bdJEG6Mds75ELwPiIUziCF+/36fUjyd+p+4HZdlmS+XYbgM5ApFAZh4yMYYKuC/alcpbL2kBUuGLFBoq7q2a9pGKUXb8vGQvtaO6Cz8Acr88r2Wl7G2a9M8z+kJjWgA4Fjyyoi2xgrxpx4Z+lQhhHGcFj/nnEnK75Bn7I6/+1TuLe+SUhrHcRgGonc9yLALEtXLGN22rTGWfalxHWJEt86lmPzi8xRxbU2u351vWvPDMEgppJR/JE2BH8qSe8G8MPZn+4f0SXLOKaXg/TTPwziTBTVJ2WxrsrDVPDwjopSqaVzbtXSmedmaASGleKpbRkXpmFIIt9BcaRSKVbwYZOe+LN5aE2O01kilOBfXIFuoJZnGcVqWJedUEEtB2HyTpPhDF/zNO5OUsmvbSO7iKQLwlHOeZsREYtHaaM7FcT4kxjCO0+VymecFkWjIKw+5797YiN3PFiCksFa37Yvchd99l3632yk4WEca0WleHjGiAUDknHwI4zi6phFCbvM2v68bfO0o5HwZxstAE5Zr2R+2MS+i3mmjT13bd1dDwNd6bngfhmGc5rmUB528/YkyVrumcc7K3Yfy0+NRX6drm2WeUloV+XejFSIvIOI4TVxIY93vDMDlqayXqn8554xlbXn+sQtesJScMKcUY/AxxBBiTDHl/fiy577UkCMxPyl52zV3d71z9uUDGKQXTYKGR9OtvU4TU/YxpJxNjZAf+dSQd0UIYZomH8Lig1KSYuouXFNKSTnHEH0IMcSCWDADK1JJa40xr589XWe91/YA6VgAKySN9ULW+CaBcUvmd1TT67oWEYNf8tbYZox5HzPej9NktBZScsGBsYIlpRRiiCHGFHFlvDMqTJ5Ofd+1UopNzO014++rEdKqCwZ8/X9fuwF+ugDMWFFSdl2bUoox+pTYA0Y0L4WnlHyMl2GQSvWiFX+iNVhK8SEMwzjPMxbGDzq6yBj5ESmlGuecu0V6Zs2wYxyneZ6XmLLgjx1/S0EpSfeqVUp+3W3FGN13bUrxfBmpzL4bXNJXTinP8zyOs1ZKG/17Uvz9XZAmKTOmnEKI67BlwXV05u0qss+Uvp8PwFhyWimvKSfMGQujOa7d1pOO5ytBCpEarqdT37a0Zl5xcCW9KrG1gcu1vAaILOUUQgoxOSyvHaz8DKfGL1I3KrQxhhBCCCylGBPpTBFXkW+qNRlTiGn1Hkckzl3XNf3a9b9lx8OyT9CyJ0ePfvHrW4ELEcsbCIxCiLZpwqlnhc2LJy3JnFmIgfNlIiMX2nIRU84xps3hFDgw4h5uj8DtB/pSaBC5wPbtyPfpl22yj1hp8r0egJUR7VzOeZkXksmmUdot/ADnAjOOwyiFNFo5a9fr8cGP0NGXI8Q4TdM0zyEmITg/qjUVhpilFG3T9Iep3Fdn2CkP4zgMQ4iJHXWvcFMqKExr3Xdd937zZH9kyyON6JjzEsI80YjetimvfpssxnQ5X6Tgd+JOyY/t/T8ynF8WP8+zX+YQQoyJ0t9tfPQgq7NXqp+gAh2/9UsjbfnJz5friGsphdF6ACjAhVj/an8WVj5TxpwTCsFd23z7djqdTkZrYLCO2v/qMu5abEpKqaQQfLMlLDtvriCmFJdlscYYrV7Vra8zZi9fmXl1ZS5bIsrI5zkjAqTVZKIgzZ6uVYpSpBRN4759u+v7/oa9ouzTrGUVzsGS99bGL09xhQbFjwG43HLq2j+w0urbt/8JIb+fL+MwhRgL9bmx5JQxI4O0f+6MyMo6FiWEdG3z7e7Udb29ep7ecB4BqvOTAWJBZFg2YSz4/av6ndMvisFt18aU52XesqLrCSjnnRHttFLiNzaDSd6PdK9ijLAptJDuFa3QUlbnn5sdf3dl6XmesRQpHvD0yBJEKuWcc85RR+fr6Ac8sblrrdu2naYlxbXEe6UqcC4BEMs8z1IJ66wUH2uUtFtrpJQWv4zjMs+zXxaiav+Q8x7JPuv2d3j8PuJzPtCYICmAnVazH1xK2bdNZKUoKVzjTnd3fd8ZY277WEKss5g5REQUh0wXAEhI5xiAKz5gcXJ6WISQWHY7ARIzwVXYGDZ7DOL6SmGN7rqu67rX5gOPi9Cb4Ojr7+9Vq/StzrtlJbcCAAMuhZjnJcZUkIqPiJsaJ2wMg03EQ1hju747nXpyHH/LIR7YXod+UqHmt0K+797HNkZ0zphS9CEQ3f3QRgW2MqIHLsQuKfJ7KEgxhuFyGYcRc+IrLXmfsGQATCvlnGvaK8nltQeilNI8z9M8hRAYA37g6RVGtoa8aV3Xter6RBXGvvCup7UmaY7hMpBJ3Hp2oXDCMKQwL/M4jVKKXU/7g+54imlepstlGKdp8RFzLoikrgvAgD33pttHfsOneqFH9FMZ5IPKOc2rANNKnbqu709d3yqt4dbUk3OhjTFGx5gS5lL4dvJb9cvmeTZWtY27eaOveP5Ga637vldKe++99+Rwl1IiicXCV5vm1f9AKmu0ta5prLX2ZrLbrhnOAKREAKaVFlzAEwzFx6UeIgMqpYyxiAyAGWOkkAD8eIi8JQ9W6u7u5Kyb52leFu/9SsUvqy8cGZ5KIaWSdB2cc8aY55n/L3oQgEnJrVZoDWNgjJFKCf7HRHPfOQOmXdVa23VpWqaYItUu4CCPVWCV5hBCWKN/m1swpb/jOMUQGMCxA018VJL367ruhQJDPy6vXVk6+EC8g133iqIvsfi6rmubhn8F3auX3HFq7cSU/eIT3XH4CSOa+LhSfNAnCSFM43QZhmG4+BA24zAuuKApew7wk2hZDtXo11oD/CyrfvZftxr4LnjHtio07l4LRjfOfTud2rbZjmtlFaR8fWnKWmutn6aFxcgO5Ax6NELwyzx538m3qQtV/OzWcw5kguucJYvJlFNKeffVY1SRI7cdKY3R1lhjzba8b6mTUUmSMZ5yJu6nteaFDVQAUEo1jeNcOGcBVh+CN1LfadNTUioptZbGGO99TJFksB8EYCnX62CtfqeDuxDcWHMqvbWWMSaEMkYrrf7Umv8QBpAQ3FrbtV2K2dMU7EExGEBgzt77aRxH54he/6FlUio+D+N4vgw+BCxMHKguewFQStn1fd93rz1nPZrz25Slr7nv9nwVJaWztmtaozVNA7Mv20g7jCeCMbprm3kaU0qIq0v2QSOaY8ZxnKQQ1hopm/e9v/THZfGXy+X+/n6e55wTY6vwziGgwkOCNDwbON9yU+DX/wpbK/iqxowFcyHpPq0b57qubRpnjRGH4ZPbcg7OudHWmqDkEEPY3ve6lyHmZVlWD2/ntqeDPTMtWnFLBsZBa02zFXh0ItwsJjcaPDxSB3qted+2FYu2ba29zmrTXPgzLwcPsnZFnJitO8M3wNv3DUqFhRDO2t0weL0OjAO/XodjH/Mtu+UuImSNxW3vfeSH/Zt344+i4Colu64jneuUEiAeW4OslJKS9/4yDFKJ/uPtY2NMwzCM45Qzci6OGmx044UQztm2dcZouOmcRb4l4ziTsvRRHIoxlnMSHJwzHQ1x/l1JBrV2+r7LKQ/jmFISB39ozkXCFEMcx9m5WSmtlXzHhU657/lyuf9+P04j9aGloLt85WTtouofWQ946V65RUAGwApjnAEH4EoKIZXWzrm2adu2If9p9h4V+1XjUGvvFyzsWKggB7AY8zAM69T2avZS9rNLjZ3vmqKIl6/tt0QFajHcXGV8dAh439ToGP8++jocLgjfqdafAe8cgPdsknPeOJdz8iTNUdjOiKZWG+c8IQ7DIIQwWlNB4CNKsoUxzLgsfprmZfGMwSr8WQr5ctB+vZkuaHjNXnPVvco4L8s4TvMSdmXpvWC46l5J03Vt37fqq+leveQ8q6Q4dV1OGEKYj8rm13wPQoiX4SKluDv1b+z977WNwor34Xy5fP/+fRwnxjbiGzx2xymF7aLHHxSGXx6AHxYnuRRCa0XVNmOdNUbKB7X6t1DDdi601tI5473x3mdEwa80CCEEIo7TAnww2nSd2GSG6OuUm9+94l2er3/8e/2t1+EDh1DJlb3tu5hxlWM85J1ciJwzSXM0jaNaxLtfZRLfW/2IvN/7skfHX9K96rqu61opxU1m1xDzOnr0pLI0InIQxpi2bYwx5Nv2Ny2pAyO6maYpxoTr2MNW2OEggJdSpnGSUlpjnePvcgVizOMwns/naZ4RUUq5hna2zvntYBuvhOof8DHX4debxZEEvU2CKinJ39dao5TZS3zv2qQoUoq2bUNYCUCw9YDZ5mRFbKzz5cI5b5rmU3mX/U144fnvXe77j+/1qpd946+/y6V473dkj1gaf3Ar/pAAfGS79f0p50I+l4UJ9oARzVgp/qgRvfsRvbnugVvRm+TfzpeBJt+vjr+b9p4QwhrTttfRo1evUcZ8CMM4zvNMMX7PV4jQCgDG2KbprHWc/7029dQM7jpyLyFeG1lL7Z14msMe3SgEf0t99WhpdX++H4exYBGrOxDd4t3tpxAPTipplFJaSyEeWdy/X7nleTP5g20U5xxWmolSVCYUj9pR77U17JP61BFMKc3LElOi48nRBphzSCmd7+/J/LVtm/132Zd3G/pH89o3vtdHf9TfH/w+lR3LB2bApRQO4KxNXZrnKaWEpcDDzhMRlMZxFJwrJa21hw3oTZ0nir7kpHG5DN4vFGuvilSMlZwZY2R59BM18BetnhDCNI6kg73bGj7MsHXbkoC++ou3lZUR3TYkI7eQf+eRAF84TacNw6AkJaviLW/nvR+HYZqmlJKQUkpBnQ7GqOWBiIwBKKW01tZqa4xSWkpxIGe9dwhmjL0gAFOoFasMIdySSd+0+RDV37mGpMFyzpwLzmGjywnM6H1gbBBCMFZID66G3oqKLxaAt/ySO2JEp0xa8+KBRjSQNMdlGIQQ376xpmm20FXglWOZO6eUdrSc8/l8/n5/vywLifkdW9SssM10oe36lmxr2euZzznn4TJcLkN6qHu1F2ABuNam79qudULwW5Slv86JnnNurcm5nec5xUS8xi36Mpr+yjmPwyCFMFq7xt0Qd3eX0GEYxvESU2IAjF39bjdXKwQAo03XNV3bWmul3HNf+AyemL9T9G6H1rrr+5Ty5XJOiZys5JEfDiBCSN/vzymlu293XdvKq2nrF3CXqqioAfi6YW6M6JxSogf+6OvEOUfMIYTL5QLASmHWGnETL3oLeESyTuM4fb8/D8NIhrU/9mUBmNG66yj9vcXxtxTmQxiGYZpHGmR6ZADAGNNKNY1zzgoh/oUlxTknW46U0jRPKScpOGPrd+dC5JRCiOM4uaaRSit1S++/lBJCHMdxnj1j7NFsInkHcU4j6f2p79rWCfFJZbePbTD4eFlWznnbuJxSDH6kIdRViAZKKdStz5i996RMhBk/jqVRUVED8AdmRaQR3TQu57QsS4ppFUQu7Fj4yhnneckZY8LTqe/aTkr+5A71goyheL+cL8P5TMXJKyvqMJVbgBWtddu4tnEbLflFzIhyGBoNKY3TNM5zDBE4PzJ7SdxDKdm2ru+u8yR/MR6ooZ36lFOMIXu/q6GRUxKpLXofLpdBctGfulcxoolYl3NeFj/PIaZMlp2H+5szIufQNO7udDqdTlqbz8wn+s0eFQBgtOq7Nvol57T48Lh1AoVzQISU0vl88T50bdN3bdM2j3ooX8LC65Pfl4oagD8cxIju+zalPC9zSknwQ1bERSkMc168X3X4sDTOKvkELeWZoz0i5owh+Gkc7y+XcVow532w/VFflhihfX+7429eda/GGGMBoM+Km6chvd2blaW/HnbF17Zt53mOKSIWeOAPzRlARhzHUQphnLH8dV3GghhCWJYlprQpCMHegEBkUIqx+tT35PW93/fPdP3/5GcBAGPU6dRnxIznEALb7JPpSnIOAAIRV43AnHJKMSVrrVJKcEFKrjVEVVR86gC8P6Jaq9PplLGknEJY1vHCjacKAFzQAx/P57MPoXGucc45a14gV4lYYoze+3max2nyfgkhbsROfjyql5X5CcbYvj91Xfdqx99tNDKGeLkM0zhhRvFDhk3R1znrmtuVpb909mCtbbsupjTPC/X+YdfGAkiYvQ/jODVNI1+mhnadukacZ794T30EuqIPr7xum7bvV8G5a+D/53E4BommbQtjKePlciGrxj0P3lU5heCI6H1IKc/LYrRxzhprSFWx8rMqKr5ABrxrRPd9Dt6fMeWcI0t8JaMWxlbXesw5xphyJudMH7w1VmsplSTmKjwMhaWUnDHEEHxYlmWe52VZckaAVYrwx9yXAVhj+r5v21bcqmtKNfNpHL33VEU/fipcDVxd296oLP0X5MFSyq7tUkohpJQ8PPSH5ryklHz0wzDQHM7LAyQihhBCiKu0y8N/IqJv23XW2k0Fhb1FPO9vrVIIIZqm+ZYyK+UyjDFGxvIepHcJdzrZ0Oiw98EHbxZjjNVaS6WkEFJydpvRzue8OIwxVjgDTpI99YhR8dUD8LbtQuNs+nZXWLk/X1JKUqz0mV2QnuzZSL4fM87zvEnjSSml4AI2iUcyKidB8xhDiimmRDkQF/zRXkMGsGR4Za25u7v79u3OGPUoLXh5BjbN0zCOIQRE3D0Htx8opRSpVNt1Xde93cHjy+ZY3DmbcjdNS0rpkT/0phGdh+EiBNevUUPLiCnFlOKj0gUZnSolmqZtmtVoqxZKn7lHUsrTqedcMIDLZYgpsow0kLxf0NVRgzj/iPPivQ9CTKuV/DZOxoFfT8gvaA3/+CMvNlz+xW+9/HefujKslFwK01I656y1dfVU/CUBmOZEu7YtiCnlcZow5xQj55wEotlWJ6SGbsqppOK9J5kC8mnhhwBMrm0pZ8yZmq+7WPjxTSnxxVLIDOTUd3d3fWPtSvu8QfcqxmGYhnGVHf5BWRpXcY/mdmXpvwNUBui6lizfc05HHRIhRM6J1NDatnk5z5bWxs7dfXR3pJTWrZYv/+yVf9HziAicSym7rsOCAHwchxhDzhlLAQbA14Gtg2cDImLMOaYEIVAaTXpepOPx+wPw23/3UQBGzIilsUYIobWpamAVXz4AHxmYSsm+7wpjgvPLMMQYBSsSYFdsOOp0r7q9haWUyc3+obrvVdd35Wo9yETXyQoiJDNg1ppv3+7uTqfGuVWP8CZvmRDCNI3LPLOtcr6b7eScOQdrTNs01hr+r+7+B41oeeq7nBOpoTG2N8tXS6JS0K+D4LLr2iNf/cltli5yzoiYKTocVwLnXCmptJJ14/zlPdqeOCn5qe+1UkaJ8+U8zZ5mByR/cB4qD2pU67wfheSUMmN/5qzzvtQKAMg5ISJnLGcsdZVU/DUZ8L7EtdanvqNG0zhOOaWcMxTcOknwQ13xgajv0bd1JzkDsN3g7/hzmLGwIqXUWvV9d3d31zp3sxkRInrvz+fLPM+Ys1jNU8ueUhx0rxol5T/uJHPt/XftvNpyEBn5gRpazng+D8BASkFqaM+nTlTQ+PHW0AtKKUTlW71mSdNFa9sWGAMOnI+L9ykj5lwYMLgmwXztCjM6bv74bP4FARiRHAJLjb4Vf1UAfuQEeTr1QkitzOVyXrzPKXMh1umGJytUG6Nqfxl4oJX1wO11P5szasc2zenu1HWtvckK8OD4m+7vL+fzOYbIt8hPZq5U9eKca226rm2aZmt2/usDhUIIa92PamiUM3HOY0zzPHNgzhpq+T95xTbVYixkK1gYwDo8s/OfORmaQw3AL8VBFYe5xgkpjXXjOA7DSDeL2FhU6Vkn4J+SswcAVlY17D+1q7zL1aBzBHBeJ4Er/sIMmK2DPKCU6nshhRAChnHydOgm8/pDDfnRH54+AG8J0OEUXoAxKaVW0jnXdX3fd8boRwH1hkgcU0pp9ZzPOW+jk4yspJUxTduQgm5dXsfz1kENLWx15j15Qrp1GV+VeMCP1QVeSatvyCNJSVQpqaWUUkzj7L1PafW2orC7D31d/5sD/C1lnpWwCcC5gFpHqfgrA/C+RwrBm9YpLV3TDsM4TZP3S0yJzFj28vK6R7CyJcdrKgSM4WH7WKN3KQBkP64aZ9u2bZrWaH2cOLrZjZJzbq0NjSuFxRgRERCFECT8L6Vs26brOqXkIUj8u3jgD924nHFZlu34UsrKWSsA4JztutY8cOP4abCg1HkDpxrpPj1T05Y35pFCiKZttNFdG8ZxmudpWZYQY85YCttoWYXz7QBEiS+1gQpj7DdlwY/Mnt/rKEJnd1bdjyv+1gB8TEM552b1qJFaqXlW3vuYIuZCOSWVkTeXGXiQ/xzV4VcGLOecU8fXWdM0TdM0Wq+JL25x/ebPLARv24ZzUFLP8+yDzzmnGCk/UEp1m7FS5d/+cOmEc65t25jSsiwhBADgXEgpjdFt23Rta19WOdh4WHkrjzBEzIjASs4iI/7+QujflAfTU6m11lorJY3RyzItS1hCzCkTSZjOTgB4LUd/Kpu3N1wDOiDmnHFzDa+o+AsD8DE+cQ5tY41RXWhD8MuyeB9CDDnltE33HlgeB3tSAA6MA3AuhJRKKWOMtcZZo5SiScXDu7y1psQ5t0ZrJZ21y7IMwzAMw7wsiFkp3TjbNo1R8qB7VdfY0R961YgOIfjFS6mM1V3fdW3rGqekfJHsKA2hldXsCDmwrd8PjFX2zHs9kgRjDNV1YoyLD8GHEHyIJJZDZ2SKUxsh64nCz1uGiV6Vuf7y1cpL3nftjCCyGoAr/uIA/MShm3OtlHPGWuu9DyGklGKMOWPO1B/Ea72JeM+ccw6ScyGFVEopbYwxxuhrEfg9eZJ7YZySbLmZqOeUjDVd0+ifEIgqSGHYORtjtywzsGKtOfWntu9b54R8ecucC6Gcc4gMAKTgjBUsSEouRiutNVTdq/d4Ktk2V80Y01oba0mfLoRIsjd5ZW3kKyP6CUfkLxaAc8ZS0DlTaRwVH37k/ZzpQikFMyltFBr5TCnT31x3B2AcOBdccC4EJ2ItcC4eCnF8KMjMeBrHlJJUyjlnneM1AP/kntLRxHt/f3+fUnTONU0jlX7V1BAZEc7zHGLc5RJZwVXQlHOtja0kuI97MDeqRcaCmbRwyELlOpT/1XvApRStlLVWa10lxCv+iQD8vNc3ZpIyLMefXAeB+BMh70Odwx+l1CnGnDPjIIUEzmsAfj4A0yw1Y8UYswtjvbxKQRVCpGnN668U2vZpZpVX59rf9WAyEoVdlWT/kgAMhcG6tfC6kCr+uQz4i25S9Vl98eXCXcSqoqKiogbgT3ru/mUi++P0/2/b2WvcrRfwn02If//jVlFRA3BFRUVFRUXFW1H5BRV/LJWql6CioqIG4IqK349at6yoqKgBuKKioqKioqIG4IqKioqKihqAKyoqKioqKmoArqioqKioqAG4oqKioqKiogbgioqKioqKGoArKioqKioqfo7/D0PMeQhpAstJAAAAAElFTkSuQmCC"


def _potpis(nalog=None):
    """Podaci za potpis u mejlu. Mogu se menjati iz Secrets, po nalogu:
    POTPIS_IME_1 / POTPIS_MEJL_1 / POTPIS_TEL_1 / POTPIS_ADRESA_1 (isto i _2),
    ili zajednički bez sufiksa."""
    n = _mail_nalog() if nalog is None else str(nalog or "")
    return {
        "ime": _smtp_kljuc("POTPIS_IME", n, "Aleksandra Apatović"),
        "mejl": _smtp_kljuc("POTPIS_MEJL", n, "nabavka@vapeshop.rs"),
        "tel": _smtp_kljuc("POTPIS_TEL", n, "+381 654 769 055"),
        "adresa": _smtp_kljuc("POTPIS_ADRESA", n, "Futoška 71, Novi Sad"),
    }


def _potpis_tekst(nalog=None):
    p = _potpis(nalog)
    _d = [str(p.get(k, "") or "") for k in ("ime", "mejl", "tel", "adresa")]
    return "\n".join([x for x in _d if x])


def _potpis_html(nalog=None):
    """Potpis u mejlu.

    Sve je u tabeli, red po red — Outlook svakom <div>-u dodaje razmak kao
    pasusu, pa se potpis „razvuče“. U redovima tabele toga nema."""
    p = _potpis(nalog)

    def _e(t):
        return (str(t or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    _ime = _e(p.get("ime"))
    _mejl = _e(p.get("mejl"))
    _tel = _e(p.get("tel"))
    _adr = _e(p.get("adresa"))
    _td = ('padding:0;margin:0;font-family:Arial,Helvetica,sans-serif;'
           'font-size:12.5px;line-height:17px;color:#444;')

    _h = ('<table cellpadding="0" cellspacing="0" border="0" '
          'style="margin:24px 0 0 0;border-collapse:collapse;">'
          '<tr><td style="padding:0;">'
          '<div style="width:190px;border-top:2px solid #e54fde;font-size:0;line-height:0;">&nbsp;</div>'
          '</td></tr>'
          '<tr><td style="padding:10px 0 8px 0;">'
          '<img src="cid:vapelogo" width="190" alt="VAPE SHOP" '
          'style="display:block;border:0;outline:none;text-decoration:none;"/>'
          '</td></tr>')
    if _ime:
        _h += ('<tr><td style="' + _td + 'font-size:13.5px;font-weight:bold;color:#111;'
               'padding-bottom:3px;">' + _ime + '</td></tr>')
    if _mejl:
        _h += ('<tr><td style="' + _td + '"><a href="mailto:' + _mejl + '" '
               'style="color:#444;text-decoration:none;">' + _mejl + '</a></td></tr>')
    if _tel:
        _h += '<tr><td style="' + _td + '">' + _tel + '</td></tr>'
    if _adr:
        _h += '<tr><td style="' + _td + 'color:#777;">' + _adr + '</td></tr>'
    _h += '</table>'
    return _h

def _ocisti_mejl(adr):
    """Sredi adresu iz šifarnika: izbaci razmake (i one pre @), < >, navodnike;
    ako ih ima više razdvojenih zarezom/tačka-zarezom, uzmi prvu."""
    import re as _re
    _a = str(adr or "").strip()
    if not _a:
        return ""
    _u = _re.search(r"<([^<>]+)>", _a)      # oblik: Ime <adresa@...>
    if _u:
        _a = _u.group(1)
    for _z in ("<", ">", '"', "'"):
        _a = _a.replace(_z, "")
    for _z in (";", ","):
        if _z in _a:
            _a = _a.split(_z)[0]
    _a = _re.sub(r"\s+", "", _a)
    return _a


def _mejl_ok(adr):
    """Da li je adresa upotrebljiva za slanje."""
    import re as _re
    _a = _ocisti_mejl(adr)
    return bool(_re.match(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$", _a))


def _napravi_poruku(cfg, to_email, subject, body, attach_bytes=None, attach_filename=None, cc_email=None):
    """Sastavi MIME poruku (telo + potpis + logo + prilog).
    Vraća (msg, spisak_primalaca, ociscena_adresa)."""
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email.mime.image import MIMEImage
    from email import encoders
    from email.utils import formataddr, formatdate, make_msgid
    import base64 as _b64s

    _sirovo = str(to_email or "").strip()
    to_email = _ocisti_mejl(to_email)
    if not _mejl_ok(to_email):
        raise RuntimeError("Neispravna email adresa u šifarniku komitenata: „" + _sirovo
                           + "“. Ispravi je u šifarniku pa pokušaj ponovo.")
    if cc_email:
        cc_email = _ocisti_mejl(cc_email) or None

    msg = MIMEMultipart("mixed")
    msg["From"] = formataddr((cfg["from_name"], cfg["from_email"]))
    msg["To"] = to_email
    if cc_email:
        msg["Cc"] = cc_email
    msg["Subject"] = subject
    try:
        msg["Date"] = formatdate(localtime=True)
        msg["Message-ID"] = make_msgid(domain=str(cfg.get("from_email", "")).split("@")[-1] or None)
    except Exception:
        pass

    # telo: obična verzija + HTML verzija sa potpisom (i logom)
    _txt = str(body or "") + "\n\n--\n" + _potpis_tekst()

    def _htm_esc(t):
        return (str(t or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    _telo_html = _htm_esc(body).replace("\r\n", "\n").replace("\n", "<br/>")
    _html = ('<div style="font-family:Arial,Helvetica,sans-serif;font-size:13.5px;'
             'line-height:1.6;color:#222;">' + _telo_html + '</div>' + _potpis_html())

    _rel = MIMEMultipart("related")
    _alt = MIMEMultipart("alternative")
    _alt.attach(MIMEText(_txt, "plain", "utf-8"))
    _alt.attach(MIMEText(_html, "html", "utf-8"))
    _rel.attach(_alt)
    try:
        _img = MIMEImage(_b64s.b64decode(VAPE_LOGO_B64), "png")
        _img.add_header("Content-ID", "<vapelogo>")
        _img.add_header("Content-Disposition", "inline", filename="vapeshop.png")
        _rel.attach(_img)
    except Exception:
        pass
    msg.attach(_rel)

    if attach_bytes and attach_filename:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attach_bytes)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", 'attachment; filename="' + str(attach_filename) + '"')
        msg.attach(part)

    recipients = [to_email] + ([cc_email] if cc_email else [])
    return (msg, recipients, to_email)


def _mem_mb():
    """Koliko memorije aplikacija trenutno zauzima, u MB (0 ako ne može da se izmeri).

    Streamlit Cloud gasi i ponovo pokreće aplikaciju kad pređe ~1000 MB —
    a to korisnika izbaci na ekran za prijavu usred posla. Zato se broj
    prikazuje tokom grupnog slanja, da se vidi da li memorija raste."""
    try:
        with open("/proc/self/status", "r") as _f:
            for _ln in _f:
                if _ln.startswith("VmRSS:"):
                    return int(int(_ln.split()[1]) / 1024)
    except Exception:
        pass
    try:
        import resource as _rs
        return int(_rs.getrusage(_rs.RUSAGE_SELF).ru_maxrss / 1024)
    except Exception:
        return 0


class MejlOgranicenje(RuntimeError):
    """Server je PRIVREMENO odbio slanje — dnevno/satno ograničenje broja
    mejlova, previše veza sa iste adrese ili slično. Nije greška u adresi;
    isti mejl će proći kasnije. Grupno slanje se u ovom slučaju zaustavlja,
    da se server ne bi dodatno zaključao."""
    pass


class MejlSesija:
    """Jedna SMTP (i jedna IMAP) veza za više mejlova zaredom.

    Grupno slanje je ranije otvaralo novu vezu za svaki mejl (prijava + TLS
    + upis u Poslato), pa je jedan mejl trajao i po minut-dva. Ovako se veza
    otvori jednom i drži otvorena, pa je slanje višestruko brže."""

    def __init__(self, nalog=None):
        self.nalog = nalog
        self.cfg = _smtp_cfg(nalog)
        self._smtp_veza = None
        self._imap_veza = None
        self._sent_folder = None
        self._imap_odustao = False      # trajno odustajanje (isključeno / nije podešeno / nema foldera)
        self._imap_razlog = ""
        self._imap_padova = 0           # koliko puta je veza privremeno pukla u ovom krugu

    # ---------- SMTP ----------
    def _smtp(self):
        import smtplib
        if self._smtp_veza is not None:
            try:
                self._smtp_veza.noop()
                return self._smtp_veza
            except Exception:
                try:
                    self._smtp_veza.close()
                except Exception:
                    pass
                self._smtp_veza = None
        cfg = self.cfg
        if cfg["use_ssl"]:
            _s = smtplib.SMTP_SSL(cfg["host"], cfg["port"], timeout=60)
        else:
            _s = smtplib.SMTP(cfg["host"], cfg["port"], timeout=60)
            _s.starttls()
        _s.login(cfg["user"], cfg["password"])
        self._smtp_veza = _s
        return _s

    # ---------- IMAP (kopija u Poslato) ----------
    def _imap(self):
        if self._imap_odustao:
            return (None, self._imap_razlog)
        if str(_cfg("SMTP_KOPIJA", "1")).strip().lower() in ("0", "false", "ne", "off"):
            self._imap_odustao = True
            self._imap_razlog = "isključeno u podešavanjima"
            return (None, self._imap_razlog)
        import imaplib
        if self._imap_veza is not None:
            try:
                self._imap_veza.noop()
                return (self._imap_veza, self._sent_folder)
            except Exception:
                try:
                    self._imap_veza.logout()
                except Exception:
                    pass
                self._imap_veza = None
        c = _imap_cfg(self.nalog)
        if not (c["host"] and c["user"] and c["password"]):
            self._imap_odustao = True
            self._imap_razlog = "IMAP nije podešen"
            return (None, self._imap_razlog)
        try:
            _im = imaplib.IMAP4_SSL(c["host"], c["port"], timeout=30)
            _im.login(c["user"], c["password"])
        except Exception as _e:
            # PRIVREMENO (mreža, timeout, server trenutno ne prima vezu) — ne odustajemo
            # zauvek, nego pokušavamo ponovo kod sledećeg mejla. Ranije je jedan ovakav
            # pad značio da NIJEDAN sledeći mejl u krugu ne dobije kopiju u Poslato.
            self._imap_padova += 1
            self._imap_razlog = str(_e)[:160]
            if self._imap_padova >= 4:
                self._imap_odustao = True
                self._imap_razlog = ("veza sa sandučetom puca (" + self._imap_razlog
                                     + ") — kopije se ne upisuju")
            return (None, self._imap_razlog)
        _f = self._sent_folder or _smtp_kljuc("IMAP_SENT", c.get("nalog") or _mail_nalog(), "") or ""
        if not _f:
            try:
                _f = _nadji_poslato_folder(_im)
            except Exception:
                _f = ""
        if not _f:
            try:
                _im.logout()
            except Exception:
                pass
            self._imap_odustao = True
            self._imap_razlog = "ne mogu da nađem folder za poslatu poštu"
            return (None, self._imap_razlog)
        self._imap_veza = _im
        self._sent_folder = _f
        return (_im, _f)

    def _kopija(self, _sirova):
        """_sirova = već serijalizovana poruka (bajtovi), da se ne pravi druga kopija u memoriji."""
        import imaplib, time as _t
        _im, _f = self._imap()
        if _im is None:
            return (False, str(_f or "kopija nije upisana"))
        try:
            _im.append(_f, "\\Seen", imaplib.Time2Internaldate(_t.time()), _sirova)
            self._imap_padova = 0
            return (True, _f)
        except Exception as _e:
            # veza je možda pukla — jedan pokušaj sa novom
            try:
                self._imap_veza.logout()
            except Exception:
                pass
            self._imap_veza = None
            _im, _f = self._imap()
            if _im is None:
                return (False, str(_e)[:160])
            try:
                _im.append(_f, "\\Seen", imaplib.Time2Internaldate(_t.time()), _sirova)
                self._imap_padova = 0
                return (True, _f)
            except Exception as _e2:
                self._imap_padova += 1
                return (False, str(_e2)[:160])

    # ---------- slanje ----------
    def posalji(self, to_email, subject, body, attach_bytes=None, attach_filename=None, cc_email=None):
        import smtplib
        if not smtp_dostupan(self.nalog):
            _n = self.cfg.get("nalog") or ""
            _suf = ("_" + _n) if _n else ""
            raise RuntimeError("Slanje mejlova nije podešeno za ovog korisnika — dodaj u Secrets: "
                               "SMTP_HOST" + _suf + " / SMTP_USER" + _suf + " / SMTP_PASSWORD" + _suf + ".")
        msg, recipients, to_email = _napravi_poruku(
            self.cfg, to_email, subject, body, attach_bytes, attach_filename, cc_email)
        # poruku serijalizujemo JEDNOM i tu istu kopiju koristimo i za slanje
        # i za upis u Poslato — ranije su se pravile dve kopije u memoriji
        _sirova = msg.as_bytes()
        del msg

        _greska = None
        for _pokusaj in (1, 2):
            try:
                _s = self._smtp()
                _s.sendmail(self.cfg["from_email"], recipients, _sirova)
                _greska = None
                break
            except smtplib.SMTPAuthenticationError:
                raise RuntimeError("Prijava na SMTP nije uspela — proveri SMTP_USER/SMTP_PASSWORD "
                                   "(za Gmail mora App Password).")
            except smtplib.SMTPRecipientsRefused as _e:
                raise RuntimeError("Server je odbio adresu primaoca: " + str(_e)[:160])
            except smtplib.SMTPResponseException as _e:
                _kod = int(getattr(_e, "smtp_code", 0) or 0)
                _tekst = getattr(_e, "smtp_error", b"")
                if isinstance(_tekst, bytes):
                    _tekst = _tekst.decode("utf-8", "replace")
                if 400 <= _kod < 500:
                    # privremeno odbijanje: ograničenje broja mejlova, previše veza…
                    raise MejlOgranicenje("Server je privremeno odbio slanje (kod " + str(_kod)
                                          + "): " + str(_tekst)[:200])
                raise RuntimeError("Server je odbio mejl (kod " + str(_kod) + "): " + str(_tekst)[:200])
            except Exception as _e:
                _greska = _e
                try:
                    self._smtp_veza.close()
                except Exception:
                    pass
                self._smtp_veza = None
        if _greska is not None:
            raise RuntimeError("Slanje mejla nije uspelo: " + str(_greska))

        # kopija u folder Poslato, da se mejl vidi i u Outlooku/webmailu
        try:
            _ok_kop, _kop = self._kopija(_sirova)
            try:
                st.session_state["_zadnja_kopija"] = (bool(_ok_kop), str(_kop))
            except Exception:
                pass
        except Exception:
            pass
        return True

    def zatvori(self):
        try:
            if self._smtp_veza is not None:
                self._smtp_veza.quit()
        except Exception:
            pass
        self._smtp_veza = None
        try:
            if self._imap_veza is not None:
                self._imap_veza.logout()
        except Exception:
            pass
        self._imap_veza = None

    # da može i u „with“ bloku
    def __enter__(self):
        return self

    def __exit__(self, *_a):
        self.zatvori()
        return False


def posalji_mejl_sa_prilogom(to_email, subject, body, attach_bytes=None, attach_filename=None,
                             cc_email=None, sesija=None):
    """Pošalji mejl preko SMTP naloga iz Secrets, sa opcionim Excel prilogom.
    Ako je prosleđena `sesija` (MejlSesija), koristi njenu već otvorenu vezu.
    Baca RuntimeError sa razumljivom porukom ako nešto fali."""
    if sesija is not None:
        return sesija.posalji(to_email, subject, body, attach_bytes, attach_filename, cc_email)
    _s = MejlSesija()
    try:
        return _s.posalji(to_email, subject, body, attach_bytes, attach_filename, cc_email)
    finally:
        _s.zatvori()

# =====================================================================
# ISTORIJA PORUDŽBINA IZ ADMINA (za detaljnu karticu)
# =====================================================================
def _admin_session():
    """Prijavi se u admin. Vraca (session, base_url, greska)."""
    import re, requests
    base = (_admin_secret("ADMIN_BASE_URL", "https://admin.vapeshop.rs") or "").rstrip("/")
    email = _admin_secret("ADMIN_LOGIN_EMAIL", "")
    pwd = _admin_secret("ADMIN_LOGIN_PASSWORD", "")
    if not email or not pwd:
        return (None, base, "Nije podešena admin prijava (Secrets).")
    _tok = re.compile(r'name="__RequestVerificationToken"[^>]*value="([^"]+)"')
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0", "Accept-Language": "sr,en;q=0.8"})
    try:
        r = s.get(base + "/login", timeout=30)
        m = _tok.search(r.text)
        if not m:
            return (None, base, "Ne mogu da otvorim login admina.")
        s.post(base + "/login",
               data={"Email": email, "Password": pwd, "__RequestVerificationToken": m.group(1)},
               headers={"Referer": base + "/login"}, timeout=30, allow_redirects=True)
        chk = s.get(base + "/orders", timeout=30)
        if ("/login" in chk.url) or ('name="Password"' in chk.text):
            return (None, base, "Prijava na admin nije uspela (proveri Secrets / pristup).")
        return (s, base, "")
    except requests.exceptions.RequestException as _e:
        return (None, base, "Greška u vezi sa adminom: " + str(_e))


def _norm_ws(t):
    import re
    return re.sub(r"\s+", " ", (t or "")).strip()


def _h_escape(t):
    import html
    return html.escape(str(t or ""))


def _clean_komitent_naziv(t):
    """Skini rep tipa ' - [mejl] VP' iz naziva komitenta (kako izgleda u padajucem
    spisku), da bi se poklopio sa imenom u listi porudzbina."""
    import re
    t = _norm_ws(t)
    t = re.sub(r'\s*-\s*\[[^\]]*\]\s*[A-Za-z]{0,3}\s*$', '', t)  # ' - [..] VP'
    t = re.sub(r'\s*\[[^\]]*\]\s*[A-Za-z]{0,3}\s*$', '', t)      # '[..] VP' bez crtice
    return _norm_ws(t)


def admin_build_komitenti(only_ids=None):
    """Napravi mapu {id_kupca: naziv} iz padajuce liste na stranici jedne porudzbine
    i sacuvaj u Supabase tabelu 'komitenti'. Ako je only_ids zadat (skup/lista id-jeva),
    cuva SAMO te (da ne prebrise nazive/kontakte iz fajla). Vraca (broj, greska)."""
    import re, html
    s, base, err = _admin_session()
    if err:
        return (0, err)
    try:
        lst = s.get(base + "/orders", timeout=45)
        _ids = re.findall(r'/order/details/(\d+)', lst.text)
        if not _ids:
            return (0, "Nema porudžbina za čitanje šifara komitenata.")
        det = s.get(base + "/order/details/" + _ids[0], timeout=45)
        _sel = re.search(r'<select[^>]*id="selectUser"[^>]*>(.*?)</select>', det.text, re.S)
        if not _sel:
            return (0, "Ne mogu da nađem listu komitenata.")
        mapa = {}
        for _v, _naz in re.findall(r'<option value="(\d+)"[^>]*>(.*?)</option>', _sel.group(1), re.S):
            mapa[int(_v)] = _clean_komitent_naziv(html.unescape(_naz))
        if only_ids is not None:
            _want = set(int(x) for x in only_ids)
            mapa = {k: v for k, v in mapa.items() if k in _want}
        if not mapa:
            return (0, "Nema naziva za povlačenje (lista prazna ili nijedan ID se ne poklapa).")
        try:
            sb_komitenti_save(mapa)
        except Exception as _e:
            return (len(mapa), "Pročitano " + str(len(mapa)) + ", ali čuvanje nije uspelo: " + str(_e))
        return (len(mapa), "")
    except Exception as _e:
        return (0, "Greška: " + str(_e))


def _parse_order_items(html_text):
    """Iz /order/details/{id} izvuci stavke: [{ida, naziv, kol, cena}] i idUser."""
    import re, html as _h
    _iduser = None
    _mu = re.search(r'data-original-user="(\d+)"', html_text)
    if _mu:
        _iduser = int(_mu.group(1))
    _blok = re.search(r'id="panelOrderItems".*?<tbody>(.*?)</tbody>', html_text, re.S)
    stavke = []
    if _blok:
        for _row in re.findall(r'<tr[^>]*>(.*?)</tr>', _blok.group(1), re.S):
            _a = re.search(r'/article/edit/(\d+)"[^>]*>(.*?)</a>', _row, re.S)
            _pc = re.findall(r'class="price-cell">\s*([^<]*?)\s*</td>', _row)
            if not _a:
                continue
            _kol = _pc[0].strip() if len(_pc) >= 1 else ""
            _cena = _pc[1].strip() if len(_pc) >= 2 else ""
            stavke.append({"ida": int(_a.group(1)),
                           "naziv": _norm_ws(_h.unescape(_a.group(2))),
                           "kol": _kol, "cena": _cena})
    return stavke, _iduser


def admin_istorija_komitenta(id_kupca, naziv, datum_od, datum_do, max_por=40):
    """Vrati listu porudzbina komitenta iz admina u periodu [datum_od, datum_do]
    (format dd.MM.yyyy), sa stavkama. Vraca (lista, greska)."""
    import re, html as _h
    import datetime as _dt
    s, base, err = _admin_session()
    if err:
        return ([], err)
    naziv_n = _clean_komitent_naziv(naziv)
    # granica: datum_od (dd.MM.yyyy) -> date za poredjenje na nasoj strani
    try:
        _cut = _dt.datetime.strptime(datum_od, "%d.%m.%Y").date()
    except Exception:
        _cut = None
    try:
        # Prazan datum na serveru = vrati sve, pa filtriramo kod nas (izbegava format datuma).
        resp = s.post(base + "/orders", data={
            "orderStatuses": ["1", "10", "20", "30", "40", "50", "60", "70", "80", "90", "95"],
            "keyword": "", "idLoad": "", "startDate": "", "endDate": ""},
            headers={"Referer": base + "/orders", "X-Requested-With": "XMLHttpRequest"}, timeout=90)
        body = resp.text or ""
        rezultat = []
        for _row in re.findall(r'<tr[^>]*>(.*?)</tr>', body, re.S):
            if '/order/details/' not in _row:
                continue
            _rowtxt = _norm_ws(_h.unescape(_row))
            if naziv_n and naziv_n not in _rowtxt:
                continue
            _mid = re.search(r'/order/details/(\d+)', _row)
            if not _mid:
                continue
            _oid = _mid.group(1)
            _md = re.search(r'(\d{2}\.\d{2}\.\d{4})(?:\s+(\d{2}:\d{2}))?', _rowtxt)
            _datum = (_md.group(1) + (" " + _md.group(2) if _md.group(2) else "")) if _md else ""
            # filter po datumu (poslednja ~3 meseca)
            if _cut and _md:
                try:
                    _dd = _dt.datetime.strptime(_md.group(1), "%d.%m.%Y").date()
                    if _dd < _cut:
                        continue
                except Exception:
                    pass
            _ms = re.search(r'labelOrderStatus[^"]*"[^>]*>([^<]+)</span>', _row)
            _status = _norm_ws(_h.unescape(_ms.group(1))) if _ms else ""
            _mc = re.search(r'price-cell">\s*([\d.]+)\s*RSD', _row)
            _cena = (_mc.group(1) + " RSD") if _mc else ""
            rezultat.append({"id": _oid, "datum": _datum, "status": _status, "cena": _cena})
        # najnovije prvo, ogranicenje
        rezultat = rezultat[:max_por]
        for _o in rezultat:
            try:
                det = s.get(base + "/order/details/" + _o["id"], timeout=45)
                _st, _iu = _parse_order_items(det.text)
                _o["stavke"] = _st
            except Exception:
                _o["stavke"] = []
        return (rezultat, "")
    except Exception as _e:
        return ([], "Greška pri čitanju istorije: " + str(_e))


def _datum_sort_key(_o):
    """Sortiranje porudžbina po datumu (najnovije prvo)."""
    import datetime as _d
    _s = (_o.get("datum") or "").split(" ")[0]
    try:
        return _d.datetime.strptime(_s, "%d.%m.%Y")
    except Exception:
        return _d.datetime.min


def _porudzbina_otkazana(_o):
    """Da li je porudžbina otkazana/stornirana/odbijena/poništena."""
    _s = (_o.get("status") or "").lower()
    return ("otkaz" in _s) or ("storn" in _s) or ("odbij" in _s) or ("ponist" in _s) or ("poništ" in _s)


def _porucio_posle_starta(idk, hist_lst, snap):
    """Da li je objekat poslao porudžbinu POSLE trenutka kad je zabeležen startni rezultat.

    Tačan način: pri beleženju starta zapamti se spisak ID-jeva porudžbina koje je
    objekat tada imao (snap['narudzbine']). Sve što se posle toga pojavi je nova
    porudžbina — bez obzira na datum, pa radi i za porudžbine istog dana.

    Rezerva (sistemi startovani starijom verzijom, gde tog spiska nema): poredi se
    datum porudžbine sa danom starta, i to STROGO kasniji dan — bolje da neko ostane
    neoznačen nego da bude pogrešno prikazan kao završen pa da ga niko ne pozove."""
    if not hist_lst or not isinstance(snap, dict):
        return False
    _baza = snap.get("narudzbine")
    # ako baš ovaj objekat nije bio u spisku na startu (dodat je kasnije), ne
    # smemo sve njegove porudžbine proglasiti novim — tada idemo na datume
    if isinstance(_baza, dict) and str(int(idk)) in _baza:
        _stare = set(str(_x) for _x in (_baza.get(str(int(idk))) or []))
        for _o in hist_lst:
            _oid = str(_o.get("id") or "")
            if _oid and (_oid not in _stare) and not _porudzbina_otkazana(_o):
                return True
        return False
    import datetime as _dt
    try:
        _sd = _dt.datetime.strptime(str(snap.get("kada") or "").strip().split(" ")[0], "%d.%m.%Y").date()
    except Exception:
        return False
    for _o in hist_lst:
        if _porudzbina_otkazana(_o):
            continue
        try:
            _dd = _dt.datetime.strptime((_o.get("datum") or "").split(" ")[0], "%d.%m.%Y").date()
        except Exception:
            continue
        if _dd > _sd:
            return True
    return False


def _treb_posle_starta(idk, hist_lst, snap):
    """Količine po artiklu SAMO iz porudžbina koje je objekat poslao POSLE starta.
    Isti kriterijum kao _porucio_posle_starta (spisak ID-jeva sa starta, pa datum
    kao rezerva). Vrati {ida: kom}."""
    out = {}
    if not hist_lst or not isinstance(snap, dict):
        return out
    import datetime as _dt

    def _dodaj(_o):
        for _s in (_o.get("stavke") or []):
            try:
                _ia = int(_s.get("ida"))
            except Exception:
                continue
            out[_ia] = out.get(_ia, 0) + _to_int_kol(_s.get("kol"))

    _baza = snap.get("narudzbine")
    if isinstance(_baza, dict) and str(int(idk)) in _baza:
        _stare = set(str(_x) for _x in (_baza.get(str(int(idk))) or []))
        for _o in hist_lst:
            _oid = str(_o.get("id") or "")
            if _oid and (_oid not in _stare) and not _porudzbina_otkazana(_o):
                _dodaj(_o)
        return out
    try:
        _sd = _dt.datetime.strptime(str(snap.get("kada") or "").strip().split(" ")[0], "%d.%m.%Y").date()
    except Exception:
        return out
    for _o in hist_lst:
        if _porudzbina_otkazana(_o):
            continue
        try:
            _dd = _dt.datetime.strptime((_o.get("datum") or "").split(" ")[0], "%d.%m.%Y").date()
        except Exception:
            continue
        if _dd > _sd:
            _dodaj(_o)
    return out


def _treb_na_startu(idk, hist_lst, snap, cutoff_date):
    """Količine po artiklu iz porudžbina koje su POSTOJALE u trenutku klika na Start
    (i datirane su posle preseka). Po tome se računa realni lager i DOPUNA porudžbine
    — i ona se posle toga više ne menja, bez obzira na kasnija ažuriranja.

    Vrati (map, ok) — ok je False ako start još nije zabeležen (tada važi živi prikaz)."""
    if not isinstance(snap, dict) or not snap:
        return ({}, False)
    import datetime as _dt
    out = {}

    def _dodaj(_o):
        for _s in (_o.get("stavke") or []):
            try:
                _ia = int(_s.get("ida"))
            except Exception:
                continue
            out[_ia] = out.get(_ia, 0) + _to_int_kol(_s.get("kol"))

    def _u_opsegu(_o):
        if _porudzbina_otkazana(_o):
            return False
        _md = (_o.get("datum") or "").strip().split(" ")[0]
        if cutoff_date and _md:
            try:
                return _dt.datetime.strptime(_md, "%d.%m.%Y").date() >= cutoff_date
            except Exception:
                return True
        return True

    _baza = snap.get("narudzbine")
    if isinstance(_baza, dict) and str(int(idk)) in _baza:
        _stare = set(str(_x) for _x in (_baza.get(str(int(idk))) or []))
        for _o in (hist_lst or []):
            if str(_o.get("id") or "") in _stare and _u_opsegu(_o):
                _dodaj(_o)
        return (out, True)
    # rezerva: nema spiska ID-jeva — uzmi sve do dana starta (uključujući taj dan)
    try:
        _sd = _dt.datetime.strptime(str(snap.get("kada") or "").strip().split(" ")[0], "%d.%m.%Y").date()
    except Exception:
        return ({}, False)
    for _o in (hist_lst or []):
        if not _u_opsegu(_o):
            continue
        try:
            _dd = _dt.datetime.strptime((_o.get("datum") or "").split(" ")[0], "%d.%m.%Y").date()
        except Exception:
            continue
        if _dd <= _sd:
            _dodaj(_o)
    return (out, True)


def _nacin_trebovanja(nase_kol, posle_map, nazivi=None):
    """Na osnovu KOLIČINA odredi da li je objekat trebovao PO NAŠEM ili PO NJIHOVOM.

    nase_kol:  {ida: tražena količina}  (dopuna porudžbine zabeležena na startu)
    posle_map: {ida: poručeno}          (šta je stvarno poručio posle starta)

    PO NAŠEM je samo ako je SVE po našem: svaki traženi artikal je poručen u punoj
    traženoj količini i nema artikala van naše liste. Ako i jedan artikal fali ili je
    poručen u manjoj količini — to je PO NJIHOVOM.

    Vrati (tip, info) — tip je 'nas' ili 'njihov'."""
    _nasi = {int(k): int(v) for k, v in (nase_kol or {}).items() if int(v or 0) > 0}
    _por = {int(k): int(v) for k, v in (posle_map or {}).items() if int(v or 0) > 0}
    _nz = {int(k): str(v) for k, v in (nazivi or {}).items()}
    _trazeno = sum(_nasi.values())
    _ukupno = sum(_por.values())
    _pokriveno = sum(min(_por.get(_ia, 0), _kol) for _ia, _kol in _nasi.items())
    _van = sum(_kol for _ia, _kol in _por.items() if _ia not in _nasi)
    _pok = (float(_pokriveno) / float(_trazeno)) if _trazeno > 0 else 0.0
    # Artikli koji NISU poručeni u punoj traženoj količini
    _fali = []
    for _ia, _kol in sorted(_nasi.items(), key=lambda x: -x[1]):
        _im = int(_por.get(_ia, 0))
        if _im < _kol:
            _fali.append({"ida": _ia, "naziv": _nz.get(_ia, ""), "trazeno": _kol, "poruceno": _im})
    if _trazeno <= 0:
        _tip = "njihov"                     # ništa nismo tražili, a oni su poručili po svom
    elif (not _fali) and _van == 0:
        _tip = "nas"                        # SVE po našem
    else:
        _tip = "njihov"
    return _tip, {"trazeno": _trazeno, "poruceno": _ukupno, "pokriveno": _pokriveno,
                  "van_liste": _van, "procenat": int(round(_pok * 100)),
                  "fali_n": len(_fali), "fali": _fali[:5]}


def admin_istorija_bulk(idk_naziv, cutoff_date, max_details=800):
    """Za ceo sistem odjednom: jedan login + jedan /orders POST, pa detalji SAMO za
    porudžbine >= cutoff_date koje se poklapaju sa objektima (po nazivu, potvrda idUser).
    idk_naziv: {idk: naziv}. Vrati ({idk: [ {id,datum,status,cena,stavke} ]}, greska)."""
    import re, html as _h, datetime as _dt
    s, base, err = _admin_session()
    if err:
        return ({}, err)
    name_items = []
    for _idk, _nz in idk_naziv.items():
        _n = _clean_komitent_naziv(_nz or "")
        if _n:
            name_items.append((int(_idk), _n))
    want_ids = set(int(x) for x in idk_naziv.keys())
    try:
        resp = s.post(base + "/orders", data={
            "orderStatuses": ["1", "10", "20", "30", "40", "50", "60", "70", "80", "90", "95"],
            "keyword": "", "idLoad": "", "startDate": "", "endDate": ""},
            headers={"Referer": base + "/orders", "X-Requested-With": "XMLHttpRequest"}, timeout=120)
        body = resp.text or ""
        cand = []
        for _row in re.findall(r'<tr[^>]*>(.*?)</tr>', body, re.S):
            if '/order/details/' not in _row:
                continue
            _rowtxt = _norm_ws(_h.unescape(_row))
            _md = re.search(r'(\d{2}\.\d{2}\.\d{4})(?:\s+(\d{2}:\d{2}))?', _rowtxt)
            if cutoff_date:
                if not _md:
                    continue
                try:
                    if _dt.datetime.strptime(_md.group(1), "%d.%m.%Y").date() < cutoff_date:
                        continue
                except Exception:
                    continue
            _hit = None
            for _idk, _n in name_items:
                if _n in _rowtxt:
                    _hit = _idk
                    break
            if _hit is None:
                continue
            _mid = re.search(r'/order/details/(\d+)', _row)
            if not _mid:
                continue
            _oid = _mid.group(1)
            _datum = (_md.group(1) + ((" " + _md.group(2)) if _md.group(2) else "")) if _md else ""
            _ms = re.search(r'labelOrderStatus[^"]*"[^>]*>([^<]+)</span>', _row)
            _status = _norm_ws(_h.unescape(_ms.group(1))) if _ms else ""
            _sl = _status.lower()
            if ("otkaz" in _sl) or ("storn" in _sl) or ("odbij" in _sl) or ("ponist" in _sl) or ("poništ" in _sl):
                continue  # otkazane/stornirane ne prikazujemo
            _mc = re.search(r'price-cell">\s*([\d.]+)\s*RSD', _row)
            _cena = (_mc.group(1) + " RSD") if _mc else ""
            cand.append((_hit, _oid, _datum, _status, _cena))
        cand = cand[:max_details]
        out = {}
        for (_hit, _oid, _datum, _status, _cena) in cand:
            try:
                det = s.get(base + "/order/details/" + _oid, timeout=45)
                _st, _iu = _parse_order_items(det.text)
                _final = _iu if (_iu in want_ids) else _hit
            except Exception:
                _st = []
                _final = _hit
            out.setdefault(_final, []).append(
                {"id": _oid, "datum": _datum, "status": _status, "cena": _cena, "stavke": _st})
        return (out, "")
    except Exception as _e:
        return ({}, "Greška pri čitanju istorije: " + str(_e))


def _to_int_kol(s):
    """Parsiraj količinu iz stringa (npr. '12', '12,00', '12.0') u int."""
    import re
    s = str(s).strip()
    if not s:
        return 0
    s = re.split(r'[.,]', s)[0]
    s = re.sub(r'[^\d-]', '', s)
    try:
        return int(s)
    except Exception:
        return 0


def _admin_presek(meta, mesec_key):
    """Datum preseka za „posle 01.": prvenstveno meta['presek'] (1. u mesecu posle
    poslednjeg meseca podataka), a ako ga nema — 1. u mesecu objave (staro ponašanje)."""
    _ps = None
    try:
        _ps = (meta or {}).get("presek")
    except Exception:
        _ps = None
    try:
        if _ps:
            _p = str(_ps).split("-")
            return datetime.date(int(_p[0]), int(_p[1]), int(_p[2]) if len(_p) > 2 else 1)
        return datetime.date(int(str(mesec_key).split("-")[0]), int(str(mesec_key).split("-")[1]), 1)
    except Exception:
        return None

_MN_KRATKO = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "maj": 5, "jun": 6,
              "jul": 7, "avg": 8, "sep": 9, "okt": 10, "nov": 11, "dec": 12}


def _lbl_u_ym(lbl):
    """'Avg 2026' -> (2026, 8). Vraća None ako oznaka nije prepoznata."""
    try:
        _d = str(lbl).strip().split()
        if len(_d) != 2:
            return None
        _m = _MN_KRATKO.get(_d[0][:3].lower())
        _g = int(_d[1])
        return (_g, _m) if _m else None
    except Exception:
        return None


def _por_otkazana(o):
    _s = str((o or {}).get("status", "") or "").lower()
    return ("otkaz" in _s) or ("storn" in _s) or ("odbij" in _s) or ("ponist" in _s) or ("poništ" in _s)


def _por_ym(o):
    """Godina i mesec porudžbine iz 'dd.mm.yyyy'."""
    try:
        _d = str((o or {}).get("datum", "") or "").split(" ")[0].split(".")
        return (int(_d[2]), int(_d[1]))
    except Exception:
        return None


def _poruceno_po_mesecima(admin_hist, mes_naz, samo_idk=None, samo_ida=None):
    """Zbir poručenih komada po mesecima iz mes_naz.
    admin_hist: {str(idk): [ {id, datum, status, stavke:[{ida, kol}]} ]}"""
    _ym = [_lbl_u_ym(l) for l in (mes_naz or [])]
    _idx = {y: i for i, y in enumerate(_ym) if y}
    out = [0] * len(_ym)
    for _k, _lst in (admin_hist or {}).items():
        if samo_idk is not None and str(_k) != str(int(samo_idk)):
            continue
        for _o in (_lst or []):
            if _por_otkazana(_o):
                continue
            _p = _por_ym(_o)
            if _p is None or _p not in _idx:
                continue
            _i = _idx[_p]
            for _s in (_o.get("stavke") or []):
                try:
                    if samo_ida is not None and int(_s.get("ida", -1)) != int(samo_ida):
                        continue
                    out[_i] += int(float(str(_s.get("kol", 0)).replace(",", ".")))
                except Exception:
                    continue
    return out


def _ucestalost_trebovanja(admin_hist, mes_naz, svi_idk):
    """Koliko je objekata poručilo 0, 1, 2, 3+ puta u posmatranom periodu.
    Broje se RAZLIČITE porudžbine (po ID-ju), ne stavke."""
    _ym = set(y for y in (_lbl_u_ym(l) for l in (mes_naz or [])) if y)
    _br = {}
    for _k, _lst in (admin_hist or {}).items():
        try:
            _ik = int(_k)
        except Exception:
            continue
        _ids = set()
        for _o in (_lst or []):
            if _por_otkazana(_o):
                continue
            if _por_ym(_o) in _ym:
                _ids.add(str(_o.get("id") or ""))
        _br[_ik] = len(_ids - {""})
    _k0 = _k1 = _k2 = _k3 = 0
    for _ik in svi_idk:
        _n = _br.get(int(_ik), 0)
        if _n == 0:
            _k0 += 1
        elif _n == 1:
            _k1 += 1
        elif _n == 2:
            _k2 += 1
        else:
            _k3 += 1
    return [("nijednom", _k0), ("1 put", _k1), ("2 puta", _k2), ("3 i više", _k3)]


def _lager_unazad(lager_sada, prodato, poruceno):
    """Rekonstruiši lager na kraju svakog meseca, unazad od današnjeg stanja.

    lager[kraj t-1] = lager[kraj t] + prodato[t] - poruceno[t]
    Vraća listu iste dužine kao prodato (nikad ispod nule)."""
    n = len(prodato)
    out = [0] * n
    _l = int(lager_sada or 0)
    for t in range(n - 1, -1, -1):
        out[t] = max(_l, 0)
        _l = _l + int(prodato[t] or 0) - int(poruceno[t] or 0)
    return out


def _primer_neredovno(kandidati, admin_hist, mes_naz):
    """Izaberi artikal koji najbolje pokazuje TRENUTNI problem.

    Uslovi (svi moraju da važe, inače primer nema smisla):
      · lager je DANAS nula — problem je sadašnji, ne istorijski,
      · artikal se prodavao bar 3 meseca — postoji stvarna potražnja,
      · prodato je više nego što je poručeno — inače objekat ima robu
        i priča o nestašici ne stoji.
    Među takvima bira se onaj sa najjačom prodajom u poslednja tri meseca.

    Namerno se NE rekonstruiše istorija lagera: to je računanje unazad iz
    današnjeg stanja, a čim se količine iz admina i prodaja ne poklope
    (povrati, roba primljena u drugom mesecu), dobiju se besmislene nule.
    Prikazuje se samo ono što pouzdano znamo: prodaja, porudžbine i lager danas."""
    _naj, _naj_sc = None, None
    for c in kandidati:
        if int(c.get("lager", 0) or 0) > 0:
            continue                              # ima robu — nije trenutni problem
        _ser = [int(x or 0) for x in (c.get("prodaja") or [])]
        if not _ser or sum(_ser) <= 0:
            continue
        _n_mes = sum(1 for x in _ser if x > 0)
        if _n_mes < 3:
            continue                              # prekratka istorija da bi se tvrdilo bilo šta
        _por = _poruceno_po_mesecima(admin_hist, mes_naz, c["idk"], c["ida"])
        if sum(_por) >= sum(_ser):
            continue                              # poručeno je koliko i prodato — nema šta da se dokazuje
        _skoro = sum(_ser[-3:]) if len(_ser) >= 3 else sum(_ser)
        _sc = _skoro * 1000 + _n_mes * 10 + min(sum(_ser), 999)
        if _naj_sc is None or _sc > _naj_sc:
            _naj_sc, _naj = _sc, {
                "obj": c["obj"], "art": c["art"], "meseci": list(mes_naz),
                "prodaja": _ser, "poruceno": _por, "lager": 0,
                "uk_prodato": sum(_ser), "uk_poruceno": sum(_por),
                "n_porudzbina": sum(1 for x in _por if x > 0),
                "mes_prosek": round(sum(_ser) / float(max(_n_mes, 1)), 1),
                "skoro3": _skoro,
            }
    return _naj


def _primer_nedovoljno(objekti_prodaja, admin_hist, mes_naz, lager_po_obj):
    """Objekat koji poručuje REDOVNO ali MANJE nego što proda.

    objekti_prodaja: {idk: (ime, [prodato po mesecima])}
    lager_po_obj: {idk: ukupan trenutni lager na problematičnim artiklima}"""
    _naj, _naj_sc = None, None
    for _ik, (_ime, _ser) in (objekti_prodaja or {}).items():
        _ser = [int(x or 0) for x in (_ser or [])]
        if not _ser or sum(_ser) <= 0:
            continue
        _por = _poruceno_po_mesecima(admin_hist, mes_naz, _ik)
        _n_por = sum(1 for x in _por if x > 0)
        if _n_por < 3:          # mora da poručuje redovno, inače je to drugi primer
            continue
        _manjak = sum(_ser) - sum(_por)
        if _manjak <= 0:
            continue
        _sc = _manjak * 10 + _n_por
        if _naj_sc is None or _sc > _naj_sc:
            _lg = int((lager_po_obj or {}).get(_ik, 0) or 0)
            _naj_sc, _naj = _sc, {
                "obj": _ime, "meseci": list(mes_naz), "prodaja": _ser, "poruceno": _por,
                "lager_niz": _lager_unazad(_lg, _ser, _por), "lager": _lg,
                "uk_prodato": sum(_ser), "uk_poruceno": sum(_por), "manjak": _manjak,
                "n_porudzbina": _n_por,
            }
    return _naj


def _treb_posle_preseka(hist_lst, cutoff_date):
    """Iz učitane istorije porudžbina saberi količine po artiklu za porudžbine
    datirane >= cutoff_date, izuzimajući otkazane/stornirane. Vrati {ida: kom}."""
    import datetime as _dt
    out = {}
    if not hist_lst:
        return out
    for _o in hist_lst:
        _dstr = (_o.get("datum") or "").strip()
        _md = _dstr.split(" ")[0] if _dstr else ""
        _ok_datum = True
        if cutoff_date and _md:
            try:
                _dd = _dt.datetime.strptime(_md, "%d.%m.%Y").date()
                _ok_datum = _dd >= cutoff_date
            except Exception:
                _ok_datum = True
        if not _ok_datum:
            continue
        _stat = (_o.get("status") or "").lower()
        if ("otkaz" in _stat) or ("storn" in _stat) or ("odbij" in _stat) or ("ponist" in _stat) or ("poništ" in _stat):
            continue
        for _s in (_o.get("stavke") or []):
            try:
                _ida = int(_s.get("ida"))
            except Exception:
                continue
            out[_ida] = out.get(_ida, 0) + _to_int_kol(_s.get("kol"))
    return out


def _statistika_agg(mesec_key):
    """Sakupi „ko je šta uradio" za mesec, preko svih sistema.
    Vrati {user: {poz, mej, dir, ubac, prijave, objekti:set}}."""
    agg = {}
    def _u(user):
        return agg.setdefault(user or "Bez oznake",
                              {"poz": 0, "mej": 0, "dir": 0, "ubac": 0, "prijave": 0, "objekti": set()})
    for _sis in sb_sisteme(mesec_key):
        _obr = sb_load_obrada(mesec_key, _sis)
        for _idk, _v in _obr.items():
            _rk = _v.get("reakcije_ko") or {}
            for _rr in (_v.get("reakcije") or []):
                _who = _rk.get(_rr) or "Bez oznake"
                _d = _u(_who)
                if _rr == "Pozvala sam": _d["poz"] += 1
                elif _rr == "Poslala sam mejl": _d["mej"] += 1
                elif _rr == "Obavestila direktorku": _d["dir"] += 1
                elif _rr == "Ubačena porudžbina": _d["ubac"] += 1
                _d["objekti"].add((_sis, int(_idk)))
        _pod = sb_ucitaj(mesec_key, _sis)
        _meta = (_pod or {}).get("meta") or {}
        for _k, _info in (_meta.get("nedeljni_prijave") or {}).items():
            _who = (_info or {}).get("ko", "") or "Bez oznake"
            _d = _u(_who); _d["prijave"] += 1
            try:
                _d["objekti"].add((_sis, int(_k)))
            except Exception:
                pass
    return agg


def render_statistika(mesec_key, sel_lbl):
    """Prikaz „ko je šta uradio" za izabrani mesec (koristi se i kod administracije i kod direktora)."""
    agg = _statistika_agg(mesec_key)
    if not agg:
        st.info("Za " + str(sel_lbl) + " još nema zabeleženih akcija administracije.")
        return
    st.markdown('<div style="font-size:16px;font-weight:800;margin:4px 0 10px;">📊 Ko je šta uradio · '
                + _h_escape(str(sel_lbl)) + '</div>', unsafe_allow_html=True)
    # poređaj: prvo imenovani (A1, A2...), pa „Bez oznake"
    _users = sorted(agg.keys(), key=lambda u: (u == "Bez oznake", u))
    _cols = st.columns(max(1, min(len(_users), 3)))
    _clr = {"Administracija 1": "#7c3aed", "Administracija 2": "#ec4899",
            str(ADMIN_IME): "#7c3aed", str(ADMIN_IME_2): "#ec4899"}
    for _i, _u in enumerate(_users):
        _d = agg[_u]
        _c = _clr.get(_u, "#64748b")
        _tot = _d["poz"] + _d["mej"] + _d["dir"] + _d["ubac"] + _d["prijave"]
        with _cols[_i % len(_cols)]:
            st.markdown(
                '<div style="background:#fff;border:1px solid #efeaf7;border-top:4px solid ' + _c + ';'
                'border-radius:14px;padding:16px 18px;margin-bottom:12px;box-shadow:0 2px 12px rgba(80,40,140,.05);">'
                '<div style="font-weight:800;font-size:15px;color:' + _c + ';margin-bottom:2px;">' + _h_escape(_u) + '</div>'
                '<div style="font-size:11.5px;color:#9aa0ad;margin-bottom:12px;">' + str(len(_d["objekti"]))
                + ' objekata · ' + str(_tot) + ' akcija</div>'
                '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">'
                + _stat_cell("📞 Pozvano", _d["poz"])
                + _stat_cell("✉️ Mejlova", _d["mej"])
                + _stat_cell("👤 Komercijali", _d["dir"])
                + _stat_cell("📦 Ubačeno", _d["ubac"])
                + _stat_cell("⚠️ Prijave (nedeljni)", _d["prijave"])
                + '</div></div>', unsafe_allow_html=True)
    if "Bez oznake" in agg:
        st.caption("Bez oznake = akcije zabeležene pre uvođenja dve prijave (nema imena ko ih je uneo).")
    # zbirna tabela
    _rows = []
    for _u in _users:
        _d = agg[_u]
        _rows.append({"Osoba": _u, "Objekata": len(_d["objekti"]), "Pozvano": _d["poz"],
                      "Mejlova": _d["mej"], "Komercijali": _d["dir"], "Ubačeno": _d["ubac"],
                      "Prijave (nedeljni)": _d["prijave"]})
    st.dataframe(pd.DataFrame(_rows), hide_index=True, use_container_width=True)


def _stat_cell(lbl, val):
    return ('<div style="background:#faf9fd;border-radius:9px;padding:7px 10px;">'
            '<div style="font-size:18px;font-weight:800;color:#1f2430;">' + str(val) + '</div>'
            '<div style="font-size:10.5px;color:#8b8fa0;margin-top:1px;">' + lbl + '</div></div>')


def prikazi_administraciju():
    st.set_page_config(page_title="VAPE — Porudžbine", page_icon="📦",
                       layout="wide", initial_sidebar_state="collapsed")
    st.markdown("""<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    section[data-testid="stSidebar"]{display:none !important;}
    header[data-testid="stHeader"]{display:none !important;}
    #MainMenu{visibility:hidden !important;} footer{visibility:hidden !important;}
    .stApp{background:#fbfbfd !important;font-family:'Inter',sans-serif;}
    div[data-testid="stMainBlockContainer"]{padding:10px 26px 40px !important;max-width:100% !important;}
    /* header */
    .adm-hdr{display:flex;align-items:center;gap:11px;padding:16px 0 16px;border-bottom:1px solid #eef0f4;margin-bottom:22px;}
    .adm-logo{width:26px;height:26px;border-radius:8px;background:linear-gradient(135deg,#a855f7,#ec4899);display:flex;align-items:center;justify-content:center;}
    .adm-logo div{width:9px;height:9px;background:#fff;border-radius:2px;}
    .adm-hdr .t1{font-weight:700;font-size:16px;letter-spacing:-.2px;color:#1f2430;}
    .adm-hdr .t2{color:#9ca3af;font-weight:500;font-size:13px;}
    /* kpi */
    .adm-kpi{display:flex;border:1px solid #eef0f4;border-radius:12px;background:#fff;overflow:hidden;margin:4px 0 14px;}
    .adm-kpi .cell{flex:1;padding:15px 20px;border-right:1px solid #f2f3f7;}
    .adm-kpi .cell:last-child{border-right:none;}
    .adm-kpi .n{font-size:22px;font-weight:700;letter-spacing:-.5px;display:flex;align-items:center;gap:8px;color:#1f2430;}
    .adm-kpi .n .d{width:8px;height:8px;border-radius:50%;}
    .adm-kpi .n .d.r{background:#e5484d;} .adm-kpi .n .d.o{background:#f2820c;} .adm-kpi .n .d.g{background:#17a34a;} .adm-kpi .n .d.p{background:#7c3aed;}
    .adm-kpi .k{font-size:12px;color:#9ca3af;margin-top:3px;font-weight:500;}
    .adm-kpi .n .d.z{background:#16a34a;}
    .adm-kpi .cell.c-p{background:#faf7ff;} .adm-kpi .cell.c-r{background:#fff7f7;} .adm-kpi .cell.c-o{background:#fffcf5;} .adm-kpi .cell.c-g{background:#f6fdf9;}
    .adm-kpi .cell.c-z{background:#dcfce7;border-left:2px solid #86efac;}
    .adm-kpi .cell.c-p .n{color:#7c3aed;} .adm-kpi .cell.c-r .n{color:#d33;} .adm-kpi .cell.c-o .n{color:#c66a00;} .adm-kpi .cell.c-g .n{color:#158a3f;}
    .adm-kpi .cell.c-z .n{color:#14532d;} .adm-kpi .cell.c-z .k{color:#166534;font-weight:700;}
    table.adm-t tr.row-red td{background:#fff8f8;} table.adm-t tr.row-org td{background:#fffcf6;}
    /* progress */
    .adm-prog{display:flex;align-items:center;gap:12px;margin-bottom:8px;}
    .adm-prog .t{font-size:12.5px;color:#6b7280;font-weight:600;white-space:nowrap;}
    .adm-prog .bar{flex:1;height:5px;background:#eef0f4;border-radius:99px;overflow:hidden;}
    .adm-prog .bar>div{height:100%;background:#7c3aed;border-radius:99px;}
    /* tables */
    table.adm-t{width:100%;border-collapse:collapse;}
    table.adm-t th{text-align:left;font-size:11px;color:#b0b4bd;font-weight:600;text-transform:uppercase;letter-spacing:.5px;padding:0 14px 12px;}
    table.adm-t td{padding:13px 14px;border-top:1px solid #f2f3f7;font-size:14px;vertical-align:middle;color:#2a2f3a;}
    table.adm-t td.idc{font-weight:700;} table.adm-t td.ce{text-align:center;} table.adm-t td.mut{color:#b0b4bd;}
    /* Lista objekata — redovi sa nazivom kao dugmetom (izgled kao ranija tabela) */
    .lst-h{font-size:11px;color:#b0b4bd;font-weight:600;text-transform:uppercase;letter-spacing:.5px;
           padding:0 0 4px;}
    .lst-c{font-size:14px;color:#2a2f3a;line-height:1.25;}
    .lst-c.ce{text-align:center;} .lst-id{font-weight:700;}
    .lst-row{display:flex;align-items:center;gap:0;width:100%;}
    div[data-testid="stHorizontalBlock"]:has(.rtag){border-top:1px solid #f2f3f7;padding:5px 0 4px;}
    div[data-testid="stHorizontalBlock"]:has(.rtag.red){background:#fff8f8;}
    div[data-testid="stHorizontalBlock"]:has(.rtag.org){background:#fffcf6;}
    div[data-testid="stHorizontalBlock"]:has(.rtag) div[data-testid="stElementContainer"]{margin:0;}
    div[data-testid="stHorizontalBlock"]:has(.rtag) button{background:transparent !important;
        border:none !important;box-shadow:none !important;padding:0 !important;min-height:0 !important;
        font-size:14px !important;font-weight:400 !important;color:#2a2f3a !important;
        text-align:left !important;justify-content:flex-start !important;width:100% !important;}
    div[data-testid="stHorizontalBlock"]:has(.rtag) button p{font-size:14px !important;
        font-weight:400 !important;margin:0 !important;text-align:left !important;}
    div[data-testid="stHorizontalBlock"]:has(.rtag) button:hover,
    div[data-testid="stHorizontalBlock"]:has(.rtag) button:hover p{color:#7c3aed !important;}
    table.adm-t td a:hover{color:#7c3aed !important;border-bottom-color:#7c3aed !important;}
    .zona{display:inline-flex;align-items:center;gap:7px;font-weight:600;font-size:13px;white-space:nowrap;}
    .zona .zd{width:8px;height:8px;border-radius:50%;}
    .z-red{color:#d33;} .z-red .zd{background:#e5484d;}
    .z-org{color:#c66a00;} .z-org .zd{background:#f2820c;}
    .z-grn{color:#158a3f;} .z-grn .zd{background:#17a34a;}
    .stat{font-size:12.5px;color:#9ca3af;}
    .stchip{display:inline-block;font-size:11.5px;color:#5b21b6;background:#f2effc;border-radius:6px;padding:2px 8px;margin:1px 3px 1px 0;}
    .tb-nas{color:#158a3f;font-weight:600;font-size:12.5px;} .tb-nj{color:#c66a00;font-weight:600;font-size:12.5px;}
    .np{color:#c4c7cf;}
    .adot{width:9px;height:9px;border-radius:50%;display:inline-block;}
    .ar-red{background:#e5484d;} .ar-yel{background:#f2820c;} .ar-grn{background:#17a34a;}
    /* empty */
    .adm-empty{text-align:center;padding:84px 20px;}
    .adm-empty .ic{font-size:40px;opacity:.5;margin-bottom:14px;}
    .adm-empty h2{font-size:19px;font-weight:600;color:#374151;margin-bottom:6px;}
    .adm-empty .w{color:#9ca3af;font-size:14px;margin-bottom:14px;}
    .adm-empty .plan{display:inline-block;color:#7c3aed;font-size:14px;font-weight:600;border:1px solid #ede9fe;background:#faf8ff;border-radius:99px;padding:8px 18px;}
    .adm-empty p{color:#b0b4bd;font-size:13px;margin-top:14px;}
    /* detail */
    .adm-dh{display:flex;align-items:center;gap:14px;margin:6px 0 18px;}
    .adm-dh .id{font-size:22px;font-weight:700;letter-spacing:-.5px;color:#1f2430;}
    .adm-dh .mut{color:#b0b4bd;font-size:13px;}
    .adm-lbl{font-size:11px;color:#b0b4bd;font-weight:600;text-transform:uppercase;letter-spacing:.5px;margin:2px 0 8px;}
    .revy{color:#158a3f;font-weight:600;font-size:12.5px;} .revn{color:#c4c7cf;font-weight:600;font-size:12.5px;}
    /* streamlit kontrole suptilnije */
    .stButton>button{border-radius:9px;font-weight:600;}
    button[data-testid="baseButton-primary"]{background:#7c3aed !important;border-color:#7c3aed !important;color:#fff !important;}
    /* dugmad za slanje u admin — jasno obojena, kompaktna */
    [class*="st-key-axn_"] button{background:#16a34a !important;border-color:#16a34a !important;color:#fff !important;font-weight:600 !important;font-size:13px !important;padding:7px 12px !important;border-radius:8px !important;box-shadow:none !important;}
    [class*="st-key-callbtn_"] button{background:#16a34a !important;border-color:#16a34a !important;color:#fff !important;font-weight:700 !important;font-size:13.5px !important;padding:8px 14px !important;border-radius:9px !important;box-shadow:none !important;}
    [class*="st-key-callbtn_"] button:hover{background:#128a3e !important;border-color:#128a3e !important;}
    [class*="st-key-axn_"] button:hover{background:#128a3e !important;border-color:#128a3e !important;}
    [class*="st-key-axj_"] button{background:#f59e0b !important;border-color:#f59e0b !important;color:#fff !important;font-weight:600 !important;font-size:13px !important;padding:7px 12px !important;border-radius:8px !important;box-shadow:none !important;}
    [class*="st-key-axj_"] button:hover{background:#d97706 !important;border-color:#d97706 !important;}
    .stButton button[kind="primary"]{background:#7c3aed !important;border-color:#7c3aed !important;color:#fff !important;font-size:13.5px !important;padding:8px 14px !important;border-radius:8px !important;font-weight:600 !important;}
    /* zaglavlje: mala, uredna dugmad */
    [class*="st-key-adm_odjava"] button{font-size:11px !important;padding:3px 8px !important;border-radius:7px !important;min-height:0 !important;line-height:1.2 !important;background:#fff !important;border:1px solid #e5e7eb !important;color:#6b7280 !important;font-weight:600 !important;box-shadow:none !important;}
    [class*="st-key-adm_odjava"] button:hover{background:#f9fafb !important;color:#374151 !important;}
    [class*="st-key-refresh_all_admin"] button{font-size:11px !important;padding:3px 8px !important;border-radius:7px !important;min-height:0 !important;line-height:1.2 !important;background:#7c3aed !important;border-color:#7c3aed !important;color:#fff !important;font-weight:600 !important;box-shadow:none !important;}
    [class*="st-key-refresh_all_admin"] button:hover{background:#6d28d9 !important;border-color:#6d28d9 !important;}
    [class*="st-key-predaj_izvestaj"] button{font-size:11px !important;padding:3px 8px !important;border-radius:7px !important;min-height:0 !important;line-height:1.2 !important;background:#0ea5e9 !important;border-color:#0ea5e9 !important;color:#fff !important;font-weight:600 !important;box-shadow:none !important;}
    [class*="st-key-predaj_izvestaj"] button:hover{background:#0284c7 !important;border-color:#0284c7 !important;}
    .stMultiSelect [data-baseweb="tag"]{background:#f2effc !important;color:#5b21b6 !important;border:none !important;}
    .stMultiSelect [data-baseweb="tag"] span{color:#5b21b6 !important;}
    </style>""", unsafe_allow_html=True)

    _adm_user = st.session_state.get("admin_user", "Administracija")
    _hc1, _hc2 = st.columns([7.5, 1.15])
    with _hc1:
        st.markdown('<div class="adm-hdr"><div class="adm-logo"><div></div></div>'
                    '<span class="t1">VAPE Porudžbine</span><span class="t2">· ' + _h_escape(_adm_user) + '</span></div>',
                    unsafe_allow_html=True)
    with _hc2:
        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
        if st.button("Odjava", key="adm_odjava", use_container_width=True):
            for _k in ("authenticated", "role", "admin_user", "mail_nalog"):
                st.session_state.pop(_k, None)
            st.rerun()
        if st.button("🔄 Ažuriraj iz admina", key="refresh_all_admin",
                     use_container_width=True, type="primary"):
            st.session_state["_req_refresh_admin"] = True
        if st.button("📝 Predaj izveštaj", key="predaj_izvestaj", use_container_width=True):
            st.session_state["_req_predaj"] = True

    if not sb_dostupan():
        st.error("Veza sa bazom trenutno nije podešena. Javi se analitičaru.")
        return

    _adm_mode = st.radio("Prikaz", ["📦 Porudžbine", "💳 Potraživanja", "⛽ Izveštaj Knez Petrol"],
                         horizontal=True, key="adm_mode", label_visibility="collapsed")
    if "Potra" in _adm_mode:
        potraz_admin_ui()
        return
    if "Knez" in _adm_mode:
        knez_admin_ui()
        return

    _pub = sb_meseci()
    _mes_keys = [m["key"] for m in _pub]
    for _k in sb_plan_meseci():
        if _k not in _mes_keys:
            _mes_keys.append(_k)
    _mes_keys = sorted(set(_mes_keys), reverse=True)
    if not _mes_keys:
        st.info("Još nema objavljenih podataka. Analitičar treba prvo da objavi bar jedan sistem.")
        return

    _c1, _c2, _c3 = st.columns([1, 1, 3])
    _mlbls = [mesec_label(k) for k in _mes_keys]
    with _c1:
        _sel_lbl = st.selectbox("Mesec", _mlbls, index=0, key="adm_mes")
    mesec_key = _mes_keys[_mlbls.index(_sel_lbl)]
    _imaju = sb_sisteme(mesec_key)
    _svi = sb_svi_sistemi()
    _sis_opts = sorted(set(_svi) | set(_imaju))
    with _c2:
        if _sis_opts:
            sistem = st.selectbox("Sistem", _sis_opts, index=0, key="adm_sis")
        else:
            sistem = None
            st.selectbox("Sistem", ["(nema)"], disabled=True)

    if not sistem:
        st.warning("Nema dostupnih sistema.")
        return

    # Rok koji je direktor postavio za ovaj mesec (administracija radi do tog roka)
    _rk_adm = sb_rokovi_get(mesec_key).get("rok_admin")
    if _rk_adm:
        _proso = _rok_je_prosao(_rk_adm)
        _bg = "#fef2f2;border-color:#fecaca;color:#b42318" if _proso else "#f0fdf4;border-color:#bbf7d0;color:#166534"
        _txt = ("Rok je istekao (" + _rok_fmt(_rk_adm) + ") — izveštaj je zaključan i predat direktoru."
                if _proso else "Rok za predaju izveštaja (" + _sel_lbl + "): " + _rok_fmt(_rk_adm))
        st.markdown('<div style="background:' + _bg + ';border:1px solid;border-radius:10px;padding:9px 14px;'
                    'font-size:12.5px;font-weight:600;margin:2px 0 12px;">⏰ ' + _txt + '</div>', unsafe_allow_html=True)

    _pcx = st.columns([1.5, 1, 2])
    with _pcx[0]:
        if st.button("📄 Napravi PDF izveštaj (" + _sel_lbl + ")", key="pdf_make", use_container_width=True):
            try:
                with st.spinner("Pravim PDF..."):
                    st.session_state["pdf_bytes"] = napravi_pdf_izvestaj(mesec_key, _sel_lbl)
                    st.session_state["pdf_mes"] = mesec_key
            except Exception as _pe:
                st.session_state["pdf_bytes"] = None
                st.error("Greška pri pravljenju PDF-a: " + str(_pe))
    with _pcx[1]:
        if st.session_state.get("pdf_bytes") and st.session_state.get("pdf_mes") == mesec_key:
            st.download_button("⬇️ Preuzmi PDF", st.session_state["pdf_bytes"],
                file_name="Izvestaj_administracije_" + mesec_key + ".pdf", mime="application/pdf",
                key="pdf_dl", use_container_width=True)

    podaci = sb_ucitaj(mesec_key, sistem)
    if not podaci or not podaci.get("stavke"):
        _plan = sb_load_plan(mesec_key)
        if _plan:
            _ph = '<div class="plan">Planirana objava do ' + str(_plan) + '.</div>'
        else:
            _ph = '<div class="plan">Datum objave još nije zakazan.</div>'
        st.markdown('<div class="adm-empty"><div class="ic">🗓️</div>'
                    '<h2>Izveštaj još nije objavljen</h2>'
                    '<div class="w">' + sistem + ' · ' + _sel_lbl + '</div>' + _ph +
                    '<p>Do tada isplanirajte obilaske.</p></div>', unsafe_allow_html=True)
        return

    stavke = podaci["stavke"]
    meta = podaci.get("meta", {}) or {}
    _mes_kol = meta.get("meseci")
    _por_lbl = "Porudžbina"

    # --- Vrati zapamćene prethodne porudžbine iz admina (posle osvežavanja stranice) ---
    # Ako u ovoj sesiji još nisu učitane, popuni iz meta.admin_hist da „posle 01." i
    # sortiranje po zonama ostanu isti kao pri poslednjem „Ažuriraj iz admina".
    # VAŽNO: objekat koji NEMA nijednu porudžbinu se ne pamti u admin_hist (prazna lista),
    # pa se popunjava iz admin_hist_idk — inače bi aplikacija za njega i dalje tražila
    # novo ažuriranje iako je već proveren.
    _saved_hist = meta.get("admin_hist") if isinstance(meta, dict) else None
    if isinstance(_saved_hist, dict):
        for _sk, _slst in _saved_hist.items():
            _hkk = "hist_" + str(sistem) + "_" + str(_sk)
            if st.session_state.get(_hkk) is None:
                st.session_state[_hkk] = {"lst": _slst or [], "err": ""}
    _cov_idk = meta.get("admin_hist_idk") if isinstance(meta, dict) else None
    if not _cov_idk and meta.get("admin_hist_at"):
        # Stariji izveštaji (pre nadogradnje) nemaju spisak — tada važi:
        # ako je ažurirano, obuhvaćeni su SVI objekti iz izveštaja.
        try:
            _cov_idk = sorted(set(str(int(s["idk"])) for s in stavke))
        except Exception:
            _cov_idk = []
    for _ck in (_cov_idk or []):
        _hkk = "hist_" + str(sistem) + "_" + str(_ck)
        if st.session_state.get(_hkk) is None:
            st.session_state[_hkk] = {"lst": [], "err": ""}

    # --- Napomena koju je analitika upisala pri objavi (dogovor sa sistemom) ---
    _nap_an = (meta.get("nap_analitika") or {}) if isinstance(meta, dict) else {}
    if _nap_an.get("tekst"):
        _nap_sub = []
        if _nap_an.get("ko"):
            _nap_sub.append(str(_nap_an["ko"]))
        if _nap_an.get("at"):
            _nap_sub.append(str(_nap_an["at"]))
        st.markdown('<div style="background:#fffbeb;border:1px solid #fcd34d;border-left:5px solid #f59e0b;'
                    'border-radius:10px;padding:11px 15px;margin:2px 0 12px;">'
                    '<div style="font-size:11.5px;font-weight:800;color:#92400e;text-transform:uppercase;'
                    'letter-spacing:.6px;">📝 Napomena iz analitike</div>'
                    '<div style="font-size:14px;color:#78350f;font-weight:600;margin-top:4px;'
                    'white-space:pre-wrap;">' + _h_escape(str(_nap_an["tekst"])) + '</div>'
                    + (('<div style="font-size:11px;color:#a16207;margin-top:5px;font-style:italic;">'
                        + _h_escape(" · ".join(_nap_sub)) + '</div>') if _nap_sub else '')
                    + '</div>', unsafe_allow_html=True)

    # --- Poslednje ažuriranje iz admina (odmah ispod dugmeta „Ažuriraj iz admina") ---
    if not (isinstance(meta, dict) and meta.get("nedeljni")):
        _ah_at = meta.get("admin_hist_at") if isinstance(meta, dict) else None
        _pk_disp = _admin_presek(meta, mesec_key)
        _pk_ds = _pk_disp.strftime("%d.%m.%Y") if _pk_disp else "01."
        if _ah_at:
            st.markdown('<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:8px 14px;'
                        'font-size:12.5px;color:#166534;font-weight:600;margin:2px 0 14px;">🔄 Podaci iz admina poslednji put ažurirani: '
                        + _h_escape(_dt_fmt(_ah_at)) + '  ·  prethodne porudžbine (posle 01.) su zapamćene i prate se od '
                        + _pk_ds + '. Ne moraš ponovo da ažuriraš ako je datum skorašnji.</div>',
                        unsafe_allow_html=True)
        else:
            st.markdown('<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:10px;padding:8px 14px;'
                        'font-size:12.5px;color:#92400e;font-weight:600;margin:2px 0 14px;">🔄 Prethodne porudžbine iz admina još nisu povučene za ovaj sistem — '
                        'klikni „Ažuriraj iz admina" gore (povuku se i trajno zapamte).</div>', unsafe_allow_html=True)

    # --- Zaključavanje meseca: predato ručno ili istekao rok (rok postavljaju direktori) ---
    _predato = bool(meta.get("predato"))
    _rok = sb_rokovi_get(mesec_key).get("rok_admin") or meta.get("rok")  # 'YYYY-MM-DD' ako je postavljen
    _rok_prosao = False
    if _rok:
        try:
            _rok_prosao = datetime.date.today() > datetime.date.fromisoformat(str(_rok))
        except Exception:
            _rok_prosao = False
    _zakljucan = _predato or _rok_prosao
    po_obj = {}
    for s in stavke:
        po_obj.setdefault(int(s["idk"]), []).append(s)
    # Presek za „posle 01." (za preračun hitnosti po DODATNOJ porudžbini kad su podaci povučeni)
    _cut_hit = _admin_presek(meta, mesec_key)
    # Startni snimak: posle klika na Start dopuna porudžbine se ZAMRZAVA — računa se
    # samo po porudžbinama koje su postojale u trenutku starta. Kasnije porudžbine ne
    # menjaju dopunu, nego objekat prebacuju u ZAVRŠENO.
    _snap_fz = meta.get("start_zone") if isinstance(meta, dict) else None
    objekti = []
    for idk, lst in po_obj.items():
        # Ako je istorija iz admina povučena za ovaj objekat — hitnost po dodatnoj porudžbini
        _hh = st.session_state.get("hist_" + str(sistem) + "_" + str(idk))
        if _hh and (_hh.get("lst") is not None) and _cut_hit:
            _por_map, _fz_ok = _treb_na_startu(idk, _hh.get("lst") or [], _snap_fz, _cut_hit)
            if not _fz_ok:
                _por_map = _treb_posle_preseka(_hh.get("lst") or [], _cut_hit)
            nivo, n_nula, izgub = hitnost_objekta_dodatna(lst, _por_map)
        else:
            nivo, n_nula, izgub = hitnost_objekta(lst)
        objekti.append({"idk": idk, "artikala": len(lst), "na_nuli": n_nula,
                        "izgub": izgub, "nivo": nivo, "lst": lst})
    objekti.sort(key=lambda r: (HIT_RANG[r["nivo"]], -r["izgub"]))
    ids_sorted = [o["idk"] for o in objekti]
    obj_by_id = {o["idk"]: o for o in objekti}

    obrada_map = sb_load_obrada(mesec_key, sistem)

    # --- Preračunaj način trebovanja po VAŽEĆEM pravilu (jednom po sesiji) ---
    # Zapisi napravljeni starijim pravilom (blaži prag) se same isprave, bez novog
    # povlačenja iz admina — koriste se već zapamćene porudžbine i startni snimak.
    _PRAVILO_V = 2
    _fixk = "_fixtip_" + str(sistem) + "_" + str(mesec_key)
    if (not st.session_state.get(_fixk)) and _snap_fz and not _zakljucan:
        _n_fix = 0
        for _ob in objekti:
            _vv = obrada_map.get(int(_ob["idk"])) or {}
            _dn0 = _vv.get("dnevnik") or {}
            _at0 = _dn0.get("trebovao_posle_starta") or {}
            if not _at0:
                continue
            if ("Ubačena porudžbina" in (_vv.get("reakcije") or [])) or _dn0.get("ubaceno"):
                continue                      # mi smo ubacili — bira se ručno
            if int((_at0.get("info") or {}).get("pravilo") or 0) >= _PRAVILO_V:
                continue                      # već po važećem pravilu
            _hh0 = st.session_state.get("hist_" + str(sistem) + "_" + str(_ob["idk"]))
            if not _hh0 or _hh0.get("err"):
                continue
            _lst0 = _hh0.get("lst") or []
            try:
                _pos0 = _treb_posle_starta(_ob["idk"], _lst0, _snap_fz)
                _st0, _ = _treb_na_startu(_ob["idk"], _lst0, _snap_fz, _cut_hit)
                _nk0 = {}; _nz0 = {}
                for _s0 in _ob["lst"]:
                    _i0 = int(_s0["ida"])
                    _nk0[_i0] = max(int(_s0.get("kol", 0) or 0) - int(_st0.get(_i0, 0) or 0), 0)
                    _nz0[_i0] = str(_s0.get("naziv", "") or "")
                _t0, _inf0 = _nacin_trebovanja(_nk0, _pos0, _nz0)
                _inf0["pravilo"] = _PRAVILO_V
                _nj0 = {str(_i0): int(_pos0.get(_i0, 0)) for _i0 in _nk0
                        if int(_pos0.get(_i0, 0)) > 0}
                if sb_oznaci_trebovao(mesec_key, sistem, _ob["idk"],
                                      kada=str(_at0.get("kada") or ""), n_por=len(_lst0),
                                      tip=_t0, njihova=_nj0, info=_inf0):
                    _n_fix += 1
            except Exception:
                pass
        st.session_state[_fixk] = True
        if _n_fix:
            obrada_map = sb_load_obrada(mesec_key, sistem)

    reviewed = set(idk for idk, v in obrada_map.items() if v.get("reakcije"))
    if st.session_state.get("_komfull") is None:
        st.session_state["_komfull"] = sb_komitenti_full()
    komfull = st.session_state.get("_komfull") or {}

    # ===== Nedeljni/sistemski sistem: pojednostavljen prikaz =====
    if isinstance(meta, dict) and meta.get("nedeljni"):
        _dani = int(meta.get("nedeljni_dani", 7) or 7)   # period pokrivenosti (7 = nedeljni, 45 = mesec i po...)
        _per_lbl = (str(_dani) + " dana")
        st.markdown('<div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:9px 14px;'
                    'font-size:12.5px;color:#1e40af;margin:2px 0 14px;">📅 Sistemski sistem (period: ' + _per_lbl
                    + ') — prikazani su samo objekti sa problemom: realni lager ne pokriva prodaju za ' + _per_lbl
                    + ' (realni lager = lager + naknadne porudžbine posle preseka).</div>',
                    unsafe_allow_html=True)

        if not (meta.get("mesec_nazivi")):
            st.warning("⚠️ Ovaj sistem je objavljen pre nadogradnje, pa u prilogu (PDF) NEMA grafika prodaje "
                       "u primeru. Da bi se primer sa grafikom pojavio, analitičar treba JEDNOM ponovo da "
                       "objavi ovaj sistem (istim Excel fajlom) — tada se upiše mesečna prodaja po artiklu.")

        _pk_ned = _admin_presek(meta, mesec_key)
        _pk_ned_lbl = _pk_ned.strftime("%d.%m.") if _pk_ned else "01."
        _ah_at_n = meta.get("admin_hist_at") if isinstance(meta, dict) else None
        _ned_start = meta.get("nedeljni_start") if isinstance(meta, dict) else None

        # Gornje dugme „Ažuriraj iz admina" (za ceo prikaz) — za sistemske ga hvatamo ovde:
        # živo osvežava realni lager (objekti prelaze zone), ali NE dira startne podatke izveštaja.
        if st.session_state.pop("_req_refresh_admin", False) and not _zakljucan:
            st.session_state["_req_ned_pull"] = "refresh"

        # --- START: povuci JEDNOM na početku, zabeleži startno stanje izveštaja i zaključaj ---
        if _ned_start:
            st.markdown('<div style="background:#eef2ff;border:1px solid #c7d2fe;border-radius:10px;padding:9px 14px;'
                        'font-size:12.5px;color:#3730a3;margin:2px 0 6px;">📌 <b>Start izveštaja zabeležen</b> ('
                        + str(_ned_start.get("kada", "")) + '): <b>' + str(_ned_start.get("problem", 0))
                        + '</b> objekata sa problemom na startu (od ' + str(_ned_start.get("n", 0))
                        + '). Ovo je START u izveštaju — naknadna ažuriranja iz admina ga ne menjaju.</div>',
                        unsafe_allow_html=True)
        elif not _zakljucan:
            _scn = st.columns([2.4, 3])
            with _scn[0]:
                if st.button("▶ Start — povuci porudžbine i zabeleži start", key="ned_start_pull",
                             use_container_width=True, type="primary"):
                    st.session_state["_req_ned_pull"] = "start"
                    st.rerun()
            with _scn[1]:
                st.caption("Povuci JEDNOM na početku rada — povuku se porudžbine do sada, zabeleži se "
                           "startno stanje izveštaja i dugme se zaključava. Kasnije za osvežavanje koristi "
                           "dugme Ažuriraj iz admina gore (menja živi prikaz zona, ne dira start).")

        if _ah_at_n:
            st.caption("🔄 Poslednji put ažurirano iz admina: " + _dt_fmt(_ah_at_n)
                       + ". Realni lager = lager sa preseka + sve što su naknadno poručili (posle "
                       + _pk_ned_lbl + ").")

        # Handler za povlačenje (start ili živo osvežavanje) — glavni refresh handler je posle return-a
        _ned_pull_mode = st.session_state.pop("_req_ned_pull", None)
        if _ned_pull_mode and not _zakljucan:
            _bez_n = [o["idk"] for o in objekti if not ((komfull.get(int(o["idk"]), {}) or {}).get("naziv"))]
            with st.spinner("Povlačim naknadne porudžbine iz admina za ceo sistem... može potrajati par minuta."):
                if _bez_n:
                    admin_build_komitenti(only_ids=_bez_n)
                _kfn = sb_komitenti_full(); st.session_state["_komfull"] = _kfn
                _idn = {o["idk"]: (_kfn.get(int(o["idk"]), {}) or {}).get("naziv", "") for o in objekti}
                try:
                    _yy6 = int(str(mesec_key).split("-")[0]); _mm6 = int(str(mesec_key).split("-")[1]) - 5
                    while _mm6 <= 0:
                        _mm6 += 12; _yy6 -= 1
                    _cut6 = datetime.date(_yy6, _mm6, 1)
                except Exception:
                    _cut6 = _pk_ned
                # Inkrementalno: ako je već jednom povučeno, traži samo razliku
                _cov_n = set(str(x) for x in (meta.get("admin_hist_idk") or []))
                _cov_n |= set(str(x) for x in ((meta.get("admin_hist") or {}).keys()))
                _cut_n = None
                if _ned_pull_mode != "start" and _ah_at_n and _cov_n:
                    try:
                        _cut_n = datetime.date.fromisoformat(str(_ah_at_n)[:10]) - datetime.timedelta(days=5)
                        if _cut6 and _cut_n < _cut6:
                            _cut_n = _cut6
                    except Exception:
                        _cut_n = None
                if _cut_n:
                    _mi = {k: v for k, v in _idn.items() if str(int(k)) in _cov_n}
                    _mn = {k: v for k, v in _idn.items() if str(int(k)) not in _cov_n}
                    _bulkn, _ben = ({}, "")
                    if _mi:
                        _bulkn, _ben = admin_istorija_bulk(_mi, _cut_n)
                    if _mn:
                        _b2n, _e2n = admin_istorija_bulk(_mn, _cut6)
                        for _k2n, _v2n in (_b2n or {}).items():
                            _bulkn[_k2n] = (_bulkn.get(_k2n) or []) + (_v2n or [])
                        if _e2n and not _b2n:
                            _ben = _ben or _e2n
                else:
                    _bulkn, _ben = admin_istorija_bulk(_idn, _cut6)
            if _ben and not _bulkn:
                st.error(_ben)
            else:
                _prevh = (meta.get("admin_hist") or {}) if isinstance(meta, dict) else {}
                _hsave = {}
                for o in objekti:
                    _old = _prevh.get(str(int(o["idk"])), []) or []
                    _new = _bulkn.get(o["idk"], []) or []
                    _byid = {}
                    for _od in (_old + _new):
                        _oid = str(_od.get("id") or "")
                        if not _oid:
                            continue
                        if _oid not in _byid:
                            _byid[_oid] = _od
                        elif (not _byid[_oid].get("stavke")) and _od.get("stavke"):
                            _byid[_oid] = _od
                    _lst = sorted(_byid.values(), key=_datum_sort_key, reverse=True)
                    st.session_state["hist_" + str(sistem) + "_" + str(o["idk"])] = {"lst": _lst, "err": ""}
                    if _lst:
                        _hsave[int(o["idk"])] = _lst
                try:
                    sb_admin_hist_set(mesec_key, sistem, _hsave, _now().isoformat(),
                                      svi_idk=[int(o["idk"]) for o in objekti])
                except Exception as _e:
                    st.warning("Povučeno, ali nije trajno sačuvano: " + str(_e))
                if _ned_pull_mode == "start":
                    # Zabeleži STARTNO stanje izveštaja: broj objekata sa problemom na realnom lageru
                    # + spisak porudžbina koje svaki objekat ima U TOM TRENUTKU. Po tom spisku se
                    # realni lager (pa i predlog) ZAMRZAVA, a sve što stigne posle znači da je
                    # objekat poručio — prelazi u ZAVRŠENO.
                    _cnt_prob = 0
                    _start_ids_n = {}
                    for o in objekti:
                        _lz0 = _hsave.get(int(o["idk"]), []) or []
                        _start_ids_n[str(int(o["idk"]))] = [str(_x.get("id") or "")
                                                            for _x in _lz0 if _x.get("id")]
                        _pm0 = _treb_posle_preseka(_lz0, _cut_hit) if _cut_hit else {}
                        _has = False
                        for s in o["lst"]:
                            _lg0 = int(s.get('lager', 0) or 0) + int(_pm0.get(int(s.get('ida', -1)), 0) or 0)
                            _pr0 = int(s.get('pred', 0) or 0)
                            if _pr0 > 0 and _lg0 < _pr0 * _dani / 30.0:
                                _has = True; break
                        if _has:
                            _cnt_prob += 1
                    try:
                        sb_nedeljni_start_set(mesec_key, sistem, {"kada": _now().strftime("%d.%m.%Y %H:%M"),
                                                                  "n": len(objekti), "problem": _cnt_prob,
                                                                  "narudzbine": _start_ids_n})
                    except Exception as _e:
                        st.warning("Start nije trajno sačuvan: " + str(_e))
                    st.success("✅ Start zabeležen — povučene porudžbine do sada, predlog je zamrznut ("
                               + str(_cnt_prob) + " objekata sa problemom). Ko posle ovoga poruči, "
                               "prelazi u ZAVRŠENO.")
                else:
                    # Ko je posle starta poslao porudžbinu — trajno u ZAVRŠENO,
                    # sa automatski prepoznatim načinom trebovanja (po količinama).
                    _n_zav_n = 0
                    if _ned_start:
                        _kada_n = _now().isoformat()
                        for o in objekti:
                            _lzn = _hsave.get(int(o["idk"]), []) or []
                            try:
                                if not _porucio_posle_starta(o["idk"], _lzn, _ned_start):
                                    continue
                                _vo_n = obrada_map.get(int(o["idk"])) or {}
                                if (("Ubačena porudžbina" in (_vo_n.get("reakcije") or []))
                                        or (_vo_n.get("dnevnik") or {}).get("ubaceno")):
                                    if sb_oznaci_trebovao(mesec_key, sistem, o["idk"],
                                                          kada=_kada_n, n_por=len(_lzn)):
                                        _n_zav_n += 1
                                    continue
                                _pos_n = _treb_posle_starta(o["idk"], _lzn, _ned_start)
                                _st_n, _ = _treb_na_startu(o["idk"], _lzn, _ned_start, _cut_hit)
                                # traženo = manjak do praga na startu (to smo tražili od sistema)
                                _trz_n = {}; _nzv_n = {}
                                for s in o["lst"]:
                                    _ia_n = int(s.get("ida", -1))
                                    _lg_n = int(s.get("lager", 0) or 0) + int(_st_n.get(_ia_n, 0) or 0)
                                    _pr_n = int(s.get("pred", 0) or 0)
                                    _pg_n = int(round(_pr_n * _dani / 30.0))
                                    _trz_n[_ia_n] = max(_pg_n - _lg_n, 0) if _pr_n > 0 else 0
                                    _nzv_n[_ia_n] = str(s.get("naziv", "") or "")
                                _tp_n, _if_n = _nacin_trebovanja(_trz_n, _pos_n, _nzv_n)
                                _if_n["pravilo"] = 2
                                _nj_n = {str(_ia_n): int(_pos_n.get(_ia_n, 0)) for _ia_n in _trz_n
                                         if int(_pos_n.get(_ia_n, 0)) > 0}
                                if sb_oznaci_trebovao(mesec_key, sistem, o["idk"], kada=_kada_n,
                                                      n_por=len(_lzn), tip=_tp_n, njihova=_nj_n,
                                                      info=_if_n):
                                    _n_zav_n += 1
                            except Exception:
                                pass
                    st.success("🔄 Osveženo iz admina — predlog ostaje zamrznut sa starta."
                               + ((" ✅ " + str(_n_zav_n) + " objekata je poručilo posle starta i "
                                   "prešlo u ZAVRŠENO.") if _n_zav_n else
                                  " Nijedan objekat nije poručio posle starta."))
                st.rerun()

        # --- Realni lager po objektu i detekcija problema ---
        # Posle klika na Start realni lager (pa i predlog) je ZAMRZNUT: računa se samo
        # po porudžbinama koje su postojale u trenutku starta. Kasnije porudžbine ga ne
        # menjaju — one objekat prebacuju u ZAVRŠENO.
        def _por_map_for(idk):
            _hh = st.session_state.get("hist_" + str(sistem) + "_" + str(idk))
            if _hh and (_hh.get("lst") is not None) and _cut_hit:
                if _ned_start:
                    _m, _ok = _treb_na_startu(idk, _hh.get("lst") or [], _ned_start, _cut_hit)
                    if _ok:
                        return _m
                return _treb_posle_preseka(_hh.get("lst") or [], _cut_hit)
            return {}

        def _zavrsen_ned(idk):
            """Objekat je ZAVRŠEN ako je posle starta poslao porudžbinu."""
            _v0 = obrada_map.get(int(idk)) or {}
            if (_v0.get("dnevnik") or {}).get("trebovao_posle_starta"):
                return True
            if (_v0.get("trebovali_tip") or "") in ("nas", "njihov"):
                return True
            if "Ubačena porudžbina" in (_v0.get("reakcije") or []):
                return True
            try:
                _hz = st.session_state.get("hist_" + str(sistem) + "_" + str(idk))
                if _hz and not _hz.get("err") and _ned_start:
                    return _porucio_posle_starta(idk, _hz.get("lst") or [], _ned_start)
            except Exception:
                pass
            return False

        def _prob_arts(lst, por_map):
            out = []
            for s in lst:
                _por = int(por_map.get(int(s.get('ida', -1)), 0) or 0)
                _lg = int(s.get('lager', 0) or 0) + _por      # REALNI lager
                _pr = int(s.get('pred', 0) or 0)
                _prag = _pr * _dani / 30.0
                if _pr > 0 and _lg < _prag:   # artikal koji se prodaje, a realni lager ispod perioda (uklj. 0)
                    out.append({"ida": int(s.get('ida', -1)), "naziv": str(s.get('naziv', '')),
                                "lager": _lg, "pred": _pr, "prag": int(round(_prag)),
                                "manjak7": max(int(round(_prag)) - _lg, 0)})
            return out
        _prob = []
        for o in objekti:
            _pm = _por_map_for(o["idk"])
            _pa = _prob_arts(o["lst"], _pm)
            if _pa:
                _prob.append({"idk": o["idk"], "arts": _pa, "lst": o["lst"],
                              "zav": _zavrsen_ned(o["idk"])})
        _n_zav_lst = sum(1 for p in _prob if p.get("zav"))

        # koji su komitenti već prijavljeni komercijali (iz meta.nedeljni_prijave)
        _prijave = dict(meta.get("nedeljni_prijave") or {}) if isinstance(meta, dict) else {}
        _n_prij = sum(1 for p in _prob if str(int(p["idk"])) in _prijave)
        # pravi ukupan broj objekata u sistemu (svi komitenti iz fajla), ako je sačuvan pri objavi
        _sis_uk = int(meta.get("n_sistem_ukupno", 0) or 0) if isinstance(meta, dict) else 0
        _uk_val = _sis_uk if _sis_uk > 0 else len(objekti)
        _uk_lbl = "Ukupno objekata u sistemu" if _sis_uk > 0 else "Objekata sa porudžbinom (u izveštaju)"

        st.markdown('<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:6px;">'
                    '<div style="background:#fff7f7;border:1px solid #fecaca;border-radius:12px;padding:15px 18px;">'
                    '<div style="font-size:22px;font-weight:800;color:#dc2626;">'
                    + str(len(_prob) - _n_zav_lst) + '</div>'
                    '<div style="font-size:12px;color:#9b6b6b;margin-top:3px;">Još u problemu (0 lagera ili < ' + _per_lbl + ')</div></div>'
                    '<div style="background:#dcfce7;border:1px solid #86efac;border-radius:12px;padding:15px 18px;">'
                    '<div style="font-size:22px;font-weight:800;color:#14532d;">' + str(_n_zav_lst) + '</div>'
                    '<div style="font-size:12px;color:#166534;margin-top:3px;font-weight:700;">✓ Završeno (poručili posle starta)</div></div>'
                    '<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:12px;padding:15px 18px;">'
                    '<div style="font-size:22px;font-weight:800;color:#b45309;">' + str(_n_prij) + '</div>'
                    '<div style="font-size:12px;color:#9a7b3a;margin-top:3px;">Prijavljeno komercijali</div></div>'
                    '<div style="background:#faf7ff;border:1px solid #e9d5ff;border-radius:12px;padding:15px 18px;">'
                    '<div style="font-size:22px;font-weight:800;color:#7c3aed;">' + str(_uk_val) + '</div>'
                    '<div style="font-size:12px;color:#8b7fa8;margin-top:3px;">' + _uk_lbl + '</div></div></div>',
                    unsafe_allow_html=True)
        if _ned_start:
            st.caption("🔒 Predlog je zamrznut na startu (" + str(_ned_start.get("kada", ""))
                       + ") — količine se više ne menjaju. Objekat koji posle toga poruči prelazi u "
                       "ZAVRŠENO i ne šalje mu se ponovo.")
        if _sis_uk > 0:
            st.caption("U sistemu je ukupno " + str(_sis_uk) + " objekata; porudžbina je generisana za "
                       + str(len(objekti)) + " (ostali su imali dovoljno zaliha, pa nemaju porudžbinu). "
                       "Objekata sa problemom (0 lagera ili ispod prodaje za " + _per_lbl + "): " + str(len(_prob)) + ".")
        else:
            st.caption("Napomena: prikazani broj su objekti kojima je generisana porudžbina. Objekti koji su imali "
                       "dovoljno zaliha u svemu se ne čuvaju u izveštaju. Za tačan ukupan broj objekata u sistemu, "
                       "ponovo objavi ovaj sistem (novi podatak se tada beleži).")

        if not _prob:
            st.success("Nema objekata sa problemom — svi imaju dovoljno zaliha za " + _per_lbl + ".")
            return

        if _n_zav_lst and _n_zav_lst >= len(_prob):
            st.success("✅ Svi objekti sa problemom su poručili posle starta — nema kome da se šalje.")
        st.markdown("<div style='margin:6px 0 8px;font-weight:700;font-size:14px;'>Objekti sa problemom "
                    "— klikni na objekat da vidiš artikle, pa štikliraj Prijavi problem komercijali."
                    + ("  <span style='font-weight:600;color:#166534;'>Zeleni (✅) su završeni — "
                       "poručili su i ne idu u prilog.</span>" if _n_zav_lst else "") + "</div>",
                    unsafe_allow_html=True)

        # sortiraj: završeni na dno, pa neprijavljeni, pa oni sa najviše problema
        _prob.sort(key=lambda p: (bool(p.get("zav")), str(int(p["idk"])) in _prijave, -len(p["arts"])))

        for p in _prob:
            _idk = int(p["idk"])
            _kinfo = komfull.get(_idk, {}) or {}
            _nz = _kinfo.get("naziv", "") or ("ID " + str(_idk))
            _tel = _kinfo.get("telefon", "") or ""
            _mail = _kinfo.get("email", "") or ""
            _is_prij = str(_idk) in _prijave
            _je_zav = bool(p.get("zav"))
            _hdr = (("✅ " if _je_zav else ("🟠 " if _is_prij else "🔴 ")) + str(_nz)
                    + "   ·   " + str(len(p["arts"])) + " art. u problemu")
            if _je_zav:
                _hdr += "   ·   ZAVRŠENO — poručio posle starta"
            if _tel:
                _hdr += "   ·   📞 " + str(_tel)
            if _mail:
                _hdr += "   ·   ✉️ " + str(_mail)
            if _is_prij:
                _hdr += "   ·   ✅ prijavljeno"
            with st.expander(_hdr, expanded=False):
                if _je_zav:
                    _vz = obrada_map.get(_idk) or {}
                    _az = (_vz.get("dnevnik") or {}).get("trebovao_posle_starta") or {}
                    _tz = str(_az.get("tip") or _vz.get("trebovali_tip") or "")
                    _iz = dict(_az.get("info") or {})
                    st.markdown('<div style="background:#dcfce7;border:2px solid #16a34a;border-radius:10px;'
                                'padding:10px 14px;margin:0 0 10px;">'
                                '<div style="font-size:15px;font-weight:900;color:#14532d;">✅ ZAVRŠENO — OBJEKAT JE PORUČIO</div>'
                                + (('<div style="font-size:12.5px;font-weight:700;color:#166534;margin-top:3px;">Trebovano '
                                    + ("PO NAŠEM SISTEMU" if _tz == "nas" else "PO NJIHOVOM SISTEMU") + '</div>')
                                   if _tz in ("nas", "njihov") else '')
                                + (('<div style="font-size:12px;color:#166534;margin-top:2px;">Poručeno '
                                    + str(_iz.get("poruceno", 0)) + ' kom od traženih ' + str(_iz.get("trazeno", 0))
                                    + ' kom (' + str(_iz.get("procenat", 0)) + '%).</div>')
                                   if _iz.get("trazeno") else '')
                                + '<div style="font-size:11.5px;color:#15803d;margin-top:5px;font-style:italic;">'
                                  'Ne šalje mu se ponovo — nije u prilogu predloga.</div></div>',
                                unsafe_allow_html=True)
                _adf_p = pd.DataFrame([{"Artikal": a["naziv"], "Realni lager": a["lager"],
                                        "Predikcija (mes.)": a["pred"], ("Za " + _per_lbl + " (~)"): a["prag"],
                                        ("Manjak (" + _per_lbl + ")"): a.get("manjak7", 0)}
                                       for a in p["arts"]])
                st.dataframe(_adf_p, hide_index=True, use_container_width=True)
                _kb = []
                if _mail:
                    _kb.append("✉️ " + _h_escape(_mail))
                if _tel:
                    _kb.append("📞 " + _h_escape(_tel))
                if _kinfo.get("mesto"):
                    _kb.append("📍 " + _h_escape(_kinfo["mesto"]))
                if _kb:
                    st.markdown('<div style="color:#6b7280;font-size:12.5px;margin:2px 0 8px;">'
                                + "&nbsp;&nbsp;·&nbsp;&nbsp;".join(_kb) + '</div>', unsafe_allow_html=True)
                _saved_nap = (_prijave.get(str(_idk), {}) or {}).get("napomena", "") if _is_prij else ""
                _chk = st.checkbox("⚠️ Prijavi problem komercijali", value=_is_prij,
                                   key="ndp_" + str(sistem) + "_" + str(mesec_key) + "_" + str(_idk),
                                   disabled=_zakljucan)
                _nap_k = st.text_input("Napomena (opciono — šta je preduzeto / dogovoreno)",
                                       value=_saved_nap,
                                       key="ndn_" + str(sistem) + "_" + str(mesec_key) + "_" + str(_idk),
                                       disabled=_zakljucan or not _chk,
                                       placeholder="npr. zvali, javiće se u ponedeljak…")
                if _is_prij and _prijave.get(str(_idk), {}).get("at"):
                    _pk_ko = _prijave[str(_idk)].get("ko", "")
                    st.caption("Prijavljeno: " + str(_prijave[str(_idk)]["at"])
                               + ((" · " + str(_pk_ko)) if _pk_ko else ""))
                if not _zakljucan:
                    # sačuvaj samo ako se stanje promenilo (checkbox ili napomena)
                    _changed = (_chk != _is_prij) or (_chk and (_nap_k or "") != (_saved_nap or ""))
                    if st.button("💾 Sačuvaj", key="ndsave_" + str(sistem) + "_" + str(mesec_key) + "_" + str(_idk),
                                 type="primary", disabled=not _changed):
                        try:
                            sb_nedeljni_prijava_set(mesec_key, sistem, _idk, _chk, _nap_k,
                                                    ko=st.session_state.get("admin_user", "Administracija"))
                            st.success("Sačuvano.")
                            st.rerun()
                        except Exception as _e:
                            st.error("Greška pri čuvanju: " + str(_e))

        # ===== Tekst mejla za glavni kontakt + PDF prilog =====
        st.markdown("<hr style='margin:18px 0 10px;border:none;border-top:1px solid #e5e7eb;'>",
                    unsafe_allow_html=True)
        st.markdown("<div style='font-weight:700;font-size:15px;margin:2px 0 2px;'>✉️ Mejl za glavni kontakt + prilog</div>",
                    unsafe_allow_html=True)
        st.caption("Ovi objekti se ne zovu pojedinačno — pošalji jedan mejl glavnom kontaktu, "
                   "a u prilogu su svi problematični objekti (predikcija, izgubljena prodaja, kritične zalihe).")

        # imena objekata + ključni brojevi + primer (artikal sa najvećom prodajom, prvenstveno na lageru 0)
        _imena = []
        _grupe = []          # za Excel predlog porudžbine: [{objekat, arts:[...]}]
        _red_obj = []        # za tabelu u PDF-u: [{ime, na_nuli, kriticnih, manjak}]
        _svi_art = []        # rezerva za PDF ako nema grafika: [{obj, art, lager, pred, manjak}]
        _art_na_nuli = 0
        _obj_sa_nulom = 0
        _izgub7 = 0
        _mes_naz = list(meta.get("mesec_nazivi", []) or []) if isinstance(meta, dict) else []
        # ako sistem nije objavljen novijom verzijom, mesečnu prodaju vadimo iz
        # analitika Excel-a koji je sačuvan pri objavi (da grafik uvek postoji)
        _ser_rez, _lab_rez = ({}, [])
        _treba_rez = not _mes_naz or not any(
            (s0.get("prodaja_mesecno") or []) for o0 in objekti for s0 in (o0.get("lst") or []))
        if _treba_rez:
            try:
                _ser_rez, _lab_rez = _serije_prodaje_iz_analitike(mesec_key, sistem)
            except Exception:
                _ser_rez, _lab_rez = ({}, [])
            if _lab_rez and not _mes_naz:
                _mes_naz = _lab_rez
        _izvor_serije = ("objava" if not _treba_rez else ("analitika" if _ser_rez else "nema"))
        _cand0 = None; _cand0_sc = -1       # kandidati sa lagerom 0
        _candA = None; _candA_sc = -1       # svi problem artikli (fallback)
        _ima_serije = False
        # Objekti koji su već poručili (ZAVRŠENO) NE ulaze u predlog ni u mejl.
        _prob_mail = [p for p in _prob if not p.get("zav")]
        for p in _prob_mail:
            _idk = int(p["idk"])
            _nz = (komfull.get(_idk, {}) or {}).get("naziv", "") or ("ID " + str(_idk))
            _nz_s = _nz.replace(str(sistem), "").strip(" -—") or _nz
            _imena.append(_nz_s)
            _ima_nulu = False
            # mapa ida -> mesečna prodaja za ovaj objekat
            _ser_map = {}
            for s in (p.get("lst") or []):
                _sid = int(s.get("ida", -1))
                _sv = [int(x or 0) for x in (s.get("prodaja_mesecno") or [])]
                if not _sv:
                    _sv = list(_ser_rez.get((_idk, _sid), []) or [])
                _ser_map[_sid] = _sv
            for a in p["arts"]:
                _izgub7 += int(a.get("manjak7", 0) or 0)
                _lg = int(a.get("lager", 0) or 0)
                if _lg <= 0:
                    _art_na_nuli += 1
                    _ima_nulu = True
                _pr = int(a.get("pred", 0) or 0)
                _ser = _ser_map.get(int(a.get("ida", -2)), [])
                if _ser and sum(_ser) > 0:
                    _ima_serije = True
                    _pl = {"obj": _nz_s, "art": a["naziv"], "meseci": _mes_naz, "prodaja": _ser,
                           "ned": int(round(_pr * 7 / 30.0)), "mes": _pr, "lager": _lg,
                           "pred7": int(round(_pr * _dani / 30.0))}
                    # najbolji primer = onaj koji se prodaje u NAJVIŠE meseci (duga, jasna istorija),
                    # pa tek onda onaj sa najvećom ukupnom prodajom
                    _n_mes = sum(1 for _x in _ser if int(_x or 0) > 0)
                    _sc = _n_mes * 10000 + int(sum(_ser))
                    if _sc > _candA_sc:
                        _candA_sc = _sc; _candA = _pl
                    if _lg <= 0 and _sc > _cand0_sc:
                        _cand0_sc = _sc; _cand0 = _pl
            _n0_obj = sum(1 for a in p["arts"] if int(a.get("lager", 0) or 0) <= 0)
            _mj_obj = sum(int(a.get("manjak7", 0) or 0) for a in p["arts"])
            _red_obj.append({"ime": _nz_s, "na_nuli": _n0_obj,
                             "kriticnih": len(p["arts"]), "manjak": _mj_obj})
            for a in p["arts"]:
                _svi_art.append({"obj": _nz_s, "art": str(a.get("naziv", "")),
                                 "lager": int(a.get("lager", 0) or 0),
                                 "pred": int(a.get("pred", 0) or 0),
                                 "manjak": int(a.get("manjak7", 0) or 0)})
            _grupe.append({"objekat": _nz_s, "arts": [
                {"naziv": str(a.get("naziv", "")), "lager": int(a.get("lager", 0) or 0),
                 "pred": int(a.get("pred", 0) or 0), "predlog": int(a.get("manjak7", 0) or 0)}
                for a in p["arts"] if int(a.get("manjak7", 0) or 0) > 0]})
            if _ima_nulu:
                _obj_sa_nulom += 1
        _primer_pl = _cand0 or _candA   # prvo artikal na 0, ako ga nema — najgori sa prodajom

        # datum lagera = poslednji dan meseca pre preseka; dodatne do = danas
        _lager_datum = ""
        try:
            if _cut_hit:
                import datetime as _dtp
                _ld = _cut_hit - _dtp.timedelta(days=1)
                _lager_datum = _ld.strftime("%d.%m.%Y.")
        except Exception:
            _lager_datum = ""
        _red_obj.sort(key=lambda r: (-int(r.get("manjak", 0) or 0), -int(r.get("na_nuli", 0) or 0),
                                     str(r.get("ime", ""))))
        # procena izgubljenog prometa objekata (prodajna cena po komadu iz analitika Excel-a)
        _izgub_rsd = 0
        try:
            _cene = _cene_iz_analitike(mesec_key, sistem)
            if _cene:
                for _pp in _prob_mail:
                    for _aa in _pp["arts"]:
                        _c1 = _cene.get(int(_aa.get("ida", -1)), 0)
                        if _c1:
                            _izgub_rsd += int(_aa.get("manjak7", 0) or 0) * _c1
                _izgub_rsd = int(round(_izgub_rsd))
        except Exception:
            _izgub_rsd = 0
        _svi_art.sort(key=lambda r: (-int(r.get("manjak", 0) or 0), -int(r.get("pred", 0) or 0)))

        # ============ DOKAZI za prilog sistemu ============
        # Sve iz podataka koje već imamo. Ako nešto fali, taj deo izveštaja se izostavi.
        _dok = {}
        try:
            _ahist_p = (meta.get("admin_hist") or {}) if isinstance(meta, dict) else {}

            def _ser_za(_ik, _s):
                _v = [int(x or 0) for x in (_s.get("prodaja_mesecno") or [])]
                if not _v:
                    _v = [int(x or 0) for x in (_ser_rez.get((int(_ik), int(_s.get("ida", -1))), []) or [])]
                return _v

            # 1) prodato po mesecima — ceo sistem (svi objekti, ne samo problematični)
            _mes_prod = [0] * len(_mes_naz)
            _obj_prod = {}          # idk -> (ime, [prodato po mesecima])
            for _o9 in objekti:
                _ik9 = int(_o9["idk"])
                _zbir = [0] * len(_mes_naz)
                for _s9 in (_o9.get("lst") or []):
                    _v9 = _ser_za(_ik9, _s9)
                    for _t9 in range(min(len(_v9), len(_zbir))):
                        _zbir[_t9] += _v9[_t9]
                        _mes_prod[_t9] += _v9[_t9]
                _nm9 = (komfull.get(_ik9, {}) or {}).get("naziv", "") or ("ID " + str(_ik9))
                _obj_prod[_ik9] = (_nm9.replace(str(sistem), "").strip(" -—") or _nm9, _zbir)

            # 2) poručeno po mesecima — iz istorije porudžbina u adminu
            _mes_por = _poruceno_po_mesecima(_ahist_p, _mes_naz)

            if _mes_naz and sum(_mes_prod) > 0 and sum(_mes_por) > 0:
                _dok["meseci"] = list(_mes_naz)
                _dok["prodato"] = _mes_prod
                _dok["poruceno"] = _mes_por

            # 2b) raspodela objekata: bez robe / na granici / u redu
            _z_crv = _z_zut = _z_zel = 0
            for _o9 in objekti:
                _pm9 = _por_map_for(_o9["idk"])
                _n9, _, _ = hitnost_objekta_dodatna(_o9["lst"], _pm9)
                if _n9 == "crveno":
                    _z_crv += 1
                elif _n9 == "zuto":
                    _z_zut += 1
                else:
                    _z_zel += 1
            _dok["zone_objekti"] = [("Roba nedostaje", _z_crv), ("Na granici", _z_zut),
                                    ("Zalihe u redu", _z_zel)]

            # 2c) stanje SVIH artikala koji se prodaju (ne samo problematičnih)
            _a_nula = _a_ispod = _a_ok = 0
            for _o9 in objekti:
                _pm9 = _por_map_for(_o9["idk"])
                for _s9 in (_o9.get("lst") or []):
                    _pr9 = int(_s9.get("pred", 0) or 0)
                    if _pr9 <= 0:
                        continue          # artikal se ne prodaje — ne ulazi u račun
                    _lg9 = int(_s9.get("lager", 0) or 0) + int(_pm9.get(int(_s9.get("ida", -1)), 0) or 0)
                    _prag9 = _pr9 * _dani / 30.0
                    if _lg9 <= 0:
                        _a_nula += 1
                    elif _lg9 < _prag9:
                        _a_ispod += 1
                    else:
                        _a_ok += 1
            if (_a_nula + _a_ispod + _a_ok) > 0:
                _dok["stanje_artikala"] = [("Lager 0", _a_nula), ("Ispod minimuma", _a_ispod),
                                           ("Dovoljno", _a_ok)]

            # 3) koliko puta je objekat uopšte poručio
            if _ahist_p and _mes_naz:
                _dok["ucestalost"] = _ucestalost_trebovanja(
                    _ahist_p, _mes_naz, [int(o["idk"]) for o in objekti])

            # 4) objekti sa najvećom izgubljenom prodajom (u dinarima)
            try:
                _cene_d = _cene_iz_analitike(mesec_key, sistem)
            except Exception:
                _cene_d = {}
            if _cene_d:
                _rang = []
                for _pp in _prob_mail:
                    _r_rsd = 0; _r_n0 = 0
                    for _aa in _pp["arts"]:
                        _c9 = _cene_d.get(int(_aa.get("ida", -1)), 0)
                        if _c9:
                            _r_rsd += int(_aa.get("manjak7", 0) or 0) * _c9
                        if int(_aa.get("lager", 0) or 0) <= 0:
                            _r_n0 += 1
                    _nm9 = (komfull.get(int(_pp["idk"]), {}) or {}).get("naziv", "") or ("ID " + str(_pp["idk"]))
                    _rang.append({"ime": _nm9.replace(str(sistem), "").strip(" -—") or _nm9,
                                  "na_nuli": _r_n0, "rsd": int(round(_r_rsd))})
                _rang.sort(key=lambda r: -r["rsd"])
                _dok["top_obj"] = [r for r in _rang if r["rsd"] > 0][:8]

            # 5) dva imenovana primera
            if _ahist_p and _mes_naz:
                _kand = []
                _lag_obj = {}
                for _pp in _prob_mail:
                    _ik9 = int(_pp["idk"])
                    _nm9 = _obj_prod.get(_ik9, ("ID " + str(_ik9), []))[0]
                    _lag_obj[_ik9] = sum(int(a.get("lager", 0) or 0) for a in _pp["arts"])
                    _po_ida = {int(s.get("ida", -1)): s for s in (_pp.get("lst") or [])}
                    for _aa in _pp["arts"]:
                        _ida9 = int(_aa.get("ida", -1))
                        _s9 = _po_ida.get(_ida9)
                        if _s9 is None:
                            continue
                        _kand.append({"idk": _ik9, "ida": _ida9, "obj": _nm9,
                                      "art": str(_aa.get("naziv", "")),
                                      "lager": int(_aa.get("lager", 0) or 0),
                                      "prodaja": _ser_za(_ik9, _s9)})
                _dok["primer_neredovno"] = _primer_neredovno(_kand, _ahist_p, _mes_naz)
                _dok["primer_nedovoljno"] = _primer_nedovoljno(
                    {k: v for k, v in _obj_prod.items() if k in _lag_obj}, _ahist_p, _mes_naz, _lag_obj)

            # 6) vrednost predloga u dinarima
            if _cene_d:
                _dok["predlog_rsd"] = int(round(sum(
                    int(_aa.get("manjak7", 0) or 0) * _cene_d.get(int(_aa.get("ida", -1)), 0)
                    for _pp in _prob_mail for _aa in _pp["arts"])))
        except Exception:
            _dok = {}

        _payload_n = {"sistem": str(sistem), "datum": _now().strftime("%d.%m.%Y."), "dani": _dani,
                      "dokazi": _dok, "n_obj_ukupno": len(objekti), "n_obj_problem": len(_prob_mail),
                      "predlog_kom": sum(int(a.get("predlog", 0) or 0) for g in _grupe for a in g.get("arts", [])),
                      "lager_datum": _lager_datum, "dodatne_do": _now().strftime("%d.%m.%Y."),
                      "objekti_imena": _imena, "objekti_red": _red_obj, "top_arts": _svi_art[:6],
                      "art_na_nuli": _art_na_nuli, "izgub_rsd": _izgub_rsd,
                      "zahtev": str(_izgub7) if _izgub7 > 0 else "",
                      "obj_sa_nulom": _obj_sa_nulom, "izgub7": _izgub7, "primer": _primer_pl}

        # ===== Izbor šta se šalje: izveštaj o problemu ili predlog porudžbine =====
        _grupe_ok = [g for g in _grupe if (g.get("arts") or [])]
        _predlog_uk = sum(int(a.get("predlog", 0) or 0) for g in _grupe_ok for a in g["arts"])

        _dat_str = _now().strftime("%d.%m.%Y.")
        # Izbor: sa prilogom (Excel) ili samo mejl sa spiskom objekata u telu.
        _bezp_key = "ned_bez_priloga_" + str(sistem) + "_" + str(mesec_key)
        _bez_priloga = st.checkbox(
            "Pošalji bez priloga — spisak objekata ide u telu mejla",
            key=_bezp_key,
            help="Kad je štiklirano, ne šalje se Excel. Umesto toga se u mejl upisuje spisak "
                 "objekata sa brojem artikala i predloženom količinom, pa primalac sve vidi "
                 "odmah u poruci.")

        _n_ob = len(_grupe_ok)
        # „Kod …“ traži genitiv: 1–4 objekta, 5 i više objekata
        _n10, _n100 = _n_ob % 10, _n_ob % 100
        _ob_rec = ("objekta" if (_n10 in (1, 2, 3, 4) and _n100 not in (11, 12, 13, 14))
                   else "objekata")

        def _art_rec(n):
            _a10, _a100 = n % 10, n % 100
            if _a10 == 1 and _a100 != 11:
                return "artikal"
            if _a10 in (2, 3, 4) and _a100 not in (12, 13, 14):
                return "artikla"
            return "artikala"

        if _bez_priloga:
            _mail_subj_n = ("Predlog dopune zaliha — " + str(sistem) + " (" + _dat_str + ")")
            _spisak = ""
            for _gi, _g in enumerate(_grupe_ok, 1):
                _kom = sum(int(a.get("predlog", 0) or 0) for a in _g["arts"])
                _na = len(_g["arts"])
                _spisak += (str(_gi) + ". " + str(_g.get("objekat", "")) + " — "
                            + str(_na) + " " + _art_rec(_na) + ", predlog " + str(_kom) + " kom\n")
            _mail_body_n = ("Poštovani,\n\n"
                            "Kod " + str(_n_ob) + " " + _ob_rec + " trenutne zalihe ne pokrivaju "
                            "prodaju ni za " + _per_lbl + ". U nastavku je spisak objekata sa brojem "
                            "artikala kojima roba nedostaje i predloženom količinom za dopunu.\n\n"
                            + _spisak + "\n"
                            "Ukupan predlog dopune: " + str(_predlog_uk) + " kom.\n\n"
                            "Molimo da se roba dopuni kako bi objekti mogli da zadrže kontinuitet "
                            "prodaje. Ako vam treba detaljan pregled po artiklima, rado ga šaljemo "
                            "u Excel tabeli.\n\n"
                            "Srdačan pozdrav")
        else:
            _mail_subj_n = "Stanje zaliha i predlog dopune — " + str(sistem) + " (" + _dat_str + ")"
            # Spisak objekata se NE nabraja u mejlu — ceo je u prilogu, na listu
            # „Predlog po objektima“. U mejlu ostaju samo dva broja.
            _mail_body_n = ("Poštovani,\n\n"
                            "U prilogu vam šaljemo pregled stanja zaliha u vašim objektima i predlog "
                            "dopune. Kod " + str(_n_ob) + " " + _ob_rec + " trenutne zalihe ne pokrivaju "
                            "prodaju ni za " + _per_lbl + ". Ukupan predlog dopune je "
                            + str(_predlog_uk) + " kom.\n\n"
                            "Prilog ima tri lista:\n"
                            "1. Predlog po objektima — spisak objekata sa artiklima i predloženim količinama\n"
                            "2. Tabela — isti podaci u ravnom obliku, za filtriranje\n"
                            "3. Izveštaj — kratak pregled stanja sa grafikonima i primerom\n\n"
                            "Molimo da se roba dopuni kako bi objekti mogli da zadrže kontinuitet prodaje.\n\n"
                            "Srdačan pozdrav")

        _sfx = ("bez" if _bez_priloga else "pre")
        _subj_key = "ned_mail_subj_" + _sfx + "_" + str(sistem) + "_" + str(mesec_key)
        _body_key = "ned_mail_body_" + _sfx + "_" + str(sistem) + "_" + str(mesec_key)
        st.text_input("Naslov mejla", value=_mail_subj_n, key=_subj_key)
        st.text_area("Tekst mejla (možeš da izmeniš pre slanja)", value=_mail_body_n,
                     height=280, key=_body_key)

        # --- Prilog (PDF izveštaj ili Excel predlog) ---
        _pdf_bytes_n = None
        _prilog_ime = ""
        _prilog_mime = ""
        # U nazivu fajla ide DATUM SLANJA (ne mesec izveštaja) — tako se u sandučetu
        # odmah vidi kad je predlog poslat.
        _osnova = str(sistem).replace(" ", "_") + "_" + _now().strftime("%d.%m.%Y")
        try:
            if _bez_priloga:
                st.caption("📧 Šalje se samo mejl, bez priloga — spisak od " + str(len(_grupe_ok))
                           + " objekata je u tekstu iznad (ukupno " + str(_predlog_uk) + " kom).")
            elif not _grupe_ok:
                st.info("Nema nijednog artikla za predlog — svi objekti imaju dovoljan lager.")
            else:
                _pdf_bytes_n = _nedeljni_predlog_xlsx(str(sistem), _dani, _dat_str, _grupe_ok,
                                                      payload=_payload_n)
                _prilog_ime = "Stanje_zaliha_i_predlog_" + _osnova + ".xlsx"
                _prilog_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if _pdf_bytes_n is not None:
                _dc1, _dc2 = st.columns([1.4, 3])
                with _dc1:
                    st.download_button(
                        "⬇️ Preuzmi prilog (Excel)", _pdf_bytes_n, file_name=_prilog_ime,
                        mime=_prilog_mime,
                        key="ned_prilog_dl_" + _sfx + "_" + str(sistem) + "_" + str(mesec_key),
                        use_container_width=True)
                with _dc2:
                    st.caption("Prilog ima tri lista: „Predlog po objektima“ (zbir po objektu), "
                               "„Tabela“ (ravan spisak sa filterom) i „Izveštaj“ (grafikoni, primer "
                               "i zahtev). Ukupno " + str(_predlog_uk) + " kom za "
                               + str(len(_grupe_ok)) + " objekata.")
                    # Jasno reci ŠTA fali u izveštaju i zašto — da se ne pogađa
                    _dd = _payload_n.get("dokazi") or {}
                    _fali = []
                    if not _dd.get("zone_objekti"):
                        _fali.append("krug „Gde je roba nestala“")
                    if not _dd.get("ucestalost"):
                        _fali.append("krug „Kako se poručuje“")
                    if not _dd.get("stanje_artikala"):
                        _fali.append("krug „Koliko artikala fali“")
                    if not _dd.get("primer_neredovno"):
                        _fali.append("primer sa grafikom")
                    if _fali:
                        _ima_hist = bool((meta.get("admin_hist") or {}) if isinstance(meta, dict) else {})
                        _ima_mes = bool(_mes_naz)
                        if not _ima_hist:
                            _zasto = ("za ovaj sistem nije povučena istorija porudžbina iz admina. "
                                      "Rešenje: klikni „🔄 Ažuriraj iz admina“ u zaglavlju, sačekaj da "
                                      "završi, pa ponovo napravi prilog.")
                        elif not _ima_mes:
                            _zasto = ("uz ovaj sistem nije sačuvana mesečna prodaja po artiklu. "
                                      "Rešenje: u kartici Objava izveštaja ponovo objavi ovaj sistem "
                                      "za ovaj mesec (isti fajl, isti parametri).")
                        else:
                            _zasto = ("nema dovoljno poklapanja između prodaje po artiklu i porudžbina "
                                      "iz admina za ovaj period.")
                        st.warning("📉 U izveštaju nedostaje: " + ", ".join(_fali) + " — " + _zasto)
        except Exception as _pe:
            st.error("Greška pri pravljenju priloga: " + str(_pe))

        # --- Slanje mejla nadležnom (glavni kontakt) sa prilogom ---
        _mail_meta = dict(meta.get("nedeljni_mail") or {}) if isinstance(meta, dict) else {}
        # Mejl koji je zadala analitika pri objavi — administracija ga NE menja
        _mfix = (meta.get("mail_to_fix") or {}) if isinstance(meta, dict) else {}
        _mfix_to = str(_mfix.get("to", "") or "").strip()
        _to_default = _mfix_to or (_mail_meta.get("to", "") or "")
        _to_key = "ned_mail_to_" + str(sistem) + "_" + str(mesec_key)
        if _mfix_to:
            st.session_state[_to_key] = _mfix_to      # uvek stoji zadati mejl
        st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)
        _mc1, _mc2 = st.columns([2.2, 1.2])
        with _mc1:
            _to_val = st.text_input(
                ("Mejl nadležnog 🔒 (zadala analitika — ne menja se)" if _mfix_to
                 else "Mejl nadležnog (glavni kontakt)"),
                value=_to_default, key=_to_key, placeholder="npr. nabavka@univerexport.rs",
                disabled=(_zakljucan or bool(_mfix_to)))
        with _mc2:
            st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
            _send_disabled = (_zakljucan or not smtp_dostupan()
                              or (_pdf_bytes_n is None and not _bez_priloga)
                              or (_bez_priloga and not _grupe_ok))
            if st.button("✉️ Pošalji mejl", key="ned_mail_send_" + str(sistem) + "_" + str(mesec_key),
                         type="primary", use_container_width=True, disabled=_send_disabled):
                _to_send = _mfix_to or (st.session_state.get(_to_key, "") or "").strip()
                _subj_send = st.session_state.get(_subj_key, _mail_subj_n)
                _body_send = st.session_state.get(_body_key, _mail_body_n)
                if not _to_send or "@" not in _to_send:
                    st.error("Upiši ispravan mejl nadležnog.")
                else:
                    try:
                        with st.spinner("✉️ Slanje mejla u toku… (može da potraje do 2 minuta, "
                                        "ne zatvaraj stranu)"):
                            posalji_mejl_sa_prilogom(
                                _to_send, _subj_send, _body_send, attach_bytes=_pdf_bytes_n,
                                attach_filename=_prilog_ime)
                        try:
                            sb_nedeljni_mail_set(mesec_key, sistem, _to_send,
                                                 ko=st.session_state.get("admin_user", "Administracija"),
                                                 poslato=True)
                        except Exception:
                            pass
                        _kk = st.session_state.get("_zadnja_kopija")
                        st.success("✅ Mejl poslat na " + _to_send
                                   + (" · bez priloga (spisak u mejlu)" if _bez_priloga
                                      else (" · prilog: " + _prilog_ime)))
                        if _kk and not _kk[0]:
                            st.warning("Mejl je poslat, ali kopija nije upisana u Poslato: "
                                       + str(_kk[1]))
                        st.rerun()
                    except Exception as _me:
                        st.error("Slanje nije uspelo: " + str(_me))
        if not smtp_dostupan():
            st.caption("ℹ️ Slanje mejlova nije podešeno za tvoj nalog (SMTP u Secrets). "
                       "Možeš da preuzmeš prilog i pošalješ ručno.")
        elif _mail_meta.get("at"):
            st.caption("📧 Poslednji put poslato: " + str(_mail_meta.get("at"))
                       + (" · " + str(_mail_meta.get("ko")) if _mail_meta.get("ko") else "")
                       + (" · " + str(_mail_meta.get("to")) if _mail_meta.get("to") else ""))
        elif _mfix_to:
            st.caption("🔒 Mejl je zadala analitika pri objavi: " + _mfix_to
                       + ((" (" + str(_mfix.get("ko")) + " · " + str(_mfix.get("at")) + ")")
                          if _mfix.get("ko") else "")
                       + ". Klikni Pošalji mejl da pošalješ sa prilogom.")
        elif _to_default:
            st.caption("Zapamćen mejl nadležnog: " + _to_default + ". Klikni Pošalji mejl da pošalješ sa prilogom.")

        if _zakljucan:
            st.caption("Izveštaj je predat/zaključan — prijave su zabeležene, izmene nisu moguće.")
        return

    n_obj = len(objekti)
    n_red = sum(1 for o in objekti if o["nivo"] == "crveno")
    n_org = sum(1 for o in objekti if o["nivo"] == "zuto")
    n_grn = sum(1 for o in objekti if o["nivo"] == "zeleno")
    n_done = len([o for o in objekti if o["idk"] in reviewed])
    _pct = int(n_done / max(n_obj, 1) * 100)

    _snap_z = meta.get("start_zone") if isinstance(meta, dict) else None

    def _zavrsen(_o, _v):
        """Da li je objekat gotov — poručio je, pa ga ne treba više zvati ni slati mu mejl."""
        _v = _v or {}
        if (_v.get("trebovali_tip") or "") in ("nas", "njihov"):
            return True
        if "Ubačena porudžbina" in (_v.get("reakcije") or []):
            return True
        if (_v.get("dnevnik") or {}).get("trebovao_posle_starta"):
            return True
        try:
            _hz = st.session_state.get("hist_" + str(sistem) + "_" + str(_o["idk"]))
            if _hz and not _hz.get("err") and _porucio_posle_starta(_o["idk"], _hz.get("lst") or [], _snap_z):
                return True
        except Exception:
            pass
        try:
            _raw = hitnost_objekta(_o["lst"])[0]
        except Exception:
            _raw = _o.get("nivo")
        return _raw in ("crveno", "zuto") and _o.get("nivo") == "zeleno"

    n_zav = 0
    for _ok2 in objekti:
        try:
            if _zavrsen(_ok2, obrada_map.get(int(_ok2["idk"])) or {}):
                n_zav += 1
        except Exception:
            pass

    st.markdown(
        '<div class="adm-kpi">'
        '<div class="cell c-p"><div class="n"><span class="d p"></span>' + str(n_obj) + '</div><div class="k">Objekata za porudžbinu</div></div>'
        '<div class="cell c-r"><div class="n"><span class="d r"></span>' + str(n_red) + '</div><div class="k">Hitno pozvati</div></div>'
        '<div class="cell c-o"><div class="n"><span class="d o"></span>' + str(n_org) + '</div><div class="k">Iskontrolisati</div></div>'
        '<div class="cell c-g"><div class="n"><span class="d g"></span>' + str(n_grn) + '</div><div class="k">Dobra</div></div>'
        '<div class="cell c-z"><div class="n"><span class="d z"></span>' + str(n_zav) + '</div><div class="k">✓ Završeno</div></div>'
        '</div>', unsafe_allow_html=True)
    st.markdown('<div class="adm-prog"><span class="t">Pregledano ' + str(n_done) + ' / ' + str(n_obj) + '</span>'
                '<div class="bar"><div style="width:' + str(_pct) + '%"></div></div></div>', unsafe_allow_html=True)

    # --- Predaja / zaključavanje izveštaja ---
    if _zakljucan:
        _pat = meta.get("predato_at")
        if _predato and _pat:
            _kada = " · predato " + str(_pat)
        elif _rok_prosao:
            _kada = " · rok istekao (" + str(_rok) + ")"
        else:
            _kada = ""
        st.markdown('<div style="background:#eef2ff;border:1px solid #c7d2fe;border-radius:10px;'
                    'padding:11px 14px;margin:4px 0 14px;color:#3730a3;font-size:13.5px;font-weight:600;">'
                    '🔒 Izveštaj za ' + _sel_lbl + ' · ' + sistem + ' je zaključan' + _kada +
                    '. Samo pregled — izmene, ažuriranje i slanje nisu mogući za ovaj mesec.</div>',
                    unsafe_allow_html=True)
        st.session_state.pop("_req_predaj", None)
    elif st.session_state.get("_req_predaj"):
        st.warning("Predajom zaključavaš " + _sel_lbl + " · " + sistem
                   + " — posle toga nema izmena, ažuriranja ni slanja u admin za taj mesec. Sigurno?")
        _pcf1, _pcf2, _pcf3 = st.columns([1, 1, 3])
        with _pcf1:
            if st.button("✅ Potvrdi predaju", key="predaj_ok", type="primary", use_container_width=True):
                try:
                    sb_predaj(mesec_key, sistem)
                    st.session_state.pop("_req_predaj", None)
                    st.rerun()
                except Exception as _e:
                    st.error("Greška pri predaji: " + str(_e))
        with _pcf2:
            if st.button("Otkaži", key="predaj_cancel", use_container_width=True):
                st.session_state.pop("_req_predaj", None)
                st.rerun()

    # --- Izvrši osvežavanje iz admina (traženo dugmetom "Ažuriraj" u zaglavlju) ---
    _bez_naziva = [o["idk"] for o in objekti if not ((komfull.get(int(o["idk"]), {}) or {}).get("naziv"))]
    if st.session_state.pop("_req_refresh_admin", False) and not _zakljucan:
        try:
            # Povuci poslednjih ~6 meseci (od 6 meseci pre meseca izveštaja, 1. u mesecu)
            _yy0 = int(str(mesec_key).split("-")[0]); _mm0 = int(str(mesec_key).split("-")[1])
            _mm6 = _mm0 - 6; _yy6 = _yy0
            while _mm6 <= 0:
                _mm6 += 12; _yy6 -= 1
            _cut_all = datetime.date(_yy6, _mm6, 1)
        except Exception:
            _cut_all = None
        # --- Inkrementalno: već zapamćeno se NE povlači ponovo ---
        # Za objekte koji su već jednom povučeni tražimo samo razliku (od poslednjeg
        # ažuriranja unazad 5 dana radi sigurnosti), a punih ~6 meseci samo za objekte
        # koji do sada nisu obuhvaćeni.
        _cov_r = set(str(x) for x in (meta.get("admin_hist_idk") or []))
        _cov_r |= set(str(x) for x in ((meta.get("admin_hist") or {}).keys()))
        _last_at_r = meta.get("admin_hist_at") if isinstance(meta, dict) else None
        _cut_inc = None
        if _last_at_r and _cov_r:
            try:
                _d_last = datetime.date.fromisoformat(str(_last_at_r)[:10])
                _cut_inc = _d_last - datetime.timedelta(days=5)
                if _cut_all and _cut_inc < _cut_all:
                    _cut_inc = _cut_all
            except Exception:
                _cut_inc = None
        _spin_txt = ("Dopunjavam iz admina samo NOVE porudžbine (od "
                     + _cut_inc.strftime("%d.%m.%Y") + ") — zapamćeno se ne povlači ponovo..."
                     ) if _cut_inc else ("Povlačim sve iz admina za ceo sistem (poslednjih ~6 meseci: "
                                         "nazivi + prethodne porudžbine + dopuna)... može potrajati par minuta.")
        with st.spinner(_spin_txt):
            if _bez_naziva:
                admin_build_komitenti(only_ids=_bez_naziva)
            _kf = sb_komitenti_full()
            st.session_state["_komfull"] = _kf
            _idk_naziv = {o["idk"]: (_kf.get(int(o["idk"]), {}) or {}).get("naziv", "") for o in objekti}
            if _cut_inc:
                _map_inc = {k: v for k, v in _idk_naziv.items() if str(int(k)) in _cov_r}
                _map_new = {k: v for k, v in _idk_naziv.items() if str(int(k)) not in _cov_r}
                _bulk, _be_all = ({}, "")
                if _map_inc:
                    _bulk, _be_all = admin_istorija_bulk(_map_inc, _cut_inc)
                if _map_new:
                    _b2, _e2 = admin_istorija_bulk(_map_new, _cut_all)
                    for _k2, _v2 in (_b2 or {}).items():
                        _bulk[_k2] = (_bulk.get(_k2) or []) + (_v2 or [])
                    if _e2 and not _b2:
                        _be_all = _be_all or _e2
            else:
                _bulk, _be_all = admin_istorija_bulk(_idk_naziv, _cut_all)
        if _be_all and not _bulk:
            st.error(_be_all)
        else:
            # Spoji sa već zapamćenim (dedupe po ID porudžbine) — novo se DODAJE, staro se ne gubi
            _prev_hist = (meta.get("admin_hist") or {}) if isinstance(meta, dict) else {}
            _n_hist = 0
            _hist_save = {}
            for o in objekti:
                _old_lst = _prev_hist.get(str(int(o["idk"])), []) or []
                _new_lst = _bulk.get(o["idk"], []) or []
                _by_id = {}
                for _od in (_old_lst + _new_lst):
                    _oid = str(_od.get("id") or "")
                    if not _oid:
                        continue
                    if _oid not in _by_id:
                        _by_id[_oid] = _od
                    elif (not (_by_id[_oid].get("stavke"))) and _od.get("stavke"):
                        _by_id[_oid] = _od  # zadrži verziju sa stavkama
                _lst = sorted(_by_id.values(), key=_datum_sort_key, reverse=True)
                st.session_state["hist_" + str(sistem) + "_" + str(o["idk"])] = {"lst": _lst, "err": ""}
                if _lst:
                    _n_hist += 1
                    _hist_save[int(o["idk"])] = _lst
            _kada_now = _now().isoformat()
            try:
                sb_admin_hist_set(mesec_key, sistem, _hist_save, _kada_now,
                                  svi_idk=[int(o["idk"]) for o in objekti])
            except Exception as _es:
                st.warning("Podaci su ažurirani, ali nisu trajno sačuvani (osvežiće se ponovo pri sledećem ažuriranju): " + str(_es))
            # Ko je poručio POSLE starta — zabeleži trajno u bazu, da status ostane
            # i kad se izađe pa ponovo uđe bez novog ažuriranja.
            _snap_r = meta.get("start_zone") if isinstance(meta, dict) else None
            _n_zav = 0
            _n_nas = _n_njih = 0
            for o in objekti:
                _lz = (_hist_save.get(int(o["idk"])) or [])
                try:
                    if _porucio_posle_starta(o["idk"], _lz, _snap_r):
                        # Ako smo MI ubacili porudžbinu kroz aplikaciju — samo zabeleži
                        # da je završeno, ali NE određuj način automatski (bira se ručno).
                        _vo = obrada_map.get(int(o["idk"])) or {}
                        if (("Ubačena porudžbina" in (_vo.get("reakcije") or []))
                                or (_vo.get("dnevnik") or {}).get("ubaceno")):
                            if sb_oznaci_trebovao(mesec_key, sistem, o["idk"],
                                                  kada=_kada_now, n_por=len(_lz)):
                                _n_zav += 1
                            continue
                        # Šta je stvarno poručio posle starta i da li je to PO NAŠEM
                        # ili PO NJIHOVOM sistemu — određuje se automatski, po količinama.
                        _pos_map = _treb_posle_starta(o["idk"], _lz, _snap_r)
                        # Poredi se sa DOPUNOM zabeleženom na startu (to smo tražili
                        # od objekta), a ne sa punom preporukom iz izveštaja.
                        _st_map, _st_ok = _treb_na_startu(o["idk"], _lz, _snap_r, _cut_hit)
                        _nase_kol = {}
                        _nase_nz = {}
                        for _s in o["lst"]:
                            try:
                                _ia0 = int(_s["ida"])
                                _nase_kol[_ia0] = max(int(_s.get("kol", 0) or 0)
                                                      - int(_st_map.get(_ia0, 0) or 0), 0)
                                _nase_nz[_ia0] = str(_s.get("naziv", "") or "")
                            except Exception:
                                pass
                        _atip, _ainfo = _nacin_trebovanja(_nase_kol, _pos_map, _nase_nz)
                        _ainfo["pravilo"] = 2
                        # u tabelu (kolona „Njihova por.") idu samo artikli iz izveštaja
                        _anjih = {str(_ia): int(_pos_map.get(_ia, 0)) for _ia in _nase_kol
                                  if int(_pos_map.get(_ia, 0)) > 0}
                        if sb_oznaci_trebovao(mesec_key, sistem, o["idk"],
                                              kada=_kada_now, n_por=len(_lz),
                                              tip=_atip, njihova=_anjih, info=_ainfo):
                            _n_zav += 1
                        if _atip == "nas":
                            _n_nas += 1
                        else:
                            _n_njih += 1
                        # osveži čekboks/izbor u detaljnom pregledu
                        st.session_state.pop("treb_" + str(o["idk"]), None)
                        st.session_state.pop("ed_" + str(o["idk"]), None)
                except Exception:
                    pass
            st.session_state["_refresh_done"] = {"sis": sistem, "mes": mesec_key,
                                                 "n": _n_hist, "zav": _n_zav,
                                                 "nas": _n_nas, "njih": _n_njih,
                                                 "inc": (_cut_inc.strftime("%d.%m.%Y") if _cut_inc else "")}
            st.rerun()
    _rf = st.session_state.get("_refresh_done")
    if _rf and _rf.get("sis") == sistem and _rf.get("mes") == mesec_key:
        _pcut = _admin_presek(meta, mesec_key)
        _pcs = _pcut.strftime("%d.%m.%Y") if _pcut else "01."
        st.success("✅ Ažurirano iz admina — prethodne porudžbine i dopuna su spremni u svakom objektu. "
                   + (("Dopunjena je samo razlika od " + str(_rf.get("inc")) + " — ranije povučeno je zapamćeno. ")
                      if _rf.get("inc") else "")
                   + "Prikazuje se šta su objekti sami poručili od " + _pcs + " (posle preseka). "
                   "Objekata sa takvim porudžbinama: " + str(_rf.get("n", 0)) + "."
                   + ((" · ✅ " + str(_rf.get("zav", 0)) + " objekata je poručilo posle starta i "
                       "trajno je označeno kao ZAVRŠENO.") if _rf.get("zav") else ""))
        if _rf.get("nas") or _rf.get("njih"):
            st.info("🔎 Način trebovanja je prepoznat automatski, po količinama: **"
                    + str(_rf.get("nas", 0)) + "** po našem sistemu · **"
                    + str(_rf.get("njih", 0)) + "** po njihovom. Po našem je samo kad je SVAKI traženi "
                    "artikal poručen u punoj količini — ako i jedan fali, ide u „po njihovom“. "
                    "Poručene količine su upisane u kolonu „Njihova por.“, a ti objekti su "
                    "zaključani (ne menjaju se ručno).")

    # --- Startni rezultat: povuci porudžbine od preseka (jednom, pa se zaključa) ---
    _pk_s = _admin_presek(meta, mesec_key)
    _pk_lbl = _pk_s.strftime("%d.%m.") if _pk_s else "01."
    if st.session_state.pop("_req_start_snap", False) and not _zakljucan and not (isinstance(meta, dict) and meta.get("start_zone")):
        with st.spinner("Povlačim porudžbine od " + _pk_lbl + " za ceo sistem (beležim startni rezultat)... može potrajati par minuta."):
            if _bez_naziva:
                admin_build_komitenti(only_ids=_bez_naziva)
            _kfs = sb_komitenti_full(); st.session_state["_komfull"] = _kfs
            _idk_naziv_s = {o["idk"]: (_kfs.get(int(o["idk"]), {}) or {}).get("naziv", "") for o in objekti}
            _bulk_s, _be_s = admin_istorija_bulk(_idk_naziv_s, _pk_s)
        if _be_s and not _bulk_s:
            st.error(_be_s)
        else:
            _cr = _zu = _ze = 0; _nh = 0
            _hist_save_s = {}
            _start_ids = {}
            for o in objekti:
                _lst_s = sorted(_bulk_s.get(o["idk"], []), key=_datum_sort_key, reverse=True)
                st.session_state["hist_" + str(sistem) + "_" + str(o["idk"])] = {"lst": _lst_s, "err": ""}
                # zapamti KOJE porudžbine objekat ima u trenutku starta — sve što se
                # kasnije pojavi znači da je objekat poručio i prelazi u „Završeno“
                _start_ids[str(int(o["idk"]))] = [str(_x.get("id") or "") for _x in _lst_s if _x.get("id")]
                if _lst_s:
                    _nh += 1
                    _hist_save_s[int(o["idk"])] = _lst_s
                _pm_s = _treb_posle_preseka(_lst_s, _pk_s) if _pk_s else {}
                _nivo_s, _, _ = hitnost_objekta_dodatna(o["lst"], _pm_s)
                if _nivo_s == "crveno":
                    _cr += 1
                elif _nivo_s == "zuto":
                    _zu += 1
                else:
                    _ze += 1
            try:
                sb_admin_hist_set(mesec_key, sistem, _hist_save_s, _now().isoformat(),
                                  svi_idk=[int(o["idk"]) for o in objekti])
            except Exception:
                pass
            try:
                sb_start_snapshot(mesec_key, sistem, {"n": len(objekti), "crveno": _cr, "zuto": _zu, "zeleno": _ze,
                                                      "kada": _now().strftime("%d.%m.%Y %H:%M"),
                                                      "narudzbine": _start_ids})
                st.session_state["_refresh_done"] = {"sis": sistem, "mes": mesec_key, "n": _nh}
                st.rerun()
            except Exception as _e:
                st.error("Greška pri čuvanju startnog rezultata: " + str(_e))

    _snap = meta.get("start_zone") if isinstance(meta, dict) else None
    if _snap:
        st.markdown('<div style="background:#eef2ff;border:1px solid #c7d2fe;border-radius:10px;padding:9px 14px;'
                    'font-size:12.5px;color:#3730a3;margin:2px 0 12px;">📌 <b>Startni rezultat zabeležen</b> ('
                    + str(_snap.get("kada", "")) + '): ' + str(_snap.get("crveno", 0)) + ' hitno · '
                    + str(_snap.get("zuto", 0)) + ' srednje · ' + str(_snap.get("zeleno", 0))
                    + ' dobro. Ovo je START u izveštaju.</div>', unsafe_allow_html=True)
    elif not _zakljucan:
        _sc1, _sc2 = st.columns([2, 3])
        with _sc1:
            if st.button("📌 Povuci porudžbine od " + _pk_lbl + " (zabeleži start)", key="pull_start", use_container_width=True):
                st.session_state["_req_start_snap"] = True
                st.rerun()
        with _sc2:
            st.caption("Povuci JEDNOM na početku rada — beleži se startno stanje (hitno/srednje/dobro) za izveštaj i zaključava se. Kasnije koristi Ažuriraj iz admina gore.")

    with st.expander("📦 Porudžbina ubačena za ceo sistem (grupna akcija)"):
        st.caption("Za sisteme gde mi direktno ubacujemo porudžbine (npr. BB TRADE, KNEZ) — jednim klikom se svi objekti označe kao Ubačena porudžbina i postaju pregledani. Ne koristiti za sisteme gde se objekti zovu pojedinačno.")
        _bc1, _bc2 = st.columns(2)
        with _bc1:
            if st.button("✅ Označi sve: porudžbina ubačena", key="bulk_ubaci", use_container_width=True, disabled=_zakljucan):
                try:
                    sb_bulk_ubaci(mesec_key, sistem, ids_sorted, ko=st.session_state.get("admin_user", "Administracija"))
                    st.success("Sve označeno kao ubačena porudžbina.")
                    st.rerun()
                except Exception as _e:
                    st.error("Greška: " + str(_e))
        with _bc2:
            if st.button("↩️ Poništi obradu celog sistema", key="bulk_reset", use_container_width=True, disabled=_zakljucan):
                try:
                    sb_bulk_reset(mesec_key, sistem)
                    st.success("Obrada sistema poništena.")
                    st.rerun()
                except Exception as _e:
                    st.error("Greška: " + str(_e))

    def _treb_list_cell(tip, rev):
        if tip == "nas":
            return '<span class="tb-nas">✓ po našem</span>'
        if tip == "njihov":
            return '<span class="tb-nj">po njihovom</span>'
        return '<span class="np">—</span>' if rev else '<span class="np">·</span>'

    _TABS = ["Lista objekata", "Detalj / obrada", "📧 Grupno slanje mejlova"]

    # --- Klik na naziv objekta u listi otvara baš taj objekat u „Detalj / obrada" ---
    # Mora da bude pravo Streamlit dugme (on_click). Link sa ?obj=... ponovo učitava
    # stranicu, pravi NOVU sesiju i izbacuje na prijavu — zato se ne koristi.
    def _otvori_objekat(_oid):
        try:
            _oo = obj_by_id.get(int(_oid))
            if not _oo:
                return
            _oz = _zona_disp(_oo["nivo"])
            _on = (komfull.get(int(_oid), {}) or {}).get("naziv", "")
            st.session_state["adm_pick"] = (_oz[2] + "  " + str(int(_oid))
                                            + (("  \u00b7  " + _on) if _on else "")
                                            + "  \u00b7  " + _oz[3])
            st.session_state["adm_tabs"] = _TABS[1]
        except Exception:
            pass

    try:
        tab_lista, tab_detalj, tab_bulk = st.tabs(_TABS, key="adm_tabs", on_change="rerun")
    except TypeError:
        tab_lista, tab_detalj, tab_bulk = st.tabs(_TABS)

    with tab_lista:
        _cut_list = _admin_presek(meta, mesec_key)
        _je_zavrseno = _zavrsen
        _rows = []
        _export_rows = []
        for o in objekti:
            z = _zona_disp(o["nivo"])
            v = obrada_map.get(o["idk"], {})
            reak = v.get("reakcije", [])
            _rko = v.get("reakcije_ko") or {}
            _dnv = v.get("dnevnik") or {}
            _pz = _dnv.get("pozivi") or []
            _mj = _dnv.get("mejlovi") or []
            _zav = _je_zavrseno(o, v)
            _nazlist = (komfull.get(int(o["idk"]), {}) or {}).get("naziv", "")
            # red za Excel izvoz
            def _koik(_e):
                return (str(_e.get("ko", "")) + " · " + _dt_kratko(_e.get("at", ""))) if _e else ""
            _export_rows.append({
                "idk": int(o["idk"]), "naziv": _nazlist or ("ID " + str(o["idk"])),
                "mejl": len(_mj) > 0, "p1": len(_pz) >= 1, "p2": len(_pz) >= 2, "p3": len(_pz) >= 3,
                "mejl_ko": _koik(_mj[-1] if _mj else None),
                "p1_ko": _koik(_pz[0] if len(_pz) >= 1 else None),
                "p2_ko": _koik(_pz[1] if len(_pz) >= 2 else None),
                "p3_ko": _koik(_pz[2] if len(_pz) >= 3 else None),
                "zavrseno": _zav,
            })
            # Prikaz u listi: samo koliko puta je zvala i da li je poslat mejl.
            # (ko je i kada — ostaje u Excel izvozu i u detaljnoj kartici.)
            if reak:
                _chips = []
                for r in reak:
                    if r == "Pozvala sam":
                        _chips.append("📞 Pozvala" + ((" " + str(len(_pz)) + "×") if _pz else ""))
                    elif r == "Poslala sam mejl":
                        _chips.append("✉️ Mejl")
                    else:
                        _chips.append(_reak_short(r))
                stat = "".join('<span class="stchip">' + _h_escape(c) + '</span>' for c in _chips)
            else:
                stat = '<span class="stat">Nepregledano</span>'
            _rc = "row-red" if o["nivo"] == "crveno" else ("row-org" if o["nivo"] == "zuto" else "")
            _treb_mark = ""
            _hf = st.session_state.get("hist_" + str(sistem) + "_" + str(o["idk"]))
            if _hf and not _hf.get("err"):
                _tt = int(sum(_treb_posle_preseka(_hf.get("lst") or [], _cut_list).values()))
                if _tt > 0:
                    _treb_mark = ' <span title="Trebovano posle 01." style="color:#b45309;font-weight:700;">⚠️ posle 01. (' + str(_tt) + ')</span>'
            _zavcell = ('<span style="background:#dcfce7;color:#14532d;font-weight:700;font-size:11.5px;'
                        'padding:3px 9px;border-radius:20px;">✓ Završeno</span>') if _zav else '<span class="np">—</span>'
            _rows.append({
                "idk": int(o["idk"]),
                "naziv": (_nazlist or "— naknadno"),
                "ima_naziv": bool(_nazlist),
                "mark_txt": (("⚠️ posle 01. (" + str(_tt) + ")") if _treb_mark else ""),
                "nula": int(o["na_nuli"]), "izgub": int(o["izgub"]),
                "zona": ('<span class="zona ' + z[0] + '"><span class="zd"></span>' + z[3] + '</span>'),
                "stat": stat,
                "treb": _treb_list_cell(v.get("trebovali_tip", ""), o["idk"] in reviewed),
                "zav": _zavcell,
                "rc": ("red" if o["nivo"] == "crveno" else ("org" if o["nivo"] == "zuto" else "")),
            })
        # Izvoz u Excel (iznad liste)
        _lc1, _lc2 = st.columns([1.4, 3])
        with _lc1:
            try:
                _zx = _zadaci_xlsx(_export_rows, sistem, _sel_lbl)
                _safe_s = "".join(ch for ch in str(sistem or "") if ch.isalnum() or ch in " _-").strip().replace(" ", "_")
                st.download_button("⬇️ Izvezi listu u Excel", _zx,
                    file_name="Zadaci_" + (_safe_s or "sistem") + "_" + str(mesec_key) + ".xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="zadaci_xlsx", use_container_width=True)
            except Exception as _ze:
                st.caption("Izvoz trenutno nije moguć: " + str(_ze))
        with _lc2:
            _nz = sum(1 for r in _export_rows if r["zavrseno"])
            st.caption("Excel: mejl + 1./2./3. poziv (Da/Ne), ko i kada, i Završeno.  ·  Završeno: "
                       + str(_nz) + " / " + str(len(_export_rows)) + " objekata.")
        # --- Lista: naziv objekta je dugme (klik otvara Detalj za taj objekat) ---
        # Ostale kolone idu kao JEDAN HTML blok po redu (fiksne \u0161irine), da lista
        # ostane brza \u2014 ina\u010de bi 238 objekata zna\u010dilo ~1900 elemenata po prikazu.
        _W = [0.62, 4.3, 8.1]
        _CW = [78, 96, 150, 172, 108, 112]      # Na nuli, Izgubljeno, Zona, Status, Trebovali, Zavr\u0161eno
        _HDR = ["Na nuli", "Izgubljeno", "Zona", "Status", "Trebovali", "Zavr\u0161eno"]

        def _lbl_md(_s):
            """Naziv ide kao labela dugmeta (markdown) — neutrališi * _ [ ] ` znakove."""
            _o = str(_s or "")
            for _ch in ("\\", "*", "_", "[", "]", "`", "~"):
                _o = _o.replace(_ch, "\\" + _ch)
            return _o

        def _celije(_vals, _cls="lst-c"):
            _o = '<div class="lst-row">'
            for _i2, _vv in enumerate(_vals):
                _ta = "center" if _i2 in (0, 1, 4, 5) else "left"
                _o += ('<div class="' + _cls + '" style="flex:0 0 ' + str(_CW[_i2])
                       + 'px;text-align:' + _ta + ';">' + _vv + '</div>')
            return _o + '</div>'

        _hc = st.columns(_W)
        _hc[0].markdown('<div class="lst-h">ID</div>', unsafe_allow_html=True)
        _hc[1].markdown('<div class="lst-h">Naziv komitenta</div>', unsafe_allow_html=True)
        _hc[2].markdown(_celije(_HDR, "lst-h"), unsafe_allow_html=True)
        for _r in _rows:
            try:
                _cc = st.columns(_W, vertical_alignment="center")
            except TypeError:
                _cc = st.columns(_W)
            _cc[0].markdown('<span class="rtag ' + _r["rc"] + '"></span>'
                            '<div class="lst-c lst-id">' + str(_r["idk"]) + '</div>',
                            unsafe_allow_html=True)
            with _cc[1]:
                try:
                    st.button(_lbl_md(_r["naziv"]) + (("   " + _r["mark_txt"]) if _r["mark_txt"] else ""),
                              key="opn_" + str(_r["idk"]), type="tertiary",
                              on_click=_otvori_objekat, args=(_r["idk"],),
                              help="Otvori detalje ovog objekta")
                except TypeError:
                    st.button(_lbl_md(_r["naziv"]) + (("   " + _r["mark_txt"]) if _r["mark_txt"] else ""),
                              key="opn_" + str(_r["idk"]),
                              on_click=_otvori_objekat, args=(_r["idk"],))
            _cc[2].markdown(_celije([
                '<span style="color:#d33;">' + str(_r["nula"]) + '</span>',
                str(_r["izgub"]), _r["zona"], _r["stat"], _r["treb"], _r["zav"]]),
                unsafe_allow_html=True)
        st.caption("\U0001F446 Klikni na naziv objekta da ti otvori ba\u0161 taj objekat u Detalj / obrada.  \u00b7  "
                   "Status i trebovanje se menjaju u kartici Detalj / obrada.  \u00b7  "
                   "Zavr\u0161eno = objekat je poru\u010dio: stigla je nova porud\u017ebina posle starta, "
                   "ili je pre\u0161ao iz crvene/narand\u017easte u zelenu, ili je trebovanje ru\u010dno zabele\u017eeno.")

    with tab_detalj:
        _labels = []
        for o in objekti:
            _zz = _zona_disp(o["nivo"])
            _nz = (komfull.get(int(o["idk"]), {}) or {}).get("naziv", "")
            _labels.append(_zz[2] + "  " + str(o["idk"]) + (("  ·  " + _nz) if _nz else "") + "  ·  " + _zz[3])
        _lab2id = {_labels[i]: ids_sorted[i] for i in range(len(objekti))}
        _id2lab = {ids_sorted[i]: _labels[i] for i in range(len(objekti))}
        if ("adm_pick" not in st.session_state) or (st.session_state.adm_pick not in _labels):
            st.session_state.adm_pick = _labels[0]

        def _next_unrev():
            _cur = _lab2id.get(st.session_state.adm_pick, ids_sorted[0])
            _idx = ids_sorted.index(_cur) if _cur in ids_sorted else -1
            for _k in range(1, len(ids_sorted) + 1):
                _cand = ids_sorted[(_idx + _k) % len(ids_sorted)]
                if _cand not in reviewed:
                    st.session_state.adm_pick = _id2lab[_cand]
                    return

        _nc1, _nc2 = st.columns([3, 1])
        with _nc1:
            st.selectbox("Izaberi / pretraži objekat (ukucaj ID ili naziv)", _labels, key="adm_pick")
        with _nc2:
            st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
            st.button("Sledeći nepregledan →", on_click=_next_unrev, use_container_width=True, key="adm_next")

        sel_id = _lab2id[st.session_state.adm_pick]
        o = obj_by_id[sel_id]
        z = _zona_disp(o["nivo"])
        v = obrada_map.get(sel_id, {"reakcije": [], "trebovali_tip": ""})
        TREB_OPT = ["— nije trebovano", "Po našem sistemu", "Po njihovom sistemu (ne po našem)"]
        TREB_CODE = {"— nije trebovano": "", "Po našem sistemu": "nas", "Po njihovom sistemu (ne po našem)": "njihov"}
        _key_treb = "treb_" + str(sel_id)
        _loaded_tip = v.get("trebovali_tip", "") or ""
        if _key_treb in st.session_state:
            _tip_now = TREB_CODE.get(st.session_state[_key_treb], "")
        else:
            _tip_now = _loaded_tip
        _auto_treb = (v.get("dnevnik") or {}).get("trebovao_posle_starta") or {}
        # Objekat koji je posle starta poslao porudžbinu je ZAVRŠEN: način trebovanja
        # je prepoznat automatski (po količinama) i ne menja se ručno — sve je zaključano.
        # Ako smo MI ubacili porudžbinu kroz aplikaciju — način trebovanja se i dalje
        # bira ručno. Zaključava se samo ono što je objekat sam uneo preko admina.
        _nas_unos = (("Ubačena porudžbina" in (v.get("reakcije") or []))
                     or bool((v.get("dnevnik") or {}).get("ubaceno")))
        _auto_lock = bool(_auto_treb) and not _nas_unos
        if _auto_lock:
            _tip_now = _loaded_tip or str(_auto_treb.get("tip") or "nas")
        _njihov_active = (_tip_now == "njihov")
        if _auto_lock:
            # widget uvek prikazuje ono što je automatski prepoznato (ne staru ručnu vrednost)
            st.session_state[_key_treb] = TREB_OPT[["", "nas", "njihov"].index(_tip_now)
                                                   if _tip_now in ["", "nas", "njihov"] else 0]
        _revb = '<span class="revy">✓ Pregledano</span>' if sel_id in reviewed else '<span class="revn">Nepregledano</span>'
        _zav_ovaj = False
        try:
            _o_sel = next((o for o in objekti if int(o["idk"]) == int(sel_id)), None)
            _zav_ovaj = bool(_o_sel is not None and _zavrsen(_o_sel, v))
        except Exception:
            _zav_ovaj = False
        if _zav_ovaj:
            _revb = ('<span style="background:#dcfce7;color:#166534;font-weight:800;font-size:12px;'
                     'padding:4px 12px;border-radius:20px;">✓ ZAVRŠENO</span>&nbsp;&nbsp;') + _revb
        _kinfo = komfull.get(int(sel_id), {}) or {}
        _knaziv = _kinfo.get("naziv", "")
        _naz_html = ('<span style="font-weight:600;color:#2a2f3a;">' + _h_escape(_knaziv) + '</span>') if _knaziv else '<span class="mut">— naziv naknadno</span>'
        st.markdown('<div class="adm-dh"><span class="id">' + str(sel_id) + '</span>'
                    '<span class="zona ' + z[0] + '"><span class="zd"></span>' + z[3] + '</span>'
                    + _naz_html +
                    '<span style="margin-left:auto;">' + _revb + '</span></div>', unsafe_allow_html=True)
        # --- VELIKA oznaka: objekat je završen i zaključan (ništa se ne dira) ---
        if _zav_ovaj:
            _ai = dict(_auto_treb.get("info") or {})
            _tip_txt = (("PO NAŠEM SISTEMU" if _tip_now == "nas" else "PO NJIHOVOM SISTEMU")
                        if _tip_now in ("nas", "njihov") else "")
            _kol_txt = ""
            if _ai.get("poruceno") or _ai.get("trazeno"):
                _kol_txt = ('Poručeno ' + str(_ai.get("poruceno", 0)) + ' kom od traženih '
                            + str(_ai.get("trazeno", 0)) + ' kom dopune (' + str(_ai.get("procenat", 0)) + '%'
                            + ((' · ' + str(_ai.get("van_liste", 0)) + ' kom van naše liste')
                               if _ai.get("van_liste") else '') + ').')
            # Zašto je „po njihovom" — koji artikli nisu poručeni kako smo tražili
            _fali_txt = ""
            if _auto_lock and _tip_now == "njihov":
                _fl = list(_ai.get("fali") or [])
                if _fl:
                    _fb = []
                    for _f in _fl[:3]:
                        _fb.append((str(_f.get("naziv", "")) or ("ID " + str(_f.get("ida", ""))))
                                   + " (traženo " + str(_f.get("trazeno", 0))
                                   + ", poručeno " + str(_f.get("poruceno", 0)) + ")")
                    _fali_txt = ("Nije po našem jer " + str(_ai.get("fali_n", len(_fl)))
                                 + " artikal/artikala nije poručeno kako smo tražili: "
                                 + "; ".join(_fb) + ("…" if _ai.get("fali_n", 0) > 3 else ""))
                elif _ai.get("van_liste"):
                    _fali_txt = ("Nije po našem jer je " + str(_ai.get("van_liste", 0))
                                 + " kom poručeno van naše liste artikala.")
            _kada_txt = _dt_kratko(_auto_treb.get("at", "")) if _auto_treb.get("at") else ""
            st.markdown(
                '<div style="background:#dcfce7;border:2px solid #16a34a;border-radius:12px;'
                'padding:14px 18px;margin:6px 0 14px;">'
                '<div style="font-size:19px;font-weight:900;color:#14532d;letter-spacing:.3px;">'
                '✅ ZAVRŠENO — OBJEKAT JE PORUČIO</div>'
                + (('<div style="font-size:14px;font-weight:800;color:#166534;margin-top:4px;">Trebovano '
                    + _tip_txt + '</div>') if _tip_txt else '')
                + (('<div style="font-size:12.5px;color:#166534;margin-top:3px;">' + _h_escape(_kol_txt)
                    + '</div>') if _kol_txt else '')
                + (('<div style="font-size:12px;color:#7c2d12;background:#fff7ed;border:1px solid #fed7aa;'
                    'border-radius:7px;padding:6px 9px;margin-top:6px;">⚠️ ' + _h_escape(_fali_txt)
                    + '</div>') if _fali_txt else '')
                + ('<div style="font-size:12px;color:#15803d;margin-top:6px;font-style:italic;">'
                   '🔒 Prepoznato automatski iz admina'
                   + ((' · ' + _h_escape(_kada_txt)) if _kada_txt else '')
                   + ' — količine su upisane u kolonu „Njihova por.“. Zaključano: ništa se ne menja ručno.'
                   '</div>' if _auto_lock else
                   '<div style="font-size:12px;color:#15803d;margin-top:6px;font-style:italic;">'
                   + ('📦 Porudžbinu smo ubacili mi, kroz aplikaciju — način trebovanja biraš ručno.'
                      if _nas_unos
                      else ('Ručno označeno trebovanje.' if _tip_now in ("nas", "njihov")
                            else 'Objekat je prešao iz crvene u zelenu zonu.'))
                   + '</div>')
                + '</div>', unsafe_allow_html=True)
        _tel_raw = str(_kinfo.get("telefon", "") or "").strip()
        # očisti broj za tel: link (zadrži cifre i vodeći +)
        import re as _reph
        _tel_clean = _reph.sub(r"[^\d+]", "", _tel_raw)
        if _tel_clean.count("+") > 1:
            _tel_clean = "+" + _tel_clean.replace("+", "")
        # Prevedi u međunarodni format (+381…) da Phone Link pouzdano okrene
        if _tel_clean.startswith("+"):
            pass
        elif _tel_clean.startswith("00"):
            _tel_clean = "+" + _tel_clean[2:]
        elif _tel_clean.startswith("0"):
            _tel_clean = "+381" + _tel_clean[1:]
        _kbits = []
        if _kinfo.get("mesto"):
            _kbits.append("📍 " + _h_escape(_kinfo["mesto"]))
        if _tel_raw:
            if _tel_clean:
                _kbits.append('<a href="tel:' + _h_escape(_tel_clean) + '" style="color:#7c3aed;text-decoration:none;font-weight:600;">📞 '
                              + _h_escape(_tel_raw) + '</a>')
            else:
                _kbits.append("📞 " + _h_escape(_tel_raw))
        if _kinfo.get("email"):
            _kbits.append('<a href="mailto:' + _h_escape(_kinfo["email"]) + '" style="color:#6b7280;text-decoration:none;">✉️ '
                          + _h_escape(_kinfo["email"]) + '</a>')
        if _kbits:
            st.markdown('<div style="margin:-8px 0 10px;color:#6b7280;font-size:13px;">'
                        + "&nbsp;&nbsp;·&nbsp;&nbsp;".join(_kbits) + '</div>', unsafe_allow_html=True)
        # Dugme za poziv: beleži poziv (ko + vreme) pa pokreće pozivanje (Phone Link)
        if _tel_clean:
            _cbtn1, _cbtn2 = st.columns([1.6, 3])
            with _cbtn1:
                if st.button("📞 Pozovi " + _tel_raw, key="callbtn_" + str(sel_id), use_container_width=True):
                    try:
                        sb_obrada_log(mesec_key, sistem, sel_id, "poziv",
                                      st.session_state.get("admin_user", "Administracija"))
                    except Exception:
                        pass
                    st.session_state["_dial_" + str(sel_id)] = _tel_clean
                    for _kk in (1, 2, 3):  # osveži čekboks „Pozvala sam N. put"
                        st.session_state.pop("callchk_" + str(sel_id) + "_" + str(_kk), None)
                    st.rerun()
            # Posle klika: pokreni pozivanje (Phone Link) + rezervni link ako se ne otvori samo
            if st.session_state.get("_dial_" + str(sel_id)):
                _dn = st.session_state.pop("_dial_" + str(sel_id))
                components.html("<script>window.location.href='tel:" + _dn + "';</script>", height=0)
                with _cbtn2:
                    st.markdown('<a href="tel:' + _h_escape(_dn) + '" style="display:inline-block;background:#16a34a;'
                                'color:#fff;text-decoration:none;font-weight:700;font-size:13px;padding:8px 16px;'
                                'border-radius:9px;">📞 Ako se poziv ne otvori sam — klikni</a>', unsafe_allow_html=True)
            # Dnevnik poziva (ko + kada), sitno sivo
            _pozivi = (v.get("dnevnik") or {}).get("pozivi") or []
            if _pozivi:
                st.markdown('<div style="font-size:12px;color:#6b7280;font-weight:600;margin:2px 0 1px;">📞 Pozvano '
                            + str(len(_pozivi)) + '× :</div>' + _dnevnik_lista_html(v.get("dnevnik") or {}, "poziv"),
                            unsafe_allow_html=True)

        # --- Presek (1. u mesecu posle poslednjeg meseca podataka) + istorija iz admina ---
        import datetime as _dtp
        _cutoff = _admin_presek(meta, mesec_key)
        _hk = "hist_" + str(sistem) + "_" + str(sel_id)
        _naziv_kom = _knaziv

        _dc1, _dc2 = st.columns([1.7, 1])
        with _dc1:
            st.markdown('<div class="adm-lbl">Porudžbina i lager'
                        + (' · 🔒 Njihova por. je upisana automatski iz admina'
                           if _auto_lock else ' · upiši Njihovu por.') + '</div>', unsafe_allow_html=True)
            _arts = sorted(o["lst"], key=lambda x: (int(x["lager"]), -int(x["kol"])))
            _njm = v.get("njihova") or {}

            _hist_cache = st.session_state.get(_hk)
            _treb_loaded = bool(_hist_cache and not _hist_cache.get("err"))
            # Posle starta dopuna je ZAMRZNUTA: računa se po porudžbinama koje su
            # postojale u trenutku klika na Start. Kasnije porudžbine je ne menjaju.
            _fz_ok = False
            _treb_map = {}
            if _treb_loaded:
                _treb_map, _fz_ok = _treb_na_startu(sel_id, _hist_cache.get("lst") or [],
                                                    (meta.get("start_zone") if isinstance(meta, dict) else None),
                                                    _cutoff)
                if not _fz_ok:
                    _treb_map = _treb_posle_preseka(_hist_cache.get("lst") or [], _cutoff)
            _treb_total = int(sum(_treb_map.values()))

            def _sd(lg):
                lg = int(lg)
                return "🔴" if lg == 0 else ("🟡" if lg <= 2 else "🟢")

            _dodatna_map = {}
            _rows_adf = []
            for a in _arts:
                _ida = int(a["ida"]); _lg = int(a["lager"]); _kol = int(a["kol"])
                _por = int(_treb_map.get(_ida, 0))
                _realni = _lg + _por
                _dod = max(_kol - _por, 0)
                _dodatna_map[_ida] = _dod
                _rows_adf.append({
                    " ": _sd(_realni),
                    "Artikal": str(a["naziv"]),
                    "Predikcija": int(round(int(a.get("pred", 0) or 0) * float(_mes_kol or 1.0))),
                    "Lager (izv.)": _lg,
                    "Posle 01.": _por,
                    "Realni lager": _realni,
                    "Naša por.": _kol,
                    "Dodatna por.": _dod,
                    "Njihova por.": int(_njm.get(str(_ida), 0)),
                })
            _adf = pd.DataFrame(_rows_adf)
            if _treb_loaded and _treb_total > 0:
                st.markdown('<div style="background:#fff4e5;border:1px solid #f0b429;border-radius:8px;'
                            'padding:8px 12px;margin:2px 0 10px;color:#8a5a00;font-size:13px;font-weight:600;">'
                            '⚠️ Već trebovano ' + str(_treb_total) + ' kom posle 01. — porudžbina je umanjena '
                            '(zelena kolona Dodatna por.).'
                            + (' <span style="font-weight:500;">Zabeleženo na startu — dopuna se više ne menja.</span>'
                               if _fz_ok else '') + '</div>', unsafe_allow_html=True)
            _colcfg = {
                " ": st.column_config.TextColumn(" ", width="small"),
                "Predikcija": st.column_config.NumberColumn(
                    "Predikcija",
                    help=("Predviđena prodaja za period porudžbine ("
                          + str(_mes_kol).replace(".", ",") + " mes)") if _mes_kol
                         else "Predviđena prodaja"),
                "Lager (izv.)": st.column_config.NumberColumn("Lager (izveštaj)", help="Lager iz izveštaja — presek na 01. u mesecu"),
                "Posle 01.": st.column_config.NumberColumn("Posle 01.", help="Koliko je već trebovano iz admina posle 01. u mesecu"),
                "Realni lager": st.column_config.NumberColumn("Realni lager", help="Lager (izveštaj) + poručeno posle 01."),
                "Naša por.": st.column_config.NumberColumn(_por_lbl, help="Preporučena porudžbina (za zadati broj meseci)"),
                "Dodatna por.": st.column_config.NumberColumn("Dodatna por.", help="Naša por. minus već poručeno posle 01. — ovo se šalje u admin"),
                "Njihova por.": st.column_config.NumberColumn("Njihova por.", help="Koliko su stvarno poručili", min_value=0, step=1),
            }
            if _njihov_active and not _zakljucan and not _auto_lock:
                _order = [" ", "Artikal", "Predikcija", "Lager (izv.)", "Posle 01.",
                          "Realni lager", "Naša por.", "Dodatna por.", "Njihova por."]
                _edited = st.data_editor(_adf[_order], hide_index=True, use_container_width=True,
                    disabled=[c for c in _order if c != "Njihova por."],
                    column_config=_colcfg, key="ed_" + str(sel_id))
                _njihova_new = {}
                for _i, _a in enumerate(_arts):
                    try:
                        _njihova_new[str(int(_a["ida"]))] = int(_edited.iloc[_i]["Njihova por."])
                    except Exception:
                        _njihova_new[str(int(_a["ida"]))] = 0
            else:
                _order = [" ", "Artikal", "Predikcija", "Lager (izv.)", "Posle 01.",
                          "Realni lager", "Naša por.", "Dodatna por."]
                if _njihov_active or _auto_lock:  # zaključano — prikaži i njihovu kolonu (samo pregled)
                    _order.append("Njihova por.")
                _sty = _adf[_order].style.set_properties(subset=["Dodatna por."], **{
                    "background-color": "#dcfce7", "color": "#14532d", "font-weight": "700"})
                if _auto_lock and "Njihova por." in _order:
                    _sty = _sty.set_properties(subset=["Njihova por."], **{
                        "background-color": "#e0e7ff", "color": "#312e81", "font-weight": "700"})
                st.dataframe(_sty, hide_index=True, use_container_width=True, column_config=_colcfg)
                _njihova_new = {str(int(a["ida"])): int(_njm.get(str(int(a["ida"])), 0)) for a in _arts}

            # --- Izvoz za objekat (mejl): kružić + naziv + lager + predikcija + dodatna por. ---
            # U aplikaciji ostaju sve kolone (gore); ovaj Excel je samo za slanje objektu.
            _exp_rows = [{"kruzic": r[" "], "naziv": r["Artikal"], "lager": r["Realni lager"],
                          "predikcija": int(r.get("Predikcija", 0) or 0),
                          "dodatna": r["Dodatna por."]}
                         for r in _rows_adf if int(r.get("Dodatna por.", 0) or 0) > 0]
            # --- Excel za preuzimanje + Prosledi mejl (automatski, sa prilogom) ---
            _exp_xlsx = None
            _fname = str(sel_id) + ".xlsx"
            _ecol1, _ecol2, _ecol3 = st.columns([1.4, 1.4, 2])
            with _ecol1:
                if _exp_rows:
                    try:
                        _exp_xlsx = _objekat_order_xlsx(_naziv_kom, sel_id, _sel_lbl, _exp_rows, meseci=_mes_kol)
                        _safe_sis = "".join(ch for ch in str(sistem or "")
                                            if ch.isalnum() or ch in " _-").strip()
                        import re as _refn
                        _mpm = _refn.search(r'MP\s*\d+', str(_naziv_kom or ""), _refn.IGNORECASE)
                        _mp = _mpm.group(0).upper().replace(" ", "") if _mpm else str(sel_id)
                        _fname = ((_safe_sis + " ") if _safe_sis else "") + _mp + ".xlsx"
                        st.download_button("⬇️ Excel za objekat", _exp_xlsx,
                            file_name=_fname,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="expobj_" + str(sel_id), use_container_width=True)
                    except Exception as _ee:
                        st.caption("Izvoz trenutno nije moguć: " + str(_ee))
                else:
                    st.button("⬇️ Excel za objekat", key="expobj_dis_" + str(sel_id),
                              use_container_width=True, disabled=True)
            with _ecol2:
                _mail_to = (_kinfo.get("email") or "").strip()
                _sk_mail = "mailsent_" + str(sistem) + "_" + str(sel_id)
                _can_mail = (bool(_exp_xlsx) and _mejl_ok(_mail_to) and smtp_dostupan()
                             and not _zakljucan and not _auto_lock)
                _vec_poslat = "Poslala sam mejl" in (v.get("reakcije") or [])
                _ck_mail = "_mailconf_" + str(sistem) + "_" + str(sel_id)
                _dk_mail = "_dosend_" + str(sistem) + "_" + str(sel_id)

                def _posalji_mejl_objektu():
                    try:
                        _subj = ("VAPE SHOP - " + str(_naziv_kom or ("ID " + str(sel_id)))
                                 + " - PORUDŽBINA - " + _now().strftime("%d.%m.%Y."))
                        _fname_mail = _fname if _exp_rows else (str(sel_id) + ".xlsx")
                        with st.spinner("✉️ Slanje mejla u toku… (može da potraje do 2 minuta, "
                                        "ne zatvaraj stranu)"):
                            posalji_mejl_sa_prilogom(_mail_to, _subj,
                                                     _mejl_tekst(_naziv_kom), _exp_xlsx,
                                                     _fname_mail)
                        _kk1 = st.session_state.get("_zadnja_kopija")
                        st.session_state[_sk_mail] = {"ok": True, "msg": "Poslato na " + _mail_to,
                                                      "kopija": _kk1}
                        # Auto: zabeleži mejl u dnevnik (ko + vreme) + upali „Poslala sam mejl"
                        try:
                            _cur_u = st.session_state.get("admin_user", "Administracija")
                            sb_obrada_log(mesec_key, sistem, sel_id, "mejl", _cur_u, kopija=_kk1)
                            st.session_state.pop("r2_" + str(sel_id), None)  # da se čekboks osveži na True
                        except Exception:
                            pass
                    except Exception as _em:
                        st.session_state[_sk_mail] = {"ok": False, "msg": str(_em)}

                if st.button("📧 Prosledi mejl", key="sendmail_" + str(sel_id),
                             use_container_width=True, disabled=not _can_mail):
                    if _vec_poslat:
                        st.session_state[_ck_mail] = True      # traži potvrdu (već je poslat)
                    else:
                        st.session_state[_dk_mail] = True
                    st.rerun()

                if st.session_state.pop(_dk_mail, False):
                    _posalji_mejl_objektu()
                    st.rerun()

                # Upozorenje preko celog ekrana kad se mejl šalje PONOVO
                if st.session_state.get(_ck_mail):
                    _mejlovi_p = (v.get("dnevnik") or {}).get("mejlovi") or []
                    _last_m = _mejlovi_p[-1] if _mejlovi_p else None
                    _kada_m = ((" (" + _ko_kratko(_last_m.get("ko", "")) + " · "
                                + _dt_kratko(_last_m.get("at", "")) + ")") if _last_m else "")

                    def _confirm_body():
                        st.warning("⚠️ Ovom objektu je mejl VEĆ poslat" + _kada_m + ".")
                        st.markdown("**Da li ste sigurni da želite da pošaljete ponovo?**")
                        st.caption(str(_naziv_kom or ("ID " + str(sel_id))) + " · " + str(_mail_to))
                        _q1, _q2 = st.columns(2)
                        with _q1:
                            if st.button("Da, pošalji ponovo", type="primary", use_container_width=True,
                                         key="mconf_yes_" + str(sel_id)):
                                st.session_state.pop(_ck_mail, None)
                                st.session_state[_dk_mail] = True
                                st.rerun()
                        with _q2:
                            if st.button("Otkaži", use_container_width=True,
                                         key="mconf_no_" + str(sel_id)):
                                st.session_state.pop(_ck_mail, None)
                                st.rerun()

                    _dlg = getattr(st, "dialog", None) or getattr(st, "experimental_dialog", None)
                    if _dlg:
                        try:
                            _dlg("Mejl je već poslat")(_confirm_body)()
                        except Exception:
                            _confirm_body()
                    else:
                        _confirm_body()
                if _auto_lock:
                    st.caption("🔒 Objekat je završen (poručio je) — slanje mejla je zaključano.")
                elif not smtp_dostupan():
                    _nn = _mail_nalog()
                    st.caption("✉️ Slanje mejla nije podešeno za tvoj nalog (Secrets: SMTP_USER"
                               + (("_" + _nn) if _nn else "") + " / SMTP_PASSWORD"
                               + (("_" + _nn) if _nn else "") + ").")
                elif not _mail_to:
                    st.caption("Objekat nema email u šifarniku komitenata.")
                elif not _mejl_ok(_mail_to):
                    st.error("✉️ Email u šifarniku nije ispravan: „" + str(_mail_to) + "“. "
                             "Najčešće je razmak pre @ ili pogrešan oblik — ispravi u šifarniku.")
                elif not _exp_rows:
                    st.caption("Nema dodatne porudžbine za slanje.")
                # Trajna zelena oznaka (iz baze) + sitni sivi dnevnik mejlova (ko + kada)
                _mejlovi = (v.get("dnevnik") or {}).get("mejlovi") or []
                # Neuspela slanja — ostaju zapisana u bazi i posle „izgubljene“ strane
                _greske = (v.get("dnevnik") or {}).get("greske") or []
                if _greske:
                    _g = _greske[-1]
                    st.error("❌ Poslednji pokušaj slanja NIJE uspeo — "
                             + _dt_kratko(_g.get("at", "")) + " · " + str(_g.get("sta", ""))[:200])
                    if len(_greske) > 1:
                        with st.expander("Raniji neuspeli pokušaji (" + str(len(_greske) - 1) + ")"):
                            for _g2 in reversed(_greske[:-1]):
                                st.caption(_dt_kratko(_g2.get("at", "")) + " · "
                                           + (_ko_kratko(_g2.get("ko", "")) or "?") + " · "
                                           + str(_g2.get("sta", ""))[:200])
                if "Poslala sam mejl" in (v.get("reakcije") or []):
                    st.success("✅ Mejl je poslat objektu" + ((" · " + _mail_to) if _mail_to else "") + ".")
                    _mhtml = _dnevnik_lista_html(v.get("dnevnik") or {}, "mejl")
                    if _mhtml:
                        st.markdown('<div style="margin:-4px 0 4px;">' + _mhtml + '</div>', unsafe_allow_html=True)
                    # --- čišćenje testnih zapisa (uključuje se u Secrets: DOZVOLI_CISCENJE = true) ---
                    _moze_cistiti = (st.session_state.get("role") == "analitika"
                                     or str(_cfg("DOZVOLI_CISCENJE", "")).strip().lower()
                                     in ("1", "true", "da", "on"))
                    if _moze_cistiti:
                        _ck_key = "_cist_pot_" + str(sistem) + "_" + str(sel_id)
                        if not st.session_state.get(_ck_key):
                            if st.button("🧹 Obriši istoriju mejlova (test)",
                                         key="cist_" + str(sistem) + "_" + str(sel_id)):
                                st.session_state[_ck_key] = True
                                st.rerun()
                        else:
                            st.warning("Obrisaće se **" + str(len(_mejlovi)) + "** zapisa o mejlovima "
                                       "za ovaj objekat i oznaka „Poslala sam mejl“. Pozivi i "
                                       "napomene ostaju. Ovo se ne može vratiti.")
                            _cc1, _cc2 = st.columns(2)
                            with _cc1:
                                if st.button("Da, obriši", key="cist_da_" + str(sistem) + "_" + str(sel_id),
                                             type="primary", use_container_width=True):
                                    try:
                                        _n_ob = sb_obrada_ocisti(mesec_key, sistem, sel_id, "mejl")
                                        sb_pregled.clear()
                                        st.session_state.pop(_ck_key, None)
                                        st.session_state.pop(_sk_mail, None)
                                        st.session_state.pop("r2_" + str(sel_id), None)
                                        st.success("Obrisano " + str(_n_ob) + " zapisa.")
                                        st.rerun()
                                    except Exception as _ce:
                                        st.error("Brisanje nije uspelo: " + str(_ce))
                            with _cc2:
                                if st.button("Odustani", key="cist_ne_" + str(sistem) + "_" + str(sel_id),
                                             use_container_width=True):
                                    st.session_state.pop(_ck_key, None)
                                    st.rerun()
                _mres = st.session_state.get(_sk_mail)
                # zapamćena odbijenica iz baze (ostaje i kad se poruka obriše iz sandučeta)
                _vr1 = (v.get("dnevnik") or {}).get("vraceno")
                if _mail_to and st.session_state.get("_vrac_v"):
                    try:
                        _zivo1 = (vraceni_mejlovi(dana=30,
                                                  _v=st.session_state.get("_vrac_v", 0)) or {}
                                  ).get(_ocisti_mejl(_mail_to).lower())
                    except Exception:
                        _zivo1 = None
                    if _zivo1:
                        _zivo1 = dict(_zivo1); _zivo1["adresa"] = _ocisti_mejl(_mail_to).lower()
                        try:
                            sb_vraceno_upisi(mesec_key, sistem, sel_id, _zivo1,
                                             postojeci_dnevnik=(v.get("dnevnik") or {}))
                        except Exception:
                            pass
                        _vr1 = _zivo1
                if _vraceno_vazi(_vr1, (v.get("dnevnik") or {}).get("mejlovi"), _mail_to):
                    st.error("↩️ Mejl na ovu adresu se **vratio** ("
                             + str(_vr1.get("kada", "")) + ") — nije stigao do objekta. "
                             "Razlog: " + str(_vr1.get("razlog", ""))
                             + "\n\nZapisano je u bazi, pa ostaje vidljivo i ako obrišeš "
                               "odbijenicu iz sandučeta.")
                _kop1 = (_mres or {}).get("kopija")
                if _kop1 and not _kop1[0]:
                    st.warning("Mejl je poslat, ali kopija nije upisana u Poslato: "
                               + str(_kop1[1]))
                if _mres and not _mres.get("ok"):
                    st.error("❌ " + _mres["msg"])
            with _ecol3:
                st.caption("Samo status (🔴🟡🟢), naziv artikla i porudžbina (dodatna / naša) — spremno da se prosledi objektu na mejl.")

            # Naše količine = preporučena porudžbina umanjena za već trebovano posle 01. (min 0)
            _nase_rows = [(sel_id, int(a["ida"]), int(_dodatna_map.get(int(a["ida"]), 0)))
                          for a in _arts if int(_dodatna_map.get(int(a["ida"]), 0)) > 0]
            _njih_rows = [(sel_id, int(a["ida"]), int(_njihova_new.get(str(int(a["ida"])), 0)))
                          for a in _arts if int(_njihova_new.get(str(int(a["ida"])), 0)) > 0]
            st.markdown('<div class="adm-lbl" style="margin-top:16px;">Ubaci u admin</div>', unsafe_allow_html=True)
            if _treb_loaded:
                if _treb_total > 0:
                    st.caption("✅ Posle 01. je već trebovano " + str(_treb_total)
                               + " kom — Naše količine se šalju umanjeno (kolona Dodatna por.).")
                else:
                    st.caption("✅ Nema trebovanja posle 01. — šalju se pune količine.")
            elif _naziv_kom:
                st.caption("ℹ️ Za ovaj objekat prethodne porudžbine još nisu povučene — klikni "
                           "„Ažuriraj iz admina“ gore (povlači se samo razlika od poslednjeg ažuriranja).")
            else:
                st.caption("Za proveru trebovanja posle 01. učitaj šifarnik komitenata (potreban je naziv).")
            _nase_items = [{"idArticle": _a, "quantity": _q} for (_k, _a, _q) in _nase_rows]
            _njih_items = [{"idArticle": _a, "quantity": _q} for (_k, _a, _q) in _njih_rows]

            def _push_admin(_tag, _items):
                _sk = "admsent_" + str(sistem) + "_" + str(sel_id) + "_" + _tag
                _prev = st.session_state.get(_sk)
                if _prev and _prev.get("ok"):
                    return  # već uspešno kreirano — ne šaljemo ponovo (bez duplih porudžbina)
                with st.spinner("Šaljem u admin..."):
                    _ok, _msg = posalji_u_admin(sel_id, _items)
                st.session_state[_sk] = {"ok": _ok, "msg": _msg}
                if _ok:
                    # Trajno zapamti da smo MI ubacili porudžbinu kroz aplikaciju —
                    # takav objekat se NE zaključava, način trebovanja se bira ručno.
                    try:
                        sb_obrada_log(mesec_key, sistem, sel_id, "ubacena",
                                      st.session_state.get("admin_user", "Administracija"),
                                      kopija=_tag)
                    except Exception:
                        pass

            _xa, _xb, _xsp = st.columns([1.3, 1.3, 1])
            with _xa:
                _clk_nas = st.button("📦 Naše → admin", key="axn_" + str(sel_id),
                                     use_container_width=True,
                                     disabled=(len(_nase_items) == 0 or _zakljucan or _auto_lock))
            with _xb:
                _clk_nj = st.button("📦 Njihove → admin", key="axj_" + str(sel_id),
                                    use_container_width=True,
                                    disabled=(len(_njih_items) == 0 or _zakljucan or _auto_lock))
            if _clk_nas:
                _push_admin("nas", _nase_items)
            if _clk_nj:
                _push_admin("njihov", _njih_items)
            for _tag, _lbl in (("nas", "Naše"), ("njihov", "Njihove")):
                _sk = "admsent_" + str(sistem) + "_" + str(sel_id) + "_" + _tag
                _res = st.session_state.get(_sk)
                if _res:
                    (st.success if _res["ok"] else st.error)(_lbl + " → admin: " + _res["msg"])
                    if st.button("↺ Pošalji ponovo (" + _lbl + ")", key="axr_" + _tag + "_" + str(sel_id)):
                        st.session_state.pop(_sk, None)
                        st.rerun()

            st.markdown('<div class="adm-lbl" style="margin-top:18px;">🧾 Prethodne porudžbine (admin · ~6 meseci)</div>', unsafe_allow_html=True)
            if not _naziv_kom:
                st.caption("Nema naziva za ovaj objekat — učitaj šifarnik komitenata u Objavi izveštaja (ili poveži iz admina).")
                if st.button("🔗 Poveži šifre komitenata iz admina", key="bk_" + str(sel_id)):
                    with st.spinner("Čitam listu komitenata iz admina..."):
                        _bn, _be = admin_build_komitenti()
                    if _bn == 0:
                        st.error(_be or "Nije uspelo.")
                    else:
                        st.session_state["_komfull"] = sb_komitenti_full()
                        st.success("Povezano " + str(_bn) + " komitenata.")
                        st.rerun()
            else:
                _hist = st.session_state.get(_hk)
                _ah_at_d = meta.get("admin_hist_at") if isinstance(meta, dict) else None
                if _hist is None:
                    st.caption("Za ovaj objekat još nije povučeno iz admina — klikni „Ažuriraj iz admina“ gore "
                               "(povlači se samo razlika od poslednjeg ažuriranja).")
                elif _hist.get("err"):
                    st.error(_hist["err"])
                elif not _hist.get("lst"):
                    st.caption("Nema porudžbina za ovaj objekat u poslednjih ~6 meseci."
                               + (" (podaci iz admina od " + _dt_fmt(_ah_at_d) + ")" if _ah_at_d else ""))
                else:
                    st.caption("Pronađeno " + str(len(_hist["lst"])) + " porudžbina — klikni na datum za sadržaj.")
                    for _o in _hist["lst"]:
                        with st.expander("📅 " + (_o["datum"] or "?") + "   ·   " + (_o["status"] or "") + "   ·   " + (_o["cena"] or "")):
                            if _o.get("stavke"):
                                import pandas as _pdh
                                _hd = _pdh.DataFrame([{"Artikal": _s["naziv"], "Kol.": _s["kol"], "Cena": _s["cena"]}
                                                      for _s in _o["stavke"]])
                                st.dataframe(_hd, hide_index=True, use_container_width=True)
                            else:
                                st.caption("Nema stavki (ili nisu učitane).")
        with _dc2:
            if "Ubačena porudžbina" in (v.get("reakcije") or []):
                st.info("📦 Porudžbina je ubačena za ceo sistem (grupno).")
            _dnv = v.get("dnevnik") or {}
            _pozivi = _dnv.get("pozivi") or []
            _ncall = len(_pozivi)
            st.markdown('<div class="adm-lbl">Pozivi <span style="font-weight:400;text-transform:none;letter-spacing:0;color:#b0b4bd;">(automatski — preko dugmeta Pozovi)</span></div>', unsafe_allow_html=True)
            # Pozivi se NE mogu ručno čekirati — samo se prikazuju; broje se preko dugmeta „Pozovi" (slušalica).
            for _n in (1, 2, 3):
                _done = _ncall >= _n
                st.checkbox("📞 Pozvala sam " + str(_n) + ". put", value=_done,
                            key="callchk_" + str(sel_id) + "_" + str(_n), disabled=True)
                if _done:
                    _e = _pozivi[_n - 1]
                    st.markdown('<div style="font-size:10.5px;color:#9ca3af;font-style:italic;margin:-6px 0 4px 26px;">'
                                + _h_escape(str(_e.get("ko", ""))) + " · " + _h_escape(_dt_kratko(_e.get("at", "")))
                                + '</div>', unsafe_allow_html=True)
            st.markdown('<div class="adm-lbl" style="margin-top:10px;">Ostale reakcije</div>', unsafe_allow_html=True)
            _loaded = list(v.get("reakcije", []))
            _mejlovi = _dnv.get("mejlovi") or []
            _r2 = ("Poslala sam mejl" in _loaded)   # automatski — označava se pri slanju mejla (pojedinačno/grupno)
            st.checkbox("✉️ Poslala sam mejl  (automatski)", value=_r2,
                        key="r2_" + str(sel_id), disabled=True)
            if _mejlovi:
                _lm = _mejlovi[-1]
                st.markdown('<div style="font-size:10.5px;color:#9ca3af;font-style:italic;margin:-6px 0 4px 26px;">✉️ '
                            + _h_escape(str(_lm.get("ko", ""))) + " · " + _h_escape(_dt_kratko(_lm.get("at", "")))
                            + '</div>', unsafe_allow_html=True)
            _r3 = st.checkbox("👤 Prosledi komercijali", value=("Obavestila direktorku" in _loaded),
                              key="r3_" + str(sel_id), disabled=(_zakljucan or _auto_lock))
            react = []
            if _ncall > 0: react.append("Pozvala sam")
            if _r2: react.append("Poslala sam mejl")
            if _r3: react.append("Obavestila direktorku")
            # Otključano i kad je porudžbina sama stigla posle starta — inače bi
            # se automatski zabeležen status obrisao pri sledećem čuvanju.
            _can = bool(react) or bool(_auto_treb) or bool(_loaded_tip)
            st.markdown('<div class="adm-lbl" style="margin-top:12px;">Trebovanje nakon reakcije'
                        + (' · 🔒 zaključano' if _auto_lock else '') + '</div>', unsafe_allow_html=True)
            _tip_idx = ["", "nas", "njihov"].index(_tip_now) if _tip_now in ["", "nas", "njihov"] else 0
            _treb_lbl = st.radio("Trebovanje", TREB_OPT, index=_tip_idx,
                                 disabled=(not _can) or _auto_lock,
                                 key=_key_treb, label_visibility="collapsed")
            _tip_val = (TREB_CODE.get(_treb_lbl, "") if _can else "") if not _auto_lock else _tip_now
            if _auto_lock:
                _ai2 = dict(_auto_treb.get("info") or {})
                st.caption("🔒 Prepoznato automatski po količinama iz admina — "
                           + ("po NAŠEM sistemu (sve je poručeno kako smo tražili)"
                              if _tip_now == "nas"
                              else ("po NJIHOVOM sistemu — poručeno "
                                    + str(_ai2.get("pokriveno", 0)) + " od "
                                    + str(_ai2.get("trazeno", 0)) + " kom dopune"
                                    + ((", " + str(_ai2.get("fali_n", 0)) + " artikal/artikala nije poručeno kako smo tražili")
                                       if _ai2.get("fali_n") else "")))
                           + ". Ne menja se ručno — objekat je završen.")
            elif not _can:
                st.caption("🔒 Otključava se kad izabereš bar jednu reakciju.")
            st.markdown('<div class="adm-lbl" style="margin-top:12px;">Napomena (interno)</div>', unsafe_allow_html=True)
            _nap = st.text_area("Napomena", value=(v.get("napomena", "") or ""), key="nap_" + str(sel_id),
                                height=72, label_visibility="collapsed",
                                placeholder="npr. zvati posle 15h, tražiti vlasnika...")
            if st.button(("💾 Sačuvaj napomenu" if _auto_lock else "💾 Sačuvaj status"),
                         key="savest_" + str(sel_id), type="primary", disabled=_zakljucan):
                if not _can:
                    st.error("Izaberi bar jednu reakciju — ne može da se sačuva samo napomena.")
                elif ("Obavestila direktorku" in react) and not (_nap or "").strip():
                    st.error("Za prosleđivanje komercijali upiši napomenu — zašto prosleđuješ (obavezno).")
                elif (not _auto_lock) and _tip_val == "njihov" and sum(int(x) for x in _njihova_new.values()) == 0:
                    st.error("Za opciju Po njihovom sistemu upiši koliko su poručili (Njihova por.) pre čuvanja.")
                else:
                    try:
                        _cur_user = st.session_state.get("admin_user", "Administracija")
                        _prev_ko = dict(v.get("reakcije_ko") or {})
                        _new_ko = {}
                        for _rk in react:
                            _new_ko[_rk] = _prev_ko.get(_rk) or _cur_user
                        sb_save_obrada(mesec_key, sistem, sel_id, react, _tip_val, _njihova_new, _nap,
                                       reakcije_ko=_new_ko, azurirao=_cur_user)
                        # ako je TEK SADA prosleđeno komercijali — zabeleži vreme prosleđivanja
                        if ("Obavestila direktorku" in react) and ("Obavestila direktorku" not in _loaded):
                            try:
                                sb_obrada_log(mesec_key, sistem, sel_id, "komercijala", _cur_user)
                            except Exception:
                                pass
                        st.success("Sačuvano ✓")
                        st.rerun()
                    except Exception as _e:
                        st.error("Greška: " + str(_e))

    with tab_bulk:
        st.markdown('<div class="adm-lbl">Grupno slanje mejlova objektima</div>', unsafe_allow_html=True)
        st.caption("Vidiš sve objekte sistema. Filtriraj po zoni / statusu mejla / nazivu, štikliraj koje želiš, "
                   "pa klikni Pošalji izabranima. Slanje radi sve isto kao pojedinačno slanje mejla "
                   "(šalje Excel prilog i automatski upali reakciju Poslala sam mejl). Šalje se samo objektima sa emailom i dodatnom porudžbinom.")
        if not smtp_dostupan():
            _n = _mail_nalog()
            _suf = ("_" + _n) if _n else ""
            st.warning("✉️ Slanje mejlova nije podešeno za tvoj nalog — dodaj u Secrets: "
                       "SMTP_HOST" + _suf + " / SMTP_USER" + _suf + " / SMTP_PASSWORD" + _suf + ".")
        else:
            _vc1, _vc2 = st.columns([3, 1.2])
            with _vc1:
                st.caption("✉️ Mejlovi se šalju sa: " + str(_smtp_cfg().get("from_email", "")))
            with _vc2:
                if st.button("↩️ Proveri vraćene mejlove", key="vrac_chk", use_container_width=True):
                    st.session_state["_vrac_v"] = st.session_state.get("_vrac_v", 0) + 1
                    try:
                        vraceni_mejlovi.clear()
                    except Exception:
                        pass
                    st.rerun()

        _selk = "bulk_sel_" + str(sistem) + "_" + str(mesec_key)
        _verk = "bulk_ver_" + str(sistem) + "_" + str(mesec_key)
        if _selk not in st.session_state:
            st.session_state[_selk] = set()
        if _verk not in st.session_state:
            st.session_state[_verk] = 0
        # Svi objekti sistema + oznaka da li je mejl već poslat (iz baze — obrada reakcije)
        _bulk_rows = []
        _n_izbaceno_zav = 0
        for o in objekti:
            _bidk = int(o["idk"])
            # Objekti koji su već poručili (ZAVRŠENO) ne ulaze u listu za slanje —
            # nema smisla tražiti porudžbinu od nekog ko je upravo poručio.
            if _zavrsen(o, obrada_map.get(_bidk, {}) or {}):
                _n_izbaceno_zav += 1
                continue
            _bkinfo = komfull.get(_bidk, {}) or {}
            _bnaziv = _bkinfo.get("naziv", "") or ("ID " + str(_bidk))
            _bemail = (_bkinfo.get("email") or "").strip()
            _bz = _zona_disp(o["nivo"])
            _bhist = st.session_state.get("hist_" + str(sistem) + "_" + str(_bidk))
            _btreb_map = _treb_posle_preseka(_bhist.get("lst") or [], _cut_hit) if (_bhist and not _bhist.get("err")) else {}
            _bstavki = 0
            for a in o["lst"]:
                _bpor = int(_btreb_map.get(int(a["ida"]), 0))
                if max(int(a["kol"]) - _bpor, 0) > 0:
                    _bstavki += 1
            _bobr = obrada_map.get(_bidk, {}) or {}
            _bmail_sent = "Poslala sam mejl" in (_bobr.get("reakcije") or [])
            _bmail_ko = (_bobr.get("reakcije_ko") or {}).get("Poslala sam mejl", "")
            _bmej_lst = ((_bobr.get("dnevnik") or {}).get("mejlovi") or [])
            _bmail_n = len(_bmej_lst) if _bmej_lst else (1 if _bmail_sent else 0)
            # kada je poslednji mejl poslat (za kolonu „Poslato“ i sortiranje po vremenu)
            _bmail_at = None
            if _bmej_lst:
                try:
                    _bmail_at = datetime.datetime.fromisoformat(str(_bmej_lst[-1].get("at", "")))
                except Exception:
                    _bmail_at = None
            _bulk_rows.append({
                "idk": _bidk, "Naziv": _bnaziv, "Email": _bemail or "(nema mejla)",
                "zona_txt": _bz[3], "zona_dot": _bz[2], "Stavki": _bstavki,
                "mail_sent": _bmail_sent, "mail_ko": _bmail_ko, "mail_n": _bmail_n,
                "mail_at": _bmail_at, "_mejlovi": _bmej_lst,
                "_vraceno_baza": (_bobr.get("dnevnik") or {}).get("vraceno"),
                "_zona": o["nivo"], "_email_ok": _mejl_ok(_bemail), "_has_rows": _bstavki > 0,
            })
        _vrac = {}
        if st.session_state.get("_vrac_v"):
            try:
                _vrac = vraceni_mejlovi(dana=30, _v=st.session_state.get("_vrac_v", 0)) or {}
            except Exception:
                _vrac = {}
        # Odbijenice se pamte u bazi, pa oznaka „VRAĆENO“ ostaje i kad se poruka
        # obriše iz sandučeta — i vidi se bez ponovnog čitanja pošte.
        _novih_vrac = 0
        for _r1 in _bulk_rows:
            _e1 = _ocisti_mejl(_r1.get("Email", "")).lower()
            _zivo = _vrac.get(_e1) if _e1 else None
            if _zivo:
                _zivo = dict(_zivo); _zivo["adresa"] = _e1
                _stara = _r1.get("_vraceno_baza") or {}
                if (str(_stara.get("kada", "")) != str(_zivo.get("kada", ""))
                        or _ocisti_mejl(_stara.get("adresa", "")).lower() != _e1):
                    try:
                        if sb_vraceno_upisi(mesec_key, sistem, _r1["idk"], _zivo,
                                            postojeci_dnevnik=(obrada_map.get(_r1["idk"], {}) or {}).get("dnevnik")):
                            _novih_vrac += 1
                    except Exception:
                        pass
                _r1["_vraceno_baza"] = _zivo
            _kand = _r1.get("_vraceno_baza")
            _r1["vraceno"] = _kand if _vraceno_vazi(_kand, _r1.get("_mejlovi"), _r1.get("Email", "")) else None
        GRUPNO_MAX = 2   # koliko puta objekat sme da dobije mejl grupnim slanjem
        for _r0 in _bulk_rows:
            _r0["_grupno_ok"] = int(_r0.get("mail_n", 0) or 0) < GRUPNO_MAX
        _n_sent_total = sum(1 for r in _bulk_rows if r["mail_sent"])
        _n_iscrpljeno = sum(1 for r in _bulk_rows if not r["_grupno_ok"])
        _n_drugi_krug = sum(1 for r in _bulk_rows if int(r.get("mail_n", 0) or 0) == 1)
        # KPI traka
        st.markdown('<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:2px 0 12px;">'
                    '<div style="background:#faf7ff;border:1px solid #e9d5ff;border-radius:12px;padding:12px 16px;">'
                    '<div style="font-size:20px;font-weight:800;color:#7c3aed;">' + str(len(_bulk_rows)) + '</div>'
                    '<div style="font-size:11.5px;color:#8b7fa8;">Ukupno objekata</div></div>'
                    '<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:12px;padding:12px 16px;">'
                    '<div style="font-size:20px;font-weight:800;color:#16a34a;">' + str(_n_sent_total) + '</div>'
                    '<div style="font-size:11.5px;color:#5b8c6b;">Mejl već poslat</div></div>'
                    '<div style="background:#fff7f7;border:1px solid #fecaca;border-radius:12px;padding:12px 16px;">'
                    '<div style="font-size:20px;font-weight:800;color:#dc2626;">' + str(sum(1 for r in _bulk_rows if r["_grupno_ok"] and r["_email_ok"] and r["_has_rows"])) + '</div>'
                    '<div style="font-size:11.5px;color:#9b6b6b;">Može grupno (1. ili 2. put)</div></div>'
                    '<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:12px;padding:12px 16px;">'
                    '<div style="font-size:20px;font-weight:800;color:#b45309;">' + str(sum(1 for r in _bulk_rows if not r["_email_ok"])) + '</div>'
                    '<div style="font-size:11.5px;color:#9a7b3a;">Bez ispravnog mejla</div></div></div>',
                    unsafe_allow_html=True)
        # Filteri (objekti kojima je mejl VEĆ poslat se ne prikazuju — ne mogu se ponovo slati grupno)
        _fc1, _fc2, _fc3 = st.columns([1, 1.35, 2])
        with _fc1:
            _f_zona = st.selectbox("Zona", ["Sve zone", "🔴 Hitno", "🟡 Iskontrolisati", "🟢 Dobra"], key="bulk_f_zona")
        with _fc2:
            _f_mejl = st.selectbox("Status mejla",
                                   ["Svi", "Nije poslat nijednom", "Poslat jednom (drugi krug)",
                                    "Samo vraćeni"], key="bulk_f_mejl")
        with _fc3:
            _f_q = st.text_input("Pretraga (naziv ili email)", value="", key="bulk_f_q", placeholder="npr. Novi Sad ili mp123@…")
        _zmap = {"🔴 Hitno": "crveno", "🟡 Iskontrolisati": "zuto", "🟢 Dobra": "zeleno"}
        def _match(r):
            if not r["_grupno_ok"]:
                return False   # već dobio 2 mejla — grupno više ne može, samo pojedinačno
            if _f_zona in _zmap and r["_zona"] != _zmap[_f_zona]:
                return False
            _mn = int(r.get("mail_n", 0) or 0)
            if _f_mejl == "Nije poslat nijednom" and _mn != 0:
                return False
            if _f_mejl == "Poslat jednom (drugi krug)" and _mn != 1:
                return False
            if _f_mejl == "Samo vraćeni" and not r.get("vraceno"):
                return False
            if _f_q.strip():
                _qq = _f_q.strip().lower()
                if _qq not in str(r["Naziv"]).lower() and _qq not in str(r["Email"]).lower():
                    return False
            return True
        _view_rows = [r for r in _bulk_rows if _match(r)]
        if _n_izbaceno_zav:
            st.caption("✅ " + str(_n_izbaceno_zav) + " objekata je ZAVRŠENO (poručili su) — "
                       "ne prikazuju se ovde i ne mogu da dobiju mejl.")
        _n_vrac = sum(1 for r in _bulk_rows if r.get("vraceno"))
        if _n_vrac:
            _lst = [r for r in _bulk_rows if r.get("vraceno")][:12]
            st.error("↩️ **" + str(_n_vrac) + " mejlova se vratilo** — nisu stigli do objekta:\n\n"
                     + "\n".join("• " + str(r["Naziv"])[:52] + " · " + str(r["Email"])
                                 + " — " + str((r["vraceno"] or {}).get("razlog", ""))[:70]
                                 for r in _lst)
                     + (("\n\n… i još " + str(_n_vrac - len(_lst))) if _n_vrac > len(_lst) else "")
                     + "\n\nOvo je zapisano u bazi — ostaje i kad obrišeš odbijenice iz sandučeta. "
                       "Oznaka sama nestaje kad ispraviš adresu ili kad mejl ponovo prođe.")
        elif st.session_state.get("_vrac_v"):
            st.success("↩️ Nema vraćenih mejlova u poslednjih 30 dana.")
        if _n_drugi_krug or _n_iscrpljeno:
            _p = []
            if _n_drugi_krug:
                _p.append(str(_n_drugi_krug) + " objekata je dobilo jedan mejl — njima se grupno "
                          "može poslati još jednom (drugi krug)")
            if _n_iscrpljeno:
                _p.append(str(_n_iscrpljeno) + " objekata je dobilo dva mejla — oni se više ne "
                          "prikazuju ovde, ponovno slanje je moguće samo pojedinačno sa kartice objekta")
            st.caption("ℹ️ " + ". ".join(_p) + ".")
        # Kad se filter promeni — poništi izbor (da izbor uvek prati ono što je trenutno prikazano)
        _sig = str(_f_zona) + "|" + str(_f_mejl) + "|" + _f_q.strip().lower()
        _sigk = "bulk_sig_" + str(sistem) + "_" + str(mesec_key)
        if st.session_state.get(_sigk) != _sig:
            st.session_state[_sigk] = _sig
            st.session_state[_selk] = set()
            st.session_state[_verk] += 1
        st.caption("Prikazano posle filtera: " + str(len(_view_rows)) + " objekata"
                   + (" (od ukupno " + str(len(_bulk_rows)) + ")" if len(_view_rows) != len(_bulk_rows) else ""))
        # Akcije nad filtriranom listom
        _ba1, _ba2, _ba3 = st.columns([1.3, 1.3, 3])
        with _ba1:
            if st.button("☑️ Izaberi sve (filtrirane)", key="bulk_pick_filt", use_container_width=True):
                for r in _view_rows:
                    st.session_state[_selk].add(r["idk"])
                st.session_state[_verk] += 1
                st.rerun()
        with _ba2:
            if st.button("✖️ Poništi izbor", key="bulk_pick_none", use_container_width=True):
                st.session_state[_selk] = set()
                st.session_state[_verk] += 1
                st.rerun()
        if not _view_rows:
            st.info("Nema objekata za izabrane filtere.")
        else:
            _bdf = pd.DataFrame([{
                "Izabrano": r["idk"] in st.session_state[_selk],
                "Zona": r["zona_dot"],
                "Naziv": r["Naziv"],
                "Email": r["Email"],
                "Mejl": ("⚠️ VRAĆENO" if r.get("vraceno")
                         else (("✅ " + str(r["mail_n"]) + "× "
                                + (_ko_kratko(r["mail_ko"]) if r["mail_ko"] else "")).strip()
                               if int(r.get("mail_n", 0) or 0) > 0 else "—")),
                "Stavki": r["Stavki"],
                "Poslato": r.get("mail_at"),
            } for r in _view_rows])
            # indeks = ID objekta, da izbor ostane tačan i kad se tabela presortira
            _bdf.index = [int(r["idk"]) for r in _view_rows]
            # kolona mora da bude pravi datum (a ne tekst) da bi sortiranje radilo
            # i kad nijedan objekat još nema poslat mejl
            _bdf["Poslato"] = pd.to_datetime(_bdf["Poslato"], errors="coerce")
            _h_editor = min(60 + 36 * len(_view_rows), 760)
            _bedited = st.data_editor(
                _bdf, hide_index=True, use_container_width=True, height=_h_editor,
                key="bulk_editor_" + str(sistem) + "_" + str(mesec_key) + "_" + str(st.session_state[_verk]),
                disabled=["Zona", "Naziv", "Email", "Mejl", "Stavki", "Poslato"],
                column_config={"Izabrano": st.column_config.CheckboxColumn("Izabrano", width="small"),
                               "Zona": st.column_config.TextColumn("Zona", width="small"),
                               "Mejl": st.column_config.TextColumn("Mejl", width="small"),
                               "Poslato": st.column_config.DatetimeColumn(
                                   "Poslato", width="medium", format="DD.MM.YYYY. HH:mm",
                                   help="Kada je poslednji mejl poslat. Klikni na zaglavlje kolone "
                                        "da sortiraš — najstariji gore znači da si te prvo slala.")})
            _new_sel_vis = set()
            for _idx, _r in _bedited.iterrows():
                if bool(_r["Izabrano"]):
                    try:
                        _new_sel_vis.add(int(_idx))
                    except Exception:
                        pass
            # izbor = tačno ono što je štiklirano u trenutno prikazanoj (filtriranoj) listi
            st.session_state[_selk] = _new_sel_vis
        _sel_ids = st.session_state[_selk]
        _view_ids = set(r["idk"] for r in _view_rows)
        _sel_ids = _sel_ids & _view_ids   # sigurnosno: samo iz trenutnog filtera
        _n_sel = len(_sel_ids)
        _n_sel_ok = sum(1 for r in _view_rows if r["idk"] in _sel_ids and r["_email_ok"] and r["_has_rows"])
        _n_sel_no_email = sum(1 for r in _view_rows if r["idk"] in _sel_ids and not r["_email_ok"])
        _n_sel_no_rows = sum(1 for r in _view_rows if r["idk"] in _sel_ids and r["_email_ok"] and not r["_has_rows"])
        st.caption("Izabrano: " + str(_n_sel) + " objekata · spremno za slanje: " + str(_n_sel_ok)
                   + ((" · bez mejla: " + str(_n_sel_no_email)) if _n_sel_no_email else "")
                   + ((" · nema dodatne porudžbine: " + str(_n_sel_no_rows)) if _n_sel_no_rows else ""))
        _mem_sada = _mem_mb()
        if _mem_sada:
            st.caption("🧠 Memorija aplikacije: " + str(_mem_sada) + " MB / ~1000 MB"
                       + ("  ·  ⚠️ blizu granice — osveži stranu (F5) pre grupnog slanja"
                          if _mem_sada > 700 else ""))
        _flash = st.session_state.pop("_bulk_flash", None)
        if _flash:
            if _flash[0] == "ok":
                st.success(_flash[1])
            else:
                st.warning(_flash[1])
        if st.button("📧 Pošalji izabranima (" + str(_n_sel_ok) + ")", key="bulk_send", type="primary",
                     use_container_width=True, disabled=(_n_sel_ok == 0 or not smtp_dostupan() or _zakljucan)):
            _bprog = st.progress(0, "✉️ Pripremam slanje…")
            _to_send_svi = [r for r in _view_rows if r["idk"] in _sel_ids and r["_email_ok"] and r["_has_rows"]]
            # Koliko mejlova po jednom kliku (da slanje sigurno stigne do kraja
            # pre nego što Streamlit prekine vezu). 0 = bez ograničenja.
            try:
                _bmax = int(str(_cfg("GRUPNO_BATCH", "8")).strip() or 0)
            except Exception:
                _bmax = 8
            _to_send = _to_send_svi[:_bmax] if _bmax > 0 else _to_send_svi
            _ostalo = len(_to_send_svi) - len(_to_send)
            # pauza između mejlova — server lakše podnosi niz mejlova zaredom
            try:
                _bpauza = float(str(_cfg("GRUPNO_PAUZA", "2")).strip() or 0)
            except Exception:
                _bpauza = 2.0
            _cur_u = st.session_state.get("admin_user", "Administracija")
            _n_ok = 0; _n_fail = 0
            _n_bez_kopije = 0; _zasto_kopija = ""
            _stop_razlog = ""
            _stop_tip = ""
            _mem_max = _mem_mb()
            _ses = MejlSesija()
            for _bi, r in enumerate(_to_send):
                _bidk2 = r["idk"]
                _o2 = obj_by_id[_bidk2]
                _bkinfo2 = komfull.get(_bidk2, {}) or {}
                _bnaziv2 = _bkinfo2.get("naziv", "") or ("ID " + str(_bidk2))
                _bemail2 = (_bkinfo2.get("email") or "").strip()
                _bhist2 = st.session_state.get("hist_" + str(sistem) + "_" + str(_bidk2))
                _btreb_map2 = _treb_posle_preseka(_bhist2.get("lst") or [], _cut_hit) if (_bhist2 and not _bhist2.get("err")) else {}
                _bmeseci = float(meta.get("meseci") or 1.0) if isinstance(meta, dict) else 1.0
                _bexp_rows2 = []
                for a in _o2["lst"]:
                    _bida2 = int(a["ida"]); _bkol2 = int(a["kol"])
                    _bpor2 = int(_btreb_map2.get(_bida2, 0))
                    _bdod2 = max(_bkol2 - _bpor2, 0)
                    if _bdod2 > 0:
                        _blg2 = int(a["lager"]) + _bpor2
                        _bsd2 = "🔴" if _blg2 == 0 else ("🟡" if _blg2 <= 2 else "🟢")
                        _bexp_rows2.append({"kruzic": _bsd2, "naziv": str(a["naziv"]), "lager": _blg2,
                                            "predikcija": int(round(int(a.get("pred", 0) or 0) * _bmeseci)),
                                            "dodatna": _bdod2})
                _bsk2 = "mailsent_" + str(sistem) + "_" + str(_bidk2)
                _mem0 = _mem_mb()
                _bprog.progress(int(_bi / len(_to_send) * 100),
                                "✉️ Šaljem " + str(_bi + 1) + "/" + str(len(_to_send))
                                + " — " + str(_bnaziv2)[:40] + " …"
                                + ((" · memorija " + str(_mem0) + " MB") if _mem0 else ""))
                try:
                    _bxlsx2 = _objekat_order_xlsx(_bnaziv2, _bidk2, _sel_lbl, _bexp_rows2, meseci=meta.get("meseci") if isinstance(meta, dict) else None)
                    import re as _refn2
                    _safe_sis2 = "".join(ch for ch in str(sistem or "") if ch.isalnum() or ch in " _-").strip()
                    _mpm2 = _refn2.search(r'MP\s*\d+', str(_bnaziv2 or ""), _refn2.IGNORECASE)
                    _mp2 = _mpm2.group(0).upper().replace(" ", "") if _mpm2 else str(_bidk2)
                    _bfname2 = ((_safe_sis2 + " ") if _safe_sis2 else "") + _mp2 + ".xlsx"
                    _bsubj2 = ("VAPE SHOP - " + str(_bnaziv2) + " - PORUDŽBINA - "
                               + _now().strftime("%d.%m.%Y."))
                    posalji_mejl_sa_prilogom(_bemail2, _bsubj2, _mejl_tekst(_bnaziv2),
                                             _bxlsx2, _bfname2, sesija=_ses)
                    st.session_state[_bsk2] = {"ok": True, "msg": "Poslato na " + _bemail2}
                    _n_ok += 1
                    _kop2 = st.session_state.get("_zadnja_kopija")
                    if _kop2 and not _kop2[0]:
                        _n_bez_kopije += 1
                        _zasto_kopija = str(_kop2[1])
                    # skini ga sa izbora, da se sledećim klikom ne pošalje ponovo
                    try:
                        st.session_state[_selk].discard(_bidk2)
                    except Exception:
                        pass
                    # Auto: zabeleži mejl u dnevnik (ko + vreme) + upali „Poslala sam mejl"
                    try:
                        sb_obrada_log(mesec_key, sistem, _bidk2, "mejl", _cur_u,
                                      kopija=st.session_state.get("_zadnja_kopija"))
                    except Exception:
                        pass
                except MejlOgranicenje as _bo:
                    # server nas privremeno koči — nema smisla nastaviti
                    st.session_state[_bsk2] = {"ok": False, "msg": str(_bo)}
                    _n_fail += 1
                    try:
                        sb_mejl_greska(mesec_key, sistem, _bidk2, str(_bo), _cur_u)
                    except Exception:
                        pass
                    _stop_razlog = str(_bo)
                    _stop_tip = "server"
                    break
                except Exception as _be:
                    st.session_state[_bsk2] = {"ok": False, "msg": str(_be)}
                    _n_fail += 1
                    # upiši grešku u bazu ODMAH, da ostane zapisana i ako se strana izgubi
                    try:
                        sb_mejl_greska(mesec_key, sistem, _bidk2, str(_be), _cur_u)
                    except Exception:
                        pass
                # oslobodi memoriju odmah — prilog ume da bude po nekoliko MB,
                # a Streamlit Cloud restartuje aplikaciju (i izbaci te na prijavu)
                # ako ukupna memorija pređe ~1000 MB
                try:
                    del _bxlsx2
                except Exception:
                    pass
                _bexp_rows2 = None
                import gc as _gc
                _gc.collect()
                _mem1 = _mem_mb()
                _mem_max = max(_mem_max, _mem1)
                _bprog.progress(int((_bi + 1) / len(_to_send) * 100),
                                "✅ Poslato " + str(_bi + 1) + "/" + str(len(_to_send))
                                + ((" · memorija " + str(_mem1) + " MB") if _mem1 else ""))
                if _mem1 and _mem1 > 800:
                    _stop_razlog = ("memorija aplikacije je na " + str(_mem1)
                                    + " MB (granica je oko 1000 MB)")
                    _stop_tip = "memorija"
                    break
                if _bpauza > 0 and _bi < len(_to_send) - 1:
                    import time as _tsl
                    _tsl.sleep(_bpauza)
            try:
                _ses.zatvori()
            except Exception:
                pass
            _bprog.empty()
            _por = []
            if _n_fail == 0:
                _tip = "ok"
                _por.append("✅ Poslato " + str(_n_ok) + " mejlova. Status je zabeležen (vidljiv i posle odjave).")
            else:
                _tip = "warn"
                _por.append("Poslato " + str(_n_ok) + " · nije uspelo " + str(_n_fail)
                            + " (proveri status u tabeli iznad).")
            if _stop_razlog and _stop_tip == "memorija":
                _tip = "warn"
                _por.append("⏸️ Slanje je zaustavljeno da aplikacija ne bi pukla — " + _stop_razlog
                            + ".\n\nOvo je isto ono što te ranije izbacivalo na ekran za prijavu. "
                            "Osveži stranu (F5), prijavi se ponovo i klikni „📧 Pošalji izabranima“ — "
                            "nastaviće tačno tamo gde je stalo, poslati objekti su već odštiklirani.")
            elif _stop_razlog:
                _tip = "warn"
                _por.append("⏸️ Slanje je zaustavljeno jer nas server privremeno koči: " + _stop_razlog
                            + "\n\nTo NIJE greška u adresama — isti mejlovi će proći kasnije. "
                            "Sačekaj sat vremena (ili do sutra ako je dnevno ograničenje) pa klikni ponovo; "
                            "poslati objekti su već odštiklirani, tako da niko neće dobiti mejl dvaput.")
            _ostalo = len(_to_send_svi) - _n_ok - _n_fail
            if _ostalo > 0 and not _stop_razlog:
                _tip = "warn"
                _por.append("Ostalo je još " + str(_ostalo) + " objekata — oni su i dalje štiklirani, "
                            "samo klikni ponovo „📧 Pošalji izabranima“ da se pošalje i taj deo.")
            if _n_bez_kopije:
                _tip = "warn"
                _por.append("📭 " + str(_n_bez_kopije) + " od " + str(_n_ok)
                            + " mejlova NIJE upisano u folder Poslato — mejlovi jesu otišli primaocima, "
                            "ali ih nećeš videti u sandučetu. Razlog: " + (_zasto_kopija or "nepoznat")
                            + ". Kod svakog objekta piše da li je kopija upisana.")
            if _mem_max:
                _por.append("🧠 Najveća zauzeta memorija tokom slanja: " + str(_mem_max)
                            + " MB (aplikacija se restartuje oko 1000 MB).")
            st.session_state["_bulk_flash"] = (_tip, "\n\n".join(_por))
            # novi ključ za tabelu, da se poslati objekti stvarno odštikliraju
            try:
                st.session_state[_verk] += 1
            except Exception:
                pass
            st.rerun()


def _direktor_blok_iz_prodaje(sistem, sales):
    """Napravi direktorski blok (prodaja_trend, poređenja, po grupama) iz tabele prodaje
    (Izveštaj prodaje), tako da ne treba ponovna objava sistema. Vrati dict ili None."""
    if not sales:
        return None
    po = sales.get("po_sistemu") or {}
    _sn = str(sistem).strip().upper()
    _key = None
    for k in po.keys():
        if str(k).strip().upper() == _sn:
            _key = k
            break
    if _key is None:
        return None
    nazivi = sales.get("nazivi") or []
    total = po[_key].get("total") or []
    grupe = po[_key].get("grupe") or {}
    _n = min(len(nazivi), len(total))
    trend = [{"mesec": nazivi[i], "kom": int(total[i])} for i in range(_n)]
    if not trend:
        return None
    d = {"prodaja_trend": trend, "prodaja_tekuci": trend[-1]}
    comp = {}
    if len(trend) >= 2:
        comp["prosli_mesec"] = trend[-2]
        _prev6 = trend[-7:-1] if len(trend) >= 7 else trend[:-1]
        if _prev6:
            comp["prosek_6m"] = {"kom": int(round(sum(t["kom"] for t in _prev6) / len(_prev6)))}
    _last = str(trend[-1]["mesec"]).split()
    if len(_last) == 2 and _last[1].isdigit():
        _lani = _last[0] + " " + str(int(_last[1]) - 1)
        for t in trend:
            if t["mesec"] == _lani:
                comp["isti_mesec_lani"] = t
                break
    d["poredjenja"] = comp
    _pg = []
    for grp, vals in grupe.items():
        if vals:
            _pg.append({"grupa": grp, "kom": int(vals[-1])})
    _pg.sort(key=lambda x: -x["kom"])
    d["po_grupama"] = _pg
    # Mesečne vrednosti po grupi (za složeni stub-grafikon) + nazivi meseca
    d["nazivi"] = [nazivi[i] for i in range(_n)]
    _gm = {}
    for grp, vals in grupe.items():
        vv = vals or []
        _gm[str(grp)] = [int(vv[i]) if i < len(vv) else 0 for i in range(_n)]
    d["grupe_mesecno"] = _gm
    return d


def prikazi_komercijalu():
    """Pregled za komercijalu (Komercijala 1 / Komercijala 2):
    1) Moja ruta (dan)  2) Objekat — popis i porudžbina  3) Kontrola od administracije."""
    st.set_page_config(page_title="VAPE — Komercijala", page_icon="🤝",
                       layout="wide", initial_sidebar_state="collapsed")
    st.markdown("""<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    section[data-testid="stSidebar"]{display:none !important;}
    header[data-testid="stHeader"]{display:none !important;}
    .stApp{background:#f4f1fb !important;font-family:'Inter',sans-serif;}
    .block-container{max-width:1180px !important;padding-top:22px !important;}
    [class*="st-key-kom_odjava"] button{font-size:11px !important;padding:4px 9px !important;border-radius:7px !important;
        min-height:0 !important;background:#fff !important;border:1px solid #e5e7eb !important;color:#6b7280 !important;font-weight:600 !important;}
    .kom-card{background:#fff;border:1px solid #efeaf7;border-radius:14px;padding:14px 18px;margin-bottom:11px;
        box-shadow:0 2px 12px rgba(80,40,140,.05);}
    .kom-nm{font-weight:700;font-size:14.5px;}
    .kom-ad{font-size:12px;color:#9aa0ad;margin-top:2px;}
    .kchip{display:inline-block;padding:3px 10px;border-radius:20px;font-size:11.5px;font-weight:700;}
    .k-red{background:#fee2e2;color:#991b1b;} .k-org{background:#ffedd5;color:#9a3412;}
    .k-grn{background:#dcfce7;color:#14532d;} .k-gry{background:#f1f2f6;color:#8b8fa0;}
    .k-vio{background:#f3e8ff;color:#6b21a8;}
    </style>""", unsafe_allow_html=True)

    def _kfmt(n):
        try:
            return f"{int(round(float(n))):,}".replace(",", ".")
        except Exception:
            return str(n)

    def _kpi(cells):
        _h = '<div style="display:grid;grid-template-columns:repeat(' + str(len(cells)) + ',1fr);gap:12px;margin:8px 0 14px;">'
        for _n, _k, _bg, _bd, _cn, _ck in cells:
            _h += ('<div style="background:' + _bg + ';border:1px solid ' + _bd + ';border-radius:12px;padding:13px 16px;">'
                   '<div style="font-size:21px;font-weight:800;color:' + _cn + ';">' + str(_n) + '</div>'
                   '<div style="font-size:11.5px;color:' + _ck + ';margin-top:2px;">' + _k + '</div></div>')
        return _h + '</div>'
    _P = ("#faf7ff", "#e9d5ff", "#7c3aed", "#8b7fa8")
    _G = ("#f0fdf4", "#bbf7d0", "#16a34a", "#5b8c6b")
    _O = ("#fffbeb", "#fde68a", "#b45309", "#9a7b3a")
    _R = ("#fff7f7", "#fecaca", "#dc2626", "#9b6b6b")

    _ku = st.session_state.get("komerc_user", "Komercijala")
    _h1, _h2 = st.columns([7.2, 1.35])
    with _h1:
        st.markdown('<div style="display:flex;align-items:center;gap:10px;margin-bottom:2px;">'
                    '<div style="width:30px;height:30px;border-radius:9px;background:linear-gradient(135deg,#7c3aed,#a855f7);"></div>'
                    '<span style="font-size:20px;font-weight:800;color:#1f2430;">VAPE Porudžbine</span>'
                    '<span style="font-size:12.5px;color:#9aa0ad;">· ' + _h_escape(_ku) + '</span></div>',
                    unsafe_allow_html=True)
    with _h2:
        if st.button("Odjava", key="kom_odjava", use_container_width=True):
            for _k in ("authenticated", "role", "komerc_user", "mail_nalog"):
                st.session_state.pop(_k, None)
            st.rerun()
        if st.button("🔄 Ažuriraj iz admina", key="kom_refresh_all",
                     use_container_width=True, type="primary"):
            st.session_state["_req_kom_refresh"] = True
            st.rerun()

    if not sb_dostupan():
        st.error("Veza sa bazom trenutno nije podešena. Javi se analitičaru.")
        return

    _tab_ok, _tab_msg = sb_ruta_tabele_ok()
    if not _tab_ok:
        st.error("⚠️ " + _tab_msg)
        st.markdown('<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:12px;padding:14px 18px;'
                    'font-size:13.5px;color:#78500a;line-height:1.6;">'
                    '<b>Šta treba uraditi (jednom):</b><br>'
                    '1. Otvori Supabase → <b>SQL Editor</b> → <b>New query</b><br>'
                    '2. Nalepi <b>KORAK 1</b> iz fajla <i>SQL_komercijala_rute.sql</i> i klikni <b>Run</b><br>'
                    '3. Za slike: Supabase → <b>Storage</b> → <b>New bucket</b> → ime <b>ruta-slike</b>, '
                    'čekiraj <b>Public bucket</b> → Save<br>'
                    '4. Osveži ovu stranicu</div>', unsafe_allow_html=True)
        return

    _pub = sb_meseci()
    _mkeys = sorted({m["key"] for m in _pub}, reverse=True)
    if not _mkeys:
        st.info("Još nema objavljenih podataka.")
        return
    try:
        _kf = sb_komitenti_full()
    except Exception:
        _kf = {}

    # tekući mesec = najnoviji objavljeni
    _mes_now = _mkeys[0]

    def _naz(idk):
        return (_kf.get(int(idk), {}) or {}).get("naziv", "") or ("ID " + str(idk))

    # --- Ažuriranje iz admina za CELU rutu odjednom (dugme je gore, ispod Odjave) ---
    _datum_r = st.session_state.get("kom_datum") or _now().date()
    if st.session_state.pop("_req_kom_refresh", False):
        _ruta_r = sb_ruta_objekti(_datum_r, _ku)
        _ids_r = [int(r.get("idk") or 0) for r in _ruta_r]
        if not _ids_r:
            st.warning("Ruta za " + _rok_fmt(str(_datum_r)) + " je prazna — nema šta da se ažurira.")
        else:
            with st.spinner("Povlačim porudžbine iz admina za sve objekte u ruti (" + str(len(_ids_r))
                            + ")… može potrajati minut."):
                try:
                    _cut_r = _now().date() - datetime.timedelta(days=190)
                    _map_r = {_i: _naz(_i) for _i in _ids_r}
                    _bulk_r, _err_r = admin_istorija_bulk(_map_r, _cut_r)
                except Exception as _e:
                    _bulk_r, _err_r = {}, str(_e)
            if _err_r and not _bulk_r:
                st.error("Ažuriranje nije uspelo: " + str(_err_r))
            else:
                _at_r = _now().isoformat()
                for _i in _ids_r:
                    _lst_r = sorted(_bulk_r.get(_i, []) or [], key=_datum_sort_key, reverse=True)
                    st.session_state["komhist_" + str(_i)] = {"lst": _lst_r, "err": "", "at": _at_r}
                st.session_state["kom_hist_at"] = _at_r
                st.success("✅ Ažurirano iz admina za sve objekte u ruti (" + str(len(_ids_r)) + ").")

    if st.session_state.get("kom_hist_at"):
        st.markdown('<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:8px 14px;'
                    'font-size:12.5px;color:#166534;font-weight:600;margin:2px 0 12px;">🔄 Podaci iz admina '
                    'poslednji put ažurirani: ' + _h_escape(_dt_kratko(st.session_state["kom_hist_at"]))
                    + '  ·  važi za sve objekte u ruti.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:10px;padding:8px 14px;'
                    'font-size:12.5px;color:#92400e;font-weight:600;margin:2px 0 12px;">🔄 Prethodne porudžbine '
                    'još nisu povučene — klikni „Ažuriraj iz admina“ gore desno (povlači odjednom za celu rutu).</div>',
                    unsafe_allow_html=True)

    def _kontakt_html(idk):
        _ki = _kf.get(int(idk), {}) or {}
        _b = []
        if _ki.get("telefon"): _b.append("📞 " + _h_escape(str(_ki["telefon"])))
        if _ki.get("email"): _b.append("✉️ " + _h_escape(str(_ki["email"])))
        if _ki.get("mesto"): _b.append("📍 " + _h_escape(str(_ki["mesto"])))
        return "&nbsp;&nbsp;·&nbsp;&nbsp;".join(_b)

    # --- Ko je prosleđen komercijali (za sve mesece koje gledamo) ---
    def _prosledjeni(mesec_key):
        """{idk: {sistem, ko, at, napomena, nedeljni}} — objekti prosleđeni komercijali."""
        out = {}
        for _sis in sb_sisteme(mesec_key):
            _pod = sb_ucitaj(mesec_key, _sis)
            if not _pod or not _pod.get("stavke"):
                continue
            _meta = _pod.get("meta") or {}
            if _meta.get("nedeljni"):
                for _k, _info in (_meta.get("nedeljni_prijave") or {}).items():
                    try:
                        out[int(_k)] = {"sistem": _sis, "ko": (_info or {}).get("ko", ""),
                                        "at": (_info or {}).get("at", ""),
                                        "napomena": (_info or {}).get("napomena", ""), "nedeljni": True}
                    except Exception:
                        pass
                continue
            for _idk, _v in sb_load_obrada(mesec_key, _sis).items():
                if "Obavestila direktorku" not in (_v.get("reakcije") or []):
                    continue
                _dn = (_v.get("dnevnik") or {}).get("komercijala") or []
                out[int(_idk)] = {"sistem": _sis,
                                  "ko": (_v.get("reakcije_ko") or {}).get("Obavestila direktorku", ""),
                                  "at": (_dn[-1].get("at", "") if _dn else ""),
                                  "napomena": _v.get("napomena", "") or "", "nedeljni": False}
        return out

    def _stavke_objekta(mesec_key, sistem, idk):
        _pod = sb_ucitaj(mesec_key, sistem)
        if not _pod or not _pod.get("stavke"):
            return [], {}
        _lst = [s for s in _pod["stavke"] if int(s.get("idk", -1)) == int(idk)]
        return _lst, (_pod.get("meta") or {})

    def _sistem_objekta(mesec_key, idk):
        for _sis in sb_sisteme(mesec_key):
            _pod = sb_ucitaj(mesec_key, _sis)
            if not _pod or not _pod.get("stavke"):
                continue
            for s in _pod["stavke"]:
                if int(s.get("idk", -1)) == int(idk):
                    return _sis
        return ""

    _tab1, _tab2, _tab3 = st.tabs(["🚚 Moja ruta", "📋 Objekat — popis i porudžbina",
                                   "✅ Kontrola od administracije"])

    # =================================================================
    # KARTICA 1 — MOJA RUTA (DAN)
    # =================================================================
    with _tab1:
        _d1, _d2, _d3, _d4 = st.columns([1.1, 0.9, 0.9, 1.4])
        with _d1:
            _datum = st.date_input("Datum rute", value=_now().date(), key="kom_datum", format="DD.MM.YYYY")
        _dan = sb_ruta_dan_get(_datum, _ku)
        with _d2:
            _st_txt = _dt_kratko(_dan.get("start_at", "")) if _dan.get("start_at") else "—"
            st.markdown('<div style="font-size:12.5px;color:#6b7280;margin:6px 0 2px;">Start rute</div>'
                        '<div style="font-weight:700;font-size:15px;">' + _h_escape(_st_txt.split(" ")[-1] if _st_txt != "—" else "—") + '</div>',
                        unsafe_allow_html=True)
            if not _dan.get("start_at"):
                if st.button("▶ Startuj rutu", key="kom_start", use_container_width=True):
                    try:
                        sb_ruta_dan_set(_datum, _ku, start_at=_now().isoformat())
                        st.rerun()
                    except Exception as _e:
                        st.error(str(_e))
        with _d3:
            _kr_txt = _dt_kratko(_dan.get("kraj_at", "")) if _dan.get("kraj_at") else "—"
            st.markdown('<div style="font-size:12.5px;color:#6b7280;margin:6px 0 2px;">Završetak</div>'
                        '<div style="font-weight:700;font-size:15px;">' + _h_escape(_kr_txt.split(" ")[-1] if _kr_txt != "—" else "—") + '</div>',
                        unsafe_allow_html=True)
            if _dan.get("start_at") and not _dan.get("kraj_at"):
                if st.button("⏹ Završi rutu", key="kom_kraj", use_container_width=True):
                    try:
                        sb_ruta_dan_set(_datum, _ku, kraj_at=_now().isoformat())
                        st.rerun()
                    except Exception as _e:
                        st.error(str(_e))
        with _d4:
            if _dan.get("kraj_at"):
                st.markdown('<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:9px 13px;'
                            'font-size:12.5px;color:#166534;font-weight:600;margin-top:22px;">✅ Ruta završena</div>',
                            unsafe_allow_html=True)
            elif _dan.get("start_at"):
                st.markdown('<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:10px;padding:9px 13px;'
                            'font-size:12.5px;color:#92400e;font-weight:600;margin-top:22px;">🚚 Ruta u toku</div>',
                            unsafe_allow_html=True)

        _ruta = sb_ruta_objekti(_datum, _ku)
        _u_ruti = set(int(r.get("idk") or 0) for r in _ruta)

        # --- Dodaj objekat u rutu ---
        st.markdown("<div style='font-weight:700;font-size:14px;margin:14px 0 4px;'>Dodaj objekat u rutu</div>",
                    unsafe_allow_html=True)
        _a1, _a2, _a3 = st.columns([1.2, 3, 1.1])
        with _a1:
            _sis_opts = ["(svi sistemi)"] + list(sb_sisteme(_mes_now))
            _a_sis = st.selectbox("Sistem", _sis_opts, key="kom_add_sis")
        with _a2:
            _kand = []
            for _s in ([_a_sis] if _a_sis != "(svi sistemi)" else sb_sisteme(_mes_now)):
                _pod = sb_ucitaj(_mes_now, _s)
                if not _pod or not _pod.get("stavke"):
                    continue
                for _idk in sorted({int(x["idk"]) for x in _pod["stavke"]}):
                    if _idk not in _u_ruti:
                        _kand.append((_idk, _s, _naz(_idk)))
            _kand.sort(key=lambda x: x[2])
            _opts = ["— izaberi objekat —"] + [x[2] for x in _kand]
            _pick = st.selectbox("Objekat", _opts, key="kom_add_obj")
        with _a3:
            st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
            if st.button("+ Dodaj u rutu", key="kom_add_btn", type="primary", use_container_width=True,
                         disabled=(_pick == "— izaberi objekat —")):
                _sel = next((x for x in _kand if x[2] == _pick), None)
                if _sel:
                    try:
                        _pr = _prosledjeni(_mes_now).get(int(_sel[0]), {})
                        sb_ruta_obj_save(_datum, _ku, _sel[0], mesec=_mes_now, sistem=_sel[1],
                                         status="nov", prosledjen_od=_pr.get("ko", ""))
                        st.success("Dodato u rutu: " + str(_sel[2]))
                        st.rerun()
                    except Exception as _e:
                        st.error("Greška: " + str(_e))

        # --- KPI + lista rute ---
        _n_zav = sum(1 for r in _ruta if r.get("status") == "zavrsen")
        _n_tok = sum(1 for r in _ruta if r.get("status") == "u_toku")
        _n_nov = len(_ruta) - _n_zav - _n_tok
        st.markdown(_kpi([(len(_ruta), "Objekata u ruti") + _P, (_n_zav, "Završeno") + _G,
                          (_n_tok, "U toku") + _O, (_n_nov, "Nije obrađeno") + _R]),
                    unsafe_allow_html=True)

        if not _ruta:
            st.info("Ruta za ovaj dan je prazna — dodaj objekte gore.")
        else:
            _prosl_now = _prosledjeni(_mes_now)
            _ruta.sort(key=lambda r: (r.get("status") == "zavrsen", _naz(r.get("idk"))))
            for r in _ruta:
                _idk = int(r.get("idk") or 0)
                _stat = r.get("status") or "nov"
                _chip = ('<span class="kchip k-grn">✓ Završeno</span>' if _stat == "zavrsen"
                         else ('<span class="kchip k-org">U toku</span>' if _stat == "u_toku"
                               else '<span class="kchip k-gry">Nije obrađeno</span>'))
                _extra = ""
                if r.get("trebovao"):
                    _extra = ('<span class="kchip k-grn" style="margin-left:6px;">porudžbina '
                              + _kfmt(r.get("poslato_kom", 0)) + ' kom</span>')
                if _idk in _prosl_now:
                    _extra += '<span class="kchip k-vio" style="margin-left:6px;">prosleđen od administracije</span>'
                _c1, _c2 = st.columns([6.2, 1.3])
                with _c1:
                    st.markdown('<div class="kom-card" style="margin-bottom:4px;">' + _chip + _extra
                                + '<div class="kom-nm" style="margin-top:6px;">' + _h_escape(_naz(_idk)) + '</div>'
                                + '<div class="kom-ad">' + (_kontakt_html(_idk) or str(r.get("sistem") or "")) + '</div>'
                                + '</div>', unsafe_allow_html=True)
                with _c2:
                    st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)
                    if st.button("Otvori", key="kom_open_" + str(_idk), use_container_width=True):
                        st.session_state["kom_sel_idk"] = _idk
                        st.session_state["kom_sel_datum"] = str(_datum)
                        st.rerun()
                    if _stat != "zavrsen":
                        if st.button("Ukloni", key="kom_del_" + str(_idk), use_container_width=True):
                            sb_ruta_obj_del(_datum, _ku, _idk)
                            st.rerun()
            if st.session_state.get("kom_sel_idk"):
                st.success('Objekat je otvoren — pređi na karticu „Objekat — popis i porudžbina“.')

    # =================================================================
    # KARTICA 2 — OBJEKAT (POPIS I PORUDŽBINA)
    # =================================================================
    with _tab2:
        _datum2 = st.session_state.get("kom_sel_datum") or str(_now().date())
        _ruta2 = sb_ruta_objekti(_datum2, _ku)
        if not _ruta2:
            st.info('Nema objekata u ruti za ' + _rok_fmt(_datum2) + '. Dodaj ih na kartici „Moja ruta“.')
        else:
            _ids = [int(r.get("idk") or 0) for r in _ruta2]
            _labels = [_naz(i) for i in _ids]
            _sel_id = st.session_state.get("kom_sel_idk")
            _idx = _ids.index(_sel_id) if _sel_id in _ids else 0
            _pick2 = st.selectbox("Objekat iz rute (" + _rok_fmt(_datum2) + ")", _labels, index=_idx, key="kom_obj_pick")
            _oid = _ids[_labels.index(_pick2)]
            _rec = next((r for r in _ruta2 if int(r.get("idk") or 0) == _oid), {})
            _zakljucano = (_rec.get("status") == "zavrsen")
            _mes_o = _rec.get("mesec") or _mes_now
            _sis_o = _rec.get("sistem") or _sistem_objekta(_mes_o, _oid)
            _stavke_o, _meta_o = _stavke_objekta(_mes_o, _sis_o, _oid)
            _mes_par = float(_meta_o.get("meseci") or 1.0)

            st.markdown('<div class="kom-card"><div class="kom-nm">' + _h_escape(_naz(_oid)) + '</div>'
                        '<div class="kom-ad">' + (_kontakt_html(_oid) or "") + '</div>'
                        '<div class="kom-ad">' + _h_escape(str(_sis_o)) + ' · ruta ' + _h_escape(_rok_fmt(_datum2))
                        + ' · ' + _h_escape(_ku) + '</div></div>', unsafe_allow_html=True)

            if _zakljucano:
                st.markdown('<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:12px;padding:11px 16px;'
                            'font-size:13px;color:#166534;font-weight:600;margin-bottom:12px;">🔒 Izveštaj je završen '
                            + (("(" + _dt_kratko(_rec.get("zavrseno_at", "")) + ")") if _rec.get("zavrseno_at") else "")
                            + " — izmene više nisu moguće.</div>", unsafe_allow_html=True)

            # --- Prethodne porudžbine iz admina (povlače se gore, za celu rutu odjednom) ---
            _hk = "komhist_" + str(_oid)
            _hh = st.session_state.get(_hk) or {}

            # --- Crveno upozorenje: trebovao POSLE prosleđivanja ---
            _pr_info = _prosledjeni(_mes_o).get(int(_oid), {})
            _pr_at = _pr_info.get("at") or ""
            if _pr_at and _hh.get("lst"):
                _pd = None
                try:
                    _pd = datetime.datetime.fromisoformat(str(_pr_at)).date()
                except Exception:
                    try:
                        _pd = datetime.datetime.strptime(str(_pr_at)[:10], "%d.%m.%Y").date()
                    except Exception:
                        _pd = None
                _posle = []
                if _pd:
                    for _o in (_hh.get("lst") or []):
                        try:
                            _dd = datetime.datetime.strptime(str(_o.get("datum", "")).split(" ")[0], "%d.%m.%Y").date()
                        except Exception:
                            continue
                        if _dd >= _pd:
                            _posle.append(_o)
                if _posle:
                    _det = "; ".join(("#" + str(_o.get("id", "")) + " od " + str(_o.get("datum", "")).split(" ")[0]
                                      + (" · " + str(_o.get("status", "")) if _o.get("status") else ""))
                                     for _o in _posle[:3])
                    st.markdown('<div style="background:#fef2f2;border:1.6px solid #fca5a5;border-radius:12px;'
                                'padding:12px 16px;font-size:13.5px;color:#b42318;font-weight:700;margin-bottom:12px;">'
                                '⚠️ Molimo vas obratite pažnju — objekat je trebovao robu <u>nakon</u> što vam je '
                                'prosleđen na pregled!'
                                '<div style="font-weight:500;font-size:12.5px;margin-top:5px;">' + _h_escape(_det)
                                + '</div></div>', unsafe_allow_html=True)

            # --- Popis lagera ---
            _bez = st.checkbox("Objekat ne dozvoljava popis lagera", value=bool(_rec.get("bez_popisa")),
                               key="kom_bez_" + str(_oid), disabled=_zakljucano)
            _popis_saved = dict(_rec.get("popis") or {})
            _por_saved = dict(_rec.get("porudzbina") or {})

            # već poručeno posle preseka (za varijantu bez popisa — kao kod administracije)
            _cut_o = _admin_presek(_meta_o, _mes_o)
            _pm_o = {}
            if _hh.get("lst") and _cut_o:
                _pm_o = _treb_posle_preseka(_hh.get("lst") or [], _cut_o)
            elif isinstance(_meta_o, dict):
                _ah = (_meta_o.get("admin_hist") or {}).get(str(int(_oid))) or []
                if _ah and _cut_o:
                    _pm_o = _treb_posle_preseka(_ah, _cut_o)

            _arts = [s for s in _stavke_o if int(s.get("pred", 0) or 0) > 0 or int(s.get("kol", 0) or 0) > 0]
            _arts.sort(key=lambda s: (-int(s.get("pred", 0) or 0)))
            if not _arts:
                st.warning("Za ovaj objekat nema artikala u izveštaju za " + mesec_label(_mes_o) + ".")
            else:
                _pop_col = "✏️ POPIS — upiši lager"
                _pred_col = "Predikcija (" + str(_mes_par).replace(".", ",") + " mes)"
                _edkey = "kom_ed_" + str(_oid) + "_" + str(_datum2) + ("_b" if _bez else "")
                _stk = "komst_" + str(_oid) + "_" + str(_datum2)
                _prev_pop = dict(st.session_state.get(_stk) or {})
                # izmene koje je korisnik upravo uneo u tabelu
                _edst = st.session_state.get(_edkey)
                _edrows = (_edst or {}).get("edited_rows", {}) if isinstance(_edst, dict) else {}

                if _bez:
                    # objekat ne dozvoljava popis — predlog = dodatna porudžbina (kao kod administracije)
                    st.markdown('<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:10px;'
                                'padding:9px 14px;font-size:12.5px;color:#92400e;margin:4px 0 8px;">Popis nije '
                                'moguć — predlog je <b>dodatna porudžbina</b> (preporuka umanjena za ono što su '
                                'već poručili posle preseka).</div>', unsafe_allow_html=True)
                    _rows_e = []
                    for s in _arts:
                        _ida = int(s.get("ida", 0))
                        _por_ranije = int(_pm_o.get(_ida, 0) or 0)
                        _rows_e.append({"ida": _ida, "Artikal": str(s.get("naziv", "")),
                                        _pred_col: int(round(int(s.get("pred", 0) or 0) * _mes_par)),
                                        "Predlog porudžbine": int(_por_saved.get(
                                            str(_ida), max(int(s.get("kol", 0) or 0) - _por_ranije, 0)) or 0)})
                    _dfe = pd.DataFrame(_rows_e)
                    _ed = st.data_editor(
                        _dfe.drop(columns=["ida"]), key=_edkey, hide_index=True, use_container_width=True,
                        disabled=(["Artikal", _pred_col] if not _zakljucano else True),
                        column_config={
                            "Artikal": st.column_config.TextColumn("Naziv artikla", width="large"),
                            _pred_col: st.column_config.NumberColumn(_pred_col, disabled=True),
                            "Predlog porudžbine": st.column_config.NumberColumn("Predlog porudžbine",
                                                                                min_value=0, step=1),
                        })
                    _popis_new, _por_new = {}, {}
                    for _i, _r in enumerate(_rows_e):
                        try:
                            _por_new[str(_r["ida"])] = int(_ed.iloc[_i]["Predlog porudžbine"] or 0)
                        except Exception:
                            _por_new[str(_r["ida"])] = int(_r["Predlog porudžbine"])
                    _n_nula = 0
                else:
                    st.markdown('<div style="background:#f5f3ff;border:1.6px solid #c4b5fd;border-radius:10px;'
                                'padding:10px 14px;font-size:13px;color:#5b21b6;margin:4px 0 8px;font-weight:600;">'
                                '✏️ Popuni samo kolonu <b>POPIS — upiši lager</b> (stanje koje si prebrojao/la na '
                                'polici). Predlog porudžbine se sam računa: <b>predikcija − popis</b>, i menja se '
                                'čim promeniš popis. Predlog možeš i ručno da prepraviš.</div>',
                                unsafe_allow_html=True)
                    _rows_e = []
                    for _i, s in enumerate(_arts):
                        _ida = int(s.get("ida", 0))
                        _pred_per = int(round(int(s.get("pred", 0) or 0) * _mes_par))
                        # popis: prvo ono što je upravo uneto, pa sačuvano, inače PRAZNO
                        _ed_r = _edrows.get(_i) or _edrows.get(str(_i)) or {}
                        _pop_v = None
                        _pop_edit = False
                        if _pop_col in _ed_r:
                            _pop_v = _ed_r.get(_pop_col)
                            _pop_edit = True
                        elif str(_ida) in _popis_saved and _popis_saved.get(str(_ida)) is not None:
                            _pop_v = _popis_saved.get(str(_ida))
                        try:
                            _pop_v = int(_pop_v) if _pop_v is not None and str(_pop_v) != "" else None
                        except Exception:
                            _pop_v = None
                        # predlog: automatski (predikcija − popis); ručna izmena se poštuje
                        # dok se popis ne promeni
                        _auto = (max(_pred_per - _pop_v, 0) if _pop_v is not None else None)
                        _pop_promenjen = (_prev_pop.get(str(_ida), "__x__") != _pop_v)
                        _man = None
                        if not _pop_promenjen:
                            if "Predlog porudžbine" in _ed_r:
                                _man = _ed_r.get("Predlog porudžbine")
                            elif str(_ida) in _por_saved:
                                _man = _por_saved.get(str(_ida))
                        elif ("Predlog porudžbine" in _ed_r) and not _pop_edit:
                            _man = _ed_r.get("Predlog porudžbine")
                        try:
                            _man = int(_man) if _man is not None and str(_man) != "" else None
                        except Exception:
                            _man = None
                        _rows_e.append({"ida": _ida, "Artikal": str(s.get("naziv", "")),
                                        _pop_col: _pop_v, _pred_col: _pred_per,
                                        "Predlog porudžbine": (_man if _man is not None else _auto)})
                    _dfe = pd.DataFrame(_rows_e)
                    for _c in (_pop_col, "Predlog porudžbine"):
                        _dfe[_c] = _dfe[_c].astype("Int64")
                    _ed = st.data_editor(
                        _dfe.drop(columns=["ida"]), key=_edkey, hide_index=True, use_container_width=True,
                        disabled=(["Artikal", _pred_col] if not _zakljucano else True),
                        column_config={
                            "Artikal": st.column_config.TextColumn("Naziv artikla", width="large"),
                            _pop_col: st.column_config.NumberColumn(
                                _pop_col, min_value=0, step=1,
                                help="Stanje prebrojano u objektu — jedina kolona koju popunjavaš."),
                            _pred_col: st.column_config.NumberColumn(_pred_col, disabled=True),
                            "Predlog porudžbine": st.column_config.NumberColumn(
                                "Predlog porudžbine", min_value=0, step=1,
                                help="Automatski: predikcija − popis. Možeš ručno da promeniš."),
                        })
                    _popis_new, _por_new, _novi_prev = {}, {}, {}
                    for _i, _r in enumerate(_rows_e):
                        _k = str(_r["ida"])
                        try:
                            _pv = _ed.iloc[_i][_pop_col]
                            _pv = None if pd.isna(_pv) else int(_pv)
                        except Exception:
                            _pv = _r[_pop_col]
                        try:
                            _ov = _ed.iloc[_i]["Predlog porudžbine"]
                            _ov = 0 if pd.isna(_ov) else int(_ov)
                        except Exception:
                            _ov = int(_r["Predlog porudžbine"] or 0)
                        if _pv is not None:
                            _popis_new[_k] = _pv
                        _por_new[_k] = _ov
                        _novi_prev[_k] = _pv
                    st.session_state[_stk] = _novi_prev
                    _n_nula = sum(1 for _v in _popis_new.values() if int(_v) == 0)
                    _n_bez = sum(1 for _r in _rows_e if _r[_pop_col] is None)
                    if _n_bez:
                        st.caption("ℹ️ Još nije popisano artikala: " + str(_n_bez)
                                   + " (predlog se pojavi čim upišeš popis).")
                _uk_kom = sum(int(v or 0) for v in _por_new.values())
                _n_art = sum(1 for x in _por_new.values() if int(x or 0) > 0)
                st.markdown(_kpi([(_kfmt(_uk_kom) + " kom", "Ukupno predlog") + _P,
                                  (_n_art, "Artikala u porudžbini") + _G,
                                  (_n_nula, "Artikala na nuli") + _R]), unsafe_allow_html=True)
                if not _zakljucano:
                    if st.button("💾 Sačuvaj popis", key="kom_save_" + str(_oid), use_container_width=False):
                        try:
                            sb_ruta_obj_save(_datum2, _ku, _oid, mesec=_mes_o, sistem=_sis_o,
                                             status="u_toku", bez_popisa=bool(_bez),
                                             popis=_popis_new, porudzbina=_por_new)
                            st.success("Popis sačuvan.")
                            st.rerun()
                        except Exception as _e:
                            st.error("Greška: " + str(_e))

            # --- Prethodne porudžbine iz admina ---
            with st.expander("📜 Prethodne porudžbine iz admina (poslednjih ~6 meseci)"):
                if _hh.get("err") and not _hh.get("lst"):
                    st.warning(_hh["err"])
                elif not _hh.get("lst"):
                    st.caption('Klikni „Ažuriraj iz admina“ gore desno — povlači odjednom za sve objekte u ruti.')
                else:
                    _hr = []
                    for _o in (_hh.get("lst") or []):
                        _hr.append({"Datum": str(_o.get("datum", "")), "Broj": str(_o.get("id", "")),
                                    "Status": str(_o.get("status", "")),
                                    "Kom": int(sum(_to_int_kol(x.get("kol")) for x in (_o.get("stavke") or []))),
                                    "Vrednost": str(_o.get("cena", ""))})
                    st.dataframe(pd.DataFrame(_hr), hide_index=True, use_container_width=True)

            # --- Slike i komentar ---
            _slike = list(_rec.get("slike") or [])
            st.markdown("<div style='font-weight:700;font-size:14px;margin:14px 0 4px;'>Slike</div>",
                        unsafe_allow_html=True)
            if _slike:
                _sc = st.columns(min(len(_slike), 4))
                for _i, _sl in enumerate(_slike[:4]):
                    with _sc[_i]:
                        st.markdown('<div style="font-size:11.5px;color:#9aa0ad;font-weight:700;">'
                                    + ("PRE" if _sl.get("tip") == "pre" else "POSLE") + '</div>',
                                    unsafe_allow_html=True)
                        if _sl.get("url"):
                            st.image(_sl["url"], use_container_width=True)
            if not _zakljucano:
                _f1, _f2 = st.columns(2)
                with _f1:
                    _up_pre = st.file_uploader("Slike PRE obilaska (do 2)", type=["jpg", "jpeg", "png"],
                                               accept_multiple_files=True, key="kom_pre_" + str(_oid))
                with _f2:
                    _up_pos = st.file_uploader("Slike NAKON obilaska (do 2)", type=["jpg", "jpeg", "png"],
                                               accept_multiple_files=True, key="kom_pos_" + str(_oid))
                if st.button("📤 Otpremi slike", key="kom_upl_" + str(_oid)):
                    _novi = list(_slike)
                    try:
                        for _tip, _files in (("pre", _up_pre or []), ("posle", _up_pos or [])):
                            for _f in list(_files)[:2]:
                                _b = _smanji_sliku(_f.getvalue())
                                _nm = (str(_oid) + "_" + str(_datum2) + "_" + _tip + "_"
                                       + _now().strftime("%H%M%S") + "_" + str(len(_novi)) + ".jpg")
                                _url = sb_ruta_slika_upload(_b, _nm)
                                _novi.append({"tip": _tip, "url": _url, "naziv": _f.name})
                        sb_ruta_obj_save(_datum2, _ku, _oid, mesec=_mes_o, sistem=_sis_o, slike=_novi)
                        st.success("Slike otpremljene.")
                        st.rerun()
                    except Exception as _e:
                        st.error("Greška pri otpremanju: " + str(_e))
                        if RUTA_BUCKET in str(_e) or "bucket" in str(_e).lower():
                            st.markdown('<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:12px;'
                                        'padding:13px 17px;font-size:13px;color:#78500a;line-height:1.6;">'
                                        '<b>Kako da se ovo reši (jednom, 3 klika):</b><br>'
                                        '1. Supabase → <b>Storage</b> (levo u meniju)<br>'
                                        '2. <b>New bucket</b> → ime: <b>ruta-slike</b><br>'
                                        '3. Čekiraj <b>Public bucket</b> → <b>Save</b><br>'
                                        'Zatim se vrati ovde i klikni Otpremi slike ponovo.</div>',
                                        unsafe_allow_html=True)

            _kom_txt = st.text_area("Komentar sa terena", value=_rec.get("komentar", "") or "",
                                    key="kom_txt_" + str(_oid), height=90, disabled=_zakljucano,
                                    placeholder="npr. gondola prazna, obećali da poručuju u ponedeljak…")

            # --- Akcije: prosledi u admin + završeno ---
            if not _zakljucano:
                _b1, _b2, _b3 = st.columns([1.6, 1.4, 3])
                with _b1:
                    _por_map = {}
                    try:
                        _por_map = _por_new
                    except Exception:
                        _por_map = dict(_rec.get("porudzbina") or {})
                    _uk_send = sum(int(v or 0) for v in _por_map.values())
                    if st.button("📤 Prosledi porudžbinu u admin", key="kom_send_" + str(_oid),
                                 type="primary", use_container_width=True, disabled=(_uk_send <= 0)):
                        _items = [{"idArticle": int(k), "quantity": int(v)} for k, v in _por_map.items() if int(v or 0) > 0]
                        with st.spinner("Šaljem porudžbinu u admin…"):
                            _ok, _msg = posalji_u_admin(int(_oid), _items)
                        if _ok:
                            try:
                                sb_ruta_obj_save(_datum2, _ku, _oid, mesec=_mes_o, sistem=_sis_o,
                                                 status="u_toku", porudzbina=_por_map, komentar=_kom_txt,
                                                 trebovao=True, poslato_kom=int(_uk_send),
                                                 poslato_admin_at=_now().isoformat(), admin_broj=str(_msg))
                            except Exception:
                                pass
                            st.success("✅ Porudžbina je uspešno prosleđena u admin — "
                                       + _kfmt(_uk_send) + " kom. " + str(_msg))
                            st.rerun()
                        else:
                            st.error("❌ " + str(_msg))
                with _b2:
                    if st.button("✓ Završeno", key="kom_done_" + str(_oid), use_container_width=True):
                        try:
                            sb_ruta_obj_save(_datum2, _ku, _oid, mesec=_mes_o, sistem=_sis_o,
                                             status="zavrsen", komentar=_kom_txt,
                                             zavrseno_at=_now().isoformat())
                            st.success("Izveštaj za ovaj objekat je završen i zaključan.")
                            st.rerun()
                        except Exception as _e:
                            st.error("Greška: " + str(_e))
                with _b3:
                    st.caption("Kada klikneš Završeno, izveštaj za ovaj objekat se zaključava i više se ne menja.")
                if _kom_txt != (_rec.get("komentar", "") or ""):
                    if st.button("💾 Sačuvaj komentar", key="kom_savec_" + str(_oid)):
                        sb_ruta_obj_save(_datum2, _ku, _oid, mesec=_mes_o, sistem=_sis_o, komentar=_kom_txt)
                        st.rerun()

    # =================================================================
    # KARTICA 3 — KONTROLA OD ADMINISTRACIJE
    # =================================================================
    with _tab3:
        _k1, _k2, _k3 = st.columns([1.1, 1.1, 2])
        _mlbls = [mesec_label(k) for k in _mkeys]
        with _k1:
            _msel = st.selectbox("Mesec", _mlbls, index=0, key="kom_kmes")
        _mk = _mkeys[_mlbls.index(_msel)]
        with _k2:
            _fsis = st.selectbox("Sistem", ["(svi sistemi)"] + list(sb_sisteme(_mk)), key="kom_ksis")
        _rok_k = sb_rokovi_get(_mk).get("rok_kontrola")
        _rok_proso = _rok_je_prosao(_rok_k) if _rok_k else False
        _zamrznuto = _rok_proso or (_mk != _mes_now)
        with _k3:
            if _rok_k:
                _bg = ("#fef2f2;border-color:#fecaca;color:#b42318" if _rok_proso
                       else "#f0fdf4;border-color:#bbf7d0;color:#166534")
                _t = ("Rok za kontrolu je istekao (" + _rok_fmt(_rok_k) + ") — mesec je zatvoren."
                      if _rok_proso else "Rok za kontrolu: " + _rok_fmt(_rok_k))
                st.markdown('<div style="background:' + _bg + ';border:1px solid;border-radius:10px;padding:9px 14px;'
                            'font-size:12.5px;font-weight:600;margin-top:22px;">⏰ ' + _t + '</div>',
                            unsafe_allow_html=True)
            else:
                st.markdown('<div style="background:#f5f3ff;border:1px solid #ddd6fe;border-radius:10px;padding:9px 14px;'
                            'font-size:12.5px;color:#5b21b6;margin-top:22px;">Rok za kontrolu još nije postavljen '
                            '(postavlja direktor).</div>', unsafe_allow_html=True)

        _prosl = _prosledjeni(_mk)
        _obil = sb_ruta_mesec(_mk)
        _redovi = []
        for _idk, _info in _prosl.items():
            if _fsis != "(svi sistemi)" and str(_info.get("sistem", "")) != _fsis:
                continue
            _ob = _obil.get(int(_idk))
            if _ob and _ob.get("status") == "zavrsen":
                _st = "zavrsen"
            elif _ob:
                _st = "u_ruti"
            else:
                _st = "neobradjen" if _zamrznuto else "ceka"
            _redovi.append({"idk": int(_idk), "info": _info, "ob": _ob or {}, "st": _st})

        _n_z = sum(1 for r in _redovi if r["st"] == "zavrsen")
        _n_c = sum(1 for r in _redovi if r["st"] in ("ceka", "u_ruti"))
        _n_n = sum(1 for r in _redovi if r["st"] == "neobradjen")
        st.markdown(_kpi([(len(_redovi), "Prosleđeno komercijali") + _P,
                          (_n_z, "Završeno (bili u ruti)") + _G,
                          (_n_c, "Čeka obilazak") + _O,
                          (_n_n, "Neobrađeno (rok istekao)") + _R]), unsafe_allow_html=True)

        if _zamrznuto:
            st.markdown('<div style="background:#fef2f2;border:1px solid #fca5a5;border-radius:12px;padding:11px 16px;'
                        'font-size:13px;color:#b42318;margin-bottom:12px;">🔒 <b>' + _h_escape(_msel)
                        + ' je zatvoren</b> — prikaz je samo za pregled. Objekti se više ne mogu dodavati u rutu, '
                        'a ono što nije obiđeno do roka trajno stoji kao <b>Neobrađeno</b>.</div>',
                        unsafe_allow_html=True)

        if not _redovi:
            st.success("Nema objekata prosleđenih komercijali za " + _msel
                       + ("" if _fsis == "(svi sistemi)" else " (sistem: " + _fsis + ")") + ".")
        else:
            _rang = {"neobradjen": 0, "ceka": 1, "u_ruti": 2, "zavrsen": 3}
            _redovi.sort(key=lambda r: (_rang.get(r["st"], 9), _naz(r["idk"])))
            for r in _redovi:
                _idk = r["idk"]; _info = r["info"]; _ob = r["ob"]
                _chip = {"zavrsen": '<span class="kchip k-grn">✓ Završeno</span>',
                         "u_ruti": '<span class="kchip k-org">U ruti (nije završen)</span>',
                         "ceka": '<span class="kchip k-org">⏳ Čeka obilazak</span>',
                         "neobradjen": '<span class="kchip k-red">✗ Neobrađeno</span>'}[r["st"]]
                if _zamrznuto:
                    _chip = _chip.replace("k-grn", "k-gry").replace("k-org", "k-gry").replace("k-red", "k-gry")
                _line2 = "Prosledila: " + _h_escape(str(_info.get("ko", "") or "administracija"))
                if _info.get("at"):
                    _line2 += " · " + _h_escape(_dt_kratko(_info["at"]) or str(_info["at"]))
                if _ob.get("datum"):
                    _line2 += " · obiđeno " + _h_escape(_rok_fmt(str(_ob.get("datum")))) + " (" + _h_escape(str(_ob.get("ko", ""))) + ")"
                elif r["st"] == "neobradjen":
                    _line2 += " · nije bio u ruti do roka" + ((" (" + _rok_fmt(_rok_k) + ")") if _rok_k else "")
                else:
                    _line2 += " · još nije bio u ruti"
                _bad = ""
                if _ob:
                    if _ob.get("trebovao"):
                        _bad += ('<span class="kchip k-grn">✓ Trebovao — ' + _kfmt(_ob.get("poslato_kom", 0))
                                 + ' kom, prosleđeno u admin</span> ')
                    elif _ob.get("status") == "zavrsen":
                        _bad += '<span class="kchip k-red">✗ Nije trebovao</span> '
                    if _ob.get("bez_popisa"):
                        _bad += '<span class="kchip k-org">objekat ne dozvoljava popis</span> '
                _nap_h = ""
                if _info.get("napomena"):
                    _nap_h = ('<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:9px;padding:8px 12px;'
                              'font-size:12.5px;color:#78500a;margin-top:9px;"><b style="color:#4b5563;">Napomena administracije:</b> '
                              + _h_escape(_info["napomena"]) + '</div>')
                if _ob.get("komentar"):
                    _nap_h += ('<div style="background:#faf9fd;border:1px solid #efeaf7;border-radius:9px;padding:8px 12px;'
                               'font-size:12.5px;color:#4b5563;margin-top:7px;"><b>Komentar sa terena:</b> '
                               + _h_escape(_ob["komentar"]) + '</div>')
                _cc1, _cc2 = st.columns([6.2, 1.3])
                with _cc1:
                    st.markdown('<div class="kom-card" style="' + ("opacity:.75;" if _zamrznuto else "") + '">'
                                + _chip + '<div class="kom-nm" style="margin-top:6px;">' + _h_escape(_naz(_idk)) + '</div>'
                                + '<div class="kom-ad">' + _h_escape(str(_info.get("sistem", ""))) + '</div>'
                                + '<div class="kom-ad">' + _line2 + '</div>'
                                + (('<div style="margin-top:8px;">' + _bad + '</div>') if _bad else "")
                                + _nap_h + '</div>', unsafe_allow_html=True)
                with _cc2:
                    if (not _zamrznuto) and (r["st"] == "ceka"):
                        st.markdown("<div style='height:22px;'></div>", unsafe_allow_html=True)
                        if st.button("+ U rutu", key="kom_k2r_" + str(_idk), use_container_width=True):
                            try:
                                sb_ruta_obj_save(_now().date(), _ku, _idk, mesec=_mk,
                                                 sistem=_info.get("sistem", ""), status="nov",
                                                 prosledjen_od=_info.get("ko", ""))
                                st.session_state["kom_sel_idk"] = _idk
                                st.session_state["kom_sel_datum"] = str(_now().date())
                                st.success("Dodato u današnju rutu.")
                                st.rerun()
                            except Exception as _e:
                                st.error("Greška: " + str(_e))


def prikazi_direktore():
    st.set_page_config(page_title="VAPE — Direktori", page_icon="📈",
                       layout="wide", initial_sidebar_state="collapsed")
    st.markdown("""<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    section[data-testid="stSidebar"]{display:none !important;}
    header[data-testid="stHeader"]{display:none !important;}
    .stApp{background:#f4f1fb !important;font-family:'Inter',sans-serif;}
    .block-container{max-width:1180px !important;padding-top:26px !important;}
    [class*="st-key-dir_odjava"] button{font-size:11px !important;padding:4px 9px !important;border-radius:7px !important;
        min-height:0 !important;background:#fff !important;border:1px solid #e5e7eb !important;color:#6b7280 !important;font-weight:600 !important;}
    </style>""", unsafe_allow_html=True)

    def _fmt(n):
        try:
            return f"{int(round(float(n))):,}".replace(",", ".")
        except Exception:
            return str(n)

    _h1, _h2 = st.columns([7.5, 1.15])
    with _h1:
        st.markdown('<div style="display:flex;align-items:center;gap:12px;padding:4px 0 2px;">'
                    '<div style="width:38px;height:38px;border-radius:10px;background:linear-gradient(135deg,#a855f7,#ec4899);"></div>'
                    '<div><span style="font-size:19px;font-weight:800;color:#1f2430;">VAPE Porudžbine</span>'
                    '<span style="font-size:13px;color:#8b8fa0;margin-left:6px;">· Direktori</span></div></div>',
                    unsafe_allow_html=True)
    with _h2:
        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
        if st.button("Odjava", key="dir_odjava", use_container_width=True):
            for _k in ("authenticated", "role", "admin_user", "mail_nalog"):
                st.session_state.pop(_k, None)
            st.rerun()

    if not sb_dostupan():
        st.error("Veza sa bazom nije podešena. Javi se analitičaru.")
        return

    _pub = sb_meseci()
    _mk = {m["key"] for m in _pub}
    # Uključi i tekući mesec (ako je počeo) + mesece za koje su postavljeni rokovi,
    # da direktor može da izabere i pre objave (i vidi poruku o roku).
    _tdy0 = datetime.date.today()
    _mk.add(str(_tdy0.year) + "-" + ("0" + str(_tdy0.month))[-2:])
    try:
        for _rmk in sb_rokovi_all().keys():
            _mk.add(_rmk)
    except Exception:
        pass
    _mes_keys = sorted(_mk, reverse=True)
    if not _mes_keys:
        st.info("Još nema objavljenih izveštaja.")
        return

    _mlbls = [mesec_label(k) for k in _mes_keys]
    # Podrazumevani mesec za prikaze = prethodni mesec (u avgustu se gleda jul); ako ga nema, najnoviji.
    _pv_y0 = _tdy0.year; _pv_m0 = _tdy0.month - 1
    if _pv_m0 <= 0:
        _pv_m0 += 12; _pv_y0 -= 1
    _prev_mk = str(_pv_y0) + "-" + ("0" + str(_pv_m0))[-2:]
    _prev_idx = _mes_keys.index(_prev_mk) if _prev_mk in _mes_keys else 0
    _view = st.session_state.get("dir_view", "dash")

    # ---- Render izveštaja po sistemu (isti prikaz za pune i delimične podatke) ----
    def _render_sistem_report(dd, puno):
        if puno:
            tek = dd.get("prodaja_tekuci", {}) or {}
            tek_kom = int(tek.get("kom", 0))
            comp = dd.get("poredjenja", {}) or {}

            def _delta_html(base):
                if base and base.get("kom"):
                    p = (tek_kom - int(base["kom"])) / int(base["kom"]) * 100.0
                    arrow = "▲" if p >= 0 else "▼"
                    col = "#16a34a" if p >= 0 else "#dc2626"
                    return ('<span style="color:' + col + ';">' + arrow + " "
                            + (("%.1f" % abs(p)).replace(".", ",")) + "%</span>")
                return '<span style="color:#c4c7cf;">— nema podataka</span>'

            _kpi = '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:22px;">'
            for (v, k) in [
                (_fmt(tek_kom) + " <span style='font-size:13px;color:#9aa0ad;'>kom</span>", "Prodaja — " + str(tek.get("mesec", ""))),
                (_delta_html(comp.get("prosli_mesec")), "vs prošli mesec"),
                (_delta_html(comp.get("isti_mesec_lani")), "vs isti mesec lani"),
                (_delta_html(comp.get("prosek_6m")), "vs 6-mesečni prosek"),
            ]:
                _kpi += ('<div style="background:#fff;border:1px solid #efeaf7;border-radius:15px;padding:16px 18px;'
                         'box-shadow:0 2px 12px rgba(80,40,140,.06);">'
                         '<div style="font-size:22px;font-weight:800;color:#1f2430;">' + v + '</div>'
                         '<div style="font-size:11.5px;color:#8b8fa0;margin-top:5px;font-weight:600;">' + k + '</div></div>')
            _kpi += '</div>'
            st.markdown(_kpi, unsafe_allow_html=True)

            # 2. Kombinovani grafikon: prodaja po mesecima i grupama (složeni stub)
            _naz = dd.get("nazivi") or [str(t.get("mesec", "")) for t in dd.get("prodaja_trend", [])]
            _gm = dd.get("grupe_mesecno") or {}
            if _naz and _gm:
                _render_stacked_chart(_naz, _gm)
            else:
                trend = dd.get("prodaja_trend", [])
                _maxk = max([t["kom"] for t in trend] or [1]) or 1
                _tr = ('<div style="background:#fff;border:1px solid #efeaf7;border-radius:16px;padding:20px 22px;margin-bottom:20px;'
                       'box-shadow:0 2px 12px rgba(80,40,140,.05);">'
                       '<div style="font-size:15px;font-weight:800;margin-bottom:16px;">Prodaja — trend po mesecima (kom)</div>')
                for t in trend:
                    _w = int(int(t["kom"]) / _maxk * 100)
                    _tr += ('<div style="display:grid;grid-template-columns:96px 1fr 96px;align-items:center;gap:12px;margin-bottom:9px;">'
                            '<span style="font-size:12.5px;color:#4b5563;font-weight:600;">' + str(t["mesec"]) + '</span>'
                            '<div style="height:12px;background:#f0ebf9;border-radius:20px;overflow:hidden;">'
                            '<div style="height:100%;width:' + str(_w) + '%;background:linear-gradient(90deg,#a855f7,#ec4899);border-radius:20px;"></div></div>'
                            '<span style="font-size:12.5px;text-align:right;font-weight:700;color:#4b5563;">' + _fmt(t["kom"]) + '</span></div>')
                _tr += '</div>'
                st.markdown(_tr, unsafe_allow_html=True)

        # Predlog porudžbine po grupama — samo za parcijalni prikaz (kad nema mesečnih grupa)
        grupe = dd.get("po_grupama", [])
        if grupe and not (dd.get("nazivi") and dd.get("grupe_mesecno")):
            _gt = "Prodaja po grupama" if puno else "Predlog porudžbine po grupama"
            _gs = ("Udeo u ukupnoj prodaji sistema (tekući mesec)." if puno
                   else "Udeo u predloženoj porudžbini po grupama (tekući mesec).")
            _gtot = sum(int(g["kom"]) for g in grupe) or 1
            _gmax = max(int(g["kom"]) for g in grupe) or 1
            _gh = ('<div style="background:#fff;border:1px solid #efeaf7;border-radius:16px;padding:20px 22px;margin-bottom:20px;'
                   'box-shadow:0 2px 12px rgba(80,40,140,.05);">'
                   '<div style="font-size:15px;font-weight:800;margin-bottom:4px;">' + _gt + '</div>'
                   '<div style="font-size:12.5px;color:#9aa0ad;margin-bottom:16px;">' + _gs + '</div>')
            for g in grupe:
                _kom = int(g["kom"]); _udeo = int(round(_kom / _gtot * 100)); _w = int(_kom / _gmax * 100)
                _gh += ('<div style="display:grid;grid-template-columns:150px 1fr 120px;align-items:center;gap:12px;margin-bottom:12px;">'
                        '<span style="font-size:13px;font-weight:600;color:#374151;">' + _h_escape(str(g["grupa"])) + '</span>'
                        '<div style="height:11px;background:#f0ebf9;border-radius:20px;overflow:hidden;">'
                        '<div style="height:100%;width:' + str(_w) + '%;background:linear-gradient(90deg,#7c3aed,#c084fc);border-radius:20px;"></div></div>'
                        '<span style="font-size:12.5px;text-align:right;font-weight:700;color:#4b5563;">' + _fmt(_kom) + ' · ' + str(_udeo) + '%</span></div>')
            _gh += '</div>'
            st.markdown(_gh, unsafe_allow_html=True)

        # 3. Prosečna prodaja po objektu (line)
        _ppo = dd.get("prosek_po_objektu") or []
        if _ppo:
            _render_prosek_line(_ppo)

        # 3b. Predlog porudžbine za sistem + pokrivenost
        _por = dd.get("porudzbina") or {}
        if _por and (_por.get("ukupno") or _por.get("po_grupi")):
            _render_porudzbina(_por)

        # 3c. Bestseleri i najslabiji artikli
        _arg = dd.get("artikli_rang") or {}
        if _arg and (_arg.get("best") or _arg.get("slab")):
            _render_artikli_rang(_arg)

        # 3d. Uspešnost akcije
        _ak = dd.get("akcija") or {}
        if _ak and _ak.get("artikli"):
            _render_akcija(_ak)

        # 4. Out of stock — po količinama (poslednji mesec)
        _okd = dd.get("oos_kom") or {}
        if _okd and (_okd.get("po_artiklu") or _okd.get("izgubljeno_kom") or _okd.get("objekata_na_0")):
            _render_oos_kom(_okd)
        else:
            oos = dd.get("oos", {}) or {}
            if oos:
                _oh = ('<div style="background:#fff;border:1px solid #efeaf7;border-radius:16px;padding:20px 22px;margin-bottom:14px;'
                       'box-shadow:0 2px 12px rgba(80,40,140,.05);">'
                       '<div style="font-size:15px;font-weight:800;margin-bottom:4px;">Out of stock — izgubljena prodaja</div>'
                       '<div style="font-size:12.5px;color:#9aa0ad;margin-bottom:16px;">Na osnovu dnevnog lagera: koliko komada je izgubljeno jer je artikal bio na nuli.</div>'
                       '<div style="display:grid;grid-template-columns:repeat(2,1fr);gap:14px;">'
                       '<div style="background:#fef2f2;border:1px solid #fecaca;border-radius:14px;padding:16px;text-align:center;">'
                       '<div style="font-size:24px;font-weight:800;color:#dc2626;">~' + _fmt(oos.get("izgubljeno_kom", 0)) + '</div>'
                       '<div style="font-size:11.5px;color:#9b6b6b;margin-top:3px;font-weight:600;">Izgubljeno (kom)</div></div>'
                       '<div style="background:#fef2f2;border:1px solid #fecaca;border-radius:14px;padding:16px;text-align:center;">'
                       '<div style="font-size:24px;font-weight:800;color:#dc2626;">' + _fmt(oos.get("kombinacija_na_0", 0)) + '</div>'
                       '<div style="font-size:11.5px;color:#9b6b6b;margin-top:3px;font-weight:600;">Kombinacija objekat × artikal na 0</div></div>'
                       '</div></div>')
                st.markdown(_oh, unsafe_allow_html=True)
                _pa = dd.get("oos_po_artiklu", [])
                if _pa:
                    _df = pd.DataFrame([{"Artikal": r["artikal"], "U koliko objekata na 0": r["objekata"],
                                         "Izgubljeno (kom)": r["izgubljeno"]} for r in _pa])
                    st.dataframe(_df, hide_index=True, use_container_width=True)

        # 5. Profitabilnost (identično kao u analitici)
        _pf_blok = dd.get("profit") or {}
        if _pf_blok and (_pf_blok.get("total_bruto") is not None):
            _render_profit_blok(_pf_blok)

    def _render_stacked_chart(nazivi, grupe_mesecno):
        # boje po grupi (stabilan raspored)
        _pal = ["#7c3aed", "#ec4899", "#f59e0b", "#0ea5e9", "#10b981", "#a855f7", "#f43f5e", "#14b8a6"]
        _keys = sorted(grupe_mesecno.keys(), key=lambda k: -sum(grupe_mesecno[k]))
        _boje = {k: _pal[i % len(_pal)] for i, k in enumerate(_keys)}
        _n = len(nazivi)
        _tot = [sum(int(grupe_mesecno[k][i]) if i < len(grupe_mesecno[k]) else 0 for k in _keys) for i in range(_n)]
        _maxt = max(_tot or [1]) or 1
        _cols = ""
        for i in range(_n):
            _hpx = int(_tot[i] / _maxt * 210)
            _segs = ""
            for k in _keys:
                _v = int(grupe_mesecno[k][i]) if i < len(grupe_mesecno[k]) else 0
                if _v > 0 and _tot[i] > 0:
                    _segs = ('<div style="width:100%;height:' + ("%.2f" % (_v / _tot[i] * 100)) + '%;background:' + _boje[k] + ';"></div>') + _segs
            _cols += ('<div style="flex:1;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;height:100%;min-width:0;">'
                      '<div title="' + _h_escape(str(nazivi[i])) + ': ' + _fmt(_tot[i]) + ' kom" '
                      'style="width:72%;height:' + str(_hpx) + 'px;border-radius:4px 4px 0 0;overflow:hidden;display:flex;flex-direction:column;justify-content:flex-end;box-shadow:0 1px 3px rgba(0,0,0,.06);">'
                      + _segs + '</div>'
                      '<div style="font-size:9px;color:#9aa0ad;margin-top:6px;font-weight:600;white-space:nowrap;transform:rotate(-30deg);transform-origin:center;">' + _h_escape(str(nazivi[i])) + '</div></div>')
        _leg = ""
        for k in _keys:
            _leg += ('<div style="display:flex;align-items:center;gap:7px;font-size:12.5px;color:#374151;font-weight:600;">'
                     '<span style="width:13px;height:13px;border-radius:4px;background:' + _boje[k] + ';"></span>' + _h_escape(str(k)) + '</div>')
        _html = ('<div style="background:#fff;border:1px solid #efeaf7;border-radius:16px;padding:20px 22px 14px;margin-bottom:20px;'
                 'box-shadow:0 2px 12px rgba(80,40,140,.05);">'
                 '<div style="font-size:15px;font-weight:800;margin-bottom:4px;">Prodaja po mesecima i grupama (kom)</div>'
                 '<div style="font-size:12.5px;color:#9aa0ad;margin-bottom:16px;">Visina stuba je ukupna prodaja meseca; boje pokazuju koliko je koja grupa prodala.</div>'
                 '<div style="display:flex;align-items:flex-end;gap:5px;height:250px;padding:8px 2px 0;">' + _cols + '</div>'
                 '<div style="display:flex;gap:18px;flex-wrap:wrap;margin-top:20px;">' + _leg + '</div></div>')
        st.markdown(_html, unsafe_allow_html=True)

    def _render_prosek_line(ppo):
        _lm = [str(p.get("mesec", "")) for p in ppo]
        _lv = [float(p.get("prosek", 0) or 0) for p in ppo]
        _n = len(_lv)
        if _n == 0:
            return
        W = 1080; H = 240; pl = 46; pr = 20; pt = 24; pb = 40
        pw = W - pl - pr; ph = H - pt - pb
        _mx = (max(_lv) or 1) * 1.15

        def _px(i):
            return pl + (i / (_n - 1) * pw if _n > 1 else pw / 2)

        def _py(v):
            return pt + ph - (v / _mx * ph if _mx else 0)

        _g = ""
        for f in (0, 0.25, 0.5, 0.75, 1):
            _y = pt + ph - f * ph
            _g += ('<line x1="' + str(pl) + '" y1="' + ("%.1f" % _y) + '" x2="' + str(W - pr) + '" y2="' + ("%.1f" % _y) + '" stroke="#f0ebf9"/>'
                   '<text x="' + str(pl - 8) + '" y="' + ("%.1f" % (_y + 4)) + '" font-size="10" fill="#b9bdc9" text-anchor="end">' + ("%.1f" % (_mx * f)).replace(".", ",") + '</text>')
        _pts = " ".join(("%.1f,%.1f" % (_px(i), _py(_lv[i]))) for i in range(_n))
        _g += '<polygon points="' + str(pl) + "," + ("%.1f" % (pt + ph)) + " " + _pts + " " + str(W - pr) + "," + ("%.1f" % (pt + ph)) + '" fill="#a855f7" fill-opacity="0.08"/>'
        _g += '<polyline points="' + _pts + '" fill="none" stroke="#7c3aed" stroke-width="2.5"/>'
        for i in range(_n):
            _x = _px(i); _y = _py(_lv[i])
            _g += ('<circle cx="' + ("%.1f" % _x) + '" cy="' + ("%.1f" % _y) + '" r="5" fill="#7c3aed" stroke="#fff" stroke-width="2"/>'
                   '<text x="' + ("%.1f" % _x) + '" y="' + ("%.1f" % (_y - 12)) + '" font-size="11" font-weight="700" fill="#6d28d9" text-anchor="middle">' + ("%.1f" % _lv[i]).replace(".", ",") + '</text>'
                   '<text x="' + ("%.1f" % _x) + '" y="' + str(H - 10) + '" font-size="10" fill="#9aa0ad" text-anchor="middle">' + _h_escape(_lm[i]) + '</text>')
        _svg = '<svg viewBox="0 0 ' + str(W) + ' ' + str(H) + '" style="width:100%;height:240px;">' + _g + '</svg>'
        _html = ('<div style="background:#fff;border:1px solid #efeaf7;border-radius:16px;padding:20px 22px;margin-bottom:20px;'
                 'box-shadow:0 2px 12px rgba(80,40,140,.05);">'
                 '<div style="font-size:15px;font-weight:800;margin-bottom:4px;">Prosečna prodaja po objektu (kom / objektu)</div>'
                 '<div style="font-size:12.5px;color:#9aa0ad;margin-bottom:10px;">Ukupna prodaja meseca podeljena brojem aktivnih objekata tog meseca.</div>'
                 + _svg + '</div>')
        st.markdown(_html, unsafe_allow_html=True)

    def _render_oos_kom(ok):
        _mes = str(ok.get("mesec", ""))
        _oh = ('<div style="background:#fff;border:1px solid #efeaf7;border-radius:16px;padding:20px 22px;margin-bottom:14px;'
               'box-shadow:0 2px 12px rgba(80,40,140,.05);">'
               '<div style="font-size:15px;font-weight:800;margin-bottom:4px;">Out of stock — po količinama · ' + _h_escape(_mes) + '</div>'
               '<div style="font-size:12.5px;color:#9aa0ad;margin-bottom:16px;">Koliko je komada izgubljeno u prethodnom mesecu jer je artikal bio na nuli.</div>'
               '<div style="display:grid;grid-template-columns:repeat(2,1fr);gap:14px;">'
               '<div style="background:#fef2f2;border:1px solid #fecaca;border-radius:14px;padding:16px;text-align:center;">'
               '<div style="font-size:24px;font-weight:800;color:#dc2626;">~' + _fmt(ok.get("izgubljeno_kom", 0)) + '</div>'
               '<div style="font-size:11.5px;color:#9b6b6b;margin-top:3px;font-weight:600;">Izgubljeno (kom) · ' + _h_escape(_mes) + '</div></div>'
               '<div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:14px;padding:16px;text-align:center;">'
               '<div style="font-size:24px;font-weight:800;color:#2563eb;">' + _fmt(ok.get("objekata_na_0", 0)) + '</div>'
               '<div style="font-size:11.5px;color:#5b7bb0;margin-top:3px;font-weight:600;">U koliko objekata je lager na 0</div></div>'
               '</div></div>')
        st.markdown(_oh, unsafe_allow_html=True)
        _pa = ok.get("po_artiklu", [])
        if _pa:
            _df = pd.DataFrame([{"Artikal": r["artikal"], "U koliko objekata na 0": r["objekata"],
                                 "Izgubljeno (kom) · " + _mes: r["izgubljeno"]} for r in _pa])
            st.dataframe(_df, hide_index=True, use_container_width=True, height=min(60 + 35 * len(_pa), 520))

    def _render_porudzbina(pr):
        _uk = int(pr.get("ukupno", 0)); _ob = int(pr.get("objekata", 0)); _dani = pr.get("dani_avg")
        _h = ('<div style="background:#fff;border:1px solid #efeaf7;border-radius:16px;padding:20px 22px;margin-bottom:20px;'
              'box-shadow:0 2px 12px rgba(80,40,140,.05);">'
              '<div style="font-size:15px;font-weight:800;margin-bottom:4px;">Predlog porudžbine za sistem</div>'
              '<div style="font-size:12.5px;color:#9aa0ad;margin-bottom:16px;">Koliko sistem treba da poruči (preporuka analitike) i za koliko dana lager traje.</div>'
              '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:' + ('16px' if pr.get("po_grupi") else '0') + ';">')
        _cards = [("#7c3aed", _fmt(_uk) + " <span style='font-size:13px;color:#9aa0ad;'>kom</span>", "Ukupno za poručivanje"),
                  ("#0ea5e9", _fmt(_ob), "Objekata poručuje")]
        if _dani is not None:
            _cards.append(("#f59e0b", _fmt(_dani) + " <span style='font-size:13px;color:#9aa0ad;'>dana</span>", "Prosečna pokrivenost lagera"))
        for (col, v, lab) in _cards:
            _h += ('<div style="background:#faf9fd;border:1px solid #efeaf7;border-left:4px solid ' + col + ';border-radius:14px;padding:15px 17px;">'
                   '<div style="font-size:20px;font-weight:800;color:' + col + ';">' + v + '</div>'
                   '<div style="font-size:11px;color:#9aa0ad;margin-top:4px;font-weight:600;">' + lab + '</div></div>')
        _h += '</div>'
        _pg = pr.get("po_grupi", [])
        if _pg:
            _gmax = max([int(g["kom"]) for g in _pg] or [1]) or 1
            _h += '<div style="font-size:12.5px;font-weight:700;color:#374151;margin:4px 0 10px;">Po grupama:</div>'
            for g in _pg:
                _k = int(g["kom"]); _w = int(_k / _gmax * 100)
                _h += ('<div style="display:grid;grid-template-columns:150px 1fr 90px;align-items:center;gap:12px;margin-bottom:9px;">'
                       '<span style="font-size:13px;font-weight:600;color:#374151;">' + _h_escape(str(g["grupa"])) + '</span>'
                       '<div style="height:11px;background:#f0ebf9;border-radius:20px;overflow:hidden;">'
                       '<div style="height:100%;width:' + str(_w) + '%;background:linear-gradient(90deg,#7c3aed,#c084fc);border-radius:20px;"></div></div>'
                       '<span style="font-size:12.5px;text-align:right;font-weight:700;color:#4b5563;">' + _fmt(_k) + ' kom</span></div>')
        _h += '</div>'
        st.markdown(_h, unsafe_allow_html=True)

    def _render_artikli_rang(ar):
        _best = ar.get("best", []); _slab = ar.get("slab", [])
        _c1, _c2 = st.columns(2)
        with _c1:
            _h = _card_open("🔝 Bestseleri", "Najprodavaniji artikli u sistemu (ceo period).")
            _mx = max([int(x["prodato"]) for x in _best] or [1]) or 1
            for x in _best:
                _p = int(x["prodato"]); _w = int(_p / _mx * 100)
                _h += ('<div style="margin-bottom:10px;">'
                       '<div style="display:flex;justify-content:space-between;font-size:12.5px;color:#374151;margin-bottom:3px;">'
                       '<span style="font-weight:600;">' + _h_escape(str(x["artikal"])[:38]) + '</span>'
                       '<span style="font-weight:700;color:#16a34a;">' + _fmt(_p) + ' kom</span></div>'
                       '<div style="height:9px;background:#f0fdf4;border-radius:20px;overflow:hidden;">'
                       '<div style="height:100%;width:' + str(_w) + '%;background:linear-gradient(90deg,#16a34a,#4ade80);border-radius:20px;"></div></div></div>')
            st.markdown(_h + "</div>", unsafe_allow_html=True)
        with _c2:
            _h = _card_open("🐌 Najslabiji artikli", "Najmanje prodaju — kandidati za smanjenje zaliha.")
            _mx = max([int(x["prodato"]) for x in _best] or [1]) or 1
            for x in _slab:
                _p = int(x["prodato"]); _w = int(_p / _mx * 100)
                _h += ('<div style="margin-bottom:10px;">'
                       '<div style="display:flex;justify-content:space-between;font-size:12.5px;color:#374151;margin-bottom:3px;">'
                       '<span style="font-weight:600;">' + _h_escape(str(x["artikal"])[:38]) + '</span>'
                       '<span style="font-weight:700;color:#dc2626;">' + _fmt(_p) + ' kom</span></div>'
                       '<div style="height:9px;background:#fef2f2;border-radius:20px;overflow:hidden;">'
                       '<div style="height:100%;width:' + str(max(_w, 2)) + '%;background:linear-gradient(90deg,#f59e0b,#fca5a5);border-radius:20px;"></div></div></div>')
            st.markdown(_h + "</div>", unsafe_allow_html=True)

    def _render_akcija(ak):
        _ua = int(ak.get("ukupno_akcija", 0)); _ur = int(ak.get("ukupno_redovna", 0)); _rz = int(ak.get("razlika", 0))
        _h = ('<div style="background:#fff;border:1px solid #efeaf7;border-radius:16px;padding:20px 22px;margin-bottom:14px;'
              'box-shadow:0 2px 12px rgba(80,40,140,.05);">'
              '<div style="font-size:15px;font-weight:800;margin-bottom:4px;">Uspešnost akcije</div>'
              '<div style="font-size:12.5px;color:#9aa0ad;margin-bottom:16px;">Koliko je akcija donela naspram da se prodavalo po redovnoj ceni, i obrt po artiklu.</div>'
              '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px;">')
        for (col, v, lab) in [("#10b981", _fmt(_ua) + " RSD", "Profit ostvaren na akciji"),
                              ("#7c3aed", _fmt(_ur) + " RSD", "Da je bila redovna cena"),
                              ("#ec4899", ("-" if _rz > 0 else "+") + _fmt(abs(_rz)) + " RSD", "Razlika (koliko je akcija „koštala“)")]:
            _h += ('<div style="background:#faf9fd;border:1px solid #efeaf7;border-left:4px solid ' + col + ';border-radius:14px;padding:15px 17px;">'
                   '<div style="font-size:18px;font-weight:800;color:' + col + ';">' + v + '</div>'
                   '<div style="font-size:11px;color:#9aa0ad;margin-top:4px;font-weight:600;">' + lab + '</div></div>')
        _h += '</div></div>'
        st.markdown(_h, unsafe_allow_html=True)
        _rows = []
        for a in ak.get("artikli", []):
            _rows.append({"Artikal": str(a["naziv"]), "Grupa": str(a.get("grupa", "")),
                          "Prodato (kom)": int(a["prodato"]), "Obrt (x)": round(float(a["obrt"]), 1),
                          "Popust %": round(float(a["popust"]), 1),
                          "Profit akcija (RSD)": int(a["profit_akcija"]),
                          "Cena akcije (RSD)": int(a["cena_akcije"]),
                          "Dani pokrivanja": int(a["dani"])})
        if _rows:
            st.dataframe(pd.DataFrame(_rows), hide_index=True, use_container_width=True,
                         height=min(60 + 35 * len(_rows), 480))

    def _card_open(naslov, podnaslov=""):
        _h = ('<div style="background:#fff;border:1px solid #efeaf7;border-radius:16px;padding:20px 22px;margin-bottom:20px;'
              'box-shadow:0 2px 12px rgba(80,40,140,.05);">'
              '<div style="font-size:15px;font-weight:800;margin-bottom:' + ("4px" if podnaslov else "16px") + ';">' + naslov + '</div>')
        if podnaslov:
            _h += '<div style="font-size:12.5px;color:#9aa0ad;margin-bottom:16px;">' + podnaslov + '</div>'
        return _h

    def _rsd(v):
        try:
            return _fmt(int(round(v))) + " RSD"
        except Exception:
            return str(v) + " RSD"

    def _bar_row(lb, val, maxv, boja):
        _w = int(abs(val) / maxv * 100) if maxv else 0
        _w = min(_w, 100)
        return ('<div style="display:grid;grid-template-columns:70px 1fr 120px;align-items:center;gap:10px;margin-bottom:7px;">'
                '<span style="font-size:11.5px;color:#888;text-align:right;">' + str(lb) + '</span>'
                '<div style="height:16px;background:#f5f0ff;border-radius:4px;overflow:hidden;">'
                '<div style="height:100%;width:' + str(_w) + '%;background:' + boja + ';border-radius:4px;"></div></div>'
                '<span style="font-size:11.5px;font-weight:700;color:#555;">' + _rsd(val) + '</span></div>')

    def _render_profit_blok(pf):
        if not pf:
            return
        n_mes = int(pf.get("n_mes", 1)) or 1
        st.markdown("<div style='margin:26px 0 2px;height:1px;background:linear-gradient(90deg,transparent,#e6def5,transparent);'></div>"
                    "<div style='font-size:16px;font-weight:800;margin:16px 0 4px;'>💰 Profitabilnost</div>"
                    "<div style='font-size:12.5px;color:#9aa0ad;margin-bottom:16px;'>Period: <b>" + _h_escape(str(pf.get("period", ""))) + "</b> · "
                    + str(pf.get("n_obj", 0)) + " objekata · " + str(n_mes) + " meseci</div>", unsafe_allow_html=True)
        # 4 KPI kartice u dinarima
        _kards = [
            ("Trošak marketinga", pf.get("total_trosak", 0), "#a855f7", ""),
            ("Bruto profit", pf.get("total_bruto", 0), "#10b981", ""),
            ("Neto profit", pf.get("total_neto", 0), "#7c3aed" if pf.get("total_neto", 0) > 0 else "#ec4899", ""),
            ("OOS izgubljen", pf.get("total_oos", 0), "#ec4899", "-"),
        ]
        _kh = '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px;">'
        for (lab, tot, col, pre) in _kards:
            _mes = int(round(tot / n_mes))
            _kh += ('<div style="background:#fff;border:1px solid #efeaf7;border-left:4px solid ' + col + ';border-radius:14px;padding:15px 17px;'
                    'box-shadow:0 2px 12px rgba(80,40,140,.06);">'
                    '<div style="font-size:10.5px;color:#9aa0ad;font-weight:700;letter-spacing:.4px;text-transform:uppercase;margin-bottom:6px;">' + lab + '</div>'
                    '<div style="font-size:19px;font-weight:800;color:' + col + ';">' + pre + _rsd(tot) + '</div>'
                    '<div style="font-size:11px;color:#aab;margin-top:3px;">' + pre + _rsd(_mes) + ' / mesec</div></div>')
        _kh += '</div>'
        st.markdown(_kh, unsafe_allow_html=True)
        # Mesečni trend bruto / neto
        _bm = pf.get("bruto_po_mes", []); _nm = pf.get("neto_po_mes", [])
        if _bm or _nm:
            _cb, _cn = st.columns(2)
            with _cb:
                _mx = max([abs(v) for _, v in _bm] or [1]) or 1
                _html = _card_open("📈 Mesečni trend bruto profita")
                for lb, v in _bm:
                    _html += _bar_row(lb, v, _mx, "#a855f7")
                st.markdown(_html + "</div>", unsafe_allow_html=True)
            with _cn:
                _mx = max([abs(v) for _, v in _nm] or [1]) or 1
                _html = _card_open("📉 Mesečni trend neto profita")
                for lb, v in _nm:
                    _html += _bar_row(lb, v, _mx, "#7c3aed" if v >= 0 else "#ec4899")
                st.markdown(_html + "</div>", unsafe_allow_html=True)
        # Profitabilnost po objektima — donut + procena uštede
        _uk = int(pf.get("obj_ukupno", 0))
        if _uk > 0:
            _pr = int(pf.get("obj_profit", 0)); _on = int(pf.get("obj_oos_neg", 0)); _pn = int(pf.get("obj_pravi_neg", 0))
            import math as _m
            _cx, _cy, _rO, _rI = 90, 90, 78, 50
            def _arc(cx, cy, r, sd, ed):
                s = _m.radians(sd - 90); e = _m.radians(ed - 90)
                lg = 1 if (ed - sd) > 180 else 0
                return (cx + r * _m.cos(s), cy + r * _m.sin(s), cx + r * _m.cos(e), cy + r * _m.sin(e), lg)
            _segs = [(_pr, "#10b981"), (_on, "#f59e0b"), (_pn, "#ec4899")]
            _tot = sum(s[0] for s in _segs) or 1
            _nonzero = [s for s in _segs if s[0] > 0]
            _paths = ""
            if len(_nonzero) == 1:
                # jedan segment = pun prsten (SVG luk od 360° se ne iscrtava)
                _col = _nonzero[0][1]
                _paths = ('<circle cx="' + str(_cx) + '" cy="' + str(_cy) + '" r="' + str(_rO) + '" fill="' + _col + '"/>'
                          '<circle cx="' + str(_cx) + '" cy="' + str(_cy) + '" r="' + str(_rI) + '" fill="#ffffff"/>')
            else:
                _ang = 0.0
                for _val, _col in _segs:
                    if _val <= 0:
                        continue
                    _sw = _val / _tot * 360.0
                    x1, y1, x2, y2, lg = _arc(_cx, _cy, _rO, _ang, _ang + _sw)
                    xi2, yi2, xi1, yi1, _ = _arc(_cx, _cy, _rI, _ang, _ang + _sw)
                    _paths += ('<path d="M' + ("%.2f" % x1) + ',' + ("%.2f" % y1) + ' A' + str(_rO) + ',' + str(_rO) + ' 0 ' + str(lg) + ' 1 '
                               + ("%.2f" % x2) + ',' + ("%.2f" % y2) + ' L' + ("%.2f" % xi2) + ',' + ("%.2f" % yi2)
                               + ' A' + str(_rI) + ',' + str(_rI) + ' 0 ' + str(lg) + ' 0 ' + ("%.2f" % xi1) + ',' + ("%.2f" % yi1)
                               + ' Z" fill="' + _col + '"/>')
                    _ang += _sw
            _svg = ('<svg width="180" height="180" viewBox="0 0 180 180" xmlns="http://www.w3.org/2000/svg">' + _paths
                    + '<text x="90" y="86" text-anchor="middle" font-size="26" font-weight="800" fill="#1f2430">' + str(_uk) + '</text>'
                    + '<text x="90" y="104" text-anchor="middle" font-size="10" fill="#9aa0ad">objekata</text></svg>')
            _leg = ('<div style="display:flex;flex-direction:column;gap:10px;">'
                    '<div style="display:flex;align-items:center;gap:8px;"><span style="width:12px;height:12px;border-radius:3px;background:#10b981;"></span>'
                    '<span style="font-size:13px;color:#374151;">Profitabilni objekti: <b>' + str(_pr) + '</b></span></div>'
                    '<div style="display:flex;align-items:center;gap:8px;"><span style="width:12px;height:12px;border-radius:3px;background:#f59e0b;"></span>'
                    '<span style="font-size:13px;color:#374151;">Neprofitabilni zbog OOS: <b>' + str(_on) + '</b></span></div>'
                    '<div style="display:flex;align-items:center;gap:8px;"><span style="width:12px;height:12px;border-radius:3px;background:#ec4899;"></span>'
                    '<span style="font-size:13px;color:#374151;">Pravi neprofitabilni: <b>' + str(_pn) + '</b></span></div>'
                    '<div style="margin-top:8px;font-size:12.5px;color:#6b7280;line-height:1.5;">Procena uštede ako se pravi neprofitabilni ugase: '
                    '<b style="color:#16a34a;">' + _rsd(pf.get("usteda_ukupno", 0)) + '</b> za period.</div></div>')
            _ph = (_card_open("🏪 Profitabilnost po objektima")
                   + '<div style="display:grid;grid-template-columns:200px 1fr;align-items:center;gap:18px;">'
                   + '<div style="text-align:center;">' + _svg + '</div>' + _leg + '</div></div>')
            st.markdown(_ph, unsafe_allow_html=True)
            # tabela objekata
            _objs = pf.get("objekti", [])
            if _objs:
                _km = sb_komitenti_map()
                _rows = []
                for o in _objs:
                    _nz = _km.get(int(o["id"]), "") or ("ID " + str(o["id"]))
                    _rows.append({"Objekat": _nz, "Neto profit (RSD)": o["neto"],
                                  "Bruto (RSD)": o["bruto"], "Trošak (RSD)": o["trosak"],
                                  "Izgubljeno OOS (RSD)": o["oos"], "Potencijal (RSD)": o["potencijal"]})
                st.markdown("<div style='font-size:13px;font-weight:700;color:#374151;margin:4px 0 8px;'>Svi objekti (od najlošijeg neto profita):</div>", unsafe_allow_html=True)
                st.dataframe(pd.DataFrame(_rows), hide_index=True, use_container_width=True, height=340)
        # OOS — izgubljena zarada (u dinarima)
        _oa = pf.get("oos_artikli", [])
        _total_oos = int(pf.get("total_oos", 0))
        if _total_oos > 0 or _oa:
            _oh = _card_open("🔴 Out of stock — izgubljena zarada", "Koliko dinara je izgubljeno jer je artikal bio na nuli.")
            _mesv = int(round(_total_oos / n_mes))
            _oh += '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:6px;">'
            for (lab, val) in [("Izgubljen profit · " + str(n_mes) + " mes.", _rsd(_total_oos)),
                               ("Prosečno mesečno", _rsd(_mesv)),
                               ("Kombinacija na 0 danas", _fmt(pf.get("oos_0_danas", 0)))]:
                _oh += ('<div style="background:#fef2f2;border:1px solid #fecaca;border-radius:14px;padding:15px;text-align:center;">'
                        '<div style="font-size:18px;font-weight:800;color:#dc2626;">' + str(val) + '</div>'
                        '<div style="font-size:11px;color:#9b6b6b;margin-top:3px;font-weight:600;">' + lab + '</div></div>')
            _oh += '</div></div>'
            st.markdown(_oh, unsafe_allow_html=True)
            if _oa:
                _df = pd.DataFrame([{"Artikal": r["naziv"], "U koliko objekata": r["objekata"],
                                     "OOS meseci": r["meseci"], "Izgubljeni profit (RSD)": r["rsd"]} for r in _oa])
                st.dataframe(_df, hide_index=True, use_container_width=True, height=320)

    def _partial_iz_stavki(stavke):
        out = {}
        g = {}
        for s in stavke:
            gg = str(s.get("grupa", "") or "—")
            g[gg] = g.get(gg, 0) + int(s.get("kol", 0) or 0)
        out["po_grupama"] = [{"grupa": k, "kom": v} for k, v in
                             sorted(g.items(), key=lambda kv: kv[1], reverse=True) if v > 0]
        _oi = [s for s in stavke if int(s.get("lager", 0) or 0) == 0 and int(s.get("pred", 0) or 0) > 0]
        out["oos"] = {"kombinacija_na_0": len(_oi), "izgubljeno_kom": sum(int(s.get("pred", 0) or 0) for s in _oi)}
        _da = {}
        for s in _oi:
            nz = str(s.get("naziv", ""))
            e = _da.setdefault(nz, {"obj": set(), "izg": 0})
            e["obj"].add(int(s["idk"])); e["izg"] += int(s.get("pred", 0) or 0)
        _top = sorted(_da.items(), key=lambda kv: kv[1]["izg"], reverse=True)[:10]
        out["oos_po_artiklu"] = [{"artikal": k, "objekata": len(v["obj"]), "izgubljeno": v["izg"]} for k, v in _top]
        return out

    # ---------- TABLA (dve kartice) ----------
    if _view == "dash":
        st.markdown('<div style="color:#6b7280;font-size:14px;margin:6px 0 20px;">Dobrodošli 👋 &nbsp;'
                    'Pregled izveštaja i efikasnosti administracije.</div>', unsafe_allow_html=True)
        st.markdown('<div style="font-size:12px;text-transform:uppercase;letter-spacing:.6px;color:#9aa0ad;'
                    'font-weight:700;margin-bottom:12px;">Kartice</div>', unsafe_allow_html=True)
        _cc1, _cc2, _cc3 = st.columns(3)
        with _cc1:
            with st.container(border=True):
                st.markdown('<div style="font-size:32px;">📊</div>'
                            '<div style="font-size:16px;font-weight:800;margin:6px 0 4px;">Izveštaj efikasnosti administracije</div>'
                            '<div style="font-size:13px;color:#8b8fa0;line-height:1.5;margin-bottom:12px;">PDF izveštaj administracije i upozorenja koja je administracija prosledila komercijali.</div>',
                            unsafe_allow_html=True)
                if st.button("Otvori →", key="dir_open_efik", use_container_width=True):
                    st.session_state["dir_view"] = "efikasnost"; st.rerun()
        with _cc2:
            with st.container(border=True):
                st.markdown('<div style="font-size:32px;">📈</div>'
                            '<div style="font-size:16px;font-weight:800;margin:6px 0 4px;">Detaljan izveštaj po sistemima</div>'
                            '<div style="font-size:13px;color:#8b8fa0;line-height:1.5;margin-bottom:12px;">Prodaja, grupe, out-of-stock i preuzimanje analitike — po sistemu i mesecu.</div>',
                            unsafe_allow_html=True)
                if st.button("Otvori →", key="dir_open_sist", use_container_width=True):
                    st.session_state["dir_view"] = "sistemi"; st.rerun()
        with _cc3:
            with st.container(border=True):
                st.markdown('<div style="font-size:32px;">💹</div>'
                            '<div style="font-size:16px;font-weight:800;margin:6px 0 4px;">Izveštaj prodaje</div>'
                            '<div style="font-size:13px;color:#8b8fa0;line-height:1.5;margin-bottom:12px;">Kompletan dashboard: prodaja, uspešnost akcije, profitabilnost, zalihe — plus izvoz u Excel.</div>',
                            unsafe_allow_html=True)
                if st.button("Otvori →", key="dir_open_prod", use_container_width=True):
                    st.session_state["dir_view"] = "prodaja"; st.rerun()
        _cd1, _cd2, _cd3 = st.columns(3)
        with _cd1:
            with st.container(border=True):
                st.markdown('<div style="font-size:32px;">🧊</div>'
                            '<div style="font-size:16px;font-weight:800;margin:6px 0 4px;">Izveštaj SYX</div>'
                            '<div style="font-size:13px;color:#8b8fa0;line-height:1.5;margin-bottom:12px;">Word izveštaji za SYX nikotinske vrećice, izlistani po mesecima za preuzimanje.</div>',
                            unsafe_allow_html=True)
                if st.button("Otvori →", key="dir_open_syx", use_container_width=True):
                    st.session_state["dir_view"] = "syx"; st.rerun()
        with _cd2:
            with st.container(border=True):
                st.markdown('<div style="font-size:32px;">💳</div>'
                            '<div style="font-size:16px;font-weight:800;margin:6px 0 4px;">Izveštaj potraživanja</div>'
                            '<div style="font-size:13px;color:#8b8fa0;line-height:1.5;margin-bottom:12px;">Potraživanja po mesecu koja je administracija dopunila — pregled i izvoz u Excel.</div>',
                            unsafe_allow_html=True)
                if st.button("Otvori →", key="dir_open_potraz", use_container_width=True):
                    st.session_state["dir_view"] = "potraz"; st.rerun()
        with _cd3:
            with st.container(border=True):
                st.markdown('<div style="font-size:32px;">📅</div>'
                            '<div style="font-size:16px;font-weight:800;margin:6px 0 4px;">Rokovi</div>'
                            '<div style="font-size:13px;color:#8b8fa0;line-height:1.5;margin-bottom:12px;">Postavi rokove po mesecu: administracija, izveštaj po sistemu i osvežavanje izveštaja prodaje.</div>',
                            unsafe_allow_html=True)
                if st.button("Otvori →", key="dir_open_rok", use_container_width=True):
                    st.session_state["dir_view"] = "rokovi"; st.rerun()
        _ce1, _ce2, _ce3 = st.columns(3)
        with _ce1:
            with st.container(border=True):
                st.markdown('<div style="font-size:32px;">👥</div>'
                            '<div style="font-size:16px;font-weight:800;margin:6px 0 4px;">Statistika — ko je šta uradio</div>'
                            '<div style="font-size:13px;color:#8b8fa0;line-height:1.5;margin-bottom:12px;">Pregled rada administracije po osobi: pozivi, mejlovi, prosleđivanja i prijave, po mesecu.</div>',
                            unsafe_allow_html=True)
                if st.button("Otvori →", key="dir_open_stat", use_container_width=True):
                    st.session_state["dir_view"] = "statistika"; st.rerun()
        return

    if st.button("← Nazad na kartice", key="dir_back"):
        st.session_state["dir_view"] = "dash"; st.rerun()

    # ---------- KARTICA: STATISTIKA (ko je šta uradio) ----------
    if _view == "statistika":
        st.markdown('<div style="font-size:20px;font-weight:800;margin:6px 0 12px;">👥 Statistika — ko je šta uradio</div>', unsafe_allow_html=True)
        if not _mes_keys:
            st.info("Još nema objavljenih podataka.")
            return
        _sl = st.selectbox("Mesec", _mlbls, index=_prev_idx if 0 <= _prev_idx < len(_mlbls) else 0, key="stat_mes_dir")
        render_statistika(_mes_keys[_mlbls.index(_sl)], _sl)
        return

    # ---------- KARTICA: IZVEŠTAJ POTRAŽIVANJA ----------
    if _view == "potraz":
        potraz_director_ui()
        return

    # ---------- KARTICA 4: ROKOVI ----------
    if _view == "rokovi":
        st.markdown('<div style="font-size:20px;font-weight:800;margin:6px 0 6px;">📅 Rokovi</div>', unsafe_allow_html=True)
        st.caption("Postavi rokove po mesecu. Administracija radi do svog roka pa se zaključava; izveštaj efikasnosti "
                   "vidiš tek kad rok prođe. Za izveštaj po sistemu direktori vide poruku o roku dok se ne objavi. "
                   "Izveštaj prodaje se uvek vidi, a rok i napomena stoje u uglu.")

        # meseci: par unazad + par unapred + objavljeni
        _today = datetime.date.today()
        _opts = set(_mes_keys)
        _yy, _mm = _today.year, _today.month - 2
        while _mm <= 0:
            _mm += 12; _yy -= 1
        for _ in range(9):
            _opts.add(str(_yy) + "-" + ("0" + str(_mm))[-2:])
            _mm += 1
            if _mm > 12:
                _mm = 1; _yy += 1
        _opts = sorted(_opts, reverse=True)
        _rsel = st.selectbox("Mesec", _opts, format_func=mesec_label, key="rok_mes_sel")
        _post = sb_rokovi_get(_rsel)

        def _dflt(v):
            try:
                return datetime.date.fromisoformat(str(v)[:10])
            except Exception:
                return _today

        _r1, _r2, _r3 = st.columns(3)
        with _r1:
            _da = st.date_input("Rok — Izveštaj administracije", value=_dflt(_post.get("rok_admin")),
                                key="rok_admin_in", format="DD.MM.YYYY")
        with _r2:
            _ds = st.date_input("Rok — Izveštaj po sistemu", value=_dflt(_post.get("rok_sistemi")),
                                key="rok_sis_in", format="DD.MM.YYYY")
        with _r3:
            _dp = st.date_input("Rok — Osvežavanje izveštaja prodaje", value=_dflt(_post.get("rok_prodaja")),
                                key="rok_prod_in", format="DD.MM.YYYY")
        _r4, _r5, _r6 = st.columns(3)
        with _r4:
            _dsx = st.date_input("Rok — Izveštaj SYX", value=_dflt(_post.get("rok_syx")),
                                 key="rok_syx_in", format="DD.MM.YYYY")
        with _r5:
            _dpz = st.date_input("Rok — Izveštaj potraživanja", value=_dflt(_post.get("rok_potraz")),
                                 key="rok_potraz_in", format="DD.MM.YYYY")
        with _r6:
            _dkn = st.date_input("Rok — Kontrola komercijale", value=_dflt(_post.get("rok_kontrola")),
                                 key="rok_kontrola_in", format="DD.MM.YYYY")
        _nap = st.text_area("Napomena (za izveštaj prodaje — šta osvežiti, na šta obratiti pažnju)",
                            value=_post.get("napomena") or "", key="rok_nap_in", height=80)
        if st.button("💾 Sačuvaj rokove", key="rok_save", type="primary"):
            try:
                sb_rokovi_set(_rsel, _da.isoformat(), _ds.isoformat(), _dp.isoformat(), _nap,
                              rok_syx=_dsx.isoformat(), rok_potraz=_dpz.isoformat(),
                              rok_kontrola=_dkn.isoformat())
                st.success("Rokovi za " + mesec_label(_rsel) + " sačuvani.")
                st.rerun()
            except Exception as _e:
                st.error("Greška pri čuvanju: " + str(_e))

        _all = sb_rokovi_all()
        if _all:
            st.markdown("<div style='margin:18px 0 6px;font-size:12px;text-transform:uppercase;letter-spacing:.6px;"
                        "color:#9aa0ad;font-weight:700;'>Postavljeni rokovi</div>", unsafe_allow_html=True)
            _rows = []
            for _mk in sorted(_all.keys(), reverse=True):
                _v = _all[_mk]
                _rows.append({"Mesec": mesec_label(_mk),
                              "Administracija": _rok_fmt(_v.get("rok_admin")),
                              "Po sistemu": _rok_fmt(_v.get("rok_sistemi")),
                              "Izveštaj prodaje": _rok_fmt(_v.get("rok_prodaja")),
                              "SYX": _rok_fmt(_v.get("rok_syx")),
                              "Potraživanja": _rok_fmt(_v.get("rok_potraz")),
                              "Kontrola komercijale": _rok_fmt(_v.get("rok_kontrola")),
                              "Napomena": (_v.get("napomena") or "")[:50]})
            st.dataframe(pd.DataFrame(_rows), hide_index=True, use_container_width=True)
        return

    # ---------- KARTICA: IZVEŠTAJ SYX ----------
    if _view == "syx":
        st.markdown('<div style="font-size:20px;font-weight:800;margin:6px 0 6px;">🧊 Izveštaj SYX (nikotinske vrećice)</div>', unsafe_allow_html=True)
        st.caption("Word izveštaji za SYX, po mesecima. Analitičar ih ubacuje u delu Objava izveštaja.")
        _syl = sb_syx_list()
        if not _syl:
            st.info("Još nema objavljenih SYX izveštaja.")
            return
        import base64 as _b64y
        for _r in _syl:
            _mk = _r.get("mesec", "")
            with st.container(border=True):
                _sc1, _sc2 = st.columns([3, 1])
                with _sc1:
                    st.markdown("<div style='font-size:15px;font-weight:800;'>" + _h_escape(mesec_label(_mk)) + "</div>"
                                "<div style='font-size:12.5px;color:#9aa0ad;margin-top:2px;'>" + _h_escape(str(_r.get("filename", ""))) + "</div>",
                                unsafe_allow_html=True)
                with _sc2:
                    _doc = sb_syx_get(_mk)
                    if _doc and _doc.get("docx_b64"):
                        try:
                            _fn = str(_r.get("filename") or ("Izvestaj_SYX_" + _mk + ".docx"))
                            st.download_button("⬇️ Preuzmi", _b64y.b64decode(_doc["docx_b64"]),
                                file_name=_fn,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key="syx_dl_" + _mk, use_container_width=True)
                        except Exception:
                            st.caption("Greška pri učitavanju dokumenta.")
        return

    # ---------- KARTICA 1: EFIKASNOST ADMINISTRACIJE ----------
    if _view == "efikasnost":
        st.markdown('<div style="font-size:20px;font-weight:800;margin:6px 0 14px;">📊 Izveštaj efikasnosti administracije</div>', unsafe_allow_html=True)
        _sel_lbl = st.selectbox("Mesec", _mlbls, index=_prev_idx, key="dir_efik_mes")
        mesec_key = _mes_keys[_mlbls.index(_sel_lbl)]

        # Zaključavanje: izveštaj efikasnosti se vidi TEK kad prođe rok administracije.
        _rok_a = sb_rokovi_get(mesec_key).get("rok_admin")
        if _rok_a and not _rok_je_prosao(_rok_a):
            st.info("Izveštaj efikasnosti za " + _sel_lbl + " biće dostupan posle roka administracije: "
                    + _rok_fmt(_rok_a) + ". Do tada administracija još obrađuje podatke.")
            return

        _pc1, _pc2 = st.columns([1, 2])
        with _pc1:
            if st.button("📄 Napravi PDF izveštaj", key="dir_pdf_make", use_container_width=True):
                try:
                    st.session_state["dir_pdf"] = napravi_pdf_izvestaj(mesec_key, _sel_lbl)
                    st.session_state["dir_pdf_for"] = mesec_key
                except Exception as _e:
                    st.session_state["dir_pdf"] = None
                    st.error("Greška pri pravljenju PDF-a: " + str(_e))
        with _pc2:
            if st.session_state.get("dir_pdf") and st.session_state.get("dir_pdf_for") == mesec_key:
                st.download_button("⬇️ Preuzmi PDF", st.session_state["dir_pdf"],
                    file_name="Izvestaj_administracije_" + mesec_key + ".pdf", mime="application/pdf",
                    key="dir_pdf_dl", use_container_width=True)

        st.markdown('<div style="margin:18px 0 4px;font-size:12px;text-transform:uppercase;letter-spacing:.6px;'
                    'color:#9aa0ad;font-weight:700;">⚠️ Upozorenja prosleđena od administracije</div>', unsafe_allow_html=True)
        st.caption("Objekti koje je administracija označila da su prosleđeni komercijali — problem, koliko treba/koliko je poručeno i šta je preduzeto.")

        _kf = st.session_state.get("_komfull_dir")
        if _kf is None:
            _kf = sb_komitenti_full(); st.session_state["_komfull_dir"] = _kf

        def _chip(on, txt):
            _bg = "#dcfce7;color:#166534" if on else "#f3f4f6;color:#9ca3af"
            return '<span style="background:' + _bg + ';padding:4px 11px;border-radius:20px;font-weight:600;font-size:12px;">' + txt + '</span>'

        _found = 0
        for _sis in sb_sisteme(mesec_key):
            _pod = sb_ucitaj(mesec_key, _sis)
            if not _pod or not _pod.get("stavke"):
                continue
            _po = {}
            for s in _pod["stavke"]:
                _po.setdefault(int(s["idk"]), []).append(s)
            # Sistemski/nedeljni sistem — prosleđuje se preko štiklice „Prijavi problem"
            _meta_d = _pod.get("meta") or {}
            if _meta_d.get("nedeljni"):
                _dani_d = int(_meta_d.get("nedeljni_dani", 7) or 7)
                _perl_d = str(_dani_d) + " dana"
                _cut_d = _admin_presek(_meta_d, mesec_key)
                _ahist_d = (_meta_d.get("admin_hist") or {}) if isinstance(_meta_d, dict) else {}
                _prij_d = dict(_meta_d.get("nedeljni_prijave") or {})
                for _idk, _info in _prij_d.items():
                    try:
                        _ik = int(_idk)
                    except Exception:
                        continue
                    _lst = _po.get(_ik) or []
                    _pm_d = _treb_posle_preseka(_ahist_d.get(str(_ik), []) or [], _cut_d) if _cut_d else {}
                    _pa = []
                    for s in _lst:
                        _por = int(_pm_d.get(int(s.get('ida', -1)), 0) or 0)
                        _lg = int(s.get('lager', 0) or 0) + _por
                        _pr = int(s.get('pred', 0) or 0)
                        _prag = _pr * _dani_d / 30.0
                        if _pr > 0 and _lg < _prag:
                            _pa.append({"naziv": str(s.get('naziv', '')), "lager": _lg, "prag": int(round(_prag))})
                    _found += 1
                    _ki = _kf.get(_ik, {}) or {}
                    _naz = _ki.get("naziv", "") or ("ID " + str(_ik))
                    _cbits = []
                    if _ki.get("telefon"): _cbits.append("📞 " + _h_escape(str(_ki["telefon"])))
                    if _ki.get("email"): _cbits.append("✉️ " + _h_escape(str(_ki["email"])))
                    _cline = ("&nbsp;&nbsp;·&nbsp;&nbsp;".join(_cbits))
                    _artrows = "".join('<div style="font-size:12.5px;color:#4b5563;padding:2px 0;">• '
                                       + _h_escape(a["naziv"]) + ' — <b>realni lager ' + str(a["lager"])
                                       + '</b> (za ' + _perl_d + ' ~' + str(a["prag"]) + ')</div>' for a in _pa) or \
                               '<div style="font-size:12.5px;color:#9aa0ad;">—</div>'
                    _napd = (_info.get("napomena", "") or "")
                    _atd = _info.get("at", "")
                    _kod = _info.get("ko", "")
                    _html = ('<div style="background:#fff;border:1px solid #efeaf7;border-left:4px solid #f59e0b;'
                             'border-radius:14px;padding:16px 18px;margin-bottom:12px;box-shadow:0 2px 12px rgba(80,40,140,.05);">'
                             '<div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:6px;">'
                             '<span style="font-weight:700;font-size:14.5px;">' + _h_escape(_naz) + '</span>'
                             '<span style="font-size:12px;color:#9aa0ad;">· ' + _h_escape(_sis) + ' (nedeljni)</span>'
                             '<span style="margin-left:auto;font-size:12px;font-weight:700;color:#b45309;">⚠️ Prijavljen problem'
                             + ((' · ' + _h_escape(str(_kod))) if _kod else '') + '</span></div>'
                             + (('<div style="font-size:12.5px;color:#6b7280;margin-bottom:9px;">' + _cline + '</div>') if _cline else '')
                             + '<div style="background:#faf9fd;border-radius:10px;padding:9px 12px;margin-bottom:'
                             + ('9px;' if _napd else '2px;') + '">'
                             '<div style="font-size:10.5px;color:#9aa0ad;text-transform:uppercase;font-weight:700;margin-bottom:3px;">Artikli u problemu</div>'
                             + _artrows + '</div>'
                             + (('<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:9px;padding:8px 12px;'
                                 'font-size:12.5px;color:#78500a;"><b style="color:#4b5563;">Napomena:</b> ' + _h_escape(_napd)
                                 + (('<span style="color:#9aa0ad;"> · ' + _h_escape(_atd) + '</span>') if _atd else '')
                                 + '</div>') if _napd else '')
                             + '</div>')
                    st.markdown(_html, unsafe_allow_html=True)
                continue
            _obr = sb_load_obrada(mesec_key, _sis)
            for _idk, _v in _obr.items():
                if "Obavestila direktorku" not in (_v.get("reakcije") or []):
                    continue
                _lst = _po.get(int(_idk))
                if not _lst:
                    continue
                _found += 1
                nivo, n_nula, izgub = hitnost_objekta(_lst)
                z = _zona_disp(nivo)
                _naz = (_kf.get(int(_idk), {}) or {}).get("naziv", "") or ("ID " + str(_idk))
                _treba = sum(int(a.get("kol", 0) or 0) for a in _lst)
                _njih = sum(int(x) for x in (_v.get("njihova") or {}).values())
                _reak = _v.get("reakcije") or []
                _rko = _v.get("reakcije_ko") or {}
                def _ko_suf(_rr):
                    _k = _rko.get(_rr, "")
                    return (" (" + _h_escape(str(_k)) + ")") if _k else ""
                _nap = _v.get("napomena", "") or ""
                _bc = "#ef4444" if nivo == "crveno" else ("#f59e0b" if nivo == "zuto" else "#22c55e")
                _html = ('<div style="background:#fff;border:1px solid #efeaf7;border-left:4px solid ' + _bc + ';'
                         'border-radius:14px;padding:16px 18px;margin-bottom:12px;box-shadow:0 2px 12px rgba(80,40,140,.05);">'
                         '<div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:8px;">'
                         '<span style="font-weight:700;font-size:14.5px;">' + _h_escape(_naz) + '</span>'
                         '<span style="font-size:12px;color:#9aa0ad;">· ' + _h_escape(_sis) + '</span>'
                         '<span style="margin-left:auto;font-size:12px;font-weight:700;color:' + _bc + ';">' + z[3] + '</span></div>'
                         '<div style="font-size:13px;color:#4b5563;margin-bottom:11px;">Problem: <b>' + str(n_nula)
                         + '</b> artikala na nuli · procenjeno izgubljeno <b>' + str(izgub) + '</b> kom/mesec.</div>'
                         '<div style="display:grid;grid-template-columns:repeat(2,1fr);gap:10px;margin-bottom:11px;">'
                         '<div style="background:#faf9fd;border-radius:10px;padding:8px 12px;">'
                         '<div style="font-size:10.5px;color:#9aa0ad;text-transform:uppercase;font-weight:700;">Treba da poruči</div>'
                         '<div style="font-size:16px;font-weight:800;color:#7c3aed;">' + _fmt(_treba) + ' kom</div></div>'
                         '<div style="background:#faf9fd;border-radius:10px;padding:8px 12px;">'
                         '<div style="font-size:10.5px;color:#9aa0ad;text-transform:uppercase;font-weight:700;">Poručio (njihovo)</div>'
                         '<div style="font-size:16px;font-weight:800;color:' + ("#16a34a" if _njih > 0 else "#dc2626") + ';">' + _fmt(_njih) + ' kom</div></div>'
                         '</div>'
                         '<div style="display:flex;gap:8px;flex-wrap:wrap;' + ('margin-bottom:9px;' if _nap else '') + '">'
                         + _chip("Pozvala sam" in _reak, "📞 Pozvano" + (_ko_suf("Pozvala sam") if "Pozvala sam" in _reak else ""))
                         + _chip("Poslala sam mejl" in _reak, "✉️ Mejl" + (_ko_suf("Poslala sam mejl") if "Poslala sam mejl" in _reak else ""))
                         + _chip(True, "👤 Prosleđeno komercijali" + (_ko_suf("Obavestila direktorku") if "Obavestila direktorku" in _reak else ""))
                         + '</div>'
                         + (('<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:9px;padding:8px 12px;'
                             'font-size:12.5px;color:#78500a;"><b style="color:#4b5563;">Napomena:</b> ' + _h_escape(_nap) + '</div>') if _nap else '')
                         + '</div>')
                st.markdown(_html, unsafe_allow_html=True)
        if _found == 0:
            st.caption("Nema upozorenja prosleđenih komercijali za ovaj mesec.")
        return

    # ---------- KARTICA 2: DETALJAN IZVEŠTAJ PO SISTEMIMA ----------
    if _view == "sistemi":
        st.markdown('<div style="font-size:20px;font-weight:800;margin:6px 0 12px;">📈 Detaljan izveštaj po sistemima</div>', unsafe_allow_html=True)
        _c1, _c2 = st.columns(2)
        with _c1:
            _sel_lbl = st.selectbox("Mesec", _mlbls, index=_prev_idx, key="dir_sis_mes")
        mesec_key = _mes_keys[_mlbls.index(_sel_lbl)]
        _sisteme = sb_sisteme(mesec_key)
        if not _sisteme:
            _rs = sb_rokovi_get(mesec_key).get("rok_sistemi")
            if _rs and not _rok_je_prosao(_rs):
                st.info("Izveštaj po sistemu za " + _sel_lbl + " još nije objavljen — biće objavljen najkasnije do "
                        + _rok_fmt(_rs) + " (rok za popunjavanje još nije istekao).")
            elif _rs:
                st.warning("Izveštaj po sistemu za " + _sel_lbl + " nije objavljen, a rok ("
                           + _rok_fmt(_rs) + ") je istekao.")
            else:
                st.info("Za " + _sel_lbl + " još nema objavljenih sistema (rok nije postavljen).")
            return
        with _c2:
            sistem = st.selectbox("Sistem", _sisteme, index=0, key="dir_sis_sis")

        st.markdown("<div style='margin:8px 0 14px;font-size:12px;text-transform:uppercase;letter-spacing:.6px;"
                    "color:#9aa0ad;font-weight:700;'>" + _h_escape(sistem) + " · " + _sel_lbl + "</div>",
                    unsafe_allow_html=True)

        podaci = sb_ucitaj(mesec_key, sistem)
        if not podaci:
            _rs = sb_rokovi_get(mesec_key).get("rok_sistemi")
            if _rs:
                st.info("Izveštaj za „" + _h_escape(str(sistem)) + "“ (" + _sel_lbl
                        + ") biće objavljen najkasnije do " + _rok_fmt(_rs) + ".")
            else:
                st.info("Izveštaj za ovaj sistem/mesec još nije objavljen.")
            return

        # Preuzimanje analitike (Excel, bez sheeta o modelu)
        _dl1, _dl2 = st.columns([1, 2])
        with _dl1:
            if st.button("⬇️ Pripremi analitiku (Excel)", key="dir_prep_xlsx", use_container_width=True):
                _b = sb_ucitaj_xlsx(mesec_key, sistem)
                if _b:
                    import base64 as _b64
                    try:
                        st.session_state["dir_xlsx_bytes"] = _b64.b64decode(_b)
                    except Exception:
                        st.session_state["dir_xlsx_bytes"] = None
                else:
                    st.session_state["dir_xlsx_bytes"] = None
                st.session_state["dir_xlsx_for"] = (mesec_key, sistem)
        with _dl2:
            if st.session_state.get("dir_xlsx_for") == (mesec_key, sistem):
                _xb = st.session_state.get("dir_xlsx_bytes")
                if _xb:
                    st.download_button("⬇️ Sačuvaj Analitika.xlsx", _xb,
                        file_name="Analitika_" + str(sistem).replace(" ", "_") + "_" + mesec_key + ".xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dir_dl_xlsx", use_container_width=True)
                else:
                    st.caption("Analitika (Excel) za ovaj sistem biće dostupna čim se sistem ponovo objavi.")
        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

        d = podaci.get("direktor") or {}
        if not isinstance(d, dict):
            d = {}
        _pf = d.get("profit")  # profitabilnost iz sačuvane analitike (ako je sistem objavljen novom verzijom)

        # Prodaju/trend/grupe uzimamo prvenstveno iz tabele prodaje (18 meseci, lepši trend);
        # profitabilnost uvek iz sačuvane analitike ovog sistema.
        _izvp = sb_ucitaj_izvestaj_prodaje()
        try:
            _sales = json.loads(_izvp["prodaja_json"]) if (_izvp and _izvp.get("prodaja_json")) else {}
        except Exception:
            _sales = {}
        _dsales = _direktor_blok_iz_prodaje(sistem, _sales)

        # Iz analitike (objava sistema): prosek po objektu, OOS, porudžbina, artikli, akcija, profit
        _extra_keys = ["profit", "prosek_po_objektu", "oos_kom", "porudzbina", "artikli_rang", "akcija"]
        _has_analitika = any(d.get(k) for k in _extra_keys)

        def _pripoji(base):
            for _k in _extra_keys:
                if d.get(_k):
                    base[_k] = d.get(_k)
            return base

        if _dsales:
            _part = _partial_iz_stavki(podaci.get("stavke") or [])
            _dsales["oos"] = _part.get("oos")
            _dsales["oos_po_artiklu"] = _part.get("oos_po_artiklu")
            _pripoji(_dsales)
            st.caption("Prodaja i grupe su iz tabele prodaje; ostalo (porudžbina, artikli, akcija, OOS, profit) je iz analitike sistema.")
            _render_sistem_report(_dsales, True)
        elif d.get("prodaja_trend"):
            _render_sistem_report(_pripoji(d), True)
        else:
            _base = _pripoji(_partial_iz_stavki(podaci.get("stavke") or []))
            if not _has_analitika:
                st.info("Prodaja i trend se pojave kad objaviš Izveštaj prodaje (ako tabela prodaje sadrži ovaj sistem), "
                        "ili kad ponovo objaviš ovaj sistem. Ispod je što je već dostupno.")
            _render_sistem_report(_base, False)

        if not _has_analitika:
            st.markdown("<div style='margin-top:10px;padding:11px 14px;background:#fff7ed;border:1px solid #fed7aa;"
                        "border-radius:10px;font-size:12.5px;color:#9a5b1e;'>Detaljni delovi (prosek po objektu, predlog "
                        "porudžbine, bestseleri, uspešnost akcije, OOS po količinama, profitabilnost) se pojave čim ponovo "
                        "objaviš ovaj sistem novom verzijom aplikacije. Fajl je isti kao i do sad.</div>", unsafe_allow_html=True)
        return

    # ---------- KARTICA 3: IZVEŠTAJ PRODAJE (dashboard, pun ekran) ----------
    if _view == "prodaja":
        # Proširi na pun ekran (samo ovaj prikaz)
        st.markdown("<style>.block-container{max-width:100% !important;"
                    "padding-left:1rem !important;padding-right:1rem !important;padding-top:10px !important;}</style>",
                    unsafe_allow_html=True)
        _izv = sb_ucitaj_izvestaj_prodaje()
        if not _izv or not _izv.get("html"):
            st.info("Izveštaj prodaje još nije objavljen. Analitičar ga pravi u delu Objava izveštaja → Izveštaj prodaje.")
            return
        _mlbl = _izv.get("mesec_label") or ""
        _gen = _izv.get("generisano") or ""
        _tc1, _tc2 = st.columns([4, 1])
        with _tc1:
            st.caption("💹 Izveštaj prodaje · poslednji mesec " + str(_mlbl)
                       + ("  ·  generisano " + str(_gen) if _gen else ""))
        with _tc2:
            _xb64 = _izv.get("xlsx_b64")
            if _xb64:
                try:
                    import base64 as _b64d
                    st.download_button("⬇️ Izvezi u Excel", _b64d.b64decode(_xb64),
                        file_name="Izvestaj_prodaje.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dir_prod_xlsx", use_container_width=True)
                except Exception:
                    pass
        # Rok za osvežavanje + napomena (u uglu) — iz rokova tekućeg meseca
        _tt = datetime.date.today()
        _curmk = str(_tt.year) + "-" + ("0" + str(_tt.month))[-2:]
        _rp = sb_rokovi_get(_curmk)
        _rpd = _rp.get("rok_prodaja")
        _rpn = (_rp.get("napomena") or "").strip()
        if _rpd or _rpn:
            _corner = ('<div style="display:flex;justify-content:flex-end;margin:-4px 0 10px;">'
                       '<div style="max-width:520px;background:#fff7ed;border:1px solid #fed7aa;border-radius:12px;'
                       'padding:10px 14px;font-size:12.5px;color:#9a5b1e;">')
            if _rpd:
                _corner += '⏰ Rok za osvežavanje podataka: <b>' + _rok_fmt(_rpd) + '</b>'
            if _rpn:
                _corner += ('<div style="margin-top:5px;color:#7c5320;"><b>Napomena:</b> ' + _h_escape(_rpn) + '</div>')
            _corner += '</div></div>'
            st.markdown(_corner, unsafe_allow_html=True)
        import streamlit.components.v1 as _components
        _components.html(_izv["html"], height=2600, scrolling=True)
        return


# =====================================================================
# ROUTER: prijava -> uloga
# =====================================================================
if not check_password():
    st.stop()

if st.session_state.get("role") == "administracija":
    prikazi_administraciju()
    st.stop()

if st.session_state.get("role") == "direktori":
    prikazi_direktore()
    st.stop()

if st.session_state.get("role") == "komercijala":
    prikazi_komercijalu()
    st.stop()

# --- od ovde nadole ide ANALITIČKI deo (pun pristup) ---

WMA_WEIGHTS = np.array([0.03, 0.07, 0.12, 0.28, 0.50])
HIST_WEIGHT = 0.03

class PredictionEngine:
    def __init__(self, file_bytes, excluded_ids, alpha, beta, min_lager, min_order, mesecni_trosak=0, analitika_meseci=None, min_per_artikal=None, meseci=1.0, max_per_artikal=None, syx_objekti=None):
        self.file_bytes = file_bytes; self.excluded = excluded_ids
        self.alpha = alpha; self.beta = beta; self.min_lager = min_lager; self.min_order = min_order
        self.min_per_artikal = min_per_artikal
        self.max_per_artikal = max_per_artikal
        self.syx_objekti = syx_objekti  # set ID-jeva koji smeju SYX; None = bez ograničenja
        self.meseci = meseci if (meseci and meseci > 0) else 1.0
        self.mesecni_trosak = mesecni_trosak
        self.analitika_meseci = analitika_meseci
        self.logs = []; self.adjustments = []; self.has_history = False
        self.has_prices = False
    def log(self, msg): self.logs.append(msg)
    def run(self, progress_bar):
        progress_bar.progress(5, "Ucitavanje..."); self._load_sheets()
        progress_bar.progress(15, "Priprema..."); self._prepare_lookups()
        progress_bar.progress(25, "Povrat/korekcija..."); self._compute_povrat()
        progress_bar.progress(40, "Mesecni pregled..."); self._build_monthly()
        progress_bar.progress(55, "Predikcija..."); self._predict_all()
        progress_bar.progress(70, "Lager..."); self._merge_lager()
        progress_bar.progress(80, "Porudzbine..."); self._compute_orders()
        progress_bar.progress(85, "Analitika..."); self._apply_min_order(); self._compute_analytics()
        progress_bar.progress(100, "Gotovo!"); return self.df_result
    def _load_sheets(self):
        xls = pd.ExcelFile(io.BytesIO(self.file_bytes))
        sm = {s.strip().lower(): s for s in xls.sheet_names}
        def find(kws):
            for kw in kws:
                for nl, no in sm.items():
                    if kw in nl: return no
            return None
        s_prod=find(['prodaja']); s_start=find(['startni']); s_pov=find(['povrat'])
        s_tl=find(['trenutni']); s_hist=find(['pre sept','pre sep','istorij'])
        if not s_prod: raise ValueError("Nema sheeta 'prodaja'!")
        if not s_start: raise ValueError("Nema sheeta 'startni lager'!")
        self.prodaja = pd.read_excel(xls, sheet_name=s_prod); self.prodaja.columns=[c.strip() for c in self.prodaja.columns]
        self.prodaja = self.prodaja[[c for c in self.prodaja.columns if 'Unnamed' not in str(c)]]
        self.log(f"Prodaja: {len(self.prodaja)} redova")
        self.region_map = {}
        if 'Region' in self.prodaja.columns:
            self.region_map = self.prodaja.drop_duplicates('ID KOMITENTA').set_index('ID KOMITENTA')['Region'].to_dict()
            self.log(f"Region: {len(set(self.region_map.values()))} regiona")
        self.startni = pd.read_excel(xls, sheet_name=s_start); self.startni.columns=[c.strip() for c in self.startni.columns]
        self.log(f"Startni: {len(self.startni)} redova")
        price_cols = ['Redovna cena','Akcijska cena','Finalna cena','Nabavna vrednost','Profit']
        if all(c in self.prodaja.columns for c in price_cols):
            self.has_prices = True; self.log("Cene i profit: DA")
        self.povrat_df = pd.DataFrame()
        if s_pov:
            self.povrat_df = pd.read_excel(xls, sheet_name=s_pov); self.povrat_df.columns=[c.strip() for c in self.povrat_df.columns]
            self.log(f"Povrat: {len(self.povrat_df)} redova")
        self.trenutni = pd.DataFrame()
        if s_tl:
            self.trenutni = pd.read_excel(xls, sheet_name=s_tl); self.trenutni.columns=[c.strip() for c in self.trenutni.columns]
            self.log(f"Trenutni lager: {len(self.trenutni)} redova")
        self.hist_df = pd.DataFrame()
        self.has_history = False
        _meseci_u_prodaji = self.prodaja[['Godina','Mesec']].drop_duplicates().values.tolist()
        _ima_pre_sept = any((int(g) < 2025) or (int(g) == 2025 and int(m) < 9) for g, m in _meseci_u_prodaji)
        if _ima_pre_sept:
            self.log("Rezim: KOMPLETAN ISTORIJAT u prodaja sheetu — istorijski sheet se ignorise")
        elif s_hist:
            _h = pd.read_excel(xls, sheet_name=s_hist)
            _h.columns = [str(c).strip() for c in _h.columns]
            # normalizuj nazive kolona na standard (razni fajlovi imaju sitne razlike)
            _cm = {}
            for c in _h.columns:
                cl = str(c).lower()
                if 'komitent' in cl: _cm[c] = 'ID KOMITENTA'
                elif 'id' in cl and 'artikl' in cl: _cm[c] = 'id artikla'
                elif 'naziv' in cl and 'artikl' in cl: _cm[c] = 'Naziv artikla'
                elif 'grup' in cl: _cm[c] = 'Grupa'
                elif 'prodat' in cl: _cm[c] = 'Prodata Kolicina'
                elif cl.startswith('kolicin') or cl.startswith('količin'): _cm[c] = 'Prodata Kolicina'
                elif cl.startswith('mesec'): _cm[c] = 'Mesec'
                elif cl.startswith('godina'): _cm[c] = 'Godina'
            _h = _h.rename(columns=_cm)
            # istorija je validna samo ako ima redove i potrebne kolone; inače se ignoriše (nije greška)
            if len(_h) > 0 and 'ID KOMITENTA' in _h.columns and 'Prodata Kolicina' in _h.columns:
                self.hist_df = _h
                self.has_history = True
                self.log(f"Istorija: {len(self.hist_df)} redova")
            else:
                self.hist_df = pd.DataFrame()
                self.has_history = False
                self.log("Istorijski sheet je prazan ili nema potrebne kolone — ignoriše se.")
        self.meseci_order = sorted(self.prodaja[['Godina','Mesec']].drop_duplicates().values.tolist())
        mn={1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'Maj',6:'Jun',7:'Jul',8:'Avg',9:'Sep',10:'Okt',11:'Nov',12:'Dec'}
        self.mesec_labels = [f"{mn.get(int(m),'?')} {int(g)}" for g,m in self.meseci_order]
        lg,lm = self.meseci_order[-1]; nm=int(lm)+1; ng=int(lg)
        if nm>12: nm=1; ng+=1
        self.pred_label = f"{mn.get(nm,'?')} {ng}"
        # Porudzbina je za isti mesec kao i predikcija (poslednji mesec podataka + 1)
        self.order_label = self.pred_label
        self.log(f"Meseci: {', '.join(self.mesec_labels)}")
        self.num_komitenti = self.prodaja['ID KOMITENTA'].nunique()
        self.trosak_po_objektu = self.mesecni_trosak / max(self.num_komitenti, 1) if self.mesecni_trosak > 0 else 0
        if self.mesecni_trosak > 0:
            self.log(f"Ukupan trosak: {self.mesecni_trosak:,.0f} / {self.num_komitenti} objekata = {self.trosak_po_objektu:,.0f} po objektu za period")
    def _prepare_lookups(self):
        kp = self.prodaja[['ID KOMITENTA','id artikla','Naziv artikla','Grupa']].drop_duplicates()
        ks = self.startni[['ID KOMITENTA','id artikla','Naziv artikla','Grupa']].drop_duplicates()
        frames = [kp, ks]
        if self.has_history:
            hcols = self.hist_df.columns.tolist()
            col_map = {}
            for c in hcols:
                cl = c.lower()
                if 'komitent' in cl: col_map[c] = 'ID KOMITENTA'
                elif 'id' in cl and 'artikl' in cl: col_map[c] = 'id artikla'
                elif 'naziv' in cl and 'artikl' in cl: col_map[c] = 'Naziv artikla'
                elif 'grup' in cl: col_map[c] = 'Grupa'
            hdf = self.hist_df.rename(columns=col_map)
            for nc in ['ID KOMITENTA','id artikla','Naziv artikla','Grupa']:
                if nc not in hdf.columns: hdf[nc] = ''
            kh = hdf[['ID KOMITENTA','id artikla','Naziv artikla','Grupa']].drop_duplicates()
            frames.append(kh)
        self.all_keys = pd.concat(frames).drop_duplicates().sort_values(['ID KOMITENTA','id artikla']).reset_index(drop=True)
        self.startni_dict = {(r['ID KOMITENTA'],r['id artikla']): r['Kolicina'] for _,r in self.startni.iterrows()}
        self.has_promet = 'PROMET KA NJIMA' in self.prodaja.columns
        self.prodaja_dict = {}
        for _,r in self.prodaja.iterrows():
            key=(r['ID KOMITENTA'],r['id artikla'],r['Godina'],r['Mesec'])
            pm = r['PROMET KA NJIMA'] if self.has_promet else 0
            self.prodaja_dict[key] = (r.get('Prodata Kolicina',r.get('Kolicina',0)), r.get('Lager',0), pm if not pd.isna(pm) else 0)
        self.hist_dict={}; self.hist_total_dict={}; self.hist_months_per_art={}
        if self.has_history:
            ha = self.hist_df.groupby(['ID KOMITENTA','id artikla'])['Prodata Kolicina'].agg(['sum','mean']).reset_index()
            for _,r in ha.iterrows():
                self.hist_dict[(int(r['ID KOMITENTA']),int(r['id artikla']))] = float(r['mean'])
                self.hist_total_dict[(int(r['ID KOMITENTA']),int(r['id artikla']))] = int(r['sum'])
            for ida in self.hist_df['id artikla'].unique():
                sub=self.hist_df[self.hist_df['id artikla']==ida]
                self.hist_months_per_art[int(ida)]=sub[['Godina','Mesec']].drop_duplicates().shape[0]
            self.log(f"Istorijski prosek za {len(self.hist_dict)} kombinacija")
        self.recent_months_per_art={}
        for ida in self.prodaja['id artikla'].unique():
            sub=self.prodaja[self.prodaja['id artikla']==ida]
            self.recent_months_per_art[int(ida)]=sub[['Godina','Mesec']].drop_duplicates().shape[0]
        self.total_months_per_art={}
        all_arts=set([int(x) for x in self.prodaja['id artikla'].unique()])
        if self.has_history: all_arts|=set([int(x) for x in self.hist_df['id artikla'].unique()])
        for ida in all_arts:
            self.total_months_per_art[ida]=self.hist_months_per_art.get(ida,0)+self.recent_months_per_art.get(ida,0)
        self.povrat_total={}
        if len(self.povrat_df)>0:
            ic=[c for c in self.povrat_df.columns if 'id' in c.lower() and 'artikl' in c.lower()]
            mc=[c for c in self.povrat_df.columns if 'mesec' in c.lower()]
            gc=[c for c in self.povrat_df.columns if 'godin' in c.lower()]
            kc=[c for c in self.povrat_df.columns if 'koli' in c.lower()]
            if ic and mc and gc and kc:
                for _,r in self.povrat_df.iterrows():
                    key=(r[ic[0]],r[gc[0]],r[mc[0]]); self.povrat_total[key]=self.povrat_total.get(key,0)+r[kc[0]]
        self.trenutni_dict={}
        if len(self.trenutni)>0:
            ikc=[c for c in self.trenutni.columns if 'komitent' in c.lower()]
            iac=[c for c in self.trenutni.columns if 'artikl' in c.lower() and 'id' in c.lower()]
            lc=[c for c in self.trenutni.columns if 'lager' in c.lower()]
            if ikc and iac and lc:
                for _,r in self.trenutni.iterrows():
                    k,a=r[ikc[0]],r[iac[0]]
                    if pd.notna(k) and pd.notna(a): self.trenutni_dict[(int(k),int(a))]=int(r[lc[0]]) if pd.notna(r[lc[0]]) else 0
        self.profit_per_unit = {}
        self.price_info = {}
        if self.has_prices:
            for ida in self.prodaja['id artikla'].unique():
                sub = self.prodaja[self.prodaja['id artikla']==ida].iloc[0]
                red, akc, fin, nab = sub['Redovna cena'], sub['Akcijska cena'], sub['Finalna cena'], sub['Nabavna vrednost']
                ppu_fin = fin/1.2/1.2 - nab
                ppu_red = red/1.2/1.2 - nab
                self.profit_per_unit[int(ida)] = ppu_fin
                self.price_info[int(ida)] = {'redovna': red, 'akcijska': akc, 'finalna': fin, 'nabavna': nab, 'profit_akcija': ppu_fin, 'profit_redovna': ppu_red}
        self.log(f"Kombinacija: {len(self.all_keys)}")
    def _compute_povrat(self):
        self.final_povrat={}; self.final_korekcija={}
        if not self.has_promet or not self.povrat_total: return
        implied={}
        for _,k in self.all_keys.iterrows():
            idk,ida=k['ID KOMITENTA'],k['id artikla']; poc=self.startni_dict.get((idk,ida),0)
            for god,mes in self.meseci_order:
                pv,lv,tv=self.prodaja_dict.get((idk,ida,god,mes),(0,0,0)); lv=lv if not pd.isna(lv) else 0
                implied[(idk,ida,god,mes)]=poc+tv-pv-lv; poc=lv
        all_art=set(list(self.prodaja['id artikla'].unique())+list(self.startni['id artikla'].unique()))
        for god,mes in self.meseci_order:
            for ida in all_art:
                ap=self.povrat_total.get((ida,god,mes),0); pi={}; ni={}
                for _,k in self.all_keys[self.all_keys['id artikla']==ida].iterrows():
                    i2=k['ID KOMITENTA']; im=implied.get((i2,ida,god,mes),0)
                    if im>0: pi[i2]=im
                    elif im<0: ni[i2]=im
                tp=sum(pi.values())
                if ap>0 and tp>0:
                    raw={i:ap*(v/tp) for i,v in pi.items()}; fl={i:int(v) for i,v in raw.items()}
                    d=ap-sum(fl.values()); rem={i:raw[i]-fl[i] for i in raw}
                    for j,i in enumerate(sorted(rem,key=rem.get,reverse=True)):
                        if j<int(d): fl[i]+=1
                    for i,pv2 in fl.items():
                        self.final_povrat[(i,ida,god,mes)]=pv2; self.final_korekcija[(i,ida,god,mes)]=pi[i]-pv2
                elif ap==0:
                    for i,v in pi.items(): self.final_korekcija[(i,ida,god,mes)]=v
                for i,v in ni.items(): self.final_korekcija[(i,ida,god,mes)]=self.final_korekcija.get((i,ida,god,mes),0)+v
    def _build_monthly(self):
        rows=[]
        for _,k in self.all_keys.iterrows():
            idk,ida=k['ID KOMITENTA'],k['id artikla']; poc=self.startni_dict.get((idk,ida),0)
            row={'ID KOMITENTA':idk,'id artikla':ida,'Naziv artikla':k['Naziv artikla'],'Grupa':k['Grupa']}
            row['Total_JanAvg']=self.hist_total_dict.get((idk,ida),0)
            for i,(god,mes) in enumerate(self.meseci_order):
                lb=self.mesec_labels[i]; pv,lv,tv=self.prodaja_dict.get((idk,ida,god,mes),(0,0,0))
                lv=lv if not pd.isna(lv) else 0; tv=tv if not pd.isna(tv) else 0
                row[f'{lb}_Pocetno']=poc; row[f'{lb}_Promet']=tv; row[f'{lb}_Prodaja']=pv
                row[f'{lb}_Povrat']=self.final_povrat.get((idk,ida,god,mes),0)
                row[f'{lb}_Korekcija']=self.final_korekcija.get((idk,ida,god,mes),0); poc=lv
            rows.append(row)
        self.df_monthly=pd.DataFrame(rows)
    def _predict_all(self):
        analysis=[]
        for _,k in self.all_keys.iterrows():
            idk,ida=k['ID KOMITENTA'],k['id artikla']; poc=self.startni_dict.get((idk,ida),0)
            sales,oos,pocs,end_lagers,promets=[],[],[],[],[]
            for god,mes in self.meseci_order:
                pv,lv,tv=self.prodaja_dict.get((idk,ida,god,mes),(0,0,0))
                lv=lv if not pd.isna(lv) else 0; tv=tv if not pd.isna(tv) else 0
                sales.append(pv); oos.append(1 if poc==0 else 0); pocs.append(poc)
                end_lagers.append(lv); promets.append(tv); poc=lv
            ha=self.hist_dict.get((idk,ida),0)
            lager_danas=self.trenutni_dict.get((idk,ida),0)
            analysis.append({'idk':idk,'ida':ida,'sales':np.array(sales,dtype=float),'oos':np.array(oos),
                'poc':np.array(pocs,dtype=float),'ha':ha,'lager_danas':lager_danas,
                'end_lagers':np.array(end_lagers,dtype=float),'promets':np.array(promets,dtype=float)})
        preds={}
        for it in analysis:
            s,o,p=it['sales'],it['oos'],it['poc']; n=len(s); ha=it['ha']
            lager_danas=it['lager_danas']
            el=it['end_lagers']; tv=it['promets']
            constrained = np.zeros(n, dtype=bool)
            for m in range(n):
                if p[m]==0 and tv[m]==0: constrained[m] = True
                elif el[m]==0 and s[m]>0: constrained[m] = True
                elif p[m]==0 and tv[m]>0 and el[m]==0: constrained[m] = True
            normal_mask = ~constrained & (p > 0)
            normal_sales = s[normal_mask]
            normal_with_sales = normal_sales[normal_sales > 0]
            if len(normal_with_sales) > 0: an = normal_with_sales.mean()
            elif len(normal_sales) > 0: an = normal_sales.mean()
            else: an = 0
            if an > 0:
                adj = s.copy().astype(float)
                for m in range(n):
                    if constrained[m]:
                        if p[m]==0 and tv[m]==0: adj[m] = an
                        elif el[m]==0 and s[m]>0: adj[m] = max(an, s[m])
                        else: adj[m] = an
                    elif p[m]>0 and p[m]<an*0.5: adj[m] = 0.5*s[m] + 0.5*an
            elif ha>0: adj=np.full(n,ha)
            else: adj=s.copy().astype(float)
            if n>=2:
                lev=adj[0]; tr=(adj[-1]-adj[0])/max(n-1,1)
                for i in range(1,n):
                    nl=self.alpha*adj[i]+(1-self.alpha)*(lev+tr); nt=self.beta*(nl-lev)+(1-self.beta)*tr; lev,tr=nl,nt
                holt=lev+tr
            else: holt=adj[0]
            w=WMA_WEIGHTS[-n:] if n<=5 else WMA_WEIGHTS; w=w/w.sum()
            wma=np.dot(adj[-len(w):],w) if n>=3 else adj.mean()
            comb = 0.4 * min(holt, wma) + 0.6 * max(holt, wma)
            ma=adj.mean()
            if ma>0 and n>=3: comb*=(1+min((np.std(adj)/ma)*0.4,0.7))
            if ha>0 and comb>0: comb=(1-HIST_WEIGHT)*comb+HIST_WEIGHT*ha
            elif ha>0 and comb==0 and s.sum()==0: comb=ha*0.20
            has_recent_sales = (s[-2:].sum() > 0) if n >= 2 else (s.sum() > 0)
            if lager_danas <= 2 and has_recent_sales:
                stocked_sales = [s[i] for i in range(n) if p[i] > 0]
                avg_when_stocked = np.mean(stocked_sales) if stocked_sales else 0
                if avg_when_stocked > 0 and comb < avg_when_stocked: comb = avg_when_stocked
            if ma > 5 and comb < ma: comb = ma
            avg_5m_raw = float(adj[-5:].mean()) if n >= 5 else float(adj.mean())
            ht=self.hist_total_dict.get((it['idk'],it['ida']),0)
            rt=float(s.sum()); tm=self.total_months_per_art.get(it['ida'],n)
            full_avg=(ht+rt)/max(tm,1)
            if comb < full_avg and comb > 0:
                if n >= 5: declining = all(adj[i] <= adj[i-1] for i in range(n-4, n))
                elif n >= 3: declining = all(adj[i] <= adj[i-1] for i in range(1, n))
                else: declining = (n >= 2 and adj[-1] <= adj[-2])
                if not declining: comb = full_avg
            if comb <= 0:
                last5 = s[-5:] if n >= 5 else s
                if last5.sum() > 0:
                    comb = 1.0
                    if s[-1] > 1: comb = s[-1]
            preds[(it['idk'],it['ida'])]=(max(0,comb),full_avg,avg_5m_raw)
        items=[{'k':k,'p':v[0],'a':v[1],'avg5':v[2]} for k,v in preds.items()]; df_p=pd.DataFrame(items)
        df_p['pr']=df_p['p'].apply(lambda x: round(x))
        df_p['ar']=df_p['a'].apply(lambda x: round(x))
        self.pred_dict={r['k']:(int(r['pr']),int(r['ar']),int(r['pr']-r['ar']),r['avg5']) for _,r in df_p.iterrows()}
        self.log(f"Predikcija: {sum(v[0] for v in self.pred_dict.values())} kom")
    def _merge_lager(self):
        for _,k in self.all_keys.iterrows():
            idk,ida=k['ID KOMITENTA'],k['id artikla']; pred,avg,razl,avg5m=self.pred_dict.get((idk,ida),(0,0,0,0))
            lager=self.trenutni_dict.get((idk,ida),None)
            idx=self.df_monthly[(self.df_monthly['ID KOMITENTA']==idk)&(self.df_monthly['id artikla']==ida)].index
            if len(idx)>0:
                ix=idx[0]; self.df_monthly.loc[ix,'Predikcija']=pred; self.df_monthly.loc[ix,'Prosek']=avg; self.df_monthly.loc[ix,'Razlika']=razl
                self.df_monthly.loc[ix,'Avg5m']=avg5m
                if lager is not None: self.df_monthly.loc[ix,'Lager_danas']=lager
                else: self.df_monthly.loc[ix,'Lager_danas']=0
        for col in ['Predikcija','Prosek','Razlika','Lager_danas']:
            if col not in self.df_monthly.columns: self.df_monthly[col]=0
            self.df_monthly[col]=self.df_monthly[col].fillna(0).astype(int)
        if 'Avg5m' not in self.df_monthly.columns: self.df_monthly['Avg5m']=0
        self.df_monthly['Avg5m']=self.df_monthly['Avg5m'].fillna(0)
    def _compute_orders(self):
        self.df_result=self.df_monthly.copy()
        def p1(row):
            if row['ID KOMITENTA'] in self.excluded: return 0
            pred_o=int(round(int(row['Predikcija'])*self.meseci))
            return max(pred_o-int(row['Lager_danas']),0)
        def p2(row):
            if row['ID KOMITENTA'] in self.excluded: return 0
            pred=int(row['Predikcija']); pred_o=int(round(pred*self.meseci)); lager=int(row['Lager_danas']); prosek=int(row['Prosek'])
            osnova=max(pred_o-lager,0)
            if self.min_lager is not None and lager < self.min_lager and pred > 0:
                dopuna = max(self.min_lager - lager, osnova)
            else:
                dopuna = osnova
            return dopuna
        self.df_result['Porudzbina_1']=self.df_result.apply(p1,axis=1).astype(int)
        self.df_result['Porudzbina_2']=self.df_result.apply(p2,axis=1).astype(int)
        last_label = self.mesec_labels[-1]
        def extra_buffer(prodaja_poslednji):
            if prodaja_poslednji <= 0: return 0
            elif prodaja_poslednji <= 5: return 2
            elif prodaja_poslednji <= 10: return 3
            elif prodaja_poslednji <= 15: return 4
            else: return 5
        def finalna_provera(row):
            if row['ID KOMITENTA'] in self.excluded: return int(row['Porudzbina_2'])
            p2_val = int(row['Porudzbina_2'])
            lager = int(row['Lager_danas'])
            prodaja_poslednji = int(row.get(f'{last_label}_Prodaja', 0))
            if (p2_val + lager) <= prodaja_poslednji:
                dodatak = extra_buffer(prodaja_poslednji)
                return p2_val + dodatak
            return p2_val
        self.df_result['Porudzbina_2'] = self.df_result.apply(finalna_provera, axis=1).astype(int)
        n_korigovano = (self.df_result['Porudzbina_2'] > self.df_result.apply(p2, axis=1)).sum()
        self.log(f"Finalna provera P2: {n_korigovano} kombinacija korigovano (porudzbina+lager <= prodaja poslednjeg meseca)")
        if self.min_per_artikal is not None and self.min_per_artikal > 1:
            mask_p2 = (
                (self.df_result['Porudzbina_2'] > 0) &
                (self.df_result['Porudzbina_2'] < self.min_per_artikal) &
                (~self.df_result['ID KOMITENTA'].isin(self.excluded))
            )
            n_podignuto_p2 = int(mask_p2.sum())
            self.df_result.loc[mask_p2, 'Porudzbina_2'] = self.min_per_artikal
            mask_p1 = (
                (self.df_result['Porudzbina_1'] > 0) &
                (self.df_result['Porudzbina_1'] < self.min_per_artikal) &
                (~self.df_result['ID KOMITENTA'].isin(self.excluded))
            )
            n_podignuto_p1 = int(mask_p1.sum())
            self.df_result.loc[mask_p1, 'Porudzbina_1'] = self.min_per_artikal
            if n_podignuto_p2 > 0 or n_podignuto_p1 > 0:
                self.log(f"Min po artiklu ({self.min_per_artikal} kom): P1={n_podignuto_p1}, P2={n_podignuto_p2} stavki podignuto na minimum")
        # Maksimum po komadu (po stavci): ograniči porudžbinu po artiklu na zadati maksimum
        if getattr(self, "max_per_artikal", None) is not None and self.max_per_artikal > 0:
            mx = int(self.max_per_artikal)
            cap2 = self.df_result['Porudzbina_2'] > mx
            cap1 = self.df_result['Porudzbina_1'] > mx
            n_cap2 = int(cap2.sum()); n_cap1 = int(cap1.sum())
            self.df_result.loc[cap2, 'Porudzbina_2'] = mx
            self.df_result.loc[cap1, 'Porudzbina_1'] = mx
            if n_cap2 > 0 or n_cap1 > 0:
                self.log(f"Max po artiklu ({mx} kom): P1={n_cap1}, P2={n_cap2} stavki ograničeno na maksimum")
        # SYX vrećice samo u zadatim objektima — u ostalima porudžbina SYX = 0
        _syx = getattr(self, "syx_objekti", None)
        if _syx is not None and 'Grupa' in self.df_result.columns:
            _g = self.df_result['Grupa'].astype(str).str.upper()
            _nz = self.df_result['Naziv artikla'].astype(str).str.upper() if 'Naziv artikla' in self.df_result.columns else _g
            _je_syx = _g.str.contains('SYX') | _nz.str.contains('SYX') | _nz.str.contains('VREĆIC') | _nz.str.contains('VRECIC')
            _van = _je_syx & (~self.df_result['ID KOMITENTA'].isin(_syx))
            _nsyx = int(((self.df_result['Porudzbina_2'] > 0) & _van).sum())
            self.df_result.loc[_van, 'Porudzbina_2'] = 0
            self.df_result.loc[_van, 'Porudzbina_1'] = 0
            self.log(f"SYX ograničenje: dozvoljeno u {len(_syx)} objekata; nulirano u ostalima ({_nsyx} stavki).")
    def _apply_min_order(self):
        self.adjustments = []
        if self.min_order is None or self.min_order <= 0: return
        grp = self.df_result.groupby('ID KOMITENTA')['Porudzbina_2'].sum()
        ima_nesto = grp[grp > 0]
        granica = self.min_order / 2
        premali = ima_nesto[ima_nesto < granica].index
        dopuni = ima_nesto[(ima_nesto >= granica) & (ima_nesto < self.min_order)].index
        mask_gasi = self.df_result['ID KOMITENTA'].isin(premali)
        n_gasi = len(premali)
        self.df_result.loc[mask_gasi, 'Porudzbina_2'] = 0
        n_dopuni = 0
        for komt_id in dopuni:
            mask_obj = (self.df_result['ID KOMITENTA'] == komt_id) & (self.df_result['Porudzbina_2'] > 0)
            ukupno = int(self.df_result.loc[self.df_result['ID KOMITENTA'] == komt_id, 'Porudzbina_2'].sum())
            nedostaje = self.min_order - ukupno
            if nedostaje <= 0 or not mask_obj.any(): continue
            idx_max = self.df_result.loc[mask_obj, 'Porudzbina_2'].idxmax()
            self.df_result.at[idx_max, 'Porudzbina_2'] += nedostaje
            n_dopuni += 1
        if n_gasi > 0:
            self.log(f"Min order ({self.min_order} kom): {n_gasi} objekata imalo premalo komada ukupno — postavljeno na 0")
        if n_dopuni > 0:
            self.log(f"Min order ({self.min_order} kom): {n_dopuni} objekata dopunjeno do minimuma {self.min_order} kom")
    def _compute_analytics(self):
        if not self.has_prices:
            self.df_oos = pd.DataFrame()
            self.df_profit_obj = pd.DataFrame()
            self.df_promo = pd.DataFrame()
            self.analitika_labels = []
            return
        df = self.df_result; ml = self.mesec_labels
        if self.analitika_meseci and len(self.analitika_meseci) > 0:
            a_meseci = self.analitika_meseci
        else:
            a_meseci = self.meseci_order
        a_indices = []
        for i, (g, m) in enumerate(self.meseci_order):
            for ag, am in a_meseci:
                if int(g) == int(ag) and int(m) == int(am):
                    a_indices.append(i); break
        if not a_indices:
            a_indices = list(range(len(self.meseci_order)))
        a_labels = [ml[i] for i in a_indices]
        a_meseci_order = [self.meseci_order[i] for i in a_indices]
        n_a = len(a_indices)
        self.analitika_labels = a_labels
        self.log(f"Analitika period: {', '.join(a_labels)} ({n_a} meseci)")
        a_set = set((int(g), int(m)) for g, m in a_meseci_order)
        prodaja_a = self.prodaja[self.prodaja.apply(lambda r: (int(r['Godina']), int(r['Mesec'])) in a_set, axis=1)]
        ppu_mesec = {}
        if self.has_prices:
            for (ida_v, god_v, mes_v), grp in self.prodaja.groupby(['id artikla','Godina','Mesec']):
                kol = grp['Prodata Kolicina'].sum()
                if kol > 0:
                    ppu_mesec[(int(ida_v), int(god_v), int(mes_v))] = grp['Profit'].sum() / kol
                else:
                    r0 = grp.iloc[0]
                    ppu_mesec[(int(ida_v), int(god_v), int(mes_v))] = r0['Finalna cena'] / 1.2 / 1.2 - r0['Nabavna vrednost']
        def get_ppu(ida_v, god_v, mes_v):
            key = (int(ida_v), int(god_v), int(mes_v))
            if key in ppu_mesec:
                return ppu_mesec[key]
            art_keys = sorted([k for k in ppu_mesec if k[0] == int(ida_v)], key=lambda x: (x[1], x[2]))
            if not art_keys:
                return self.profit_per_unit.get(int(ida_v), 0)
            target = int(god_v) * 12 + int(mes_v)
            best = min(art_keys, key=lambda x: abs(x[1] * 12 + x[2] - target))
            return ppu_mesec[best]
        oos_rows = []
        for _, k in self.all_keys.iterrows():
            idk, ida = k['ID KOMITENTA'], k['id artikla']
            poc = self.startni_dict.get((idk, ida), 0)
            month_sales = []
            month_poc = []
            month_ulaz = []
            month_kraj = []
            month_gm = []
            for i, (god, mes) in enumerate(self.meseci_order):
                lb = ml[i]
                pv_arr = df[(df['ID KOMITENTA']==idk)&(df['id artikla']==ida)][f'{lb}_Prodaja'].values
                pv = int(pv_arr[0]) if len(pv_arr) > 0 else 0
                tv_arr = df[(df['ID KOMITENTA']==idk)&(df['id artikla']==ida)][f'{lb}_Promet'].values
                tv = int(tv_arr[0]) if len(tv_arr) > 0 else 0
                lv_col = self.prodaja_dict.get((idk, ida, god, mes), (0, 0, 0))
                kraj = lv_col[1] if not pd.isna(lv_col[1]) else 0
                if i in a_indices:
                    month_sales.append(pv)
                    month_poc.append(poc)
                    month_ulaz.append(tv)
                    month_kraj.append(kraj)
                    month_gm.append((god, mes))
                poc = kraj
            month_constrained_type = []
            for j in range(len(month_sales)):
                p = month_poc[j]
                u = month_ulaz[j]
                kr = month_kraj[j]
                pr = month_sales[j]
                if p == 0 and u == 0:
                    month_constrained_type.append(1)
                elif p == 0 and u > 0 and kr == 0:
                    month_constrained_type.append(3)
                elif kr == 0 and pr > 0:
                    month_constrained_type.append(2)
                else:
                    month_constrained_type.append(0)
            normal_sales = [month_sales[j] for j in range(len(month_sales))
                            if month_constrained_type[j] == 0 and month_sales[j] > 0]
            avg_stocked = np.mean(normal_sales) if normal_sales else 0
            month_oos_kom = []
            month_oos_flag = []
            total_lost_kom = 0
            for j in range(len(month_sales)):
                t = month_constrained_type[j]
                pr = month_sales[j]
                if avg_stocked == 0:
                    izgub_kom = 0
                elif t == 1:
                    izgub_kom = avg_stocked
                elif t == 2 or t == 3:
                    izgub_kom = max(0, avg_stocked - pr)
                else:
                    izgub_kom = 0
                month_oos_kom.append(izgub_kom)
                month_oos_flag.append(1 if izgub_kom >= 0.5 else 0)
                total_lost_kom += izgub_kom
            oos_count = sum(month_oos_flag)
            if total_lost_kom > 0 and avg_stocked > 0:
                row = {
                    'ID KOMITENTA': idk, 'id artikla': ida,
                    'Naziv artikla': k['Naziv artikla'], 'Grupa': k['Grupa'],
                    'Prosek_kad_ima': round(avg_stocked, 1),
                    'Lager_danas': self.trenutni_dict.get((idk, ida), 0)
                }
                total_lost_rsd = 0
                for j in range(len(month_sales)):
                    god_j, mes_j = month_gm[j]
                    lb_j = a_labels[j]
                    if month_oos_kom[j] > 0:
                        ppu_j = get_ppu(ida, god_j, mes_j)
                        izgub_rsd = round(month_oos_kom[j] * ppu_j, 0)
                        row[f'OOS_{lb_j}'] = round(month_oos_kom[j], 1)
                        row[f'Izgub_{lb_j}'] = izgub_rsd
                        total_lost_rsd += izgub_rsd
                    else:
                        row[f'OOS_{lb_j}'] = 0
                        row[f'Izgub_{lb_j}'] = 0
                row['OOS_meseci'] = oos_count
                row['Izgubljeni_profit'] = round(total_lost_rsd, 0)
                oos_rows.append(row)
        self.df_oos = pd.DataFrame(oos_rows)
        if len(self.df_oos) > 0:
            self.df_oos = self.df_oos.sort_values('Izgubljeni_profit', ascending=False)
            self.log(f"OOS analiza (nova logika - kom izgubljeno): {len(self.df_oos)} kombinacija, izgubljeno {self.df_oos['Izgubljeni_profit'].sum():,.0f} RSD")
        trosak_mes_po_obj = self.trosak_po_objektu / max(n_a, 1) if self.trosak_po_objektu > 0 else 0
        profit_rows = []
        for idk in self.prodaja['ID KOMITENTA'].unique():
            sub = prodaja_a[prodaja_a['ID KOMITENTA'] == idk]
            total_prod = int(sub['Prodata Kolicina'].sum())
            total_profit = sub['Profit'].sum()
            n_art = self.all_keys[self.all_keys['ID KOMITENTA'] == idk]['id artikla'].nunique()
            mes_data = {}
            for _, r in sub.iterrows():
                key = f"{int(r['Godina'])}/{int(r['Mesec'])}"
                mes_data[key] = mes_data.get(key, 0) + r['Profit']
            mes_data_neto = {k: v - trosak_mes_po_obj for k, v in mes_data.items()}
            oos_sub = self.df_oos[self.df_oos['ID KOMITENTA'] == idk] if len(self.df_oos) > 0 else pd.DataFrame()
            lost = oos_sub['Izgubljeni_profit'].sum() if len(oos_sub) > 0 else 0
            trosak_total = self.trosak_po_objektu
            neto = total_profit - trosak_total
            row_dict = {
                'ID KOMITENTA': int(idk), 'Artikala': n_art,
                'Prodato_kom': total_prod, 'Bruto_profit': round(total_profit, 0),
                'Trosak_mkt': round(trosak_total, 0),
                'Neto_profit': round(neto, 0),
                'Izgubljeno_OOS': round(lost, 0),
                'Potencijalni_profit': round(neto + lost, 0),
            }
            for j in range(n_a):
                key_j = f"{int(a_meseci_order[j][0])}/{int(a_meseci_order[j][1])}"
                row_dict[f'Neto_{a_labels[j]}'] = round(mes_data_neto.get(key_j, -trosak_mes_po_obj), 0)
                row_dict[f'Bruto_{a_labels[j]}'] = round(mes_data.get(key_j, 0), 0)
            profit_rows.append(row_dict)
        self.trosak_mes_po_obj = trosak_mes_po_obj
        self.df_profit_obj = pd.DataFrame(profit_rows).sort_values('Neto_profit', ascending=True)
        promo_rows = []
        for ida in self.prodaja['id artikla'].unique():
            pi = self.price_info.get(int(ida), {})
            if not pi: continue
            sub = prodaja_a[prodaja_a['id artikla'] == ida]
            total_prod = int(sub['Prodata Kolicina'].sum())
            if total_prod == 0: continue
            profit_akcija = sub['Profit'].sum()
            profit_redovna = pi['profit_redovna'] * total_prod
            razlika = profit_redovna - profit_akcija
            prihod_akcija = (sub['Finalna cena'] * sub['Prodata Kolicina']).sum()
            prihod_redovna = (sub['Redovna cena'] * sub['Prodata Kolicina']).sum()
            first_a_idx = a_indices[0]
            if first_a_idx == 0:
                start_lager = self.startni[self.startni['id artikla']==ida]['Kolicina'].sum() if 'Kolicina' in self.startni.columns else 0
            else:
                prev_god, prev_mes = self.meseci_order[first_a_idx - 1]
                prev_sub = self.prodaja[(self.prodaja['id artikla']==ida) & (self.prodaja['Godina']==prev_god) & (self.prodaja['Mesec']==prev_mes)]
                start_lager = prev_sub['Lager'].sum() if len(prev_sub) > 0 else 0
                start_lager = start_lager if not pd.isna(start_lager) else 0
            lageri = [start_lager]
            for god, mes in a_meseci_order:
                msub = self.prodaja[(self.prodaja['id artikla']==ida) & (self.prodaja['Godina']==god) & (self.prodaja['Mesec']==mes)]
                lager_kraj = msub['Lager'].sum() if len(msub) > 0 else 0
                lageri.append(lager_kraj if not pd.isna(lager_kraj) else 0)
            avg_lager = np.mean(lageri)
            obrt = total_prod / avg_lager if avg_lager > 0 else 0
            dani_pokrivanja = (avg_lager / (total_prod / (n_a * 30))) if total_prod > 0 else 999
            n_obj_aktiv = sub[sub['Prodata Kolicina']>0]['ID KOMITENTA'].nunique()
            n_obj_total = sub['ID KOMITENTA'].nunique()
            prod_po_obj = total_prod / n_obj_aktiv if n_obj_aktiv > 0 else 0
            mes_prod = {}
            for _, r in sub.iterrows():
                key = f"{int(r['Godina'])}/{int(r['Mesec'])}"
                mes_prod[key] = mes_prod.get(key, 0) + int(r['Prodata Kolicina'])
            promo_rows.append({
                'id artikla': int(ida),
                'Naziv': sub.iloc[0]['Naziv artikla'],
                'Grupa': sub.iloc[0]['Grupa'],
                'Redovna': pi['redovna'], 'Akcijska': pi['akcijska'],
                'Popust_%': round((1 - pi['akcijska']/pi['redovna'])*100, 1),
                'Prodato_kom': total_prod,
                'Prihod_akcija': round(prihod_akcija, 0),
                'Prihod_redovna': round(prihod_redovna, 0),
                'Profit_akcija': round(profit_akcija, 0),
                'Profit_da_je_redovna': round(profit_redovna, 0),
                'Cena_akcije': round(razlika, 0),
                'Avg_lager': round(avg_lager, 0),
                'Obrt_x': round(obrt, 1),
                'Dani_pokrivanja': round(dani_pokrivanja, 0),
                'Obj_aktivnih': n_obj_aktiv,
                'Obj_ukupno': n_obj_total,
                'Prod_po_obj': round(prod_po_obj, 1),
                **{f'Prod_{a_labels[j]}': mes_prod.get(f"{int(a_meseci_order[j][0])}/{int(a_meseci_order[j][1])}", 0) for j in range(n_a)}
            })
        self.df_promo = pd.DataFrame(promo_rows).sort_values('Obrt_x', ascending=False)

def create_excel(engine, ukljuci_model=True):
    df=engine.df_result; ml=engine.mesec_labels; wb=Workbook()
    hf=PatternFill('solid',fgColor='2F5496'); hfn=Font(bold=True,color='FFFFFF',name='Arial',size=10)
    sfnt=Font(bold=True,name='Arial',size=9); dfn=Font(name='Arial',size=9)
    tb=Border(left=Side('thin','B4C6E7'),right=Side('thin','B4C6E7'),top=Side('thin','B4C6E7'),bottom=Side('thin','B4C6E7'))
    ca=Alignment(horizontal='center',vertical='center'); caw=Alignment(horizontal='center',vertical='center',wrap_text=True)
    sf_poc=PatternFill('solid',fgColor='D6E4F0'); sf_prom=PatternFill('solid',fgColor='C6EFCE')
    sf_prod=PatternFill('solid',fgColor='FFF2CC'); sf_pov=PatternFill('solid',fgColor='FCE4EC')
    sf_kor=PatternFill('solid',fgColor='E8E8E8'); sf_pred=PatternFill('solid',fgColor='D5A6E6')
    sf_avg=PatternFill('solid',fgColor='B4D7E8'); sf_razl=PatternFill('solid',fgColor='FFD699')
    sf_lager=PatternFill('solid',fgColor='DAEEF3'); sf_p1=PatternFill('solid',fgColor='92D050')
    sf_p2=PatternFill('solid',fgColor='00B050'); pred_hdr=PatternFill('solid',fgColor='7030A0')
    ord_hdr=PatternFill('solid',fgColor='375623'); sf_hist=PatternFill('solid',fgColor='E2D5F1')
    nf_money='#,##0'
    SC=5; sub_h=['Pocetno stanje','Promet (ulaz)','Prodaja','Povrat','Korekcija']
    sub_f=[sf_poc,sf_prom,sf_prod,sf_pov,sf_kor]; col_suf=['_Pocetno','_Promet','_Prodaja','_Povrat','_Korekcija']
    ws1=wb.active; ws1.title="Pregled po objektima"
    for c,t in enumerate(['ID Komitenta','ID Artikla','Naziv Artikla','Grupa'],1):
        cell=ws1.cell(1,c,t); cell.font=hfn; cell.fill=hf; cell.alignment=ca; cell.border=tb
        ws1.merge_cells(start_row=1,end_row=2,start_column=c,end_column=c)
    hist_col=5; month_start=5
    if engine.has_history:
        cell=ws1.cell(1,hist_col,'Jan-Avg 2025'); cell.font=hfn; cell.fill=PatternFill('solid',fgColor='6B3FA0')
        cell.alignment=ca; cell.border=tb
        ws1.merge_cells(start_row=1,end_row=1,start_column=hist_col,end_column=hist_col)
        c2=ws1.cell(2,hist_col,'Total prodaja'); c2.font=sfnt; c2.fill=sf_hist; c2.alignment=caw; c2.border=tb
        month_start=6
    for i,label in enumerate(ml):
        sc=month_start+i*SC
        ws1.merge_cells(start_row=1,end_row=1,start_column=sc,end_column=sc+SC-1)
        cell=ws1.cell(1,sc,label); cell.font=hfn; cell.fill=hf; cell.alignment=ca
        for cc in range(sc,sc+SC): ws1.cell(1,cc).border=tb; ws1.cell(1,cc).fill=hf
        for j,(sh,sfill) in enumerate(zip(sub_h,sub_f)):
            cell=ws1.cell(2,sc+j,sh); cell.font=sfnt; cell.fill=sfill; cell.border=tb; cell.alignment=caw
    ps=month_start+len(ml)*SC
    ws1.merge_cells(start_row=1,end_row=1,start_column=ps,end_column=ps+2)
    cell=ws1.cell(1,ps,f'{engine.pred_label} - PREDIKCIJA'); cell.font=hfn; cell.fill=pred_hdr; cell.alignment=ca
    for cc in range(ps,ps+3): ws1.cell(1,cc).border=tb; ws1.cell(1,cc).fill=pred_hdr
    for j,(sh,sfill) in enumerate(zip(['Predikcija','Prosek (svi mes.)','Razlika'],[sf_pred,sf_avg,sf_razl])):
        cell=ws1.cell(2,ps+j,sh); cell.font=sfnt; cell.fill=sfill; cell.border=tb; cell.alignment=caw
    os_c=ps+3
    ws1.merge_cells(start_row=1,end_row=1,start_column=os_c,end_column=os_c+2)
    cell=ws1.cell(1,os_c,f'PORUDZBINA - {engine.order_label}'); cell.font=hfn; cell.fill=ord_hdr; cell.alignment=ca
    for cc in range(os_c,os_c+3): ws1.cell(1,cc).border=tb; ws1.cell(1,cc).fill=ord_hdr
    ll="Lager danas"
    if len(engine.trenutni)>0:
        dc=[c for c in engine.trenutni.columns if 'dan' in c.lower()]
        if dc:
            try: d=pd.to_datetime(engine.trenutni[dc[0]].iloc[0]); ll=f"Lager na dan\n{d.strftime('%d.%m.%Y')}"
            except: pass
    for j,(sh,sfill) in enumerate(zip([ll,'Porudzbina\n(osnovna)',f'Porudzbina\n(min. {engine.min_lager} na stanju)'],[sf_lager,sf_p1,sf_p2])):
        cell=ws1.cell(2,os_c+j,sh); cell.font=sfnt; cell.fill=sfill; cell.border=tb; cell.alignment=caw
    for idx,row in df.iterrows():
        r=idx+3
        for c2,col in enumerate(['ID KOMITENTA','id artikla','Naziv artikla','Grupa'],1):
            ws1.cell(r,c2,row[col]).font=dfn; ws1.cell(r,c2).border=tb
        if engine.has_history:
            v=int(row.get('Total_JanAvg',0)); cell=ws1.cell(r,hist_col,v); cell.font=dfn
            cell.alignment=ca; cell.border=tb
            if v>0: cell.fill=PatternFill('solid',fgColor='F3EAFA')
        for i,label in enumerate(ml):
            cb=month_start+i*SC
            for j,suf in enumerate(col_suf):
                cn=f'{label}{suf}'; v=row.get(cn,0)
                cell=ws1.cell(r,cb+j,int(v) if not pd.isna(v) else 0); cell.font=dfn; cell.alignment=ca; cell.border=tb
        for j,cn in enumerate(['Predikcija','Prosek','Razlika']):
            v=int(row.get(cn,0)); cell=ws1.cell(r,ps+j,v); cell.alignment=ca; cell.border=tb
            if cn=='Razlika':
                if v>0: cell.font=Font(name='Arial',size=9,color='006100',bold=True)
                elif v<0: cell.font=Font(name='Arial',size=9,color='9C0006',bold=True)
                else: cell.font=dfn
            else: cell.font=dfn
        for j,cn in enumerate(['Lager_danas','Porudzbina_1','Porudzbina_2']):
            v=int(row.get(cn,0)); cell=ws1.cell(r,os_c+j,v); cell.alignment=ca; cell.border=tb
            if cn!='Lager_danas' and v>0: cell.font=Font(name='Arial',size=9,bold=True,color='375623')
            else: cell.font=dfn
    ws1.column_dimensions['A'].width=14; ws1.column_dimensions['B'].width=11; ws1.column_dimensions['C'].width=50; ws1.column_dimensions['D'].width=12
    if engine.has_history: ws1.column_dimensions[get_column_letter(hist_col)].width=14
    for i in range(len(ml)):
        for j in range(SC): ws1.column_dimensions[get_column_letter(month_start+i*SC+j)].width=14
    for j in range(3): ws1.column_dimensions[get_column_letter(ps+j)].width=14
    for j in range(3): ws1.column_dimensions[get_column_letter(os_c+j)].width=18
    ws1.freeze_panes=f'{get_column_letter(month_start)}3'
    ws1.auto_filter.ref=f"A2:{get_column_letter(ws1.max_column)}{ws1.max_row}"
    ws2=wb.create_sheet("Totali po mesecima")
    for c,h in enumerate(['Mesec','Promet (ulaz)','Prodaja','Stvarni povrat','Korekcija','Neto (Promet-Povrat)'],1):
        cell=ws2.cell(1,c,h); cell.font=hfn; cell.fill=hf; cell.alignment=caw; cell.border=tb
    ro=2
    if engine.has_history:
        ws2.cell(ro,1,'Jan-Avg 2025 (UKUPNO)').font=Font(bold=True,name='Arial',size=10,color='6B3FA0')
        ws2.cell(ro,1).alignment=ca; ws2.cell(ro,1).border=tb
        cell=ws2.cell(ro,3,int(df['Total_JanAvg'].sum())); cell.font=Font(bold=True,name='Arial',size=10,color='6B3FA0')
        cell.fill=sf_hist; cell.alignment=ca; cell.border=tb; cell.number_format=nf_money
        for c in [2,4,5,6]: ws2.cell(ro,c,'-').font=dfn; ws2.cell(ro,c).alignment=ca; ws2.cell(ro,c).border=tb
        ro+=2
    for ri,label in enumerate(ml,ro):
        ws2.cell(ri,1,label).font=Font(bold=True,name='Arial',size=10); ws2.cell(ri,1).alignment=ca; ws2.cell(ri,1).border=tb
        vals=[int(df[f'{label}_Promet'].sum()),int(df[f'{label}_Prodaja'].sum()),int(df[f'{label}_Povrat'].sum()),int(df[f'{label}_Korekcija'].sum())]
        vals.append(vals[0]-vals[2])
        fills=[sf_prom,sf_prod,sf_pov,sf_kor,sf_poc]
        for c2,(v,f) in enumerate(zip(vals,fills),2):
            cell=ws2.cell(ri,c2,v); cell.font=dfn; cell.fill=f; cell.alignment=ca; cell.border=tb; cell.number_format=nf_money
    fr=ro+len(ml)+1
    ws2.cell(fr,1,f'PORUDZBINA {engine.order_label.upper()}').font=Font(bold=True,name='Arial',size=11,color='375623'); ws2.cell(fr,1).border=tb
    ir=[(f'Predikcija {engine.pred_label}',int(df['Predikcija'].sum()),sf_pred),('Prosek (svi meseci)',int(df['Prosek'].sum()),sf_avg),
        ('Trenutni lager',int(df['Lager_danas'].sum()),sf_lager),
        ('Porudzbina (osnovna)',int(df[~df['ID KOMITENTA'].isin(engine.excluded)]['Porudzbina_1'].sum()),sf_p1),
        (f'Porudzbina (min. {engine.min_lager})',int(df[~df['ID KOMITENTA'].isin(engine.excluded)]['Porudzbina_2'].sum()),sf_p2)]
    for i,(label,val,fill) in enumerate(ir,fr+1):
        ws2.cell(i,1,label).font=Font(bold=True,name='Arial',size=10); ws2.cell(i,1).alignment=ca; ws2.cell(i,1).border=tb
        cell=ws2.cell(i,2,val); cell.font=Font(bold=True,name='Arial',size=11); cell.fill=fill; cell.alignment=ca; cell.border=tb; cell.number_format=nf_money
    ws2.column_dimensions['A'].width=32; ws2.column_dimensions['B'].width=18
    for c in 'CDEF': ws2.column_dimensions[c].width=18
    if engine.has_prices and len(engine.df_oos) > 0:
        ws_oos = wb.create_sheet("OOS Izgubljeni profit")
        oos_hdr = PatternFill('solid', fgColor='C00000')
        oos_fill = PatternFill('solid', fgColor='FCE4EC')
        a_labels_oos = engine.analitika_labels if engine.analitika_labels else engine.mesec_labels
        fixed_h = ['ID Komitenta','ID Artikla','Naziv','Grupa','Prosek kad ima','Lager danas']
        mes_h = []
        for lb in a_labels_oos: mes_h += [f'OOS {lb} (kom)', f'Izgub {lb} (RSD)']
        all_h = fixed_h + mes_h + ['OOS meseci ukupno','Izgubljeni profit (RSD)']
        for c, h in enumerate(all_h, 1):
            cell = ws_oos.cell(1, c, h)
            cell.font=Font(bold=True,color='FFFFFF',name='Arial',size=9)
            cell.fill=oos_hdr; cell.alignment=caw; cell.border=tb
        for idx, (_, row) in enumerate(engine.df_oos.iterrows(), 2):
            vals = [row['ID KOMITENTA'], row['id artikla'], row['Naziv artikla'], row['Grupa'],
                    row.get('Prosek_kad_ima',0), row.get('Lager_danas',0)]
            for lb in a_labels_oos:
                vals.append(row.get(f'OOS_{lb}', 0))
                vals.append(row.get(f'Izgub_{lb}', 0))
            vals += [row.get('OOS_meseci',0), row.get('Izgubljeni_profit',0)]
            for c, v in enumerate(vals, 1):
                cell = ws_oos.cell(idx, c, v); cell.font=dfn; cell.border=tb; cell.alignment=ca
                col_name = all_h[c-1]
                if col_name.startswith('OOS ') and isinstance(v, (int, float)) and v > 0:
                    cell.fill = oos_fill; cell.font = Font(name='Arial',size=9,bold=True,color='C00000')
                if col_name.startswith('Izgub ') or col_name == 'Izgubljeni profit (RSD)':
                    cell.number_format = nf_money
                if col_name == 'Lager danas' and v == 0:
                    cell.fill = oos_fill; cell.font = Font(name='Arial',size=9,bold=True,color='C00000')
        ws_oos.column_dimensions['A'].width=13; ws_oos.column_dimensions['B'].width=10
        ws_oos.column_dimensions['C'].width=45; ws_oos.column_dimensions['D'].width=12
        ws_oos.column_dimensions['E'].width=14; ws_oos.column_dimensions['F'].width=12
        for i in range(len(a_labels_oos)*2):
            ws_oos.column_dimensions[get_column_letter(7+i)].width=13
        last_col = 7 + len(a_labels_oos)*2
        ws_oos.column_dimensions[get_column_letter(last_col)].width=14
        ws_oos.column_dimensions[get_column_letter(last_col+1)].width=18
        ws_oos.freeze_panes='E2'
        ws_oos.auto_filter.ref=f"A1:{get_column_letter(len(all_h))}{len(engine.df_oos)+1}"
    if engine.has_prices and len(engine.df_profit_obj) > 0:
        ws_prof = wb.create_sheet("Profitabilnost objekata")
        prof_hdr = PatternFill('solid', fgColor='1F4E79')
        bad_fill = PatternFill('solid', fgColor='FCE4EC')
        good_fill = PatternFill('solid', fgColor='E2EFDA')
        headers = ['ID Komitenta','Artikala','Prodato kom','Bruto profit (RSD)','Trosak mkt (RSD)','Neto profit (RSD)','Izgubljeno OOS (RSD)','Potencijal (RSD)']
        for lb in (engine.analitika_labels if engine.analitika_labels else ml): headers.append(f'Neto {lb}')
        for c, h in enumerate(headers, 1):
            cell = ws_prof.cell(1, c, h); cell.font=Font(bold=True,color='FFFFFF',name='Arial',size=9); cell.fill=prof_hdr; cell.alignment=caw; cell.border=tb
        for idx, (_, row) in enumerate(engine.df_profit_obj.iterrows(), 2):
            vals = [row['ID KOMITENTA'], row['Artikala'], row['Prodato_kom'], row['Bruto_profit'],
                    row['Trosak_mkt'], row['Neto_profit'], row['Izgubljeno_OOS'], row['Potencijalni_profit']]
            for lb in (engine.analitika_labels if engine.analitika_labels else ml): vals.append(row.get(f'Neto_{lb}', 0))
            for c, v in enumerate(vals, 1):
                cell = ws_prof.cell(idx, c, v); cell.font=dfn; cell.border=tb; cell.alignment=ca
                if c >= 4: cell.number_format=nf_money
                if c == 6:
                    if v <= 0: cell.fill = bad_fill; cell.font = Font(name='Arial', size=9, bold=True, color='C00000')
                    elif v > 0: cell.fill = good_fill
                if c >= 9:
                    if v < 0: cell.font = Font(name='Arial', size=9, color='C00000')
                    elif v > 0: cell.font = Font(name='Arial', size=9, color='006100')
        for cl in 'AB': ws_prof.column_dimensions[cl].width=13
        ws_prof.column_dimensions['C'].width=12
        for cl in 'DEFGH': ws_prof.column_dimensions[cl].width=18
        a_ml = engine.analitika_labels if engine.analitika_labels else ml
        for i in range(len(a_ml)): ws_prof.column_dimensions[get_column_letter(9+i)].width=14
        ws_prof.freeze_panes='B2'
        ws_prof.auto_filter.ref=f"A1:{get_column_letter(len(headers))}{len(engine.df_profit_obj)+1}"
    if engine.has_prices and len(engine.df_promo) > 0:
        ws_akc = wb.create_sheet("Analiza akcije")
        akc_hdr = PatternFill('solid', fgColor='BF8F00')
        good_obrt = PatternFill('solid', fgColor='E2EFDA')
        bad_obrt = PatternFill('solid', fgColor='FCE4EC')
        headers = ['ID Artikla','Naziv','Grupa','Redovna\ncena','Akcijska\ncena','Popust\n%',
                   'Prodato\nkom','Prihod\nakcija (RSD)','Prihod da je\nredovna (RSD)',
                   'Profit\nakcija (RSD)','Profit da je\nredovna (RSD)','Cena akcije\n(RSD)',
                   'Prosecni\nlager','Obrt\n(x)','Dani\npokrivanja',
                   'Aktivnih\nobjekata','Ukupno\nobjekata','Prod.\npo objektu']
        for lb in (engine.analitika_labels if engine.analitika_labels else ml): headers.append(f'Prod.\n{lb}')
        for c, h in enumerate(headers, 1):
            cell = ws_akc.cell(1, c, h); cell.font=Font(bold=True,color='FFFFFF',name='Arial',size=9); cell.fill=akc_hdr; cell.alignment=caw; cell.border=tb
        for idx, (_, row) in enumerate(engine.df_promo.iterrows(), 2):
            vals = [row['id artikla'], row['Naziv'], row['Grupa'], row['Redovna'], row['Akcijska'],
                    row['Popust_%'], row['Prodato_kom'],
                    row['Prihod_akcija'], row['Prihod_redovna'],
                    row['Profit_akcija'], row['Profit_da_je_redovna'], row['Cena_akcije'],
                    row['Avg_lager'], row['Obrt_x'], row['Dani_pokrivanja'],
                    row['Obj_aktivnih'], row['Obj_ukupno'], row['Prod_po_obj']]
            for lb in (engine.analitika_labels if engine.analitika_labels else ml): vals.append(row.get(f'Prod_{lb}', 0))
            for c, v in enumerate(vals, 1):
                cell = ws_akc.cell(idx, c, v); cell.font=dfn; cell.border=tb; cell.alignment=ca
                if c in [4,5,8,9,10,11,12]: cell.number_format=nf_money
                if c == 14:
                    if v >= 2.0: cell.fill = good_obrt; cell.font = Font(name='Arial',size=9,bold=True,color='006100')
                    elif v < 1.0: cell.fill = bad_obrt; cell.font = Font(name='Arial',size=9,bold=True,color='C00000')
                if c == 15 and v > 120: cell.fill = bad_obrt
        ws_akc.column_dimensions['A'].width=10; ws_akc.column_dimensions['B'].width=45; ws_akc.column_dimensions['C'].width=12
        for cl in 'DEFG': ws_akc.column_dimensions[cl].width=12
        for cl in 'HIJKL': ws_akc.column_dimensions[cl].width=16
        for cl in 'MNOPQR': ws_akc.column_dimensions[cl].width=13
        a_ml2 = engine.analitika_labels if engine.analitika_labels else ml
        for i in range(len(a_ml2)): ws_akc.column_dimensions[get_column_letter(19+i)].width=11
        ws_akc.auto_filter.ref=f"A1:{get_column_letter(len(headers))}{len(engine.df_promo)+1}"
    ws3=wb.create_sheet("O modelu"); ws3.column_dimensions['A'].width=100
    info=["OPIS MODELA PREDIKCIJE I PORUDZBINE","",f"=== PREDIKCIJA ZA {engine.pred_label.upper()} ===","",
        "Model predvidja POTENCIJAL PRODAJE.","",
        f"  1. Constrained sales korekcija:",
        f"     - Kraj meseca lager=0 i prodaja>0: rasprodato, potraznja veca — zameni prosekom normalnih meseci",
        f"     - Pocetno=0 i promet=0: cist OOS — zameni prosekom normalnih meseci",
        f"     - Pocetno=0 i promet>0 i kraj=0: dobili i rasprodali — zameni prosekom",
        f"     - Normalni meseci = ostalo robe na kraju (lager>0)",
        f"  2. Holt DES (alpha={engine.alpha}, beta={engine.beta}) + WMA (50/28/12/7/3%)",
        "  3. Kombinacija: 60% veci + 40% manji od Holt/WMA",
        "  4. Varijansa boost (faktor 0.4, max 70%)",
        "  5. Niska zaliha (0-2): predikcija minimum prosek kad je na stanju",
        "  6. Prodaja 5+ mesecno: predikcija minimum prosek",
        "  7. Donje ogranicenje: predikcija < prosek samo ako poslednjih 5 meseci pada ili stagnira (<=)",
        "  8. Sigurnosna mreza: predikcija=0 samo ako nista prodato u poslednjih 5 meseci; ako poslednji mesec >1 onda min taj broj",
        "  9. Zaokruzivanje: round (predikcija i prosek)",
        ]
    if engine.has_history: info+=[f"  10. Istorijski podaci: {HIST_WEIGHT*100:.0f}% tezina"]
    info+=["",f"=== PORUDZBINA ZA {engine.order_label.upper()} ===","",
        f"P1 (osnovna): max(Pred-Lager, 0)",
        f"P2 (sa dopunom): Za lager<=2: dopuna do max(predikcija, prosek, min porudzbina={engine.min_order}); Za lager>2: dopuna do min {engine.min_lager}",
        f"P2 finalna provera: ako (P2+lager) <= prodaja_poslednjeg_meseca, dodaje se buffer (1-5 kom: +2, 6-10: +3, 11-15: +4, 16+: +5)",
        f"Min kom po artiklu (po stavci): {engine.min_per_artikal if engine.min_per_artikal else 'nije zadat'} — ako je porudzbina > 0 ali manja od minimuma, podize se na minimum. Nule ostaju 0.",
        f"Iskljuceni: {', '.join(str(x) for x in sorted(engine.excluded))}"]
    if engine.has_prices:
        info+=["",f"=== ANALITIKA ===","",
            f"Profit formula: (Finalna cena / 1.2 / 1.2 - Nabavna) x Kolicina",
            f"OOS izgubljeni profit (NOVA LOGIKA - 3 uslova):",
            f"  USLOV 1 - Cist OOS (poc=0 i ulaz=0): izgubljeno_kom = avg_stocked",
            f"  USLOV 2 - Rasprodato (kraj=0 i prodaja>0): izgubljeno_kom = max(0, avg_stocked - prodaja)",
            f"  USLOV 3 - Dobili i rasprodali (poc=0, ulaz>0, kraj=0): izgubljeno_kom = max(0, avg_stocked - prodaja)",
            f"  Ako mesec nije constrained (kraj > 0): OOS = 0 (imali su robe i nisu rasprodali)",
            f"  avg_stocked = prosek prodaje u mesecima koji NISU constrained",
            f"  Izgubljeni profit = izgubljeno_kom x profit/kom po mesecu",
            f"Ukupan trosak marketinga: {engine.mesecni_trosak:,.0f} RSD / {engine.num_komitenti} objekata = {engine.trosak_po_objektu:,.0f} RSD po objektu za period",
            f"Mesecni trosak po objektu: {engine.trosak_po_objektu / max(len(engine.analitika_labels), 1):,.0f} RSD",
            f"Neto po mesecu = Bruto profit meseca - mesecni trosak po objektu"]
    info+=[f"","Generisano: {_now().strftime('%d.%m.%Y. u %H:%M')}"]
    for i,line in enumerate(info,1):
        cell=ws3.cell(i,1,line)
        if i==1: cell.font=Font(bold=True,name='Arial',size=14,color='375623')
        elif '===' in line: cell.font=Font(bold=True,name='Arial',size=12,color='7030A0')
        else: cell.font=Font(name='Arial',size=10)
    if not ukljuci_model:
        try:
            if "O modelu" in wb.sheetnames:
                del wb["O modelu"]
        except Exception:
            pass
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

DEFAULT_EXCLUDED = "1023, 1027, 1034, 1043, 1057, 1060, 1061, 1076, 1315, 1347, 1349, 1359"
st.set_page_config(page_title="VAPE Analitika", page_icon="\U0001f4a8", layout="wide", initial_sidebar_state="collapsed")
st.markdown("""<style>
section[data-testid="stSidebar"] { display: none !important; }
header[data-testid="stHeader"] { display: none !important; }
#MainMenu { visibility: hidden !important; }
footer { visibility: hidden !important; }
.main .block-container,
div[data-testid="block-container"],
div[data-testid="stMainBlockContainer"] {
    padding: 12px 16px 0 16px !important;
    max-width: 100% !important;
}
</style>""", unsafe_allow_html=True)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
    .stApp {
        background: #f5f0ff !important;
        font-family: 'Poppins', sans-serif;
    }
    .main .stTextInput > div > div > input,
    .main .stNumberInput > div > div > input {
        background: white !important;
        border: 1px solid rgba(168,85,247,0.25) !important;
        color: #1a0533 !important;
        border-radius: 8px !important;
    }
    .main .stTextInput > div > div > input::placeholder,
    .main .stNumberInput > div > div > input::placeholder {
        color: #9ca3af !important;
    }
    .main .stTextInput > div > div > input:focus,
    .main .stNumberInput > div > div > input:focus {
        border-color: #a855f7 !important;
        box-shadow: 0 0 0 2px rgba(168,85,247,0.15) !important;
    }
    .metric-card {
        background: white;
        border-radius: 14px;
        padding: 16px 20px;
        box-shadow: 0 2px 12px rgba(124,58,237,0.07);
        border: 1px solid rgba(168,85,247,0.12);
        text-align: center;
    }
    .metric-value {
        font-size: 26px; font-weight: 700;
        background: linear-gradient(135deg, #7c3aed, #ec4899);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    .metric-value-red { font-size: 26px; font-weight: 700; color: #dc2626; }
    .metric-value-green { font-size: 26px; font-weight: 700; color: #059669; }
    .metric-label { font-size: 11px; color: #888; margin-top: 4px; }
    .stButton > button {
        background: linear-gradient(135deg, #a855f7 0%, #ec4899 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 14px 32px !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        box-shadow: 0 4px 15px rgba(168,85,247,0.3) !important;
        transition: opacity 0.2s !important;
    }
    .stButton > button:hover { opacity: 0.88 !important; }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 14px 32px !important;
        font-weight: 700 !important;
        box-shadow: 0 4px 15px rgba(16,185,129,0.25) !important;
    }
    .stMultiSelect [data-baseweb="tag"] {
        background: linear-gradient(135deg, #a855f7, #ec4899) !important;
        border: none !important;
        border-radius: 99px !important;
        color: white !important;
        font-weight: 600 !important;
        font-size: 12px !important;
    }
    .stMultiSelect [data-baseweb="tag"] span { color: white !important; }
    .stMultiSelect [data-baseweb="tag"] button { color: rgba(255,255,255,0.8) !important; }
    .stMultiSelect [data-baseweb="select"] > div {
        border: 1px solid rgba(168,85,247,0.3) !important;
        border-radius: 10px !important;
        background: white !important;
    }
    .stMultiSelect [data-baseweb="select"] > div:focus-within {
        border-color: #a855f7 !important;
        box-shadow: 0 0 0 2px rgba(168,85,247,0.15) !important;
    }
    .success-box {
        background: linear-gradient(135deg, rgba(16,185,129,0.08), rgba(5,150,105,0.04));
        border: 1px solid rgba(16,185,129,0.2);
        border-radius: 10px;
        padding: 12px 16px;
    }
    .warn-box {
        background: linear-gradient(135deg, rgba(220,38,38,0.07), rgba(220,38,38,0.02));
        border: 1px solid rgba(220,38,38,0.18);
        border-radius: 10px;
        padding: 10px 14px;
        margin: 6px 0;
    }
    .section-title {
        font-size: 17px; font-weight: 600; color: #4c1d95; margin: 16px 0 8px 0;
    }
</style>
""", unsafe_allow_html=True)
alpha = 0.4
beta = 0.2
min_lager = None
min_order = None
min_per_artikal = None
max_per_artikal = None
syx_objekti = None
mesecni_trosak = 0
excluded_str = DEFAULT_EXCLUDED
excluded = set()
for part in excluded_str.replace('\n', ',').split(','):
    p = part.strip()
    if p.isdigit(): excluded.add(int(p))
def render_header(subtitle):
    st.markdown(f'''<div style="background:#12002a;border-radius:16px;padding:0 28px;height:60px;
        display:flex;align-items:center;justify-content:space-between;margin-bottom:24px;
        border-bottom:3px solid;border-image:linear-gradient(90deg,#a855f7,#ec4899) 1;
        box-shadow:0 4px 20px rgba(18,0,42,0.18);">
        <div style="display:flex;align-items:center;gap:12px;">
            <div style="width:30px;height:30px;background:linear-gradient(135deg,#a855f7,#ec4899);
                border-radius:8px;display:flex;align-items:center;justify-content:center;">
                <div style="width:11px;height:11px;background:white;border-radius:3px;"></div>
            </div>
            <span style="font-size:18px;font-weight:700;color:white;">VAPE</span>
            <span style="font-size:18px;font-weight:300;color:rgba(255,255,255,0.4);">Analitika</span>
            <span style="font-size:11px;color:rgba(255,255,255,0.25);margin-left:8px;">·</span>
            <span style="font-size:12px;color:rgba(255,255,255,0.35);">{subtitle}</span>
        </div>
        <div style="display:flex;gap:12px;align-items:center;">
            <div style="display:flex;gap:6px;align-items:center;">
                <div style="width:8px;height:8px;border-radius:50%;background:rgba(168,85,247,0.7);"></div>
                <div style="width:8px;height:8px;border-radius:50%;background:rgba(236,72,153,0.5);"></div>
                <div style="width:8px;height:8px;border-radius:50%;background:rgba(255,255,255,0.15);"></div>
            </div>
        </div>
    </div>''', unsafe_allow_html=True)
render_header("Predikcija prodaje · Profitabilnost · OOS analiza · Efekti akcije")
_co = st.columns([6, 1])
with _co[1]:
    if st.button("🔓 Odjava", key="ana_odjava"):
        for _k in ("authenticated", "role", "admin_user", "mail_nalog"):
            st.session_state.pop(_k, None)
        st.rerun()

st.markdown("""<style>
/* --- Tabovi kao pilule --- */
div[data-baseweb="tab-list"] { gap: 8px; border-bottom: none !important; }
button[data-baseweb="tab"] {
    background:#fff; border:1px solid rgba(168,85,247,0.22); border-radius:12px;
    padding:8px 20px; color:#6b21a8; font-weight:600; }
button[data-baseweb="tab"][aria-selected="true"] {
    background:linear-gradient(135deg,#a855f7,#ec4899); color:#fff; border-color:transparent;
    box-shadow:0 4px 14px rgba(168,85,247,0.3); }
div[data-baseweb="tab-highlight"], div[data-baseweb="tab-border"] { display:none !important; }
/* --- Kartice (bordered container) --- */
div[data-testid="stVerticalBlockBorderWrapper"] {
    background:#fff !important; border:1px solid rgba(168,85,247,0.14) !important;
    border-radius:14px !important; box-shadow:0 2px 12px rgba(124,58,237,0.07) !important;
    padding:14px 18px !important; margin-bottom:10px !important; }
/* --- File uploader dropzone --- */
[data-testid="stFileUploaderDropzone"] {
    background:#faf5ff !important; border:2px dashed rgba(168,85,247,0.35) !important;
    border-radius:12px !important; }
/* --- Objavi dugme zeleno --- */
.st-key-obj_btn button {
    background:linear-gradient(135deg,#10b981,#059669) !important;
    box-shadow:0 4px 15px rgba(16,185,129,0.25) !important; }
/* --- Naslovi kartica --- */
.obj-title { font-size:15px; font-weight:600; color:#4c1d95; margin:0 0 10px 0; display:flex; align-items:center; gap:8px; }
.obj-badge { background:linear-gradient(135deg,#a855f7,#ec4899); color:#fff; width:22px; height:22px;
    border-radius:50%; display:inline-flex; align-items:center; justify-content:center; font-size:12px; font-weight:700; }
.obj-badge.green { background:linear-gradient(135deg,#10b981,#059669); }
.chip-ok { display:inline-block; background:#eafaf0; color:#059669; border:1px solid rgba(16,185,129,.25);
    border-radius:99px; padding:5px 13px; font-size:12.5px; font-weight:600; margin:4px 5px 0 0; }
.chip-no { display:inline-block; background:#f3f4f6; color:#9ca3af; border:1px dashed #d1d5db;
    border-radius:99px; padding:5px 13px; font-size:12.5px; font-weight:600; margin:4px 5px 0 0; }
</style>""", unsafe_allow_html=True)

tab_obj, tab_ana = st.tabs(["📤 Objava izveštaja", "📊 Analitika"])

with tab_obj:
    with st.container(border=True):
        st.markdown('<div class="obj-title">📋 Objavljeno za koleginice</div>', unsafe_allow_html=True)
        if not sb_dostupan():
            st.info("Supabase nije podešen — dodaj SUPABASE_URL i SUPABASE_KEY u Streamlit Secrets da bi objava radila.")
        else:
            _pmeseci = sb_meseci()
            if not _pmeseci:
                st.caption("Još nijedan sistem nije objavljen.")
            else:
                _plabels = [m["label"] for m in _pmeseci]
                _pkeys = [m["key"] for m in _pmeseci]
                _psel = st.selectbox("Mesec", _plabels, index=0, key="obj_preg_mes")
                _pmk = _pkeys[_plabels.index(_psel)]
                _imaju = sb_sisteme(_pmk)
                _svi = sb_svi_sistemi()
                _chip = ""
                for _s in _imaju:
                    _chip += '<span class="chip-ok">\u2713 ' + str(_s) + '</span>'
                for _s in _svi:
                    if _s not in _imaju:
                        _chip += '<span class="chip-no">\u25cb ' + str(_s) + ' \u00b7 nije objavljen</span>'
                if not _chip:
                    _chip = '<span style="color:#9ca3af;font-size:13px;">Nema objavljenih sistema za ovaj mesec.</span>'
                st.markdown(_chip, unsafe_allow_html=True)

                # --- Brisanje objavljenih izve\u0161taja za ovaj mesec ---
                if _imaju:
                    with st.expander("Obri\u0161i objavljeni izve\u0161taj (" + _psel + ")"):
                        st.caption("Izaberi sistem(e) za ovaj mesec koje \u017Eeli\u0161 da obri\u0161e\u0161, pa potvrdi. "
                                   "Brisanje je trajno \u2014 posle mo\u017Ee\u0161 da objavi\u0161 nove.")
                        _del_sel = st.multiselect("Sistemi za brisanje", _imaju, key="del_sis_" + _pmk)
                        _del_ok = st.checkbox("Potvr\u0111ujem brisanje izabranih", key="del_ok_" + _pmk)
                        if st.button("Obri\u0161i izabrano", key="del_btn_" + _pmk,
                                     disabled=not (_del_sel and _del_ok), use_container_width=True):
                            _nbr = 0
                            for _s in _del_sel:
                                try:
                                    sb_obrisi(_pmk, _s); _nbr += 1
                                except Exception as _e:
                                    st.error("Gre\u0161ka pri brisanju \u201E" + str(_s) + "\u201C: " + str(_e))
                            if _nbr:
                                st.success("Obrisano: " + str(_nbr) + " izve\u0161taj(a). Sad mo\u017Ee\u0161 da objavi\u0161 nove.")
                                st.rerun()

    with st.container(border=True):
        st.markdown('<div class="obj-title">\U0001F5D3\uFE0F Plan objave (za koleginice)</div>', unsafe_allow_html=True)
        if not sb_dostupan():
            st.caption("Nedostupno dok Supabase nije podešen.")
        else:
            _today = datetime.date.today()
            _mopts = []
            _yy = _today.year; _mm = _today.month - 3  # uključi i prethodne mesece
            while _mm <= 0:
                _mm += 12; _yy -= 1
            for _i in range(9):
                _mopts.append(str(_yy) + "-" + ("0" + str(_mm))[-2:])
                _mm += 1
                if _mm > 12:
                    _mm = 1; _yy += 1
            # podrazumevano: prethodni mesec (u avgustu se radi izveštaj za jul)
            _pv_y = _today.year; _pv_m = _today.month - 1
            if _pv_m <= 0:
                _pv_m += 12; _pv_y -= 1
            _prev_key = str(_pv_y) + "-" + ("0" + str(_pv_m))[-2:]
            _def_idx = _mopts.index(_prev_key) if _prev_key in _mopts else 0
            _pc1, _pc2, _pc3 = st.columns([1.3, 1, 1])
            with _pc1:
                _pmes = st.selectbox("Mesec", _mopts, index=_def_idx, format_func=mesec_label, key="plan_mes")
            with _pc2:
                _pdat = st.date_input("Objaviću do", value=_today, key="plan_dat", format="DD.MM.YYYY")
            with _pc3:
                st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
                if st.button("\U0001F4BE Sačuvaj plan", key="plan_btn", use_container_width=True):
                    try:
                        sb_save_plan(_pmes, _pdat.strftime("%d.%m.%Y"))
                        st.success("Plan sačuvan: " + mesec_label(_pmes) + " \u2192 do " + _pdat.strftime("%d.%m.%Y"))
                    except Exception as _e:
                        st.error("Greška: " + str(_e))
            _curplan = sb_load_plan(_pmes)
            if _curplan:
                st.caption("Trenutni plan za " + mesec_label(_pmes) + ": do " + str(_curplan))

    with st.container(border=True):
        st.markdown('<div class="obj-title">📄 Izveštaj SYX (nikotinske vrećice)</div>', unsafe_allow_html=True)
        st.caption("Ubaci Word (.docx) izveštaj za SYX po mesecu. Kod direktora se pojavljuje u kartici Izveštaj SYX, izlistan po mesecima.")
        if not sb_dostupan():
            st.caption("Nedostupno dok Supabase nije podešen.")
        else:
            _sy_today = datetime.date.today()
            _sy_opts = set(m["key"] for m in sb_meseci())
            for _r in sb_syx_list():
                _sy_opts.add(_r.get("mesec"))
            _yy, _mm = _sy_today.year, _sy_today.month - 2
            while _mm <= 0:
                _mm += 12; _yy -= 1
            for _ in range(9):
                _sy_opts.add(str(_yy) + "-" + ("0" + str(_mm))[-2:])
                _mm += 1
                if _mm > 12:
                    _mm = 1; _yy += 1
            _sy_opts = sorted([o for o in _sy_opts if o], reverse=True)
            _sy_pvy = _sy_today.year; _sy_pvm = _sy_today.month - 1
            if _sy_pvm <= 0:
                _sy_pvm += 12; _sy_pvy -= 1
            _sy_prev = str(_sy_pvy) + "-" + ("0" + str(_sy_pvm))[-2:]
            _sy_idx = _sy_opts.index(_sy_prev) if _sy_prev in _sy_opts else 0
            _syc1, _syc2 = st.columns([1, 2])
            with _syc1:
                _sy_mes = st.selectbox("Mesec", _sy_opts, index=_sy_idx, format_func=mesec_label, key="syx_mes")
            with _syc2:
                _sy_file = st.file_uploader("Word dokument (.docx)", type=["docx"], key="syx_up")
            _rok_sy = sb_rokovi_get(_sy_mes).get("rok_syx")
            if _rok_sy:
                st.caption(("⏰ Rok za SYX (" + mesec_label(_sy_mes) + "): " + _rok_fmt(_rok_sy))
                           + ("  ·  rok istekao" if _rok_je_prosao(_rok_sy) else ""))
            if st.button("📤 Sačuvaj SYX izveštaj", key="syx_save", use_container_width=True,
                         disabled=(_sy_file is None)):
                try:
                    import base64 as _b64s
                    _bytes = _sy_file.getvalue()
                    _b64 = _b64s.b64encode(_bytes).decode("ascii")
                    sb_syx_set(_sy_mes, _sy_file.name, _b64)
                    st.success("SYX izveštaj za " + mesec_label(_sy_mes) + " sačuvan (" + _sy_file.name + ").")
                    st.rerun()
                except Exception as _e:
                    st.error("Greška pri čuvanju: " + str(_e))
            _sy_all = sb_syx_list()
            if _sy_all:
                st.markdown("<div style='margin-top:6px;font-size:12px;text-transform:uppercase;letter-spacing:.5px;"
                            "color:#9aa0ad;font-weight:700;'>Postavljeni SYX izveštaji</div>", unsafe_allow_html=True)
                for _r in _sy_all:
                    _rc1, _rc2, _rc3 = st.columns([2, 3, 1])
                    with _rc1:
                        st.markdown("**" + mesec_label(_r.get("mesec", "")) + "**")
                    with _rc2:
                        st.caption(str(_r.get("filename", "")))
                    with _rc3:
                        if st.button("Obriši", key="syx_del_" + str(_r.get("mesec")), use_container_width=True):
                            try:
                                sb_syx_obrisi(_r.get("mesec"))
                                st.rerun()
                            except Exception as _e:
                                st.error("Greška: " + str(_e))

    with st.container(border=True):
        st.markdown('<div class="obj-title">💳 Izveštaj potraživanja</div>', unsafe_allow_html=True)
        st.caption("Ubaci Excel izveštaja potraživanja po mesecu. Administracija ga dopunjava direktno u aplikaciji "
                   "(iznosi, statusi, komentari) i prosleđuje; direktor ga vidi i izvozi u identičan Excel.")
        if not sb_dostupan():
            st.caption("Nedostupno dok Supabase nije podešen.")
        else:
            _pz_today = datetime.date.today()
            _pz_opts = set(m["key"] for m in sb_meseci())
            for _r in sb_potraz_list():
                _pz_opts.add(_r.get("mesec"))
            _yy, _mm = _pz_today.year, _pz_today.month - 3
            while _mm <= 0:
                _mm += 12; _yy -= 1
            for _ in range(9):
                _pz_opts.add(str(_yy) + "-" + ("0" + str(_mm))[-2:])
                _mm += 1
                if _mm > 12:
                    _mm = 1; _yy += 1
            _pz_opts = sorted([o for o in _pz_opts if o], reverse=True)
            _pz_pvy = _pz_today.year; _pz_pvm = _pz_today.month - 1
            if _pz_pvm <= 0:
                _pz_pvm += 12; _pz_pvy -= 1
            _pz_prev = str(_pz_pvy) + "-" + ("0" + str(_pz_pvm))[-2:]
            _pz_idx = _pz_opts.index(_pz_prev) if _pz_prev in _pz_opts else 0
            _pzc1, _pzc2 = st.columns([1, 2])
            with _pzc1:
                _pz_mes = st.selectbox("Mesec", _pz_opts, index=_pz_idx, format_func=mesec_label, key="pz_mes")
            with _pzc2:
                _pz_file = st.file_uploader("Excel potraživanja (.xlsx)", type=["xlsx"], key="pz_up")
            _rok_pzc = sb_rokovi_get(_pz_mes).get("rok_potraz")
            if _rok_pzc:
                st.caption(("⏰ Rok za potraživanja (" + mesec_label(_pz_mes) + "): " + _rok_fmt(_rok_pzc))
                           + ("  ·  rok istekao" if _rok_je_prosao(_rok_pzc) else ""))
            if st.button("📤 Objavi za administraciju", key="pz_save", use_container_width=True,
                         disabled=(_pz_file is None)):
                try:
                    import base64 as _b64p
                    _pb = _pz_file.getvalue()
                    _struct = potraz_parse(_pb)
                    _pop = potraz_init_popuna(_struct)
                    sb_potraz_set(_pz_mes, _pz_file.name, _b64p.b64encode(_pb).decode("ascii"),
                                  json.dumps(_struct), json.dumps(_pop))
                    _nred = sum(len(s["redovi"]) for L in _struct["listovi"] for s in L["sekcije"])
                    st.success("Objavljeno za administraciju: " + mesec_label(_pz_mes) + " · "
                               + str(len(_struct["listovi"])) + " listova, " + str(_nred) + " stavki.")
                    st.rerun()
                except Exception as _e:
                    st.error("Greška pri obradi Excela: " + str(_e))
            _pz_all = sb_potraz_list()
            if _pz_all:
                st.markdown("<div style='margin-top:6px;font-size:12px;text-transform:uppercase;letter-spacing:.5px;"
                            "color:#9aa0ad;font-weight:700;'>Objavljeni izveštaji potraživanja</div>", unsafe_allow_html=True)
                for _r in _pz_all:
                    _rc1, _rc2, _rc3 = st.columns([2, 3, 1])
                    with _rc1:
                        st.markdown("**" + mesec_label(_r.get("mesec", "")) + "**"
                                    + ("  ✅ predato" if _r.get("predato") else "  ⏳ u obradi"))
                    with _rc2:
                        st.caption("📄 " + str(_r.get("naziv", "")) + "  ·  poslednje ažuriranje: "
                                   + (_dt_fmt(_r.get("azurirano")) or "—"))
                    with _rc3:
                        if st.button("Obriši", key="pz_del_" + str(_r.get("mesec")), use_container_width=True):
                            try:
                                sb_potraz_obrisi(_r.get("mesec"))
                                st.rerun()
                            except Exception as _e:
                                st.error("Greška: " + str(_e))

    with st.container(border=True):
        st.markdown('<div class="obj-title">👥 Šifarnik komitenata (nazivi + kontakt)</div>', unsafe_allow_html=True)
        st.caption("Učitaj Excel sa kolonama: ID, Naziv, e-mail, Kontakt, Mesto, Adresa. Koristi se za nazive objekata i istoriju porudžbina. Ubaci pre objave da se pokupe i nove radnje.")
        if not sb_dostupan():
            st.caption("Nedostupno dok Supabase nije podešen.")
        else:
            _dc1, _dc2 = st.columns([2, 1])
            with _dc2:
                if st.button("🔍 Proveri šifarnik u bazi", key="kom_check", use_container_width=True):
                    _all = _sb_select_all("komitenti", "idk,naziv,email")
                    _sa_naz = sum(1 for r in _all if (r.get("naziv") or "").strip())
                    _sa_mail = sum(1 for r in _all if (r.get("email") or "").strip())
                    st.session_state["_kom_check"] = {"uk": len(_all), "naz": _sa_naz, "mail": _sa_mail}
            _chk = st.session_state.get("_kom_check")
            if _chk:
                with _dc1:
                    st.caption("U bazi: " + str(_chk["uk"]) + " komitenata · sa nazivom " + str(_chk["naz"])
                               + " · sa mejlom " + str(_chk["mail"]) + ".")
            up_k = st.file_uploader("Excel šifarnika komitenata (.xlsx)", type=['xlsx', 'xls'],
                                    key="kom_upl", label_visibility="collapsed")
            if up_k is not None:
                try:
                    _dfk = pd.read_excel(up_k, dtype=str)
                    _cmap = {}
                    for _c in _dfk.columns:
                        _cl = str(_c).strip().lower()
                        if _cl == "id" or _cl == "idk":
                            _cmap["idk"] = _c
                        elif _cl.startswith("naziv"):
                            _cmap["naziv"] = _c
                        elif ("mail" in _cl) or ("mejl" in _cl):
                            _cmap["email"] = _c
                        elif _cl.startswith("kontakt") or _cl.startswith("telefon") or _cl.startswith("tel"):
                            _cmap["telefon"] = _c
                        elif _cl.startswith("mesto") or _cl.startswith("grad"):
                            _cmap["mesto"] = _c
                        elif _cl.startswith("adresa"):
                            _cmap["adresa"] = _c
                    if "idk" not in _cmap or "naziv" not in _cmap:
                        st.error("Fajlu fale kolone. Potrebne su bar 'ID' i 'Naziv'. Pronađene kolone: " + ", ".join(str(c) for c in _dfk.columns))
                    else:
                        _krows = []
                        for _, _rr in _dfk.iterrows():
                            _idraw = _rr.get(_cmap["idk"])
                            if _idraw is None or str(_idraw).strip() == "" or str(_idraw).strip().lower() == "nan":
                                continue
                            try:
                                _idv = int(float(str(_idraw).strip()))
                            except Exception:
                                continue
                            def _cell(_key):
                                if _key not in _cmap:
                                    return ""
                                _val = _rr.get(_cmap[_key])
                                if _val is None or str(_val).strip().lower() == "nan":
                                    return ""
                                return str(_val).strip()
                            _krows.append({"idk": _idv,
                                           "naziv": _clean_komitent_naziv(_cell("naziv")),
                                           "email": _cell("email").replace(" ", ""),
                                           "telefon": _cell("telefon"),
                                           "mesto": _cell("mesto"),
                                           "adresa": _cell("adresa")})
                        st.caption("Pronađeno " + str(len(_krows)) + " komitenata u fajlu.")
                        if st.button("💾 Sačuvaj šifarnik komitenata", key="kom_save", use_container_width=True):
                            try:
                                _n = sb_komitenti_upsert_rows(_krows)
                                st.session_state["_komitenti_map"] = None
                                st.session_state["_komfull"] = None
                                if _n and _n > 0:
                                    st.success("Sačuvano ✓ U bazi sada ima " + str(_n) + " komitenata. "
                                               "Nazivi će se videti u administraciji (osveži F5).")
                                else:
                                    st.error("Ništa nije upisano u bazu. Pokreni SQL setup za tabelu 'komitenti' (vidi uputstvo).")
                            except Exception as _e:
                                st.error("Greška pri čuvanju: " + str(_e))
                except Exception as _e:
                    st.error("Ne mogu da pročitam fajl: " + str(_e))

    with st.container(border=True):
        st.markdown('<div class="obj-title">📊 Izveštaj prodaje (za direktore)</div>', unsafe_allow_html=True)
        st.caption("Ubaci dve tabele (tabela sistemi + tabela troškova). Pravi se dashboard koji direktori vide kao treću karticu. Čuva se samo poslednji.")
        if not sb_dostupan():
            st.caption("Nedostupno dok Supabase nije podešen.")
        else:
            _ip_c1, _ip_c2 = st.columns(2)
            with _ip_c1:
                _up_sis = st.file_uploader("Tabela sistemi (.xlsx)", type=['xlsx'], key="izp_sis")
            with _ip_c2:
                _up_tro = st.file_uploader("Tabela troškova (.xlsx)", type=['xlsx'], key="izp_tro")
            _ip_q1 = st.radio("Tip izveštaja", ["Potpun (sve 4 kartice)", "Nepotpun (Prodaja + Uspešnost akcije)"],
                              key="izp_potpun", horizontal=True)
            _potpun = _ip_q1.startswith("Potpun")
            _iskljuci = False
            if _potpun:
                _ip_q2 = st.radio("Poslednji mesec u profitabilnosti?", ["Uključi", "Isključi"],
                                  key="izp_iskljuci", horizontal=True)
                _iskljuci = (_ip_q2 == "Isključi")
            if _up_sis is not None and _up_tro is not None:
                if st.button("📊 Generiši i objavi izveštaj prodaje", key="izp_gen", use_container_width=True):
                    try:
                        import io as _io2, base64 as _b64i
                        import izvestaj_prodaje as _izp
                        with st.spinner("Generišem izveštaj prodaje (može par sekundi)..."):
                            _html, _xlsx, _mes, _spr = _izp.generisi_izvestaj_prodaje(
                                _io2.BytesIO(_up_sis.getvalue()), _io2.BytesIO(_up_tro.getvalue()),
                                potpun=_potpun, iskljuci_poslednji=_iskljuci)
                            _xb64 = _b64i.b64encode(_xlsx).decode("ascii") if _xlsx else ""
                            _pj = json.dumps(_spr, ensure_ascii=False) if _spr else ""
                            sb_objavi_izvestaj_prodaje(_html, _xb64, _mes, prodaja_json=_pj)
                        st.success("✅ Izveštaj prodaje objavljen (" + str(_mes) + "). Direktori ga vide u trećoj kartici.")
                    except ModuleNotFoundError:
                        st.error("Nedostaje fajl izvestaj_prodaje.py u projektu — dodaj ga na GitHub pored streamlit_app.py.")
                    except Exception as _e:
                        st.error("Greška: " + str(_e))
                        import traceback as _tb2
                        st.code(_tb2.format_exc())

    with st.container(border=True):
        st.markdown('<div class="obj-title"><span class="obj-badge">1</span> Učitaj Excel jednog sistema</div>', unsafe_allow_html=True)
        up_o = st.file_uploader("Excel fajl (.xlsx)", type=['xlsx', 'xls'], key="obj_upl", label_visibility="collapsed")

    if up_o is not None:
        _obytes = up_o.read()
        st.markdown(f'<div class="success-box">\u2705 Fajl <strong>{up_o.name}</strong> učitan ({len(_obytes)//1024} KB)</div>', unsafe_allow_html=True)
        _oy = None; _omn = None
        try:
            _x = pd.ExcelFile(io.BytesIO(_obytes))
            _sm = {s.strip().lower(): s for s in _x.sheet_names}
            _sp = None
            for _nl, _no in _sm.items():
                if 'prodaja' in _nl:
                    _sp = _no; break
            _pdf = pd.read_excel(_x, sheet_name=_sp)
            _pdf.columns = [c.strip() for c in _pdf.columns]
            _ms = sorted(_pdf[['Godina', 'Mesec']].drop_duplicates().values.tolist())
            _oy = int(_ms[-1][0]); _omn = int(_ms[-1][1]) + 1
            if _omn > 12:
                _omn = 1; _oy += 1
        except Exception:
            _oy = None; _omn = None
        _mk = (str(_oy) + "-" + ("0" + str(_omn))[-2:]) if _oy else None
        _mlbl = mesec_label(_mk) if _mk else "automatski"

        with st.container(border=True):
            st.markdown('<div class="obj-title"><span class="obj-badge">2</span> Parametri porudžbine</div>', unsafe_allow_html=True)
            _oc1, _oc2, _oc3 = st.columns(3)
            with _oc1:
                _o_mes = st.number_input("Broj meseci za porudžbinu", min_value=0.5, value=1.5, step=0.5, format="%.1f", key="obj_mes_num", help="Porudžbina pokriva ovoliko meseci predviđene prodaje (može 1.0, 1.5, 2.0...).")
                _o_ml = st.text_input("Min. lager po artiklu", value="", placeholder="prazno = bez ograničenja", key="obj_ml")
                _o_mo = st.text_input("Min. ukupna porudžbina po objektu", value="", placeholder="prazno = bez ograničenja", key="obj_mo")
            with _oc2:
                _o_mpa = st.text_input("Min. kom po artiklu (po stavci)", value="", placeholder="prazno = bez ograničenja", key="obj_mpa")
                _o_maxpa = st.text_input("Maksimum po komadu (po stavci)", value="", placeholder="prazno = bez ograničenja", key="obj_maxpa")
                _o_tr = st.number_input("Ukupan trosak mkt (RSD)", min_value=0, value=0, step=10000, key="obj_tr")
            with _oc3:
                _o_excl = st.text_area("Isključeni komitenti (ID, zarez)", value=DEFAULT_EXCLUDED, height=110, key="obj_excl")
                _o_syx = st.text_area("Objekti koji prodaju SYX (ID, zarez)", value="", height=90, key="obj_syx",
                                      placeholder="prazno = SYX ide svima; ako upišeš ID-jeve, SYX se predlaže samo u tim objektima")
        _o_min_lager = int(_o_ml) if _o_ml.strip().isdigit() else None
        _o_min_order = int(_o_mo) if _o_mo.strip().isdigit() else None
        _o_min_pa = int(_o_mpa) if _o_mpa.strip().isdigit() else None
        _o_max_pa = int(_o_maxpa) if _o_maxpa.strip().isdigit() else None
        _o_excluded = set()
        for _part in _o_excl.replace('\n', ',').split(','):
            _p = _part.strip()
            if _p.isdigit():
                _o_excluded.add(int(_p))
        _o_syx_set = set()
        for _part in _o_syx.replace('\n', ',').split(','):
            _p = _part.strip()
            if _p.isdigit():
                _o_syx_set.add(int(_p))
        _o_syx_obj = _o_syx_set if _o_syx_set else None

        with st.container(border=True):
            st.markdown('<div class="obj-title"><span class="obj-badge green">\u2713</span> Objavi za koleginice</div>', unsafe_allow_html=True)
            _od1, _od2 = st.columns(2)
            with _od1:
                _osist = st.text_input("Naziv sistema (kako će koleginice videti)", value=os.path.splitext(up_o.name)[0].strip(), key="obj_sist")
            with _od2:
                # Auto mesec = poslednji mesec u fajlu + 1; ali dozvoli izbor (npr. objaviti pod Jul)
                if _mk:
                    _auto_y, _auto_m = int(_mk[:4]), int(_mk[5:7])
                else:
                    _tt2 = datetime.date.today(); _auto_y, _auto_m = _tt2.year, _tt2.month
                _mo_keys = []
                _yy2, _mm2 = _auto_y, _auto_m - 3
                while _mm2 <= 0:
                    _mm2 += 12; _yy2 -= 1
                for _ in range(6):
                    _mo_keys.append(str(_yy2) + "-" + ("0" + str(_mm2))[-2:])
                    _mm2 += 1
                    if _mm2 > 12:
                        _mm2 = 1; _yy2 += 1
                _auto_key = str(_auto_y) + "-" + ("0" + str(_auto_m))[-2:]
                _def_i = _mo_keys.index(_auto_key) if _auto_key in _mo_keys else len(_mo_keys) - 1
                _o_mes_key = st.selectbox("Mesec porudžbine", _mo_keys, index=_def_i,
                                          format_func=mesec_label, key="obj_mes_sel",
                                          help="Automatski je poslednji mesec iz fajla + 1, ali možeš da objaviš pod drugim mesecom (npr. Jul).")
            _o_nedeljni = st.checkbox("Sistemski sistem (trebuju u fiksnom periodu) — pojednostavljen prikaz", value=False, key="obj_nedeljni",
                                      help="Za sisteme koji trebuju sami u fiksnom ritmu i ne možemo da utičemo. Administracija vidi samo objekte sa problemom (realni lager ispod prodaje za zadati period) i jedno polje Napomena.")
            _o_sist_dani = 7
            if _o_nedeljni:
                _cpd = st.columns([1, 2])
                with _cpd[0]:
                    _o_sist_dani = st.number_input("Na koliko dana trebuju (period pokrivenosti)",
                                                   min_value=1, max_value=120, value=7, step=1,
                                                   key="obj_sist_dani",
                                                   help="Npr. 7 = nedeljni, 14 = na dve nedelje, 30 = mesečni, 45 = mesec i po. "
                                                        "Ne trebuju svi isto — ovde upiši koliko dana taj sistem pokriva jednim trebovanjem.")
                with _cpd[1]:
                    st.caption("Objekat je u problemu ako mu realni lager (lager + naknadne porudžbine) ne pokriva "
                               "prodaju za " + str(int(_o_sist_dani)) + " dana. Prag = predikcija × (" + str(int(_o_sist_dani)) + " ÷ 30).")
            # --- Napomena i mejl koje zadaje analitika (administracija ih samo vidi) ---
            # Ako za ovaj sistem/mesec već postoji objava, popuni postojeće vrednosti
            # (jednom), da se ponovnom objavom slučajno ne obrišu.
            _pf_k = "obj_pref_" + str(_o_mes_key) + "_" + str(_osist).strip()
            if _osist.strip() and not st.session_state.get(_pf_k):
                try:
                    _old_p = sb_ucitaj(_o_mes_key, _osist.strip())
                    _old_m = (_old_p or {}).get("meta") or {}
                    _on = (_old_m.get("nap_analitika") or {}).get("tekst", "")
                    _om = (_old_m.get("mail_to_fix") or {}).get("to", "")
                    if _on and not st.session_state.get("obj_napomena"):
                        st.session_state["obj_napomena"] = _on
                    if _om and not st.session_state.get("obj_mail_fix"):
                        st.session_state["obj_mail_fix"] = _om
                except Exception:
                    pass
                st.session_state[_pf_k] = True
            with st.expander("📝 Napomena i mejl za administraciju (opciono)",
                             expanded=bool(st.session_state.get("obj_napomena")
                                           or st.session_state.get("obj_mail_fix"))):
                _o_napomena = st.text_area(
                    "Napomena za administraciju", key="obj_napomena", height=80,
                    placeholder="npr. Mejl se šalje petkom. Kontakt osoba je Marko, zvati posle 10h.",
                    help="Prikazuje se administraciji na vrhu ovog sistema. Koristi za dogovor "
                         "sa sistemom — kada se šalje, kome, šta da paze.")
                _o_mail_fix = st.text_input(
                    "Mejl nadležnog (zaključan za administraciju)", key="obj_mail_fix",
                    placeholder="npr. nabavka@medius.rs — ostavi prazno da ga upiše administracija",
                    help="Ako ovde upišeš mejl, administracija ga NE može menjati — samo šalje na njega. "
                         "Ako ostaviš prazno, administracija sama upisuje mejl.")
                if (_o_mail_fix or "").strip() and "@" not in _o_mail_fix:
                    st.warning("Mejl ne izgleda ispravno (nema @).")

            if not sb_dostupan():
                st.info("Objava nije moguća dok Supabase nije podešen.")
            elif st.button("📤 Objavi za koleginice", use_container_width=True, key="obj_btn"):
                if not _osist.strip():
                    st.error("Upiši naziv sistema.")
                else:
                    try:
                        _pb = st.progress(0, "Računam porudžbinu...")
                        _eng = PredictionEngine(_obytes, _o_excluded, alpha, beta, _o_min_lager, _o_min_order, _o_tr, None, _o_min_pa, meseci=float(_o_mes), max_per_artikal=_o_max_pa, syx_objekti=_o_syx_obj)
                        _res = _eng.run(_pb)
                        _pb.empty()
                        # mesec pod kojim se objavljuje = izabrani mesec (podrazumevano auto)
                        _mk2 = _o_mes_key
                        _mlbl2 = mesec_label(_mk2)
                        # Presek za „posle 01." = 1. u mesecu POSLE poslednjeg meseca podataka
                        # (nezavisno od izabranog meseca objave). Npr. podaci do jula -> presek 01.08.
                        try:
                            _lg3, _lm3 = _eng.meseci_order[-1]
                            _pm3 = int(_lm3) + 1; _py3 = int(_lg3)
                            if _pm3 > 12:
                                _pm3 = 1; _py3 += 1
                            _presek_iso = str(_py3) + "-" + ("0" + str(_pm3))[-2:] + "-01"
                        except Exception:
                            _presek_iso = None
                        _stavke = stavke_iz_rezultata(_res, _eng)
                        _payload = {"mesec_label": _mlbl2, "meta": {"pred_label": _eng.pred_label, "order_label": _eng.order_label, "min_lager": _eng.min_lager, "meseci": round(float(_o_mes), 1), "presek": _presek_iso, "nedeljni": bool(_o_nedeljni), "nedeljni_dani": int(_o_sist_dani), "mesec_nazivi": list(getattr(_eng, "mesec_labels", []) or []), "generisano": _now().strftime("%d.%m.%Y %H:%M"), "n_objekata": int(len({_s2['idk'] for _s2 in _stavke})), "n_sistem_ukupno": int(getattr(_eng, "num_komitenti", 0)), "ukupno_kom": int(sum(_s2['kol'] for _s2 in _stavke))}, "stavke": _stavke}
                        # Napomena + zaključan mejl koje zadaje analitika
                        _nap_an = (st.session_state.get("obj_napomena", "") or "").strip()
                        _mf_an = (st.session_state.get("obj_mail_fix", "") or "").strip()
                        if _nap_an:
                            _payload["meta"]["nap_analitika"] = {
                                "tekst": _nap_an[:1500],
                                "ko": st.session_state.get("admin_user", "Analitika"),
                                "at": _now().strftime("%d.%m.%Y %H:%M")}
                        if _mf_an:
                            _payload["meta"]["mail_to_fix"] = {
                                "to": _mf_an[:200],
                                "ko": st.session_state.get("admin_user", "Analitika"),
                                "at": _now().strftime("%d.%m.%Y %H:%M")}
                        try:
                            _payload["direktor"] = direktor_blok(_eng, _res)
                        except Exception:
                            _payload["direktor"] = None
                        _xb64 = None
                        try:
                            import base64 as _b64
                            _xbuf = create_excel(_eng, ukljuci_model=False)
                            _xb64 = _b64.b64encode(_xbuf.getvalue()).decode("ascii")
                        except Exception:
                            _xb64 = None
                        sb_objavi(_mk2, _osist, _payload, xlsx_b64=_xb64)
                        st.success(f"\u2705 Objavljeno: {_osist} \u2014 {_mlbl2} \u00b7 {len(_stavke)} stavki, {_payload['meta']['n_objekata']} objekata. Osveži (F5) da se ažurira lista gore.")
                    except Exception as _e:
                        st.error(f"Greška pri objavi: {_e}")
                        import traceback as _tb
                        st.code(_tb.format_exc())
    else:
        st.caption("\u2191 Učitaj Excel fajl da bi objavio porudžbinu za koleginice.")

with tab_ana:
    with st.expander("⚙️ Parametri analize", expanded=False):
        pc1, pc2, pc3 = st.columns(3)
        with pc1:
            st.markdown("**📦 Porudžbina**")
            meseci_ana = st.number_input("Broj meseci za porudžbinu", min_value=0.5, value=1.5, step=0.5, format="%.1f",
                                         help="Porudžbina pokriva ovoliko meseci predviđene prodaje (može 1.0, 1.5, 2.0...).")
            _ml_str = st.text_input("Minimalni lager po artiklu", value="", placeholder="prazno = bez ograničenja")
            min_lager = int(_ml_str) if _ml_str.strip().isdigit() else None
            _mo_str = st.text_input("Min. ukupna porudžbina po objektu", value="", placeholder="prazno = bez ograničenja")
            min_order = int(_mo_str) if _mo_str.strip().isdigit() else None
            _mpa_str = st.text_input("Min. kom po artiklu (po stavci)", value="", placeholder="prazno = bez ograničenja",
                                      help="Ako je porudžbina za jedan artikal manja od ovog broja (ali > 0), podiže se na minimum. Nule ostaju nule.")
            min_per_artikal = int(_mpa_str) if _mpa_str.strip().isdigit() else None
            _maxpa_str = st.text_input("Maksimum po komadu (po stavci)", value="", placeholder="prazno = bez ograničenja",
                                       help="Gornja granica porudžbine po jednom artiklu (po stavci). Ako je predlog veći, spušta se na ovaj maksimum.")
            max_per_artikal = int(_maxpa_str) if _maxpa_str.strip().isdigit() else None
        with pc2:
            st.markdown("**💰 Troškovi**")
            mesecni_trosak = st.number_input(
                "Ukupan trosak mkt/ulistavanja (RSD)",
                min_value=0, value=0, step=10000)
        with pc3:
            st.markdown("**⛔ Isključeni komitenti**")
            excluded_str = st.text_area("ID-evi razdvojeni zarezom", value=DEFAULT_EXCLUDED, height=80)
            st.markdown("**🧊 Objekti koji prodaju SYX**")
            syx_str = st.text_area("ID-evi (prazno = SYX ide svima)", value="", height=70, key="ana_syx",
                                   help="Ako upišeš ID-jeve, SYX vrećice se predlažu u porudžbini samo u tim objektima; u ostalima se stave na 0.")
    excluded = set()
    for part in excluded_str.replace('\n', ',').split(','):
        p = part.strip()
        if p.isdigit(): excluded.add(int(p))
    syx_set = set()
    for part in syx_str.replace('\n', ',').split(','):
        p = part.strip()
        if p.isdigit(): syx_set.add(int(p))
    syx_objekti = syx_set if syx_set else None
    uploaded = st.file_uploader("Učitaj Excel fajl sa podacima", type=['xlsx','xls'])
    if uploaded:
        file_bytes = uploaded.read()
        st.markdown(f'<div class="success-box">✅ Fajl <strong>{uploaded.name}</strong> učitan ({len(file_bytes)//1024} KB)</div>', unsafe_allow_html=True)
        st.markdown("")
        try:
            _xls = pd.ExcelFile(io.BytesIO(file_bytes))
            _sm = {s.strip().lower(): s for s in _xls.sheet_names}
            _sp = None
            for kw in ['prodaja']:
                for nl, no in _sm.items():
                    if kw in nl: _sp = no; break
            if _sp:
                _prod = pd.read_excel(_xls, sheet_name=_sp); _prod.columns=[c.strip() for c in _prod.columns]
                _meseci = sorted(_prod[['Godina','Mesec']].drop_duplicates().values.tolist())
                _mn={1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'Maj',6:'Jun',7:'Jul',8:'Avg',9:'Sep',10:'Okt',11:'Nov',12:'Dec'}
                _labels = [f"{_mn.get(int(m),'?' )} {int(g)}" for g,m in _meseci]
                st.markdown('**📅 Period za analizu** (OOS, Profitabilnost, Akcija — ne utiče na predikciju):')
                selected_labels = st.multiselect("Odaberi mesece", _labels, default=_labels, help="Predikcija uvek koristi sve mesece. Ovaj filter se odnosi samo na analitiku.")
                if not selected_labels:
                    st.warning("⚠️ Mora biti odabran bar jedan mesec za analizu. Automatski je odabran poslednji mesec.")
                    selected_labels = [_labels[-1]] if _labels else []
                selected_meseci = [_meseci[i] for i, lb in enumerate(_labels) if lb in selected_labels]
            else:
                selected_labels = []; selected_meseci = []
        except:
            selected_labels = []; selected_meseci = []
        if st.button("🚀 POKRENI ANALIZU", use_container_width=True):
            progress_bar = st.progress(0)
            try:
                engine = PredictionEngine(file_bytes, excluded, alpha, beta, min_lager, min_order, mesecni_trosak, selected_meseci, min_per_artikal, meseci=float(meseci_ana), max_per_artikal=max_per_artikal, syx_objekti=syx_objekti)
                result = engine.run(progress_bar)
                st.session_state["last_engine"] = engine
                st.session_state["last_result"] = result
                st.session_state["last_filename"] = uploaded.name
                st.markdown("---")
                tp = int(result['Predikcija'].sum()); tl = int(result['Lager_danas'].sum())
                t1 = int(result[~result['ID KOMITENTA'].isin(excluded)]['Porudzbina_1'].sum())
                t2 = int(result[~result['ID KOMITENTA'].isin(excluded)]['Porudzbina_2'].sum())
                if engine.has_prices:
                    tab1, tab2 = st.tabs(["📦 Porudžbina", "💰 Profitabilnost objekata & OOS"])
                else:
                    tab1, = st.tabs(["📦 Porudžbina"])
                with tab1:
                    n_obj_salji = int(result[result['Porudzbina_2'] > 0]['ID KOMITENTA'].nunique())
                    tp_prosek = int(result['Prosek'].sum())
                    m1,m2,m3,m4,m5 = st.columns(5)
                    m1.markdown(f'<div class="metric-card"><div class="metric-value">{tp:,}</div><div class="metric-label">Predikcija (kom)</div></div>', unsafe_allow_html=True)
                    m2.markdown(f'<div class="metric-card"><div class="metric-value">{tp_prosek:,}</div><div class="metric-label">Prosek (kom)</div></div>', unsafe_allow_html=True)
                    m3.markdown(f'<div class="metric-card"><div class="metric-value-green">{t2:,}</div><div class="metric-label">Porudžbina (kom)</div></div>', unsafe_allow_html=True)
                    m4.markdown(f'<div class="metric-card"><div class="metric-value">{n_obj_salji:,}</div><div class="metric-label">Objekata prima robu</div></div>', unsafe_allow_html=True)
                    m5.markdown(f'<div class="metric-card"><div class="metric-value">{tl:,}</div><div class="metric-label">Lager danas</div></div>', unsafe_allow_html=True)
                    st.markdown("")
                    st.markdown("<div style='margin:24px 0 4px 0;'></div>", unsafe_allow_html=True)
                    ml = engine.mesec_labels
                    df_r = engine.df_result.copy()
                    kom_mes = {}
                    for lb in ml:
                        col_lb = f'{lb}_Prodaja'
                        if col_lb in df_r.columns:
                            grp = df_r.groupby('ID KOMITENTA')[col_lb].sum()
                            for kid, v in grp.items():
                                if kid not in kom_mes: kom_mes[kid] = {}
                                kom_mes[kid][lb] = int(v)
                    import numpy as _np2
                    def _is_rastuci(vals5, dozvoljeni_sum=1):
                        padovi = sum(1 for i in range(1, len(vals5)) if vals5[i] < vals5[i-1])
                        return padovi <= dozvoljeni_sum and vals5[-1] > vals5[0] and sum(vals5) >= 10
                    def _is_padajuci(vals5, dozvoljeni_sum=1):
                        rasti = sum(1 for i in range(1, len(vals5)) if vals5[i] > vals5[i-1])
                        return rasti <= dozvoljeni_sum and vals5[-1] < vals5[0] and sum(vals5) >= 10
                    def _rast_pct(vals5):
                        first = vals5[0] if vals5[0] > 0 else 1
                        return (vals5[-1] - vals5[0]) / first * 100
                    rastuci_list = []
                    padajuci_list = []
                    for kid, mes_vals in kom_mes.items():
                        vals_all = [mes_vals.get(lb, 0) for lb in ml]
                        vals5 = vals_all[-5:] if len(vals_all) >= 5 else vals_all
                        if len(vals5) < 3: continue
                        if _is_rastuci(vals5):
                            rastuci_list.append({
                                'ID': kid, 'Ukupno': sum(vals_all),
                                'Vals': vals_all, 'Vals5': vals5,
                                'Rast': _rast_pct(vals5),
                                'Zadnji': vals5[-1], 'Prvi': vals5[0],
                            })
                        elif _is_padajuci(vals5):
                            padajuci_list.append({
                                'ID': kid, 'Ukupno': sum(vals_all),
                                'Vals': vals_all, 'Vals5': vals5,
                                'Pad': _rast_pct(vals5),
                                'Zadnji': vals5[-1], 'Prvi': vals5[0],
                            })
                    rastuci_list = sorted(rastuci_list, key=lambda x: x['Rast'], reverse=True)[:10]
                    padajuci_list = sorted(padajuci_list, key=lambda x: x['Pad'])[:10]
                    def _render_trend_section(title, icon, color, items, is_rast):
                        label_color = "#10b981" if is_rast else "#ef4444"
                        label_bg = "#f0fdf4" if is_rast else "#fef2f2"
                        if not items:
                            components.html(f"""<!DOCTYPE html><html><body style="margin:0;padding:4px 0;font-family:'DM Sans',sans-serif;">
                            <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">
                                <span style="font-size:17px;">{icon}</span>
                                <span style="font-size:13px;font-weight:700;color:#111;">{title}</span>
                            </div>
                            <div style="color:#aaa;font-size:13px;padding:12px 0;">Nema podataka za prikaz</div>
                            </body></html>""", height=80)
                            return
                        rows_html = ""
                        for r in items:
                            vals5 = r['Vals5']
                            mx = max(vals5) if max(vals5) > 0 else 1
                            bars = "".join(
                                f'<div style="flex:1;display:flex;flex-direction:column;justify-content:flex-end;gap:0;">'
                                f'<div style="height:{int(v/mx*28)}px;background:{"linear-gradient(180deg,#a855f7,#c084fc)" if is_rast else "linear-gradient(180deg,#ec4899,#f9a8d4)"};border-radius:2px 2px 0 0;min-height:2px;"></div></div>'
                                for v in vals5
                            )
                            sign = "+" if is_rast else ""
                            pct = r['Rast'] if is_rast else r['Pad']
                            rows_html += f"""<div style="display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid #f3f4f6;">
                                <div style="font-family:'DM Mono',monospace;font-size:14px;font-weight:500;color:#111;width:46px;flex-shrink:0;">{int(r["ID"])}</div>
                                <div style="display:flex;align-items:flex-end;gap:2px;height:32px;width:90px;flex-shrink:0;">{bars}</div>
                                <div style="flex:1;font-size:11px;color:#aaa;">{int(r["Ukupno"]):,} kom</div>
                                <div style="font-size:12px;font-weight:700;color:{label_color};white-space:nowrap;">{sign}{pct:.0f}% &nbsp;<span style="font-weight:400;color:#bbb;font-size:11px;">({int(r["Prvi"])}→{int(r["Zadnji"])})</span></div>
                            </div>"""
                        h_px = len(items) * 48 + 56
                        components.html(f"""<!DOCTYPE html><html>
                        <head><link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@400;600;700&display=swap" rel="stylesheet"></head>
                        <body style="margin:0;padding:4px 0;font-family:'DM Sans',sans-serif;background:white;">
                            <div style="display:flex;align-items:center;gap:8px;margin-bottom:14px;">
                                <span style="font-size:17px;">{icon}</span>
                                <span style="font-size:13px;font-weight:700;color:#111;">{title}</span>
                                <span style="font-size:10px;font-weight:700;color:{label_color};background:{label_bg};border-radius:20px;padding:2px 8px;">zadnjih 5 mes.</span>
                            </div>
                            <div style="font-size:9px;color:#ccc;display:flex;gap:10px;margin-bottom:4px;">
                                <span style="width:46px;"></span>
                                <span style="width:90px;text-align:center;text-transform:uppercase;letter-spacing:.5px;">trend</span>
                                <span style="flex:1;text-transform:uppercase;letter-spacing:.5px;">ukupno</span>
                                <span style="text-transform:uppercase;letter-spacing:.5px;">rast (prvi→zadnji)</span>
                            </div>
                            {rows_html}
                        </body></html>""", height=h_px)
                    def _render_oos_section(items, max_val):
                        if not items:
                            components.html('''<!DOCTYPE html><html><body style="margin:0;padding:4px 0;font-family:sans-serif;">
                            <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">
                                <span style="font-size:17px;">🔴</span>
                                <span style="font-size:13px;font-weight:700;color:#111;">OOS — Lager 0, najveći potencijal</span>
                            </div>
                            <div style="color:#aaa;font-size:13px;">Nema OOS podataka</div>
                            </body></html>''', height=80)
                            return
                        rows_html = ""
                        for r in items:
                            pct = int(r['Izgubljeno'] / max_val * 100)
                            rows_html += f"""<div style="padding:9px 0;border-bottom:1px solid #f9f9f9;">
                                <div style="display:flex;align-items:center;gap:10px;margin-bottom:5px;">
                                    <div style="font-family:'DM Mono',monospace;font-size:14px;font-weight:500;color:#111;width:46px;flex-shrink:0;">{int(r["ID KOMITENTA"])}</div>
                                    <div style="font-size:10px;font-weight:700;color:#ec4899;background:#fdf2f8;border-radius:4px;padding:2px 7px;">{int(r["Artikala"])} artikala bez robe</div>
                                    <div style="margin-left:auto;font-family:'DM Mono',monospace;font-size:13px;font-weight:700;color:#7c3aed;">{int(r["Izgubljeno"]):,} RSD</div>
                                </div>
                                <div style="height:5px;background:#f5f0ff;border-radius:99px;overflow:hidden;">
                                    <div style="width:{pct}%;height:100%;background:linear-gradient(90deg,#a855f7,#ec4899);border-radius:99px;"></div>
                                </div>
                            </div>"""
                        h_px = len(items) * 54 + 56
                        components.html(f"""<!DOCTYPE html><html>
                        <head><link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@400;600;700&display=swap" rel="stylesheet"></head>
                        <body style="margin:0;padding:4px 0;font-family:'DM Sans',sans-serif;background:white;">
                            <div style="display:flex;align-items:center;gap:8px;margin-bottom:14px;">
                                <span style="font-size:17px;">🔴</span>
                                <span style="font-size:13px;font-weight:700;color:#111;">OOS — Lager 0, najveći potencijal</span>
                                <span style="font-size:10px;font-weight:700;color:#ec4899;background:#fdf2f8;border-radius:20px;padding:2px 8px;">top 10</span>
                            </div>
                            <div style="font-size:9px;color:#ccc;display:flex;gap:10px;margin-bottom:4px;align-items:center;">
                                <span style="width:46px;"></span>
                                <span style="flex:1;text-transform:uppercase;letter-spacing:.5px;"></span>
                                <span style="text-transform:uppercase;letter-spacing:.5px;">izgubljen profit</span>
                            </div>
                            {rows_html}
                        </body></html>""", height=h_px)
                    col_rast, col_pad = st.columns(2)
                    with col_rast:
                        _render_trend_section("Rastući trendovi", "📈", "#a855f7", rastuci_list, True)
                    with col_pad:
                        _render_trend_section("Padajući trendovi", "📉", "#ec4899", padajuci_list, False)
                    st.markdown("<div style='margin:20px 0 4px 0;'></div>", unsafe_allow_html=True)
                    if engine.has_prices and len(engine.df_oos) > 0:
                        oos_k = engine.df_oos.copy()
                        if 'Lager_danas' in oos_k.columns:
                            oos_k = oos_k[oos_k['Lager_danas'] == 0]
                        oos_top = oos_k.groupby('ID KOMITENTA').agg(
                            Izgubljeno=('Izgubljeni_profit','sum'),
                            Artikala=('id artikla','nunique')
                        ).reset_index().sort_values('Izgubljeno', ascending=False).head(10)
                        oos_items = oos_top.to_dict('records')
                        oos_max = int(oos_top['Izgubljeno'].max()) if len(oos_top) > 0 else 1
                    else:
                        oos_items = []; oos_max = 1
                    col_oos2, col_empty = st.columns(2)
                    with col_oos2:
                        _render_oos_section(oos_items, oos_max)
                if engine.has_prices:
                    with tab2:
                        period_str2 = ", ".join(engine.analitika_labels) if engine.analitika_labels else "svi meseci"
                        n_mes = len(engine.analitika_labels) if engine.analitika_labels else len(engine.mesec_labels)
                        n_obj = engine.num_komitenti
                        prof = engine.df_profit_obj.copy()
                        total_bruto = int(prof['Bruto_profit'].sum())
                        total_neto = int(prof['Neto_profit'].sum())
                        total_trosak = int(prof['Trosak_mkt'].sum())
                        total_oos_izgubljen = int(engine.df_oos['Izgubljeni_profit'].sum()) if len(engine.df_oos) > 0 else 0
                        mes_trosak = total_trosak / max(n_mes, 1)
                        mes_bruto = total_bruto / max(n_mes, 1)
                        mes_neto = total_neto / max(n_mes, 1)
                        mes_oos = total_oos_izgubljen / max(n_mes, 1)
                        st.caption(f"📅 Period analize: **{period_str2}** · {n_obj} objekata · {n_mes} meseci")
                        ka, kb, kc, kd = st.columns(4)
                        def _kard(col, label, total, mes, color, prefix=""):
                            col.markdown(f"""
                            <div style="background:white;border-radius:12px;padding:16px 18px;
                                border-left:4px solid {color};box-shadow:0 2px 8px rgba(0,0,0,0.07);height:100%;">
                                <div style="font-size:10px;color:#999;font-weight:600;letter-spacing:.5px;text-transform:uppercase;margin-bottom:6px;">{label}</div>
                                <div style="font-size:22px;font-weight:700;color:{color};">{prefix}{total:,.0f} RSD</div>
                                <div style="font-size:11px;color:#aaa;margin-top:3px;">{prefix}{mes:,.0f} RSD / mesec</div>
                            </div>""", unsafe_allow_html=True)
                        _kard(ka, f"Ukupan trosak · {n_mes} meseci", total_trosak, mes_trosak, "#a855f7")
                        _kard(kb, f"Bruto profit · {n_mes} meseci", total_bruto, mes_bruto, "#10b981")
                        _kard(kc, f"Neto profit · {n_mes} meseci", total_neto, mes_neto, "#7c3aed" if total_neto > 0 else "#ec4899")
                        _kard(kd, f"OOS izgubljen · {n_mes} meseci", total_oos_izgubljen, mes_oos, "#ec4899", prefix="-")
                        st.markdown("<div style='margin:20px 0 4px 0;'></div>", unsafe_allow_html=True)
                        a_labels_trend = engine.analitika_labels if engine.analitika_labels else engine.mesec_labels
                        a_meseci_trend = engine.analitika_meseci if (engine.analitika_meseci and len(engine.analitika_meseci) > 0) else engine.meseci_order
                        bruto_po_mes = []
                        neto_po_mes = []
                        for i, lb in enumerate(a_labels_trend):
                            col_bruto = f'Bruto_{lb}'
                            col_neto = f'Neto_{lb}'
                            bruto_val = prof[col_bruto].sum() if col_bruto in prof.columns else 0
                            neto_val = prof[col_neto].sum() if col_neto in prof.columns else 0
                            bruto_po_mes.append((lb, bruto_val))
                            neto_po_mes.append((lb, neto_val))
                        def _trend_recenica(podaci, naziv):
                            vals = [v for _, v in podaci]
                            if len(vals) < 2: return ""
                            prvi_lb, prvi_v = podaci[0]
                            posl_lb, posl_v = podaci[-1]
                            if prvi_v == 0: return ""
                            promena_pct = ((posl_v - prvi_v) / abs(prvi_v)) * 100
                            smer = "porastao" if promena_pct > 0 else "pao"
                            boja = "#10b981" if promena_pct > 0 else "#ec4899"
                            return f'<span style="color:{boja};font-weight:600;">{naziv} je {smer} za {abs(promena_pct):.0f}%</span> — od <b>{prvi_v:,.0f} RSD</b> ({prvi_lb}) do <b>{posl_v:,.0f} RSD</b> ({posl_lb}).'
                        def _bar_chart_html(podaci, max_val, color_pos, color_neg):
                            bars = ""
                            for lb, val in podaci:
                                pct = abs(val) / max_val * 100 if max_val > 0 else 0
                                pct = min(pct, 100)
                                color = color_pos if val >= 0 else color_neg
                                val_fmt = f"{val:,.0f} RSD"
                                bars += f"""
                                <div style="display:flex;align-items:center;margin-bottom:5px;gap:8px;">
                                    <div style="width:52px;font-size:11px;color:#888;text-align:right;flex-shrink:0;">{lb}</div>
                                    <div style="flex:1;background:#f5f0ff;border-radius:3px;height:18px;position:relative;">
                                        <div style="width:{pct:.1f}%;background:{color};height:100%;border-radius:3px;transition:width .3s;"></div>
                                    </div>
                                    <div style="width:110px;font-size:11px;color:#555;font-weight:600;flex-shrink:0;">{val_fmt}</div>
                                </div>"""
                            return f'<div style="padding:4px 0;">{bars}</div>'
                        max_bruto = max(abs(v) for _, v in bruto_po_mes) if bruto_po_mes else 1
                        max_neto = max(abs(v) for _, v in neto_po_mes) if neto_po_mes else 1
                        col_bruto, col_neto = st.columns(2)
                        with col_bruto:
                            st.markdown('<div class="section-title">📈 Mesečni trend bruto profita</div>', unsafe_allow_html=True)
                            rec_b = _trend_recenica(bruto_po_mes, "Bruto profit")
                            if rec_b: st.markdown(f'<p style="font-size:13px;color:#555;margin-bottom:6px;">{rec_b}</p>', unsafe_allow_html=True)
                            chart_b = _bar_chart_html(bruto_po_mes, max_bruto, "#a855f7", "#ec4899")
                            components.html(f'<!DOCTYPE html><html><body style="margin:0;padding:8px 12px;font-family:sans-serif;">{chart_b}</body></html>', height=len(bruto_po_mes)*28+20)
                        with col_neto:
                            st.markdown('<div class="section-title">📉 Mesečni trend neto profita</div>', unsafe_allow_html=True)
                            rec_n = _trend_recenica(neto_po_mes, "Neto profit")
                            if rec_n: st.markdown(f'<p style="font-size:13px;color:#555;margin-bottom:6px;">{rec_n}</p>', unsafe_allow_html=True)
                            chart_n = _bar_chart_html(neto_po_mes, max_neto, "#7c3aed", "#ec4899")
                            components.html(f'<!DOCTYPE html><html><body style="margin:0;padding:8px 12px;font-family:sans-serif;">{chart_n}</body></html>', height=len(neto_po_mes)*28+20)
                        st.markdown("<div style='margin:20px 0 4px 0;'></div>", unsafe_allow_html=True)
                        st.markdown('<div class="section-title">🏪 Profitabilnost po objektima</div>', unsafe_allow_html=True)
                        ukupno_obj = len(prof)
                        neto_neg = prof[prof['Neto_profit'] <= 0]
                        n_neto_neg = len(neto_neg)
                        oos_neg = prof[(prof['Neto_profit'] <= 0) & (prof['Potencijalni_profit'] > 0)]
                        n_oos_neg = len(oos_neg)
                        pravi_neg = prof[(prof['Neto_profit'] <= 0) & (prof['Potencijalni_profit'] <= 0)]
                        n_pravi_neg = len(pravi_neg)
                        pct_pravi = round(n_pravi_neg / max(ukupno_obj, 1) * 100)
                        trosak_po_obj = engine.trosak_po_objektu
                        trosak_mes_obj = trosak_po_obj / max(n_mes, 1)
                        usteda_trosak = n_pravi_neg * trosak_po_obj
                        usteda_gubitak = abs(pravi_neg['Neto_profit'].sum()) if n_pravi_neg > 0 else 0
                        usteda_mes = (usteda_trosak + usteda_gubitak) / max(n_mes, 1)
                        n_profitabilni = ukupno_obj - n_neto_neg
                        pct_prof = n_profitabilni / max(ukupno_obj, 1)
                        pct_oos_neg_v = n_oos_neg / max(ukupno_obj, 1)
                        pct_pravi_v = n_pravi_neg / max(ukupno_obj, 1)
                        cx, cy, r_out, r_in = 110, 110, 90, 60
                        def _arc_path(cx, cy, r, start_deg, end_deg):
                            s = math.radians(start_deg - 90)
                            e = math.radians(end_deg - 90)
                            large = 1 if (end_deg - start_deg) > 180 else 0
                            x1,y1 = cx+r*math.cos(s), cy+r*math.sin(s)
                            x2,y2 = cx+r*math.cos(e), cy+r*math.sin(e)
                            return f"M {x1:.1f} {y1:.1f} A {r} {r} 0 {large} 1 {x2:.1f} {y2:.1f}"
                        def _donut_seg(cx, cy, ro, ri, start_deg, end_deg, color):
                            if end_deg - start_deg < 0.5: return ""
                            oa = _arc_path(cx, cy, ro, start_deg, end_deg)
                            s2 = math.radians(end_deg - 90); s1 = math.radians(start_deg - 90)
                            x_ie, y_ie = cx+ri*math.cos(s2), cy+ri*math.sin(s2)
                            x_is, y_is = cx+ri*math.cos(s1), cy+ri*math.sin(s1)
                            large = 1 if (end_deg - start_deg) > 180 else 0
                            x2o,y2o = cx+ro*math.cos(s2), cy+ro*math.sin(s2)
                            x1o,y1o = cx+ro*math.cos(s1), cy+ro*math.sin(s1)
                            return f'<path d="{oa} L {x_ie:.1f} {y_ie:.1f} A {ri} {ri} 0 {large} 0 {x_is:.1f} {y_is:.1f} Z" fill="{color}"/>'
                        deg_prof = pct_prof * 360
                        deg_oos = pct_oos_neg_v * 360
                        deg_pravi = pct_pravi_v * 360
                        seg1 = _donut_seg(cx, cy, r_out, r_in, 0, deg_prof, "#10b981")
                        seg2 = _donut_seg(cx, cy, r_out, r_in, deg_prof, deg_prof+deg_pravi, "#ec4899")
                        seg3 = _donut_seg(cx, cy, r_out, r_in, deg_prof+deg_pravi, deg_prof+deg_pravi+deg_oos, "#a855f7")
                        donut_svg = f"""<svg width="220" height="220" xmlns="http://www.w3.org/2000/svg">
                            {seg1}{seg2}{seg3}
                            <circle cx="{cx}" cy="{cy}" r="{r_in}" fill="white"/>
                            <text x="{cx}" y="{cy-8}" text-anchor="middle" font-size="26" font-weight="700" fill="#111" font-family="sans-serif">{n_profitabilni}</text>
                            <text x="{cx}" y="{cy+14}" text-anchor="middle" font-size="12" fill="#888" font-family="sans-serif">profitabilnih</text>
                        </svg>
                        <div style="margin-top:8px;font-size:12px;font-family:sans-serif;">
                            <div style="display:flex;align-items:center;gap:6px;margin-bottom:5px;">
                                <span style="width:12px;height:12px;background:#10b981;border-radius:2px;display:inline-block;flex-shrink:0;"></span>
                                <span style="color:#555;"><strong>{n_profitabilni} profitabilnih</strong> ({round(pct_prof*100)}% mreže)</span>
                            </div>
                            <div style="display:flex;align-items:center;gap:6px;margin-bottom:5px;">
                                <span style="width:12px;height:12px;background:#ec4899;border-radius:2px;display:inline-block;flex-shrink:0;"></span>
                                <span style="color:#555;"><strong>{n_pravi_neg} neprofitabilnih</strong> ({round(pct_pravi_v*100)}% mreže)</span>
                            </div>
                            <div style="display:flex;align-items:center;gap:6px;">
                                <span style="width:12px;height:12px;background:#a855f7;border-radius:2px;display:inline-block;flex-shrink:0;"></span>
                                <span style="color:#555;"><strong>{n_oos_neg} neto-neg. OOS</strong> potencijal</span>
                            </div>
                        </div>"""
                        tekst = f"""
    <div style="background:white;border-radius:12px;padding:20px 24px;box-shadow:0 2px 8px rgba(0,0,0,0.06);margin-bottom:16px;font-size:14px;line-height:1.8;color:#333;">
    <p>Od <strong>{ukupno_obj} objekata</strong>, <strong>{n_neto_neg}</strong> je neto negativno.
    Medjutim, <strong>{n_oos_neg}</strong> od njih ima negativan neto isključivo zbog OOS-a — kada se uračuna izgubljena zarada,
    njihov potencijal je pozitivan. Ovi objekti nisu problem, samo nisu imali robu.</p>
    <p>Pravih neprofitabilnih je <strong>{n_pravi_neg}</strong> ({pct_pravi}% ukupne mreže) — negativni čak i po potencijalu.
    Trošak po objektu je <strong>{trosak_po_obj:,.0f} RSD</strong> za {n_mes} {'mesec' if n_mes==1 else 'meseci'} /
    <strong>{trosak_mes_obj:,.0f} RSD</strong> mesečno.</p>
    <p>Zatvaranjem <strong>{n_pravi_neg} pravih neprofitabilnih</strong> skidamo trošak
    <strong>{n_pravi_neg} × {trosak_po_obj:,.0f} RSD = {usteda_trosak:,.0f} RSD</strong>
    ({usteda_trosak/max(n_mes,1):,.0f} RSD/mes) i prestajemo da gubimo
    <strong>{usteda_gubitak:,.0f} RSD</strong> ({usteda_gubitak/max(n_mes,1):,.0f} RSD/mes) na negativnim objektima.
    Ostaju samo objekti koji su u plusu.</p>
    </div>"""
                        col_tekst, col_donut = st.columns([3, 1])
                        with col_tekst:
                            st.markdown(tekst, unsafe_allow_html=True)
                        with col_donut:
                            components.html(f"""<!DOCTYPE html><html><body style="margin:0;padding:12px 8px;font-family:sans-serif;background:transparent;">
                                {donut_svg}
                            </body></html>""", height=310)
                        a_labels_trend2 = engine.analitika_labels if engine.analitika_labels else engine.mesec_labels
                        a_meseci_trend2 = engine.analitika_meseci if (engine.analitika_meseci and len(engine.analitika_meseci) > 0) else engine.meseci_order
                        chart_mes_data = []
                        for i, (lb, (g, m)) in enumerate(zip(a_labels_trend2, a_meseci_trend2)):
                            col_neto_lb = f'Neto_{lb}'
                            if col_neto_lb in prof.columns:
                                n_prof_mes = int((prof[col_neto_lb] > 0).sum())
                                n_nepr_mes = int((prof[col_neto_lb] <= 0).sum())
                            else:
                                n_prof_mes = 0; n_nepr_mes = 0
                            chart_mes_data.append((lb, n_prof_mes, n_nepr_mes))
                        if chart_mes_data:
                            max_obj_mes = max(a + b for _, a, b in chart_mes_data) if chart_mes_data else 1
                            bar_w = max(30, min(60, 700 // max(len(chart_mes_data), 1)))
                            bars_html = ""
                            for lb, np_v, nn_v in chart_mes_data:
                                h_p = int(np_v / max(max_obj_mes, 1) * 140)
                                h_n = int(nn_v / max(max_obj_mes, 1) * 140)
                                bars_html += f"""
                                <div style="display:flex;flex-direction:column;align-items:center;gap:2px;">
                                    <div style="display:flex;align-items:flex-end;gap:3px;height:160px;">
                                        <div style="width:{bar_w}px;height:{h_p}px;background:#a855f7;border-radius:3px 3px 0 0;position:relative;">
                                            <span style="position:absolute;top:-18px;left:50%;transform:translateX(-50%);font-size:10px;font-weight:700;color:#7c3aed;white-space:nowrap;">{np_v}</span>
                                        </div>
                                        <div style="width:{bar_w}px;height:{h_n}px;background:#ec4899;border-radius:3px 3px 0 0;position:relative;">
                                            <span style="position:absolute;top:-18px;left:50%;transform:translateX(-50%);font-size:10px;font-weight:700;color:#be185d;white-space:nowrap;">{nn_v}</span>
                                        </div>
                                    </div>
                                    <div style="font-size:10px;color:#888;margin-top:4px;text-align:center;width:{bar_w*2+3}px;">{lb}</div>
                                </div>"""
                            chart_html = f"""<!DOCTYPE html><html><body style="margin:0;padding:0;font-family:sans-serif;background:white;">
                            <div style="padding:16px 20px;">
                                <div style="display:flex;gap:16px;margin-bottom:14px;">
                                    <span style="display:flex;align-items:center;gap:5px;font-size:12px;color:#555;">
                                        <span style="width:12px;height:12px;background:#a855f7;border-radius:2px;display:inline-block;"></span> Profitabilni taj mesec (neto &gt; 0)
                                    </span>
                                    <span style="display:flex;align-items:center;gap:5px;font-size:12px;color:#555;">
                                        <span style="width:12px;height:12px;background:#ec4899;border-radius:2px;display:inline-block;"></span> Neprofitabilni taj mesec (neto ≤ 0)
                                    </span>
                                </div>
                                <div style="display:flex;gap:6px;align-items:flex-end;overflow-x:auto;padding-bottom:4px;">
                                    {bars_html}
                                </div>
                            </div>
                            </body></html>"""
                            components.html(chart_html, height=220)
                            st.markdown('''<p style="font-size:12px;color:#9ca3af;margin-top:4px;">
                            ℹ️ Grafikon prikazuje profitabilnost po potencijalu <strong>za svaki mesec posebno</strong> — razlikuje se od ukupnih brojeva iznad koji se odnose na <strong>ceo analizirani period</strong>. Na primer, objekat koji je u poslednjem mesecu neprofitabilan može biti profitabilan gledano kroz ceo period.
                            </p>''', unsafe_allow_html=True)
                        st.markdown("<div style='margin:20px 0 4px 0;'></div>", unsafe_allow_html=True)
                        st.markdown('<div class="section-title">🔴 OOS — Izgubljena zarada zbog nedostatka robe</div>', unsafe_allow_html=True)
                        if len(engine.df_oos) > 0:
                            a_labels_oos = engine.analitika_labels if engine.analitika_labels else engine.mesec_labels
                            oos_ukupno = int(engine.df_oos['Izgubljeni_profit'].sum())
                            oos_mes_avg = oos_ukupno // max(n_mes, 1)
                            oos_kombinacija = int((engine.df_oos['OOS_meseci'] > 0).sum()) if 'OOS_meseci' in engine.df_oos.columns else len(engine.df_oos)
                            oos_0_danas = int((engine.df_oos.get('Lager_danas', 0) == 0).sum()) if 'Lager_danas' in engine.df_oos.columns else oos_kombinacija
                            o1, o2, o3 = st.columns(3)
                            def _oos_kard(col, label, val, suffix=""):
                                col.markdown(f"""<div style="background:white;border-radius:12px;padding:16px 18px;
                                    border-top:3px solid #ec4899;box-shadow:0 2px 8px rgba(0,0,0,0.07);text-align:center;">
                                    <div style="font-size:22px;font-weight:700;color:#ec4899;">{val:,}{suffix}</div>
                                    <div style="font-size:11px;color:#aaa;margin-top:4px;text-transform:uppercase;letter-spacing:.5px;">{label}</div>
                                </div>""", unsafe_allow_html=True)
                            _oos_kard(o1, f"Izgubljen profit · {n_mes} meseci (RSD)", oos_ukupno)
                            _oos_kard(o2, "Prosečno mesečno (RSD)", oos_mes_avg)
                            _oos_kard(o3, "Kombinacija na 0 lagera danas", oos_0_danas)
                            st.markdown("<div style='margin:18px 0 4px 0;'></div>", unsafe_allow_html=True)
                            mes_izgub = []
                            mes_oos_count = []
                            for lb in a_labels_oos:
                                col_izgub = f'Izgub_{lb}'
                                col_oos = f'OOS_{lb}'
                                v_izgub = int(engine.df_oos[col_izgub].sum()) if col_izgub in engine.df_oos.columns else 0
                                v_oos = int((engine.df_oos[col_oos] > 0).sum()) if col_oos in engine.df_oos.columns else 0
                                mes_izgub.append(v_izgub)
                                mes_oos_count.append(v_oos)
                            if any(v > 0 for v in mes_izgub):
                                max_izgub = max(mes_izgub) if mes_izgub else 1
                                chart_w = 860
                                chart_h = 220
                                pad_l, pad_r, pad_t, pad_b = 60, 20, 30, 40
                                plot_w = chart_w - pad_l - pad_r
                                plot_h = chart_h - pad_t - pad_b
                                n_pts = len(a_labels_oos)
                                def px(i): return pad_l + int(i / max(n_pts-1,1) * plot_w)
                                def py(v): return pad_t + plot_h - int(v / max(max_izgub,1) * plot_h)
                                pts_area = " ".join(f"{px(i)},{py(v)}" for i, v in enumerate(mes_izgub))
                                pts_area = f"{px(0)},{pad_t+plot_h} " + pts_area + f" {px(n_pts-1)},{pad_t+plot_h}"
                                pts_line = " ".join(f"{px(i)},{py(v)}" for i, v in enumerate(mes_izgub))
                                dots = ""
                                labels_svg = ""
                                x_labels = ""
                                for i, (lb, v, vc) in enumerate(zip(a_labels_oos, mes_izgub, mes_oos_count)):
                                    x, y = px(i), py(v)
                                    v_k = f"{v//1000}k" if v >= 1000 else str(v)
                                    dots += f'<circle cx="{x}" cy="{y}" r="5" fill="#a855f7" stroke="white" stroke-width="2"/>'
                                    labels_svg += f'<text x="{x}" y="{y-10}" text-anchor="middle" font-size="10" font-weight="700" fill="#7c3aed">{v_k}</text>'
                                    labels_svg += f'<text x="{x}" y="{y+20}" text-anchor="middle" font-size="9" fill="#999">({vc})</text>'
                                    x_labels += f'<text x="{x}" y="{chart_h-6}" text-anchor="middle" font-size="9" fill="#aaa">{lb}</text>'
                                svg = f"""<svg width="{chart_w}" height="{chart_h}" xmlns="http://www.w3.org/2000/svg" style="font-family:sans-serif;">
                                    <text x="{pad_l-5}" y="{pad_t-8}" font-size="10" fill="#888">Izgubljen profit (RSD)</text>
                                    <text x="{chart_w-pad_r}" y="{pad_t-8}" font-size="10" fill="#aaa" text-anchor="end">Broj OOS kombinacija u zagradama</text>
                                    <polygon points="{pts_area}" fill="#a855f7" fill-opacity="0.08"/>
                                    <polyline points="{pts_line}" fill="none" stroke="#a855f7" stroke-width="2.5"/>
                                    {dots}{labels_svg}{x_labels}
                                </svg>"""
                                components.html(f'<!DOCTYPE html><html><body style="margin:0;padding:0;background:white;">{svg}</body></html>', height=chart_h+10)
                            oos_art = engine.df_oos.groupby(['id artikla','Naziv artikla']).agg(
                                Izgubljeni_profit=('Izgubljeni_profit','sum')
                            ).reset_index().sort_values('Izgubljeni_profit', ascending=False).head(5)
                            bar_colors = ["#a855f7","#ec4899","#7c3aed","#c084fc","#f472b6"]
                            top5_max = int(oos_art['Izgubljeni_profit'].max()) if len(oos_art) > 0 else 1
                            bars5 = ""
                            for i, (_, row) in enumerate(oos_art.iterrows()):
                                naziv = str(row['Naziv artikla'])[:35]
                                val = int(row['Izgubljeni_profit'])
                                pct = val / top5_max * 100
                                color = bar_colors[i % len(bar_colors)]
                                bars5 += f"""
                                <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;">
                                    <div style="width:200px;font-size:12px;color:#444;text-align:right;flex-shrink:0;">{naziv}</div>
                                    <div style="flex:1;background:#f5f0ff;border-radius:4px;height:22px;position:relative;">
                                        <div style="width:{pct:.1f}%;background:{color};height:100%;border-radius:4px;"></div>
                                    </div>
                                    <div style="width:110px;font-size:12px;font-weight:700;color:{color};flex-shrink:0;">{val:,} RSD</div>
                                </div>"""
                            st.markdown("**Top 5 artikala po izgubljenom profitu:**")
                            components.html(f"""<!DOCTYPE html><html><body style="margin:0;padding:8px 12px;font-family:sans-serif;background:white;">
                                {bars5}
                            </body></html>""", height=len(oos_art)*42+20)
                            with st.expander("📋 Svi artikli po izgubljenom profitu"):
                                oos_art_all = engine.df_oos.groupby(['id artikla','Naziv artikla']).agg(
                                    Objekata=('ID KOMITENTA','nunique'),
                                    OOS_meseci=('OOS_meseci','sum'),
                                    Izgubljeni_profit=('Izgubljeni_profit','sum')
                                ).reset_index().sort_values('Izgubljeni_profit', ascending=False)
                                oos_art_all.columns = ['ID Art.','Naziv','Objekata','OOS meseci','Izg. profit (RSD)']
                                st.dataframe(oos_art_all, use_container_width=True, height=300)
                        else:
                            st.success("Nema OOS problema!")
                        st.markdown("<div style='margin:24px 0 4px 0;'></div>", unsafe_allow_html=True)
                        st.markdown('<div class="section-title">⚡ Scenario: Optimalna mreža</div>', unsafe_allow_html=True)
                        prof2 = engine.df_profit_obj.copy()
                        oos_ukupno2 = int(engine.df_oos['Izgubljeni_profit'].sum()) if len(engine.df_oos) > 0 else 0
                        pozitivni = prof2[prof2['Potencijalni_profit'] > 0]
                        neto_pozitivnih = int(pozitivni['Neto_profit'].sum())
                        pravi_neg2 = prof2[(prof2['Neto_profit'] <= 0) & (prof2['Potencijalni_profit'] <= 0)]
                        n_pravi_neg2 = len(pravi_neg2)
                        usteda_trosak2 = int(n_pravi_neg2 * engine.trosak_po_objektu)
                        usteda_gubitak2 = int(abs(pravi_neg2['Neto_profit'].sum()))
                        ukupni_potencijal = neto_pozitivnih + usteda_trosak2 + usteda_gubitak2 + oos_ukupno2
                        stvarni_neto = int(prof2['Neto_profit'].sum())
                        razlika = ukupni_potencijal - stvarni_neto
                        period_sc = period_str2
                        def _red(label, val, color="#10b981", bold_val=True):
                            val_str = f"+{val:,} RSD" if val >= 0 else f"{val:,} RSD"
                            v_style = f"font-weight:{'700' if bold_val else '400'};color:{color};"
                            return f"""<div style="display:flex;justify-content:space-between;align-items:center;
                                padding:8px 0;border-bottom:1px solid #f3f4f6;">
                                <span style="font-size:13px;color:#555;">{label}</span>
                                <span style="{v_style}font-size:13px;">{val_str}</span>
                            </div>"""
                        def _red_bold(label, val, color="#111"):
                            val_str = f"= {val:,} RSD"
                            return f"""<div style="display:flex;justify-content:space-between;align-items:center;
                                padding:10px 0;border-top:2px solid #e5e7eb;margin-top:4px;">
                                <span style="font-size:14px;font-weight:700;color:#111;">{label}</span>
                                <span style="font-size:14px;font-weight:700;color:{color};">{val_str}</span>
                            </div>"""
                        scenario_html = f"""
                        <div style="background:white;border-radius:12px;padding:20px 24px;
                            box-shadow:0 2px 8px rgba(0,0,0,0.07);font-family:sans-serif;">
                            <div style="font-size:12px;font-weight:600;color:#a855f7;margin-bottom:12px;
                                text-transform:uppercase;letter-spacing:.5px;">
                                Period: {period_sc} ({n_mes} meseci)
                            </div>
                            <p style="font-size:13px;color:#666;margin-bottom:14px;">
                                Ako se istovremeno zatvore neprofitabilni objekti i eliminiše OOS, mreža ide sa
                                <strong>{stvarni_neto:,} RSD</strong> neto profita na
                                <strong style="color:#10b981;">+{ukupni_potencijal:,} RSD</strong> za {n_mes} meseci.
                            </p>
                            {_red(f"Neto profit pozitivnih objekata (potencijal > 0)", neto_pozitivnih, "#10b981")}
                            {_red(f"Ušteda: zatvaranje {n_pravi_neg2} neprofitabilnih obj.", usteda_trosak2 + usteda_gubitak2, "#10b981")}
                            {_red(f"Povraćaj izgub. zarade (OOS eliminacija)", oos_ukupno2, "#10b981")}
                            {_red_bold(f"UKUPNI POTENCIJAL ({n_mes} meseci)", ukupni_potencijal, "#10b981")}
                            <div style="height:8px;"></div>
                            {_red(f"Stvarni neto profit ({n_mes} meseci)", stvarni_neto, "#555", False)}
                            {_red(f"Razlika — potencijal koji još nije ostvaren", razlika, "#a855f7")}
                        </div>"""
                        st.markdown(scenario_html, unsafe_allow_html=True)
                        if engine.region_map:
                            st.markdown("<div style='margin:28px 0 6px 0;'></div>", unsafe_allow_html=True)
                            st.markdown('<div class="section-title">🗺️ Profitabilnost po okruzima</div>', unsafe_allow_html=True)
                            prof_reg = prof.copy()
                            prof_reg['Region'] = prof_reg['ID KOMITENTA'].map(engine.region_map).fillna('Ostalo')
                            prof_reg['Profitabilan'] = prof_reg['Neto_profit'] > 0
                            reg_grp = prof_reg.groupby('Region').agg(
                                Ukupno=('ID KOMITENTA','count'),
                                Ostaje=('Profitabilan','sum'),
                            ).reset_index()
                            reg_grp['Zatvara'] = reg_grp['Ukupno'] - reg_grp['Ostaje']
                            reg_grp = reg_grp.sort_values('Ukupno', ascending=False).reset_index(drop=True)
                            mali_okruzi_df = reg_grp[reg_grp['Ostaje'] < 5]
                            mali_okruzi = mali_okruzi_df['Region'].tolist()
                            rows_html = ""
                            for _, r in reg_grp.iterrows():
                                okrug = r['Region']
                                ukupno = int(r['Ukupno'])
                                ostaje = int(r['Ostaje'])
                                zatvara = int(r['Zatvara'])
                                mali = " *" if okrug in mali_okruzi else ""
                                mali_color = "#a855f7" if mali else "#111"
                                pct_o = ostaje / max(ukupno, 1) * 100
                                pct_z = zatvara / max(ukupno, 1) * 100
                                bar = f"""<div style="display:flex;width:120px;height:14px;border-radius:3px;overflow:hidden;">
                                    <div style="width:{pct_o:.0f}%;background:#a855f7;"></div>
                                    <div style="width:{pct_z:.0f}%;background:#ec4899;"></div>
                                </div>"""
                                rows_html += f"""<tr style="border-bottom:1px solid #f3f4f6;">
                                    <td style="padding:7px 10px;font-size:13px;color:{mali_color};font-weight:600;">{okrug}{mali}</td>
                                    <td style="padding:7px 10px;font-size:13px;font-weight:700;text-align:center;">{ukupno}</td>
                                    <td style="padding:7px 10px;font-size:13px;text-align:center;">
                                        <span style="color:#a855f7;font-weight:700;">{ostaje}</span>
                                        <span style="color:#999;"> / </span>
                                        <span style="color:#ec4899;font-weight:700;">{zatvara}</span>
                                    </td>
                                    <td style="padding:7px 16px;">{bar}</td>
                                </tr>"""
                            uk_ukupno = int(reg_grp['Ukupno'].sum())
                            uk_ostaje = int(reg_grp['Ostaje'].sum())
                            uk_zatvara = int(reg_grp['Zatvara'].sum())
                            rows_html += f"""<tr style="border-top:2px solid #e5e7eb;background:#f9fafb;">
                                <td style="padding:9px 10px;font-size:13px;font-weight:700;">UKUPNO</td>
                                <td style="padding:9px 10px;font-size:13px;font-weight:700;text-align:center;">{uk_ukupno}</td>
                                <td style="padding:9px 10px;font-size:13px;text-align:center;">
                                    <span style="color:#a855f7;font-weight:700;">{uk_ostaje}</span>
                                    <span style="color:#999;"> / </span>
                                    <span style="color:#ec4899;font-weight:700;">{uk_zatvara}</span>
                                </td>
                                <td></td>
                            </tr>"""
                            header_html = """<tr style="background:#f9fafb;border-bottom:2px solid #e5e7eb;">
                                <th style="padding:9px 10px;font-size:11px;color:#888;font-weight:600;text-align:left;text-transform:uppercase;letter-spacing:.4px;">Okrug</th>
                                <th style="padding:9px 10px;font-size:11px;color:#888;font-weight:600;text-align:center;text-transform:uppercase;letter-spacing:.4px;">Ukupno obj.</th>
                                <th style="padding:9px 10px;font-size:11px;color:#888;font-weight:600;text-align:center;text-transform:uppercase;letter-spacing:.4px;">✓ Ostaje / ✗ Zatvara</th>
                                <th style="padding:9px 10px;font-size:11px;color:#888;font-weight:600;text-transform:uppercase;letter-spacing:.4px;"></th>
                            </tr>"""
                            tbl_height = len(reg_grp) * 34 + 80
                            components.html(f"""<!DOCTYPE html><html><body style="margin:0;padding:0;font-family:sans-serif;background:white;">
                            <table style="width:100%;border-collapse:collapse;">
                                <thead>{header_html}</thead>
                                <tbody>{rows_html}</tbody>
                            </table>
                            </body></html>""", height=tbl_height)
                            if mali_okruzi:
                                mali_str = ", ".join(mali_okruzi)
                                st.markdown(f'<div style="font-size:12px;color:#a855f7;padding:6px 4px;">* Okruzi sa manje od 5 profitabilnih objekata ({mali_str}): Ne preporučuje se angazovanje komercijalistu isključivo za ove okruge — broj preostalih objekata premali je da bi opravdao redovne obilaske.</div>', unsafe_allow_html=True)
                            if len(mali_okruzi_df) > 0:
                                st.markdown("<div style='margin:20px 0 6px 0;'></div>", unsafe_allow_html=True)
                                prof_reg_mali = prof_reg[prof_reg['Region'].isin(mali_okruzi) & (prof_reg['Neto_profit'] > 0)]
                                n_mali_prof = len(prof_reg_mali)
                                neto_mali_prof = int(prof_reg_mali['Neto_profit'].sum())
                                usteda_mali_trosak = int(n_mali_prof * engine.trosak_po_objektu)
                                scA_potencijal = ukupni_potencijal
                                scB_potencijal = scA_potencijal + usteda_mali_trosak - neto_mali_prof
                                period_label = period_str2
                                def _sc_red(label, val, color="#555", bold=False):
                                    sign = "+" if val >= 0 else ""
                                    fw = "700" if bold else "400"
                                    return f"""<div style="display:flex;justify-content:space-between;padding:7px 0;border-bottom:1px solid #f3f4f6;">
                                        <span style="font-size:13px;color:#555;">{label}</span>
                                        <span style="font-size:13px;font-weight:{fw};color:{color};">{sign}{val:,} RSD</span>
                                    </div>"""
                                def _sc_total(label, val, color="#10b981"):
                                    return f"""<div style="display:flex;justify-content:space-between;padding:9px 0;border-top:2px solid #e5e7eb;margin-top:4px;">
                                        <span style="font-size:14px;font-weight:700;color:#111;">{label}</span>
                                        <span style="font-size:14px;font-weight:700;color:{color};">= {val:,} RSD</span>
                                    </div>"""
                                sc_html = f"""<div style="font-family:sans-serif;background:white;border-radius:12px;
                                    padding:20px 24px;box-shadow:0 2px 8px rgba(0,0,0,0.07);">
                                    <div style="font-size:12px;font-weight:600;color:#a855f7;text-transform:uppercase;
                                        letter-spacing:.5px;margin-bottom:14px;">
                                        Uticaj zatvaranja objekata u malim okruzima ({period_label})
                                    </div>
                                    <p style="font-size:13px;color:#666;margin-bottom:14px;">
                                        Zatvaranjem {n_mali_prof} profitabilnih objekata u {len(mali_okruzi)} malih okruga
                                        štedimo trošak, ali gubimo deo zarade. Poređenje dva scenarija:
                                    </p>
                                    <div style="font-size:12px;font-weight:600;color:#7c3aed;margin:10px 0 6px 0;">
                                        Scenario A: Zatvaramo samo {n_pravi_neg2} neprofitabilnih + OOS eliminacija
                                    </div>
                                    {_sc_red(f"Neto profit pozitivnih objekata ({n_mes}m)", neto_pozitivnih, "#10b981", False)}
                                    {_sc_red(f"Ušteda: zatvaranje {n_pravi_neg2} neprofitabilnih ({n_mes}m)", usteda_trosak2 + usteda_gubitak2, "#10b981", False)}
                                    {_sc_red(f"Povraćaj OOS izgubljene zarade ({n_mes}m)", oos_ukupno2, "#10b981", False)}
                                    {_sc_total(f"POTENCIJAL SCENARIO A", scA_potencijal)}
                                    <div style="font-size:12px;font-weight:600;color:#ec4899;margin:16px 0 6px 0;">
                                        Scenario B: Scenario A + zatvaramo i {n_mali_prof} obj. iz malih okruga
                                    </div>
                                    {_sc_red(f"Potencijal Scenario A", scA_potencijal, "#10b981", False)}
                                    {_sc_red(f"Ušteda troška: {n_mali_prof} obj. × {engine.trosak_po_objektu:,.0f} RSD × {n_mes} mes", usteda_mali_trosak, "#10b981", False)}
                                    {_sc_red(f"Izgubljen profit zatvorenih {n_mali_prof} obj. ({n_mes}m)", -neto_mali_prof, "#ec4899", False)}
                                    {_sc_total(f"POTENCIJAL SCENARIO B", scB_potencijal, "#10b981" if scB_potencijal >= scA_potencijal else "#a855f7")}
                                </div>"""
                                components.html(f'<!DOCTYPE html><html><body style="margin:0;padding:0;">{sc_html}</body></html>', height=420)

                st.markdown("---")
                excel_buf = create_excel(engine)
                fname_xl = f"ANALITIKA_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
                st.download_button(f"📥 Preuzmi Excel — {fname_xl}", data=excel_buf, file_name=fname_xl,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except Exception as e:
                st.error(f"Greska: {str(e)}")
                import traceback; st.code(traceback.format_exc())
    else:
        components.html("""<!DOCTYPE html><html><head>
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap" rel="stylesheet">
    </head><body style="margin:0;padding:0;background:transparent;font-family:'Poppins',sans-serif;">
    <div style="max-width:680px;margin:32px auto 0 auto;padding:0 16px;">
      <p style="font-size:11px;color:#9ca3af;font-weight:600;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:14px;">
        AMAN d.o.o. &middot; Analiticki sistem
      </p>
      <h1 style="font-size:36px;font-weight:700;color:#1a0533;line-height:1.2;margin-bottom:12px;margin-top:0;">
        Predikcija prodaje<br>
        <span style="background:linear-gradient(135deg,#a855f7,#ec4899);-webkit-background-clip:text;-webkit-text-fill-color:transparent;">
          &amp; Porudzbine
        </span>
      </h1>
      <p style="font-size:15px;color:#6b7280;margin-bottom:28px;line-height:1.6;">
        Profitabilnost objekata &middot; OOS analiza &middot; Trendovi komitenata &middot; Analiza akcije
      </p>
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:36px;">
        <span style="font-size:12px;background:rgba(168,85,247,0.10);color:#7c3aed;border-radius:99px;padding:5px 14px;font-weight:600;">Predikcija</span>
        <span style="font-size:12px;background:rgba(236,72,153,0.09);color:#be185d;border-radius:99px;padding:5px 14px;font-weight:600;">Profitabilnost</span>
        <span style="font-size:12px;background:rgba(239,68,68,0.09);color:#b91c1c;border-radius:99px;padding:5px 14px;font-weight:600;">OOS analiza</span>
        <span style="font-size:12px;background:rgba(16,185,129,0.09);color:#065f46;border-radius:99px;padding:5px 14px;font-weight:600;">Trendovi</span>
      </div>
      <div style="height:1px;background:linear-gradient(90deg,rgba(168,85,247,0.3),rgba(236,72,153,0.2),transparent);margin-bottom:28px;"></div>
      <p style="font-size:14px;color:#9ca3af;text-align:center;margin-top:8px;">
        &#8593; Učitaj Excel fajl iznad da pocnes analizu
      </p>
    </div>
    </body></html>""", height=340)
