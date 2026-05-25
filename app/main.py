import os
import uuid
import shutil
import logging
import subprocess
from datetime import datetime, timedelta
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, UploadFile, File, HTTPException, Request, Depends, BackgroundTasks, Form, Cookie
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

import pandas as pd
from docxtpl import DocxTemplate
from sqlalchemy import create_engine, Column, Integer, String, DateTime
from sqlalchemy.orm import declarative_base, sessionmaker
from pydantic_settings import BaseSettings
from apscheduler.schedulers.background import BackgroundScheduler
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired

# --- Конфигурация ---
class Settings(BaseSettings):
    ADMIN_USER: str = "admin"
    ADMIN_PASS: str = "secure_password"
    SESSION_SECRET: str = "change_me"
    DB_URL: str = "sqlite:////app/data/history.db"
    KEEP_FILES_DAYS: int = 30
    ENABLE_PDF: bool = True

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"

settings = Settings()
serializer = URLSafeTimedSerializer(settings.SESSION_SECRET)

UPLOAD_DIR = Path("/app/volumes/uploads")
OUTPUT_DIR = Path("/app/volumes/outputs")
BACKUP_DIR = Path("/app/data/backups")
for d in (UPLOAD_DIR, OUTPUT_DIR, BACKUP_DIR): d.mkdir(parents=True, exist_ok=True)

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# --- БД ---
engine = create_engine(settings.DB_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

class Generation(Base):
    __tablename__ = "generations"
    id = Column(Integer, primary_key=True, autoincrement=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    status = Column(String, default="processing")
    total = Column(Integer, default=0)
    created = Column(Integer, default=0)
    skipped = Column(Integer, default=0)
    report_file = Column(String, nullable=True)

class InventoryItem(Base):
    __tablename__ = "inventory_items"
    id = Column(Integer, primary_key=True, autoincrement=True)
    inv_number = Column(String, index=True)
    config = Column(String)
    gen_id = Column(Integer, index=True)
    status = Column(String)
    variant = Column(Integer)

Base.metadata.create_all(bind=engine)

# --- Сессия / Авторизация ---
def get_current_user(request: Request, session_cookie: str = Cookie(None)):
    if not session_cookie:
        return RedirectResponse(url="/login", status_code=302)
    try:
        data = serializer.loads(session_cookie, max_age=3600)
        return data["username"]
    except (BadSignature, SignatureExpired):
        return RedirectResponse(url="/login", status_code=302)

# --- Логика генерации ---
VARIANTS = [
    {"conclusion": "По результатам технического освидетельствования установлено, что оборудование выработало нормативный срок эксплуатации и имеет критический физический и моральный износ. Выявлены неустранимые неисправности основных узлов. Стоимость восстановительного ремонта превышает 80% рыночной стоимости нового аналога, что экономически нецелесообразно в соответствии с требованиями к эффективному использованию бюджетных средств. Оригинальные комплектующие сняты с производства.", "recommendation": "Списать оборудование с балансового учёта учреждения в порядке, установленном действующим законодательством РФ и учётной политикой ГКУ, с последующей утилизацией в соответствии с экологическими нормами и требованиями к обращению ТКО."},
    {"conclusion": "В ходе диагностики выявлена полная утрата работоспособности оборудования вследствие выхода из строя критически важных компонентов. Восстановление технически невозможно по причине отсутствия совместимых запасных частей на рынке (оборудование и его узлы сняты с производства). Дальнейшая эксплуатация невозможна и не соответствует требованиям информационной безопасности и технической надёжности ГКУ. Затраты на ремонт превышают экономически обоснованный лимит восстановления.", "recommendation": "Признать оборудование непригодным для дальнейшего использования, снять с инвентарного учёта и оформить акт на списание в установленном порядке. Освободившееся рабочее место обеспечить за счёт замены на современное оборудование."},
    {"conclusion": "Техническое освидетельствование подтвердило несоответствие оборудования современным требованиям к аппаратному обеспечению информационных систем ГКУ. Оборудование имеет критический моральный износ, не поддерживает необходимые обновления программного обеспечения и средства защиты информации. Физические неисправности носят неустранимый характер, запасные части отсутствуют в продаже. Ремонт экономически нецелесообразен, так как затраты превышают стоимость нового оборудования, соответствующего текущим стандартам учреждения.", "recommendation": "Инициировать процедуру списания оборудования с баланса ГКУ в связи с полной утратой потребительских свойств и технической несовместимостью с действующей ИТ-инфраструктурой. Документально оформить передачу на утилизацию уполномоченной организации."}
]

def process_row(i, row, variants_manual, template_path, out_dir):
    var_idx = i % 3
    if variants_manual is not None and i < len(variants_manual) and pd.notna(variants_manual.iloc[i]):
        try:
            v = int(str(variants_manual.iloc[i]).strip())
            var_idx = v - 1 if 1 <= v <= 3 else i % 3
        except: pass

    safe_inv = "".join(c for c in str(row["inv_number"]) if c.isalnum() or c in "-_")
    filename = f"Акт_списания_{safe_inv}.docx"
    filepath = out_dir / filename

    if filepath.exists():
        return {"inv_number": row["inv_number"], "config": row["config"], "status": "skipped", "variant": var_idx+1, "filename": filename}
    
    try:
        tpl = DocxTemplate(str(template_path))
        ctx = {
            "config": row["config"], "inv_number": row["inv_number"],
            "conclusion_text": VARIANTS[var_idx]["conclusion"],
            "recommendation_text": VARIANTS[var_idx]["recommendation"]
        }
        tpl.render(ctx)
        tpl.save(str(filepath))
        return {"inv_number": row["inv_number"], "config": row["config"], "status": "created", "variant": var_idx+1, "filename": filename}
    except Exception as e:
        return {"inv_number": row["inv_number"], "config": row["config"], "status": f"error: {e}", "variant": var_idx+1, "filename": filename}

def convert_to_pdf(docx_path: Path, out_dir: Path):
    if not settings.ENABLE_PDF or not docx_path.exists(): return None
    pdf_path = out_dir / docx_path.with_suffix(".pdf")
    if pdf_path.exists(): return pdf_path.name
    try:
        subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx_path)],
                       check=True, capture_output=True, timeout=30)
        return pdf_path.name if pdf_path.exists() else None
    except Exception: return None

# --- Фоновые задачи ---
def cleanup_old_files():
    threshold = datetime.utcnow() - timedelta(days=settings.KEEP_FILES_DAYS)
    removed = 0
    for f in OUTPUT_DIR.iterdir():
        if f.is_file() and datetime.utcfromtimestamp(f.stat().st_mtime) < threshold:
            f.unlink()
            removed += 1
    if removed: logging.info(f"🗑 Удалено {removed} старых файлов")

def backup_db():
    ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    backup = BACKUP_DIR / f"history_{ts}.db"
    shutil.copy2("/app/data/history.db", backup)
    logging.info(f"💾 Бэкап БД создан: {backup.name}")

scheduler = BackgroundScheduler()
scheduler.add_job(cleanup_old_files, "cron", hour=3)
scheduler.add_job(backup_db, "cron", hour=4)

# --- FastAPI App ---
app = FastAPI(title="Акт Списания ГКУ", docs_url=None, redoc_url=None)
templates = Jinja2Templates(directory="app/templates")

@app.on_event("startup")
def startup():
    scheduler.start()

@app.on_event("shutdown")
def shutdown():
    scheduler.shutdown()

# --- Роуты авторизации ---
@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request):
    return templates.TemplateResponse("login.html", {"request": request})

@app.post("/login")
async def login(request: Request, username: str = Form(...), password: str = Form(...)):
    if username == settings.ADMIN_USER and password == settings.ADMIN_PASS:
        token = serializer.dumps({"username": username})
        response = RedirectResponse(url="/", status_code=302)
        response.set_cookie(
            key="session",
            value=token,
            httponly=True,
            secure=False,  # ⚠️ Измените на True при подключении HTTPS
            samesite="lax",
            max_age=3600
        )
        return response
    return templates.TemplateResponse("login.html", {
        "request": request,
        "error": "Неверный логин или пароль"
    })

@app.get("/logout")
async def logout():
    response = RedirectResponse(url="/login", status_code=302)
    response.delete_cookie(key="session")
    return response

# --- Защищённые роуты ---
@app.get("/", response_class=HTMLResponse)
async def index(request: Request, username: str = Depends(get_current_user)):
    with SessionLocal() as s:
        history = s.query(Generation).order_by(Generation.id.desc()).all()
        items = s.query(InventoryItem).order_by(InventoryItem.id.desc()).limit(200).all()
    return templates.TemplateResponse("index.html", {"request": request, "history": history, "items": items, "username": username})

@app.post("/generate")
async def generate(excel: UploadFile, template: UploadFile, background: BackgroundTasks, username: str = Depends(get_current_user)):
    if not excel.filename.endswith(".xlsx") or not template.filename.endswith(".docx"):
        raise HTTPException(400, "Только .xlsx и .docx")

    uid = uuid.uuid4().hex[:8]
    excel_path = UPLOAD_DIR / f"{uid}_{excel.filename}"
    tpl_path = UPLOAD_DIR / f"{uid}_{template.filename}"
    
    # 1. Сохраняем загруженные файлы
    excel_path.write_bytes(await excel.read())
    tpl_path.write_bytes(await template.read())

    with SessionLocal() as s:
        gen = Generation(status="processing")
        s.add(gen)
        s.commit()
        gen_id = gen.id

    def run_worker():
        try:
            df_full = pd.read_excel(excel_path, header=None, engine="openpyxl")
            df = df_full.iloc[:, 1:3].copy()  # Берём строго столбцы B и C
            df.columns = ["config", "inv_number"]
            variants_manual = df_full.iloc[:, 3].copy() if df_full.shape[1] > 3 else None
            
            # Очистка данных
            df['config'] = df['config'].astype(str).str.strip()
            df['inv_number'] = df['inv_number'].astype(str).str.strip()
            
            # 🔥 ФИЛЬТР 1: Убираем пустые и явные заголовки
            # Отсекаем строки, где инв. номер похож на заголовок таблицы
            invalid_keywords = ['инвентарный', 'номер', 'наименование', 'inventory', 'config', '№']
            def is_valid_inv(val):
                if not val or val.lower() in invalid_keywords:
                    return False
                # Опционально: если инв. номера всегда цифровые, можно добавить проверку isdigit()
                return True
            
            df = df[df['inv_number'].apply(is_valid_inv)]
            total = len(df)
            logging.info(f"🔍 После фильтрации осталось {total} валидных записей")

            items_data = []
            created = skipped = 0
            
            with ThreadPoolExecutor(max_workers=2) as pool:
                futures = [pool.submit(process_row, i, row, variants_manual, tpl_path, OUTPUT_DIR) for i, row in df.iterrows()]
                for f in futures:
                    res = f.result()
                    items_data.append({
                        "inv_number": res["inv_number"],
                        "config": res["config"],
                        "status": res["status"],
                        "variant": res["variant"]
                    })
                    if res["status"] == "created": created += 1
                    elif res["status"] == "skipped": skipped += 1

            if settings.ENABLE_PDF:
                for f in OUTPUT_DIR.iterdir():
                    if f.suffix == ".docx" and f.name.startswith("Акт_списания_"):
                        convert_to_pdf(f, OUTPUT_DIR)

            # 🔥 ФИЛЬТР 2: Безопасная запись в БД с обработкой дублей
            with SessionLocal() as s2:
                for data in items_data:
                    try:
                        item = InventoryItem(
                            inv_number=data["inv_number"],
                            config=data["config"],
                            gen_id=gen_id,
                            status=data["status"],
                            variant=data["variant"]
                        )
                        s2.add(item)
                    except Exception as db_err:
                        # Если дубль (UNIQUE constraint), просто логируем и идём дальше
                        if "UNIQUE constraint" in str(db_err):
                            logging.warning(f"⏭️ Пропущен дубль инв. номера: {data['inv_number']}")
                        else:
                            logging.error(f"❌ Ошибка БД для {data['inv_number']}: {db_err}")
                
                gen_db = s2.query(Generation).get(gen_id)
                if gen_db:
                    gen_db.status = "completed"
                    gen_db.total = total
                    gen_db.created = created
                    gen_db.skipped = skipped
                    gen_db.report_file = f"Статус_отчет_{uid}.xlsx"
                
                s2.commit()

                df_rep = pd.DataFrame([
                    {"Инв. номер": d["inv_number"], "Конфигурация": d["config"], "Вариант": d["variant"], "Статус": d["status"]} 
                    for d in items_data
                ])
                df_rep.to_excel(OUTPUT_DIR / gen_db.report_file, index=False, engine="openpyxl")
                
            logging.info(f"✅ Завершено #{gen_id}: Создано {created}, Пропущено {skipped}")
            
        except Exception as e:
            logging.error(f"❌ Ошибка в фоновой задаче: {e}")
            with SessionLocal() as s_err:
                gen_db = s_err.query(Generation).get(gen_id)
                if gen_db:
                    gen_db.status = f"error: {str(e)[:100]}"
                    s_err.commit()
        finally:
            if excel_path.exists(): excel_path.unlink()
            if tpl_path.exists(): tpl_path.unlink()

    background.add_task(run_worker)
    return {"status": "accepted", "message": "Генерация запущена в фоне"}

@app.get("/download/{filename}")
async def download(filename: str, username: str = Depends(get_current_user)):
    filepath = OUTPUT_DIR / filename
    if not filepath.is_file(): raise HTTPException(404, "Не найдено")
    media = "application/pdf" if filepath.suffix == ".pdf" else "application/octet-stream"
    return FileResponse(filepath, media_type=media, filename=filename)

@app.post("/backup")
async def manual_backup(username: str = Depends(get_current_user)):
    try:
        backup_db()
        return {"status": "ok", "message": "Бэкап создан"}
    except Exception as e:
        raise HTTPException(500, str(e))